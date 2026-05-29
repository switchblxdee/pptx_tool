"""
LangChain-инструмент для GigaChat.

Tool принимает путь к xlsx и текстовый промпт пользователя,
возвращает путь к сгенерированному .pptx.

Архитектура:
1. ExcelReader читает xlsx → DataContext
2. Промпт + DataContext отправляются в GigaChat
3. Ответ парсится в DigestSpec (с валидацией и ретраями)
4. DigestBuilder собирает .pptx

Ошибки на каждом этапе оборачиваются в осмысленные сообщения —
LangChain-агент должен видеть, что именно пошло не так, чтобы
правильно действовать (повторить, спросить пользователя, и т.д.).
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Optional, Type

from langchain_core.callbacks import CallbackManagerForToolRun
from langchain_core.language_models import BaseChatModel
from langchain_core.messages import HumanMessage, SystemMessage
from langchain_core.tools import BaseTool
from pydantic import BaseModel, Field, ValidationError

from .builder import DigestBuilder
from .excel_reader import ExcelReader
from .json_repair import repair_json
from .prompts import (
    analysis_user_prompt,
    data_only_user_prompt,
    system_prompt,
)
from .schemas import DigestSpec


logger = logging.getLogger(__name__)


# --------------------------------------------------------------------------- #
# Схема входных параметров для tool
# --------------------------------------------------------------------------- #

class GenerateDigestInput(BaseModel):
    """Входные параметры инструмента."""
    xlsx_path: str = Field(
        description=(
            "Абсолютный путь к xlsx-файлу с двумя листами: "
            "'Сырье' (табличные данные) и 'Суммаризация' (текст про проблемы)."
        )
    )
    style_prompt: str = Field(
        description=(
            "Описание желаемого стиля и содержания дайджеста на естественном языке. "
            "Например: 'Корпоративный дайджест в пастельной палитре для топ-менеджмента'."
        )
    )
    output_path: Optional[str] = Field(
        default=None,
        description="Куда сохранить .pptx. Если не указан — рядом с входным файлом.",
    )
    slide_count: Optional[int] = Field(
        default=None,
        description=(
            "Сколько topic-слайдов (тем) сделать. Если не указано — определяется "
            "автоматически: LLM решает по важности на основе числа уникальных тем "
            "в данных."
        ),
    )
    grouping_column: Optional[str] = Field(
        default=None,
        description=(
            "Имя колонки, по которой группировать данные в темы (1 уникальное "
            "значение = 1 topic-слайд). Если не указано — ищется автоматически "
            "среди кандидатов ('Объект сигнала', 'Продукт', 'Тема' и т.п.)."
        ),
    )


# --------------------------------------------------------------------------- #
# Извлечение и валидация LLM-ответа
# --------------------------------------------------------------------------- #

def parse_spec(raw_response: str) -> DigestSpec:
    """
    Парсит и валидирует ответ LLM.

    Использует многоуровневый ремонт JSON (repair_json), который чинит
    типичные дефекты слабых моделей: markdown-обёртки, трейлинг-запятые,
    одинарные кавычки, обрезанный хвост, Python-литералы и т.п. — БЕЗ
    дополнительного обращения к LLM.
    """
    try:
        data = repair_json(raw_response)
    except ValueError as e:
        raise ValueError(
            f"LLM вернул JSON, который не удалось распарсить даже после ремонта.\n"
            f"{e}"
        )

    try:
        return DigestSpec.model_validate(data)
    except ValidationError as e:
        raise ValueError(
            f"JSON распарсился, но не соответствует схеме DigestSpec.\n"
            f"Ошибки валидации:\n{_format_validation_errors(e)}"
        )


def _format_validation_errors(e: ValidationError) -> str:
    """
    Превращает ValidationError в компактный, понятный для LLM список.

    Полный traceback Pydantic слишком длинный и зашумлённый — даём
    модели короткие указания «поле X: что не так», чтобы ретрай был
    точнее.
    """
    lines = []
    for err in e.errors()[:15]:  # не больше 15, чтобы не раздувать промпт
        loc = " → ".join(str(x) for x in err["loc"])
        msg = err["msg"]
        lines.append(f"  • {loc}: {msg}")
    if len(e.errors()) > 15:
        lines.append(f"  ... и ещё {len(e.errors()) - 15} ошибок")
    return "\n".join(lines)


# --------------------------------------------------------------------------- #
# Сам LangChain Tool
# --------------------------------------------------------------------------- #

class GenerateDigestTool(BaseTool):
    """
    LangChain-инструмент для генерации дайджеста .pptx из xlsx + промпта.

    Использование:
        from langchain_gigachat import GigaChat
        from pptx_generator import GenerateDigestTool

        llm = GigaChat(credentials="...", model="GigaChat-Pro")
        tool = GenerateDigestTool(llm=llm)

        # Через invoke:
        result = tool.invoke({
            "xlsx_path": "/path/to/data.xlsx",
            "style_prompt": "Корпоративный дайджест, пастельные цвета",
        })

        # Или дать агенту:
        agent = create_react_agent(llm, [tool], ...)
    """

    name: str = "generate_digest_pptx"
    description: str = (
        "Создаёт корпоративный аналитический дайджест в PowerPoint (.pptx) "
        "на основе xlsx-файла и текстового описания стиля. Excel должен содержать "
        "листы 'Сырье' (табличные данные с цитатами/упоминаниями) и 'Суммаризация' "
        "(текст про темы). Инструмент структурирует данные в формат дайджеста: "
        "обложка с общими KPI и pill-тегами источников + детальные слайды по темам "
        "с цитатами сотрудников, бейджами 'new', и KPI по каждой теме. "
        "Стилизация (палитра, шрифты) адаптируется под промпт пользователя. "
        "Возвращает путь к сгенерированному .pptx."
    )
    args_schema: Type[BaseModel] = GenerateDigestInput

    llm: BaseChatModel = Field(description="LLM (обычно GigaChat) для генерации структуры")
    max_retries: int = Field(default=3, description="Сколько раз повторить при невалидном JSON")
    inject_system_prompt: bool = Field(
        default=False,
        description=(
            "Если True — инструкция по JSON-схеме идёт в SystemMessage. "
            "Если False (по умолчанию) — вся инструкция в одном HumanMessage, "
            "что НЕ конфликтует с уже заданным системным промптом (например, "
            "'Привет, ты Агата'). Рекомендуется False, если у LLM уже есть "
            "системная персона."
        ),
    )
    use_structured_output: bool = Field(
        default=False,
        description=(
            "Использовать ли llm.with_structured_output() через function calling. "
            "ПО УМОЛЧАНИЮ ВЫКЛЮЧЕНО. Причина: модели работают через внутренний "
            "GigaChat-прокси, и транспорт function-calling к GLM/DeepSeek "
            "непредсказуем (часто возвращает None). Текстовый режим с few-shot "
            "промптом и ремонтом JSON надёжнее для сильных моделей (GLM, DeepSeek) "
            "и не зависит от деталей прокси. Включай только если уверен, что "
            "structured output корректно проброшен к модели."
        ),
    )

    def _run(
        self,
        xlsx_path: str,
        style_prompt: str,
        output_path: Optional[str] = None,
        slide_count: Optional[int] = None,
        grouping_column: Optional[str] = None,
        run_manager: Optional[CallbackManagerForToolRun] = None,
    ) -> str:
        """Синхронный запуск."""
        try:
            return self._generate(
                xlsx_path, style_prompt, output_path,
                slide_count, grouping_column,
            )
        except Exception as e:
            logger.exception("Ошибка при генерации дайджеста")
            # Возвращаем читабельную ошибку агенту, чтобы он мог отреагировать
            return f"ОШИБКА: {type(e).__name__}: {e}"

    def _generate(
        self,
        xlsx_path: str,
        style_prompt: str,
        output_path: Optional[str],
        slide_count: Optional[int] = None,
        grouping_column: Optional[str] = None,
    ) -> str:
        # 1. Читаем Excel с автопоиском группировки
        logger.info("Читаю Excel: %s", xlsx_path)
        data_context = ExcelReader(
            xlsx_path,
            grouping_column=grouping_column,
            requested_slide_count=slide_count,
        ).read()
        logger.info(
            "Прочитано: %d строк × %d колонок. Группировка: %s (%d тем)",
            data_context.row_count, data_context.column_count,
            data_context.grouping.column_name if data_context.grouping else "—",
            data_context.grouping.group_count if data_context.grouping else 0,
        )

        # 2. Просим LLM сгенерить структуру (с ретраями)
        spec = self._call_llm_with_retries(data_context, style_prompt)

        # 3. Определяем путь вывода
        if output_path is None:
            input_path = Path(xlsx_path)
            output_path = str(input_path.with_suffix(".pptx"))

        # 4. Собираем .pptx
        logger.info(
            "Собираю дайджест: обложка + %d тем → %s",
            len(spec.topics), output_path,
        )
        result_path = DigestBuilder(spec).build(output_path)

        return str(result_path)

    def _call_llm_with_retries(
        self, data_context, style_prompt: str
    ) -> DigestSpec:
        """
        Получает DigestSpec от LLM. Двухступенчатая стратегия:

        1. Если use_structured_output=True — пробуем function calling через
           llm.with_structured_output(DigestSpec). Самый надёжный путь для
           GigaChat: модель отдаёт структурированный объект, а не текст.

        2. Если структурный режим недоступен/упал — откатываемся на
           текстовый режим с многоуровневым ремонтом JSON и ретраями.
        """
        # --- Попытка 1: structured output ---
        if self.use_structured_output:
            spec = self._try_structured_output(data_context, style_prompt)
            if spec is not None:
                logger.info("Структурный вывод успешен")
                return spec
            logger.info("Структурный вывод недоступен, откат на текстовый режим")

        # --- Попытка 2: текстовый режим с ремонтом + ретраи ---
        return self._text_mode_with_retries(data_context, style_prompt)

    def _try_structured_output(
        self, data_context, style_prompt: str
    ) -> Optional[DigestSpec]:
        """
        Пробует получить DigestSpec через with_structured_output.

        langchain-gigachat поддерживает два метода:
        - function_calling (по умолчанию) — через tool-calling транспорт
        - json_mode — модель обязана вернуть валидный JSON по схеме

        Пробуем оба. Возвращает DigestSpec при успехе или None.

        ВАЖНО: None — это НОРМАЛЬНЫЙ исход, не ошибка. DigestSpec —
        сложная вложенная схема (массивы объектов, опциональные блоки),
        и GigaChat часто не вытягивает её через structured output,
        возвращая None или пустой объект. В этом случае управление
        корректно переходит к текстовому режиму, который для GigaChat
        на практике надёжнее за счёт few-shot примера и ремонта JSON.
        """
        if self.inject_system_prompt:
            messages = [
                SystemMessage(content=system_prompt()),
                HumanMessage(content=data_only_user_prompt(data_context, style_prompt)),
            ]
        else:
            messages = [
                HumanMessage(content=analysis_user_prompt(data_context, style_prompt)),
            ]

        # Пробуем доступные методы по очереди
        for method_kwargs in ({}, {"method": "json_mode"}):
            method_name = method_kwargs.get("method", "function_calling")
            try:
                structured_llm = self.llm.with_structured_output(
                    DigestSpec, **method_kwargs
                )
            except (AttributeError, NotImplementedError, TypeError) as e:
                logger.info(
                    "structured_output[%s] недоступен у этой модели: %s",
                    method_name, e,
                )
                continue

            try:
                result = structured_llm.invoke(messages)
            except Exception as e:
                logger.info(
                    "structured_output[%s] не справился (%s) — пробую дальше",
                    method_name, type(e).__name__,
                )
                continue

            # Нормализуем результат
            if isinstance(result, DigestSpec):
                logger.info("structured_output[%s] успешен", method_name)
                return result
            if isinstance(result, dict) and result:
                try:
                    spec = DigestSpec.model_validate(result)
                    logger.info("structured_output[%s] успешен (из dict)", method_name)
                    return spec
                except ValidationError as e:
                    logger.info(
                        "structured_output[%s] вернул dict не по схеме — пробую дальше",
                        method_name,
                    )
                    continue

            # None или пустой объект — ожидаемо для сложной схемы
            logger.info(
                "structured_output[%s] вернул %s — перехожу к следующему методу",
                method_name, type(result).__name__,
            )

        logger.info("structured_output не дал результата — откат на текстовый режим")
        return None

    def _text_mode_with_retries(
        self, data_context, style_prompt: str
    ) -> DigestSpec:
        """
        Текстовый режим: LLM возвращает JSON-текст, мы чиним и валидируем,
        при неудаче — ретраим с точечным фидбеком об ошибке.
        """
        if self.inject_system_prompt:
            messages = [
                SystemMessage(content=system_prompt()),
                HumanMessage(content=data_only_user_prompt(data_context, style_prompt)),
            ]
        else:
            messages = [
                HumanMessage(content=analysis_user_prompt(data_context, style_prompt)),
            ]

        last_error: Optional[Exception] = None
        for attempt in range(self.max_retries + 1):
            logger.info("Текстовый запрос к LLM (попытка %d/%d)",
                        attempt + 1, self.max_retries + 1)
            response = self.llm.invoke(messages)
            raw = response.content if isinstance(response.content, str) else str(response.content)

            try:
                return parse_spec(raw)
            except ValueError as e:
                last_error = e
                logger.warning("Парсинг неуспешен (попытка %d): %s", attempt + 1, e)

                # Точечный фидбэк: показываем модели именно её ошибки
                messages.append(response)
                messages.append(HumanMessage(content=(
                    f"Твой предыдущий ответ невалиден. Конкретные ошибки:\n{e}\n\n"
                    f"Исправь ИМЕННО эти поля и верни полный JSON заново. "
                    f"Только сырой JSON, начинающийся с {{ и заканчивающийся }}. "
                    f"Без markdown-обёрток ```, без пояснений, без приветствий."
                )))

        raise RuntimeError(
            f"LLM не смог вернуть валидный JSON за {self.max_retries + 1} попыток. "
            f"Последняя ошибка:\n{last_error}"
        )

    async def _arun(self, *args: Any, **kwargs: Any) -> str:
        # Простая async-обёртка через invoke; для production стоит сделать
        # настоящую async через self.llm.ainvoke и aiofiles для Excel
        return self._run(*args, **kwargs)


# Backward compatibility aliases
GeneratePresentationTool = GenerateDigestTool
GeneratePresentationInput = GenerateDigestInput
