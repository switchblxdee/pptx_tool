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

import json
import logging
import re
from pathlib import Path
from typing import Any, Optional, Type

from langchain_core.callbacks import CallbackManagerForToolRun
from langchain_core.language_models import BaseChatModel
from langchain_core.messages import HumanMessage, SystemMessage
from langchain_core.tools import BaseTool
from pydantic import BaseModel, Field, ValidationError

from .builder import DigestBuilder
from .excel_reader import ExcelReader
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
            "Описание желаемого стиля и содержания презентации на естественном языке. "
            "Например: 'Корпоративная презентация в тёмной палитре, акцент на проблемах "
            "качества, для топ-менеджмента, 10-12 слайдов'."
        )
    )
    output_path: Optional[str] = Field(
        default=None,
        description="Куда сохранить .pptx. Если не указан — рядом с входным файлом.",
    )


# --------------------------------------------------------------------------- #
# Извлечение и валидация LLM-ответа
# --------------------------------------------------------------------------- #

_JSON_FENCE_RE = re.compile(r"```(?:json)?\s*(.*?)```", re.DOTALL)


def extract_json(raw: str) -> str:
    """
    Достаёт JSON из ответа LLM.

    LLM иногда оборачивает JSON в ```json``` или добавляет преамбулу типа
    'Вот ваша презентация:'. Чистим это аккуратно.
    """
    raw = raw.strip()

    # 1. Пробуем найти fenced код-блок
    fence_match = _JSON_FENCE_RE.search(raw)
    if fence_match:
        return fence_match.group(1).strip()

    # 2. Берём первый { и последний } — самый надёжный способ
    first = raw.find("{")
    last = raw.rfind("}")
    if first != -1 and last != -1 and last > first:
        return raw[first:last + 1]

    return raw


def parse_spec(raw_response: str) -> DigestSpec:
    """Парсит и валидирует ответ LLM."""
    json_str = extract_json(raw_response)

    try:
        data = json.loads(json_str)
    except json.JSONDecodeError as e:
        raise ValueError(f"LLM вернул невалидный JSON: {e}. Ответ:\n{raw_response[:500]}")

    try:
        return DigestSpec.model_validate(data)
    except ValidationError as e:
        raise ValueError(
            f"LLM вернул JSON, не соответствующий схеме DigestSpec.\n"
            f"Ошибки валидации:\n{e}"
        )


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

    def _run(
        self,
        xlsx_path: str,
        style_prompt: str,
        output_path: Optional[str] = None,
        run_manager: Optional[CallbackManagerForToolRun] = None,
    ) -> str:
        """Синхронный запуск."""
        try:
            return self._generate(xlsx_path, style_prompt, output_path)
        except Exception as e:
            logger.exception("Ошибка при генерации презентации")
            # Возвращаем читабельную ошибку агенту, чтобы он мог отреагировать
            return f"ОШИБКА: {type(e).__name__}: {e}"

    def _generate(
        self,
        xlsx_path: str,
        style_prompt: str,
        output_path: Optional[str],
    ) -> str:
        # 1. Читаем Excel
        logger.info("Читаю Excel: %s", xlsx_path)
        data_context = ExcelReader(xlsx_path).read()
        logger.info(
            "Прочитано: %d строк × %d колонок, %d символов суммаризации",
            data_context.row_count, data_context.column_count,
            len(data_context.summary_text),
        )

        # 2. Просим LLM сгенерить структуру (с ретраями)
        spec = self._call_llm_with_retries(data_context, style_prompt)

        # 3. Определяем путь вывода
        if output_path is None:
            input_path = Path(xlsx_path)
            output_path = str(input_path.with_suffix(".pptx"))

        # 4. Собираем .pptx
        logger.info("Собираю .pptx: %d слайдов → %s", len(spec.slides), output_path)
        result_path = DigestBuilder(spec).build(output_path)

        return str(result_path)

    def _call_llm_with_retries(
        self, data_context, style_prompt: str
    ) -> DigestSpec:
        """
        Дёргает LLM, ретраит при невалидном JSON, добавляя ошибку в контекст.

        Два режима работы:

        1. inject_system_prompt=False (по умолчанию):
           Всё уезжает в одном HumanMessage. Не трогает системный промпт.
           Используй этот режим, если у LLM уже есть системная персона
           («Привет, ты Агата...») — она останется работать.

        2. inject_system_prompt=True:
           Инструкция в SystemMessage + данные в HumanMessage.
           Используй, если делаешь отдельный изолированный вызов LLM
           без предустановленной персоны.
        """
        if self.inject_system_prompt:
            messages = [
                SystemMessage(content=system_prompt()),
                HumanMessage(content=data_only_user_prompt(data_context, style_prompt)),
            ]
        else:
            # Одно сообщение — самодостаточная инструкция + данные.
            # Не перезаписывает системный промпт LLM.
            messages = [
                HumanMessage(content=analysis_user_prompt(data_context, style_prompt)),
            ]

        last_error: Optional[Exception] = None
        for attempt in range(self.max_retries + 1):
            logger.info("Запрос к LLM (попытка %d/%d)", attempt + 1, self.max_retries + 1)
            response = self.llm.invoke(messages)
            raw = response.content if isinstance(response.content, str) else str(response.content)

            try:
                return parse_spec(raw)
            except ValueError as e:
                last_error = e
                logger.warning("Парсинг неуспешен (попытка %d): %s", attempt + 1, e)

                # Добавляем ошибку в контекст и просим исправить
                messages.append(response)
                messages.append(HumanMessage(content=(
                    f"Твой предыдущий ответ невалиден. Ошибка валидации:\n{e}\n\n"
                    f"Верни ИСПРАВЛЕННЫЙ JSON, строго соответствующий формату из примера. "
                    f"Только сырой JSON, начинающийся с {{ и заканчивающийся }}. "
                    f"Без markdown, без пояснений, без обёрток."
                )))

        raise RuntimeError(
            f"LLM не смог вернуть валидный JSON за {self.max_retries + 1} попыток. "
            f"Последняя ошибка: {last_error}"
        )

    async def _arun(self, *args: Any, **kwargs: Any) -> str:
        # Простая async-обёртка через invoke; для production стоит сделать
        # настоящую async через self.llm.ainvoke и aiofiles для Excel
        return self._run(*args, **kwargs)


# Backward compatibility aliases
GeneratePresentationTool = GenerateDigestTool
GeneratePresentationInput = GenerateDigestInput
