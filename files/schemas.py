"""
Pydantic-схемы для дайджеста.

Модель данных полностью переориентирована под формат «корпоративный
аналитический дайджест» (как на референсе «Голос IT»).

Структура:
    DigestSpec
    ├── meta (период, номер выпуска, дата)
    ├── style (палитра, шрифты)
    ├── cover (обложка с KPI и тегами источников)
    └── topics: [TopicSlide]  (детальные слайды по темам)
"""
from __future__ import annotations

from typing import List, Optional

from pydantic import BaseModel, Field, field_validator


# --------------------------------------------------------------------------- #
# Стиль
# --------------------------------------------------------------------------- #

class ColorPalette(BaseModel):
    """
    Палитра дайджеста.

    На референсе видны: бирюзово-персиковый градиент (gradient_start/end),
    мягкие пастельные карточки тем (card_bg), белые KPI-карточки (kpi_bg),
    акцентный оранжевый для бейджа `new` (badge).
    """
    gradient_start: str = Field(description="Левый цвет градиента фона (обычно бирюзовый)")
    gradient_end: str = Field(description="Правый цвет градиента фона (обычно персиковый)")
    card_bg: str = Field(description="Фон карточек темы (мягкий пастельный)")
    kpi_bg: str = Field(description="Фон KPI-карточек (обычно белый)")
    text_dark: str = Field(description="Основной тёмный текст")
    text_muted: str = Field(description="Приглушённый текст (метаданные, цитаты)")
    accent: str = Field(description="Акцентный цвет для заголовков тем и плашек")
    badge: str = Field(description="Цвет бейджа «new» (обычно оранжевый/коралл)")

    @field_validator("*")
    @classmethod
    def validate_hex(cls, v: str) -> str:
        v = v.strip().lstrip("#").upper()
        if len(v) != 6 or not all(c in "0123456789ABCDEF" for c in v):
            raise ValueError(f"Некорректный hex-цвет: {v!r}, ожидается 6 hex-символов")
        return v


class Typography(BaseModel):
    heading_font: str = Field(default="Calibri")
    body_font: str = Field(default="Calibri")


class DigestStyle(BaseModel):
    palette: ColorPalette
    typography: Typography = Field(default_factory=Typography)


# --------------------------------------------------------------------------- #
# Мета-информация дайджеста (футер)
# --------------------------------------------------------------------------- #

class DigestMeta(BaseModel):
    """
    Метаданные дайджеста — отображаются в шапке cover-слайда
    и в футере каждого topic-слайда.
    """
    issue_date: str = Field(max_length=30, description="Дата выпуска, напр. '28 мая 2026'")
    period: str = Field(max_length=40, description="Период, напр. '21–27 мая 2026'")
    issue_number: str = Field(max_length=30, description="Номер выпуска, напр. '№ 1 / еженедельный'")
    next_issue: Optional[str] = Field(
        default=None, max_length=40,
        description="Когда следующий выпуск, напр. '1 июня 2026'"
    )
    note: Optional[str] = Field(
        default="Подготовлен автоматически", max_length=60,
        description="Финальная пометка в футере"
    )


# --------------------------------------------------------------------------- #
# KPI-карточки (используются на cover и на topic-слайдах)
# --------------------------------------------------------------------------- #

class KPICard(BaseModel):
    """Одна KPI-карточка."""
    value: str = Field(max_length=10, description="Число/значение, крупно")
    label: str = Field(max_length=40, description="Подпись под числом")
    icon_hint: Optional[str] = Field(
        default=None, max_length=20,
        description="Подсказка по иконке: 'signal', 'lens', 'arrow_up', 'info' и т.п."
    )


# --------------------------------------------------------------------------- #
# Cover-слайд (обложка дайджеста)
# --------------------------------------------------------------------------- #

class CoverSlide(BaseModel):
    """
    Обложка дайджеста.

    На референсе:
    - Крупный заголовок («Голос IT»)
    - Подзаголовок («дайджест для руководства Блока T»)
    - Описание мелким текстом («Темы, волнующие сотрудников...»)
    - 4 pill-метки источников
    - 4 KPI-карточки внизу
    """
    title: str = Field(max_length=60, description="Главный заголовок («Голос IT»)")
    subtitle: str = Field(max_length=120, description="Подзаголовок дайджеста")
    description: Optional[str] = Field(
        default=None, max_length=200,
        description="Описание мелким текстом под подзаголовком"
    )
    source_tags: List[str] = Field(
        min_length=1, max_length=6,
        description="Источники в виде pill-меток"
    )
    kpis: List[KPICard] = Field(
        min_length=2, max_length=4,
        description="Ключевые показатели на обложке"
    )

    @field_validator("source_tags")
    @classmethod
    def limit_tag_length(cls, v: List[str]) -> List[str]:
        # Pills сильно растягиваются — короткие тексты обязательны
        return [t[:50].rstrip() for t in v]


# --------------------------------------------------------------------------- #
# Topic-слайд (детальный слайд по теме)
# --------------------------------------------------------------------------- #

class TopicItem(BaseModel):
    """
    Одна тема-проблема внутри topic-слайда.

    На референсе у каждого item:
    - Номер
    - Заголовок темы жирным
    - Цитата сотрудника курсивом
    - Период справа (или «на анализе у команды»)
    - Количество упоминаний справа
    - Опциональный бейдж `new`
    """
    title: str = Field(max_length=110, description="Краткое название проблемы")
    quote: Optional[str] = Field(
        default=None, max_length=200,
        description="Прямая цитата сотрудника (курсивом)"
    )
    period: str = Field(
        max_length=30,
        description="Период, напр. '2Q 2026' или 'на анализе у команды'"
    )
    mentions: int = Field(ge=0, le=99999, description="Количество упоминаний")
    is_new: bool = Field(default=False, description="Показать ли бейдж 'new'")


class TopicSlide(BaseModel):
    """
    Детальный слайд по теме/продукту.

    На референсе:
    - Слева плашка с названием темы («PDLC: GigaCode CLI»)
    - 4 KPI-карточки в линию: Источники / Сигналов / Активных тем / Новая тема
    - Список TopicItem (1-3 элемента)
    - Внизу pill-метки источников этой темы
    """
    title: str = Field(max_length=60, description="Название темы/продукта")
    kpis: List[KPICard] = Field(
        min_length=2, max_length=4,
        description="KPI-карточки в линию вверху"
    )
    items: List[TopicItem] = Field(
        min_length=1, max_length=4,
        description="Список выявленных тем/проблем"
    )
    source_tags: List[str] = Field(
        default_factory=list, max_length=6,
        description="Источники этой темы в виде pill-меток (#Ai in Dev Community и т.п.)"
    )

    @field_validator("source_tags")
    @classmethod
    def limit_tag_length(cls, v: List[str]) -> List[str]:
        return [t[:50].rstrip() for t in v]


# --------------------------------------------------------------------------- #
# Корневая спецификация
# --------------------------------------------------------------------------- #

class DigestSpec(BaseModel):
    """
    Полная спецификация дайджеста.

    Это то, что возвращает LLM и принимает на вход DigestBuilder.
    """
    style: DigestStyle
    meta: DigestMeta
    cover: CoverSlide
    topics: List[TopicSlide] = Field(
        min_length=1, max_length=20,
        description="Темы для детальных слайдов"
    )

    @field_validator("topics")
    @classmethod
    def sanitize_recommendations(cls, topics: List[TopicSlide]) -> List[TopicSlide]:
        """
        Программная защита от рекомендаций.

        Дайджест — это аналитический формат: только наблюдения, без советов.
        Если LLM прислал тему с маркерами рекомендаций — переименуем или
        отфильтруем.
        """
        sanitized = []
        for topic in topics:
            # Чистим title темы
            if any(m in topic.title.lower() for m in _RECOMMENDATION_TITLE_MARKERS):
                topic = topic.model_copy(update={"title": "Анализ темы"})

            # Чистим items: фильтруем те, что начинаются с императива
            clean_items = [
                item for item in topic.items
                if not _is_imperative(item.title)
            ]
            if not clean_items:
                clean_items = topic.items[:1]  # оставим хоть один
            topic = topic.model_copy(update={"items": clean_items})

            sanitized.append(topic)
        return sanitized


# --------------------------------------------------------------------------- #
# Санитайзер «рекомендательного» контента
#
# Module-level: подчёркивание здесь — обычное соглашение Python.
# Внутри BaseModel такие атрибуты превратились бы в ModelPrivateAttr.
# --------------------------------------------------------------------------- #

_RECOMMENDATION_TITLE_MARKERS = (
    "рекоменд", "следующие шаги", "что делать", "план действий",
    "меры по устранению", "предложения",
    "плана внедрения", "дорожная карта",
)

_IMPERATIVE_MARKERS = (
    "следует ", "необходимо ", "нужно ", "рекомендуется",
    "стоит внедрить", "предлагается", "требуется ", "должен быть внедрён",
)


def _is_imperative(text: str) -> bool:
    """True если текст начинается с императивного слова."""
    lower = text.lower().strip()
    return any(lower.startswith(m) or m in lower[:30] for m in _IMPERATIVE_MARKERS)
