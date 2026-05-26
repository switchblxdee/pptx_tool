"""
Pydantic-схемы для структурированного вывода от LLM.

Эти модели — контракт между GigaChat и генератором презентаций.
LLM возвращает JSON, который валидируется в PresentationSpec,
а дальше PresentationBuilder детерминированно строит .pptx.

Такое разделение даёт:
- предсказуемость (LLM не управляет рендерингом напрямую);
- надёжность (валидация ловит галлюцинации до сборки);
- тестируемость (генератор тестируется без LLM).
"""
from __future__ import annotations

from enum import Enum
from typing import List, Literal, Optional, Union

from pydantic import BaseModel, Field, field_validator, model_validator


# --------------------------------------------------------------------------- #
# Базовые типы
# --------------------------------------------------------------------------- #

class ChartType(str, Enum):
    """Поддерживаемые типы нативных PPTX-графиков."""
    BAR = "bar"           # горизонтальные столбцы (категории слева)
    COLUMN = "column"     # вертикальные столбцы (классический bar chart)
    LINE = "line"         # тренды во времени
    PIE = "pie"           # доли целого
    SCATTER = "scatter"   # корреляции между двумя величинами


class SlideLayout(str, Enum):
    """Лэйауты слайдов. Каждый имеет свой рендер в builder.py."""
    TITLE = "title"                       # обложка
    SECTION_HEADER = "section_header"     # разделитель раздела
    BULLETS = "bullets"                   # заголовок + буллеты
    TWO_COLUMN = "two_column"             # текст слева, текст справа
    CHART = "chart"                       # заголовок + один график на весь слайд
    CHART_WITH_TEXT = "chart_with_text"   # график + аналитика рядом
    KPI = "kpi"                           # 3-4 крупных числа с подписями
    QUOTE = "quote"                       # цитата / ключевой инсайт
    CLOSING = "closing"                   # финальный слайд


# --------------------------------------------------------------------------- #
# Стиль (адаптируется под промпт пользователя)
# --------------------------------------------------------------------------- #

class ColorPalette(BaseModel):
    """
    Палитра из 6 цветов в hex (без #).

    LLM подбирает её под тему презентации.
    Палитра валидируется на корректность hex.
    """
    primary: str = Field(description="Основной цвет — фоны заголовков, акценты")
    secondary: str = Field(description="Вторичный — поддерживающий цвет")
    accent: str = Field(description="Контрастный акцент для важных элементов")
    background: str = Field(description="Фон контентных слайдов (обычно светлый)")
    text_dark: str = Field(description="Тёмный текст для светлого фона")
    text_light: str = Field(description="Светлый текст для тёмного фона")

    @field_validator("*")
    @classmethod
    def validate_hex(cls, v: str) -> str:
        v = v.strip().lstrip("#").upper()
        if len(v) != 6 or not all(c in "0123456789ABCDEF" for c in v):
            raise ValueError(f"Некорректный hex-цвет: {v!r}, ожидается 6 hex-символов")
        return v


class Typography(BaseModel):
    """Шрифтовая пара. LLM подбирает под настроение презентации."""
    heading_font: str = Field(default="Calibri", description="Шрифт заголовков")
    body_font: str = Field(default="Calibri", description="Шрифт основного текста")


class PresentationStyle(BaseModel):
    """Цельный стилевой контракт презентации."""
    palette: ColorPalette
    typography: Typography = Field(default_factory=Typography)
    mood: str = Field(
        default="professional",
        description="Настроение: professional, energetic, calm, technical и т.п."
    )


# --------------------------------------------------------------------------- #
# Контент слайдов (дискриминированные модели)
# --------------------------------------------------------------------------- #

class ChartSeries(BaseModel):
    """Один ряд данных для графика."""
    name: str = Field(description="Название серии (легенда)")
    values: List[float] = Field(description="Значения серии")

    @field_validator("values")
    @classmethod
    def non_empty(cls, v: List[float]) -> List[float]:
        if not v:
            raise ValueError("Серия не может быть пустой")
        return v


class ChartData(BaseModel):
    """Данные графика, типонезависимое представление."""
    chart_type: ChartType
    title: str = Field(description="Заголовок графика")
    categories: List[str] = Field(description="Категории (ось X или сектора pie)")
    series: List[ChartSeries] = Field(description="Один или больше рядов данных")
    x_axis_title: Optional[str] = None
    y_axis_title: Optional[str] = None

    @model_validator(mode="after")
    def validate_consistency(self) -> "ChartData":
        if not self.series:
            raise ValueError("Нужен хотя бы один ряд данных")

        # Pie принимает только одну серию
        if self.chart_type == ChartType.PIE and len(self.series) > 1:
            raise ValueError("Pie chart поддерживает только одну серию данных")

        # Длины серий должны совпадать с категориями (кроме scatter)
        if self.chart_type != ChartType.SCATTER:
            n = len(self.categories)
            for s in self.series:
                if len(s.values) != n:
                    raise ValueError(
                        f"Серия '{s.name}': {len(s.values)} значений, "
                        f"но категорий {n}. Длины должны совпадать."
                    )
        return self


class KPIItem(BaseModel):
    """Крупный показатель для KPI-слайда."""
    value: str = Field(description="Само число со знаком, напр. '+24%', '1.2M'")
    label: str = Field(description="Что означает это число (краткое описание)")


# --- Дискриминированные слайды ----------------------------------------------

class _BaseSlide(BaseModel):
    layout: SlideLayout
    speaker_notes: Optional[str] = Field(
        default=None, description="Заметки докладчика (не отображаются на слайде)"
    )


class TitleSlide(_BaseSlide):
    layout: Literal[SlideLayout.TITLE] = SlideLayout.TITLE
    title: str
    subtitle: Optional[str] = None


class SectionHeaderSlide(_BaseSlide):
    layout: Literal[SlideLayout.SECTION_HEADER] = SlideLayout.SECTION_HEADER
    title: str
    description: Optional[str] = None


class BulletsSlide(_BaseSlide):
    layout: Literal[SlideLayout.BULLETS] = SlideLayout.BULLETS
    title: str
    bullets: List[str] = Field(min_length=1, max_length=7)


class TwoColumnSlide(_BaseSlide):
    layout: Literal[SlideLayout.TWO_COLUMN] = SlideLayout.TWO_COLUMN
    title: str
    left_heading: str
    left_bullets: List[str] = Field(min_length=1, max_length=6)
    right_heading: str
    right_bullets: List[str] = Field(min_length=1, max_length=6)


class ChartSlide(_BaseSlide):
    layout: Literal[SlideLayout.CHART] = SlideLayout.CHART
    title: str
    chart: ChartData


class ChartWithTextSlide(_BaseSlide):
    layout: Literal[SlideLayout.CHART_WITH_TEXT] = SlideLayout.CHART_WITH_TEXT
    title: str
    chart: ChartData
    insights: List[str] = Field(
        min_length=1, max_length=5,
        description="Ключевые выводы из графика"
    )


class KPISlide(_BaseSlide):
    layout: Literal[SlideLayout.KPI] = SlideLayout.KPI
    title: str
    kpis: List[KPIItem] = Field(min_length=2, max_length=4)


class QuoteSlide(_BaseSlide):
    layout: Literal[SlideLayout.QUOTE] = SlideLayout.QUOTE
    quote: str
    attribution: Optional[str] = None


class ClosingSlide(_BaseSlide):
    layout: Literal[SlideLayout.CLOSING] = SlideLayout.CLOSING
    title: str = "Спасибо за внимание"
    subtitle: Optional[str] = None


Slide = Union[
    TitleSlide,
    SectionHeaderSlide,
    BulletsSlide,
    TwoColumnSlide,
    ChartSlide,
    ChartWithTextSlide,
    KPISlide,
    QuoteSlide,
    ClosingSlide,
]


# --------------------------------------------------------------------------- #
# Корневая спецификация
# --------------------------------------------------------------------------- #

class PresentationSpec(BaseModel):
    """
    Полная спецификация презентации.

    Это то, что возвращает LLM и принимает на вход PresentationBuilder.
    """
    title: str = Field(description="Название презентации")
    author: Optional[str] = None
    style: PresentationStyle
    slides: List[Slide] = Field(min_length=1, description="Слайды по порядку")

    @field_validator("slides")
    @classmethod
    def first_slide_is_title(cls, v: List[Slide]) -> List[Slide]:
        # Не строгое требование, но логичное по UX
        if v and not isinstance(v[0], TitleSlide):
            # не валим — просто предупреждение в логике билдера
            pass
        return v
