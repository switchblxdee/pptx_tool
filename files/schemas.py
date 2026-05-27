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
    title: str = Field(max_length=80, description="Заголовок графика")
    categories: List[str] = Field(description="Категории (ось X или сектора pie)")
    series: List[ChartSeries] = Field(description="Один или больше рядов данных")
    x_axis_title: Optional[str] = Field(default=None, max_length=40)
    y_axis_title: Optional[str] = Field(default=None, max_length=40)

    @field_validator("categories")
    @classmethod
    def limit_category_text(cls, v: List[str]) -> List[str]:
        # Длинные подписи категорий ломают вёрстку оси X
        return [str(c)[:25].rstrip() for c in v]

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
    value: str = Field(max_length=10, description="Само число со знаком, напр. '+24%', '1.2M'")
    label: str = Field(max_length=40, description="Что означает это число (краткое описание)")


# --- Дискриминированные слайды ----------------------------------------------

class _BaseSlide(BaseModel):
    layout: SlideLayout
    speaker_notes: Optional[str] = Field(
        default=None, description="Заметки докладчика (не отображаются на слайде)"
    )


class TitleSlide(_BaseSlide):
    layout: Literal[SlideLayout.TITLE] = SlideLayout.TITLE
    title: str = Field(max_length=80)
    subtitle: Optional[str] = Field(default=None, max_length=140)


class SectionHeaderSlide(_BaseSlide):
    layout: Literal[SlideLayout.SECTION_HEADER] = SlideLayout.SECTION_HEADER
    title: str = Field(max_length=60)
    description: Optional[str] = Field(default=None, max_length=200)


class BulletsSlide(_BaseSlide):
    layout: Literal[SlideLayout.BULLETS] = SlideLayout.BULLETS
    title: str = Field(max_length=90)
    bullets: List[str] = Field(min_length=1, max_length=6)

    @field_validator("bullets")
    @classmethod
    def limit_bullet_length(cls, v: List[str]) -> List[str]:
        # Каждый буллет не длиннее 140 символов — иначе перенос съест слайд
        return [b[:140].rstrip() for b in v]


class TwoColumnSlide(_BaseSlide):
    layout: Literal[SlideLayout.TWO_COLUMN] = SlideLayout.TWO_COLUMN
    title: str = Field(max_length=90)
    left_heading: str = Field(max_length=40)
    left_bullets: List[str] = Field(min_length=1, max_length=5)
    right_heading: str = Field(max_length=40)
    right_bullets: List[str] = Field(min_length=1, max_length=5)

    @field_validator("left_bullets", "right_bullets")
    @classmethod
    def limit_bullet_length(cls, v: List[str]) -> List[str]:
        return [b[:110].rstrip() for b in v]


class ChartSlide(_BaseSlide):
    layout: Literal[SlideLayout.CHART] = SlideLayout.CHART
    title: str = Field(max_length=90)
    chart: ChartData


class ChartWithTextSlide(_BaseSlide):
    layout: Literal[SlideLayout.CHART_WITH_TEXT] = SlideLayout.CHART_WITH_TEXT
    title: str = Field(max_length=90)
    chart: ChartData
    insights: List[str] = Field(
        min_length=1, max_length=4,
        description="Ключевые выводы из графика (только наблюдения, не рекомендации)"
    )

    @field_validator("insights")
    @classmethod
    def limit_insight_length(cls, v: List[str]) -> List[str]:
        return [i[:130].rstrip() for i in v]


class KPISlide(_BaseSlide):
    layout: Literal[SlideLayout.KPI] = SlideLayout.KPI
    title: str = Field(max_length=90)
    kpis: List[KPIItem] = Field(min_length=2, max_length=4)


class QuoteSlide(_BaseSlide):
    layout: Literal[SlideLayout.QUOTE] = SlideLayout.QUOTE
    quote: str = Field(max_length=220)
    attribution: Optional[str] = Field(default=None, max_length=80)


class ClosingSlide(_BaseSlide):
    layout: Literal[SlideLayout.CLOSING] = SlideLayout.CLOSING
    title: str = Field(default="Спасибо за внимание", max_length=80)
    subtitle: Optional[str] = Field(default=None, max_length=140)


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
    title: str = Field(max_length=100, description="Название презентации")
    author: Optional[str] = Field(default=None, max_length=60)
    style: PresentationStyle
    slides: List[Slide] = Field(min_length=1, description="Слайды по порядку")

    # Маркеры "рекомендательного" контента. Если LLM проигнорировал запрет
    # в промпте и всё равно прислал такое — мы либо переименуем слайд,
    # либо отфильтруем явно императивные буллеты.
    _RECOMMENDATION_TITLE_MARKERS = (
        "рекоменд", "следующие шаги", "что делать", "план действий",
        "необходимо", "меры по устранению", "предложения",
        "плана внедрения", "дорожная карта",
    )
    _IMPERATIVE_BULLET_MARKERS = (
        "следует ", "необходимо ", "нужно ", "рекомендуется",
        "стоит внедрить", "предлагается", "требуется ", "должен быть внедрён",
    )

    @field_validator("slides")
    @classmethod
    def sanitize_recommendations(cls, slides: List[Slide]) -> List[Slide]:
        """
        Программная защита: если LLM всё-таки прислал «рекомендации»,
        мы их обезвреживаем.

        Стратегия:
        - Заголовки слайдов с маркерами рекомендаций переименовываются
          в нейтральный «Итоги анализа».
        - Императивные буллеты («следует внедрить...») отфильтровываются.
        - Если после фильтрации в слайде не осталось буллетов — добавляем
          placeholder, чтобы не валить рендер.
        """
        sanitized: List[Slide] = []
        for slide in slides:
            sanitized.append(cls._sanitize_one(slide))
        return sanitized

    @classmethod
    def _sanitize_one(cls, slide: Slide) -> Slide:
        # Переименование заголовков, похожих на рекомендации
        if hasattr(slide, "title") and isinstance(slide.title, str):
            title_lower = slide.title.lower()
            if any(m in title_lower for m in cls._RECOMMENDATION_TITLE_MARKERS):
                slide = slide.model_copy(update={"title": "Итоги анализа"})

        # Фильтрация буллетов от императивных формулировок
        if isinstance(slide, BulletsSlide):
            filtered = cls._filter_imperative(slide.bullets)
            if not filtered:
                filtered = ["Анализ данных представлен на предыдущих слайдах"]
            slide = slide.model_copy(update={"bullets": filtered})

        elif isinstance(slide, TwoColumnSlide):
            left = cls._filter_imperative(slide.left_bullets) or ["—"]
            right = cls._filter_imperative(slide.right_bullets) or ["—"]
            slide = slide.model_copy(update={
                "left_bullets": left, "right_bullets": right,
            })

        elif isinstance(slide, ChartWithTextSlide):
            filtered = cls._filter_imperative(slide.insights)
            if not filtered:
                filtered = ["См. данные на графике"]
            slide = slide.model_copy(update={"insights": filtered})

        return slide

    @classmethod
    def _filter_imperative(cls, items: List[str]) -> List[str]:
        """Убирает буллеты, начинающиеся с императивных слов."""
        result = []
        for item in items:
            lower = item.lower().strip()
            if not any(lower.startswith(m) or m in lower[:30]
                       for m in cls._IMPERATIVE_BULLET_MARKERS):
                result.append(item)
        return result
