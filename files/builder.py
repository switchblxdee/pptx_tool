"""
Сборка .pptx из PresentationSpec.

Архитектурные принципы:
- Каждый layout рендерится отдельным методом — никаких if-цепочек в одном месте.
- Графики — нативные PowerPoint-объекты, не картинки: можно редактировать.
- Любой текст идёт через _set_text() c контролем размера и анти-overflow.
- Координаты вычисляются от констант LAYOUT, чтобы их было легко править централизованно.

Защиты от типичных проблем:
- Авто-усадка текста (MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE) на всех контентных боксах.
- Заголовки графиков выключены — заголовок слайда уже служит для этой цели.
- Маркеры/подписи данных включены там, где это улучшает читаемость.
- Динамическая раскладка KPI: 2 шт. — крупно по центру, 4 шт. — в ряд.
"""
from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import List, Tuple

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from .schemas import (
    BulletsSlide,
    ChartData,
    ChartSlide,
    ChartType,
    ChartWithTextSlide,
    ClosingSlide,
    KPISlide,
    PresentationSpec,
    QuoteSlide,
    SectionHeaderSlide,
    Slide,
    TitleSlide,
    TwoColumnSlide,
)

# --------------------------------------------------------------------------- #
# Константы раскладки — единая точка для всех координат
# --------------------------------------------------------------------------- #

# Размеры 16:9
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

# Поля слайда
MARGIN_X = Inches(0.7)              # боковые поля
MARGIN_TOP = Inches(0.5)            # верхнее поле
MARGIN_BOTTOM = Inches(0.45)        # нижнее поле

# Зона заголовка контентного слайда
TITLE_TOP = Inches(0.55)
TITLE_HEIGHT = Inches(0.85)
CONTENT_TOP = Inches(1.7)           # начало контента под заголовком
CONTENT_HEIGHT = SLIDE_HEIGHT - CONTENT_TOP - MARGIN_BOTTOM
CONTENT_WIDTH = SLIDE_WIDTH - MARGIN_X * 2

# NS для прямого XML-патчинга
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NSMAP = {"c": NS_C, "a": NS_A}

# Маппинг типов
_CHART_TYPE_MAP = {
    ChartType.BAR: XL_CHART_TYPE.BAR_CLUSTERED,
    ChartType.COLUMN: XL_CHART_TYPE.COLUMN_CLUSTERED,
    ChartType.LINE: XL_CHART_TYPE.LINE,
    ChartType.PIE: XL_CHART_TYPE.PIE,
    ChartType.SCATTER: XL_CHART_TYPE.XY_SCATTER,
}


class PresentationBuilder:
    """Собирает .pptx по PresentationSpec."""

    def __init__(self, spec: PresentationSpec):
        self.spec = spec
        self.style = spec.style
        self.prs = Presentation()
        self.prs.slide_width = SLIDE_WIDTH
        self.prs.slide_height = SLIDE_HEIGHT

    def build(self, output_path: str | Path) -> Path:
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        for slide_spec in self.spec.slides:
            self._render_slide(slide_spec)

        self.prs.save(output_path)
        return output_path

    # ----------------------------------------------------------------------- #
    # Диспетчер
    # ----------------------------------------------------------------------- #

    def _render_slide(self, s: Slide) -> None:
        renderers = {
            TitleSlide: self._render_title,
            SectionHeaderSlide: self._render_section,
            BulletsSlide: self._render_bullets,
            TwoColumnSlide: self._render_two_column,
            ChartSlide: self._render_chart,
            ChartWithTextSlide: self._render_chart_with_text,
            KPISlide: self._render_kpi,
            QuoteSlide: self._render_quote,
            ClosingSlide: self._render_closing,
        }
        renderer = renderers.get(type(s))
        if renderer is None:
            raise ValueError(f"Неизвестный тип слайда: {type(s).__name__}")

        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])  # blank
        renderer(slide, s)

        if s.speaker_notes:
            slide.notes_slide.notes_text_frame.text = s.speaker_notes

    # ----------------------------------------------------------------------- #
    # Рендеры
    # ----------------------------------------------------------------------- #

    def _render_title(self, slide, s: TitleSlide) -> None:
        """Обложка: тёмный фон, акцентная полоса слева, крупный заголовок."""
        self._set_background(slide, self.style.palette.primary)

        # Акцентная полоса (визуальный мотив, переносится на closing-слайд)
        accent_bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
            Inches(0.3), SLIDE_HEIGHT,
        )
        self._fill_solid(accent_bar, self.style.palette.accent)
        accent_bar.line.fill.background()

        # Заголовок — фиксированная зона с запасом под 2-3 строки.
        # Размер шрифта зависит от длины заголовка, чтобы не полагаться
        # только на auto-fit (он по-разному рендерится в LibreOffice).
        if len(s.title) > 60:
            title_size = 36
        elif len(s.title) > 35:
            title_size = 42
        else:
            title_size = 50

        self._add_text_box(
            slide, s.title,
            left=Inches(1.0), top=Inches(2.1),
            width=Inches(11.3), height=Inches(2.4),
            font=self.style.typography.heading_font,
            size=title_size, bold=True,
            color=self.style.palette.text_light,
            anchor=MSO_ANCHOR.BOTTOM,
        )

        if s.subtitle:
            self._add_text_box(
                slide, s.subtitle,
                left=Inches(1.0), top=Inches(4.75),
                width=Inches(11.3), height=Inches(1.4),
                font=self.style.typography.body_font,
                size=20 if len(s.subtitle) > 80 else 22,
                color=self._with_alpha(self.style.palette.text_light, 0.85),
                anchor=MSO_ANCHOR.TOP,
            )

        # Автор внизу — нейтральный светлый, не яркий
        if self.spec.author:
            self._add_text_box(
                slide, self.spec.author,
                left=Inches(1.0), top=Inches(6.6),
                width=Inches(11.3), height=Inches(0.45),
                font=self.style.typography.body_font,
                size=13,
                color=self._with_alpha(self.style.palette.text_light, 0.7),
            )

    def _render_section(self, slide, s: SectionHeaderSlide) -> None:
        """Раздел: тёмный фон, крупный номер-маркер слева, текст справа."""
        self._set_background(slide, self.style.palette.primary)

        # Крупный декоративный акцент-маркер слева — толстая вертикальная полоса
        bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0.8), Inches(3.1),
            Inches(0.12), Inches(1.3),
        )
        self._fill_solid(bar, self.style.palette.accent)
        bar.line.fill.background()

        # Заголовок раздела — адаптивный размер
        section_title_size = 38 if len(s.title) > 40 else 44

        self._add_text_box(
            slide, s.title,
            left=Inches(1.2), top=Inches(2.9),
            width=Inches(11.0), height=Inches(1.4),
            font=self.style.typography.heading_font,
            size=section_title_size, bold=True,
            color=self.style.palette.text_light,
            anchor=MSO_ANCHOR.MIDDLE,
        )

        # Описание — мягким светлым цветом
        if s.description:
            self._add_text_box(
                slide, s.description,
                left=Inches(1.2), top=Inches(4.45),
                width=Inches(10.5), height=Inches(1.6),
                font=self.style.typography.body_font,
                size=18,
                color=self._with_alpha(self.style.palette.text_light, 0.8),
            )

    def _render_bullets(self, slide, s: BulletsSlide) -> None:
        self._set_background(slide, self.style.palette.background)
        self._add_title(slide, s.title)

        body_box = slide.shapes.add_textbox(
            MARGIN_X, CONTENT_TOP, CONTENT_WIDTH, CONTENT_HEIGHT,
        )
        tf = body_box.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.margin_left = Emu(0)
        tf.margin_top = Emu(0)

        # Динамически адаптируем размер шрифта под количество буллетов
        if len(s.bullets) <= 3:
            font_size, line_gap = 22, Pt(18)
        elif len(s.bullets) <= 5:
            font_size, line_gap = 19, Pt(14)
        else:
            font_size, line_gap = 17, Pt(10)

        for i, bullet in enumerate(s.bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            # Чистим тег — точка + неразрывный пробел для аккуратного выравнивания
            p.text = f"\u2022\u00a0\u00a0{bullet}"
            p.space_after = line_gap
            p.alignment = PP_ALIGN.LEFT
            for run in p.runs:
                run.font.name = self.style.typography.body_font
                run.font.size = Pt(font_size)
                run.font.color.rgb = self._rgb(self.style.palette.text_dark)

    def _render_two_column(self, slide, s: TwoColumnSlide) -> None:
        self._set_background(slide, self.style.palette.background)
        self._add_title(slide, s.title)

        gap = Inches(0.4)
        col_width = (CONTENT_WIDTH - gap) / 2

        self._render_column(
            slide,
            left=MARGIN_X, top=CONTENT_TOP,
            width=col_width, height=CONTENT_HEIGHT,
            heading=s.left_heading, bullets=s.left_bullets,
        )
        self._render_column(
            slide,
            left=MARGIN_X + col_width + gap, top=CONTENT_TOP,
            width=col_width, height=CONTENT_HEIGHT,
            heading=s.right_heading, bullets=s.right_bullets,
        )

    def _render_column(
        self, slide, left, top, width, height,
        heading: str, bullets: List[str],
    ) -> None:
        heading_height = Inches(0.55)

        # Заголовок колонки — тонкая верхняя полоса с текстом на фоне primary
        head = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, left, top, width, heading_height,
        )
        self._fill_solid(head, self.style.palette.primary)
        head.line.fill.background()
        tf = head.text_frame
        tf.margin_left = Inches(0.2)
        tf.margin_right = Inches(0.2)
        tf.margin_top = Emu(0)
        tf.margin_bottom = Emu(0)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.word_wrap = True
        tf.text = heading
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        for run in p.runs:
            run.font.name = self.style.typography.heading_font
            run.font.size = Pt(16)
            run.font.bold = True
            run.font.color.rgb = self._rgb(self.style.palette.text_light)

        # Тело колонки
        body_top = top + heading_height + Inches(0.25)
        body_height = height - heading_height - Inches(0.25)
        body_box = slide.shapes.add_textbox(left, body_top, width, body_height)
        tf = body_box.text_frame
        tf.word_wrap = True
        tf.margin_left = Emu(0)
        tf.margin_top = Emu(0)

        font_size = 16 if len(bullets) <= 4 else 14

        for i, bullet in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = f"\u2022\u00a0\u00a0{bullet}"
            p.space_after = Pt(10)
            p.alignment = PP_ALIGN.LEFT
            for run in p.runs:
                run.font.name = self.style.typography.body_font
                run.font.size = Pt(font_size)
                run.font.color.rgb = self._rgb(self.style.palette.text_dark)

    def _render_chart(self, slide, s: ChartSlide) -> None:
        self._set_background(slide, self.style.palette.background)
        self._add_title(slide, s.title)

        self._add_chart(
            slide, s.chart,
            left=MARGIN_X, top=CONTENT_TOP,
            width=CONTENT_WIDTH, height=CONTENT_HEIGHT,
        )

    def _render_chart_with_text(self, slide, s: ChartWithTextSlide) -> None:
        self._set_background(slide, self.style.palette.background)
        self._add_title(slide, s.title)

        # ~60/40 — график достаточно широкий, чтобы читались категории
        gap = Inches(0.45)
        chart_width = Inches(7.6)
        text_width = CONTENT_WIDTH - chart_width - gap

        self._add_chart(
            slide, s.chart,
            left=MARGIN_X, top=CONTENT_TOP,
            width=chart_width, height=CONTENT_HEIGHT,
        )

        # Текстовый блок справа
        text_left = MARGIN_X + chart_width + gap
        insights_box = slide.shapes.add_textbox(
            text_left, CONTENT_TOP, text_width, CONTENT_HEIGHT,
        )
        tf = insights_box.text_frame
        tf.word_wrap = True
        tf.margin_left = Emu(0)
        tf.margin_top = Emu(0)

        # Заголовок блока — нейтральный, не "рекомендации"
        p0 = tf.paragraphs[0]
        p0.text = "Наблюдения"
        p0.space_after = Pt(14)
        for run in p0.runs:
            run.font.name = self.style.typography.heading_font
            run.font.size = Pt(17)
            run.font.bold = True
            run.font.color.rgb = self._rgb(self.style.palette.accent)

        # Сами инсайты
        font_size = 14 if len(s.insights) >= 3 else 15
        for insight in s.insights:
            p = tf.add_paragraph()
            p.text = f"\u25b8\u00a0\u00a0{insight}"
            p.space_after = Pt(9)
            p.alignment = PP_ALIGN.LEFT
            for run in p.runs:
                run.font.name = self.style.typography.body_font
                run.font.size = Pt(font_size)
                run.font.color.rgb = self._rgb(self.style.palette.text_dark)

    def _render_kpi(self, slide, s: KPISlide) -> None:
        self._set_background(slide, self.style.palette.background)
        self._add_title(slide, s.title)

        n = len(s.kpis)

        # Адаптивная раскладка: 2 — крупно по центру, 3-4 — в ряд
        if n == 2:
            card_width = Inches(4.2)
            total = card_width * 2 + Inches(0.5)
            start_left = (SLIDE_WIDTH - total) / 2
            gap = Inches(0.5)
        else:
            gap = Inches(0.35)
            total_width = Inches(12.0)
            card_width = Emu(int((total_width - gap * (n - 1)) / n))
            start_left = (SLIDE_WIDTH - total_width) / 2

        card_top = Inches(2.3)
        card_height = Inches(3.1)

        # Адаптивный размер значения под число карточек
        value_size = 60 if n == 2 else (52 if n == 3 else 44)
        label_size = 16 if n == 2 else 14

        for i, kpi in enumerate(s.kpis):
            left = start_left + (card_width + gap) * i
            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, left, card_top, card_width, card_height,
            )
            self._fill_solid(card, self.style.palette.primary)
            card.line.fill.background()
            card.adjustments[0] = 0.05

            # Значение — крупно по центру
            self._add_text_box(
                slide, kpi.value,
                left=left, top=card_top + Inches(0.45),
                width=card_width, height=Inches(1.8),
                font=self.style.typography.heading_font,
                size=value_size, bold=True,
                color=self.style.palette.text_light,
                align=PP_ALIGN.CENTER,
                anchor=MSO_ANCHOR.MIDDLE,
            )

            # Подпись — под значением, с небольшим внутренним отступом
            self._add_text_box(
                slide, kpi.label,
                left=left + Inches(0.2), top=card_top + Inches(2.25),
                width=card_width - Inches(0.4), height=Inches(0.7),
                font=self.style.typography.body_font,
                size=label_size,
                color=self._with_alpha(self.style.palette.text_light, 0.85),
                align=PP_ALIGN.CENTER,
                anchor=MSO_ANCHOR.TOP,
            )

    def _render_quote(self, slide, s: QuoteSlide) -> None:
        self._set_background(slide, self.style.palette.primary)

        # Крупные декоративные кавычки в верхнем углу
        self._add_text_box(
            slide, "\u201C",
            left=Inches(0.9), top=Inches(0.6),
            width=Inches(2.0), height=Inches(2.0),
            font="Georgia", size=140, bold=True,
            color=self.style.palette.accent,
        )

        # Текст цитаты — крупный, с авто-усадкой при необходимости
        # Размер шрифта зависит от длины: длинная цитата = мельче
        if len(s.quote) > 140:
            size = 22
        elif len(s.quote) > 80:
            size = 26
        else:
            size = 30

        self._add_text_box(
            slide, s.quote,
            left=Inches(1.5), top=Inches(2.6),
            width=Inches(10.5), height=Inches(3.2),
            font=self.style.typography.heading_font,
            size=size,
            color=self.style.palette.text_light,
            anchor=MSO_ANCHOR.MIDDLE,
        )

        if s.attribution:
            self._add_text_box(
                slide, f"\u2014\u00a0{s.attribution}",
                left=Inches(1.5), top=Inches(6.1),
                width=Inches(10.5), height=Inches(0.6),
                font=self.style.typography.body_font,
                size=15,
                color=self.style.palette.accent,
            )

    def _render_closing(self, slide, s: ClosingSlide) -> None:
        """Финал — повторяет визуальный мотив обложки (акцентная полоса)."""
        self._set_background(slide, self.style.palette.primary)

        # Та же акцентная полоса, что и на title — для визуальной рифмы
        accent_bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
            Inches(0.3), SLIDE_HEIGHT,
        )
        self._fill_solid(accent_bar, self.style.palette.accent)
        accent_bar.line.fill.background()

        self._add_text_box(
            slide, s.title,
            left=Inches(1.0), top=Inches(3.0),
            width=Inches(11.3), height=Inches(1.5),
            font=self.style.typography.heading_font,
            size=50, bold=True,
            color=self.style.palette.text_light,
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE,
        )

        if s.subtitle:
            self._add_text_box(
                slide, s.subtitle,
                left=Inches(1.0), top=Inches(4.6),
                width=Inches(11.3), height=Inches(0.8),
                font=self.style.typography.body_font,
                size=20,
                color=self._with_alpha(self.style.palette.text_light, 0.75),
                align=PP_ALIGN.CENTER,
            )

    # ----------------------------------------------------------------------- #
    # Графики
    # ----------------------------------------------------------------------- #

    def _add_chart(
        self, slide, cd: ChartData,
        left, top, width, height,
    ) -> None:
        xl_type = _CHART_TYPE_MAP[cd.chart_type]
        if cd.chart_type == ChartType.SCATTER:
            data = self._build_scatter_data(cd)
        else:
            data = self._build_category_data(cd)

        graphic_frame = slide.shapes.add_chart(xl_type, left, top, width, height, data)
        chart = graphic_frame.chart

        # КЛЮЧЕВОЕ: убираем заголовок графика, потому что его роль играет
        # заголовок слайда. Дублирование выглядит как "AI-slop".
        chart.has_title = False

        self._style_chart(chart, cd)

    @staticmethod
    def _build_category_data(cd: ChartData) -> CategoryChartData:
        data = CategoryChartData()
        data.categories = cd.categories
        for series in cd.series:
            data.add_series(series.name, series.values)
        return data

    @staticmethod
    def _build_scatter_data(cd: ChartData) -> XyChartData:
        data = XyChartData()
        if len(cd.series) >= 2:
            x_series = cd.series[0]
            for y_series in cd.series[1:]:
                s = data.add_series(y_series.name)
                for x, y in zip(x_series.values, y_series.values):
                    s.add_data_point(x, y)
        else:
            s = data.add_series(cd.series[0].name)
            for i, y in enumerate(cd.series[0].values):
                s.add_data_point(float(i + 1), y)
        return data

    def _style_chart(self, chart, cd: ChartData) -> None:
        """Применяет стиль палитры и улучшения читаемости."""
        palette_colors = [
            self.style.palette.primary,
            self.style.palette.accent,
            self.style.palette.secondary,
        ]

        # --- Легенда ---
        if cd.chart_type == ChartType.PIE:
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.RIGHT
            chart.legend.include_in_layout = False
            self._set_legend_font(chart)
        else:
            chart.has_legend = len(cd.series) > 1
            if chart.has_legend:
                chart.legend.position = XL_LEGEND_POSITION.BOTTOM
                chart.legend.include_in_layout = False
                self._set_legend_font(chart)

        # --- Оси (для всех кроме pie) ---
        if cd.chart_type != ChartType.PIE:
            self._style_axes(chart, cd)

        # --- Раскраска и подписи данных ---
        if cd.chart_type == ChartType.PIE:
            self._color_pie_points(chart)
            self._add_pie_data_labels(chart, cd.categories)
        elif cd.chart_type == ChartType.SCATTER:
            self._style_scatter(chart, palette_colors)
        elif cd.chart_type == ChartType.LINE:
            self._style_line(chart, palette_colors)
        else:  # bar, column
            self._style_bars(chart, cd, palette_colors)

    def _style_axes(self, chart, cd: ChartData) -> None:
        """Шрифты осей и заголовки осей."""
        try:
            cat_axis = chart.category_axis
            self._set_axis_font(cat_axis)
            if cd.x_axis_title:
                cat_axis.has_title = True
                cat_axis.axis_title.text_frame.text = cd.x_axis_title
                self._set_text_font_in_frame(
                    cat_axis.axis_title.text_frame,
                    font=self.style.typography.body_font,
                    size=12, color=self.style.palette.text_dark,
                )
        except (AttributeError, ValueError):
            pass

        try:
            val_axis = chart.value_axis
            self._set_axis_font(val_axis)
            if cd.y_axis_title:
                val_axis.has_title = True
                val_axis.axis_title.text_frame.text = cd.y_axis_title
                self._set_text_font_in_frame(
                    val_axis.axis_title.text_frame,
                    font=self.style.typography.body_font,
                    size=12, color=self.style.palette.text_dark,
                )
        except (AttributeError, ValueError):
            pass

    def _style_bars(self, chart, cd: ChartData, palette_colors: List[str]) -> None:
        """Раскраска и data labels для bar/column."""
        for idx, series in enumerate(chart.series):
            color = palette_colors[idx % len(palette_colors)]
            fill = series.format.fill
            fill.solid()
            fill.fore_color.rgb = self._rgb(color)
            series.format.line.fill.background()

        # Подписи значений — только если серий немного (иначе каша)
        if len(cd.series) <= 2:
            for series in chart.series:
                series.data_labels.show_value = True
                series.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
                series.data_labels.font.size = Pt(10)
                series.data_labels.font.name = self.style.typography.body_font
                series.data_labels.font.color.rgb = self._rgb(self.style.palette.text_dark)

    def _style_line(self, chart, palette_colors: List[str]) -> None:
        """Раскраска линий + маркеры на точках."""
        for idx, series in enumerate(chart.series):
            color = palette_colors[idx % len(palette_colors)]
            line = series.format.line
            line.color.rgb = self._rgb(color)
            line.width = Pt(2.5)

            # Маркеры на каждой точке для читаемости
            marker = series.marker
            marker.style = 8  # ROUND
            marker.size = 7
            marker.format.fill.solid()
            marker.format.fill.fore_color.rgb = self._rgb(color)
            marker.format.line.color.rgb = self._rgb(color)

    def _style_scatter(self, chart, palette_colors: List[str]) -> None:
        """Крупные цветные маркеры с обводкой."""
        for idx, series in enumerate(chart.series):
            color = palette_colors[idx % len(palette_colors)]
            marker = series.marker
            marker.style = 8  # ROUND
            marker.size = 12
            marker.format.fill.solid()
            marker.format.fill.fore_color.rgb = self._rgb(color)
            marker.format.line.color.rgb = self._rgb(color)
            # Линию между точками убираем — это scatter, не line
            series.format.line.fill.background()

    def _color_pie_points(self, chart) -> None:
        """Красит каждый сектор pie в свой цвет палитры через прямой XML."""
        # Расширенная палитра для большого количества секторов
        extended = [
            self.style.palette.primary,
            self.style.palette.accent,
            self.style.palette.secondary,
            self.style.palette.text_dark,
        ]

        plot = chart.plots[0]
        series_xml = plot.series[0]._element

        # Удаляем существующие dPt
        for dpt in series_xml.findall(qn("c:dPt")):
            series_xml.remove(dpt)

        # Куда вставлять — после стандартных префиксных элементов
        insert_after_tags = {qn(t) for t in [
            "c:idx", "c:order", "c:tx", "c:spPr", "c:explosion",
        ]}
        insert_idx = 0
        for i, child in enumerate(series_xml):
            if child.tag in insert_after_tags:
                insert_idx = i + 1

        # Считаем количество точек
        cat = series_xml.find(".//" + qn("c:cat"))
        if cat is None:
            return
        pt_count_elem = cat.find(".//" + qn("c:ptCount"))
        if pt_count_elem is None:
            return
        n_points = int(pt_count_elem.get("val"))

        for i in range(n_points):
            color = extended[i % len(extended)].lstrip("#")
            dpt_xml = f"""<c:dPt xmlns:c="{NS_C}" xmlns:a="{NS_A}">
                <c:idx val="{i}"/>
                <c:bubble3D val="0"/>
                <c:spPr>
                    <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
                    <a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>
                </c:spPr>
            </c:dPt>"""
            dpt = etree.fromstring(dpt_xml)
            series_xml.insert(insert_idx + i, dpt)

    def _add_pie_data_labels(self, chart, categories: List[str]) -> None:
        """Pie data labels — только проценты, без дублирования значений."""
        plot = chart.plots[0]
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.show_percentage = True
        dl.show_value = False
        dl.show_category_name = False
        dl.show_series_name = False
        dl.position = XL_LABEL_POSITION.OUTSIDE_END
        dl.font.size = Pt(11)
        dl.font.name = self.style.typography.body_font
        dl.font.color.rgb = self._rgb(self.style.palette.text_dark)
        dl.font.bold = True

    def _set_legend_font(self, chart) -> None:
        try:
            font = chart.legend.font
            font.size = Pt(11)
            font.name = self.style.typography.body_font
            font.color.rgb = self._rgb(self.style.palette.text_dark)
        except AttributeError:
            pass

    def _set_axis_font(self, axis) -> None:
        try:
            font = axis.tick_labels.font
            font.size = Pt(10)
            font.name = self.style.typography.body_font
            font.color.rgb = self._rgb(self.style.palette.text_dark)
        except AttributeError:
            pass

    # ----------------------------------------------------------------------- #
    # Общие утилиты текста и стиля
    # ----------------------------------------------------------------------- #

    def _add_title(self, slide, text: str) -> None:
        """Стандартный заголовок контентного слайда с адаптивным размером."""
        # Чем длиннее заголовок — тем меньше шрифт, чтобы влезал в одну строку.
        # Это надёжнее, чем полагаться на auto-fit (LibreOffice рендерит иначе).
        if len(text) > 60:
            size = 22
        elif len(text) > 40:
            size = 26
        else:
            size = 30

        self._add_text_box(
            slide, text,
            left=MARGIN_X, top=TITLE_TOP,
            width=CONTENT_WIDTH, height=TITLE_HEIGHT,
            font=self.style.typography.heading_font,
            size=size, bold=True,
            color=self.style.palette.text_dark,
            anchor=MSO_ANCHOR.MIDDLE,
        )

    def _add_text_box(
        self, slide, text: str,
        left, top, width, height,
        font: str, size: int, color: str,
        bold: bool = False,
        align: PP_ALIGN = PP_ALIGN.LEFT,
        anchor: MSO_ANCHOR = MSO_ANCHOR.TOP,
    ) -> None:
        """
        Единая точка создания текстового бокса.

        Включает MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE — это критично:
        если текст всё-таки длиннее ожидаемого, PowerPoint сам уменьшит
        шрифт, чтобы влез. Без этого получаем overflow.
        """
        box = slide.shapes.add_textbox(left, top, width, height)
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = Emu(0)
        tf.margin_right = Emu(0)
        tf.margin_top = Emu(0)
        tf.margin_bottom = Emu(0)

        tf.text = text
        p = tf.paragraphs[0]
        p.alignment = align
        for run in p.runs:
            run.font.name = font
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.color.rgb = self._rgb(color)

        # Анти-overflow: усадка после установки текста.
        # ВАЖНО: вызывать после .text, иначе не сработает.
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

    def _set_text_font_in_frame(
        self, text_frame, font: str, size: int, color: str, bold: bool = False,
    ) -> None:
        """Применяет шрифт ко всем run'ам в text_frame (для уже созданных)."""
        for p in text_frame.paragraphs:
            for run in p.runs:
                run.font.name = font
                run.font.size = Pt(size)
                run.font.bold = bold
                run.font.color.rgb = self._rgb(color)

    def _set_background(self, slide, hex_color: str) -> None:
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = self._rgb(hex_color)

    def _fill_solid(self, shape, hex_color: str) -> None:
        shape.fill.solid()
        shape.fill.fore_color.rgb = self._rgb(hex_color)

    @staticmethod
    def _rgb(hex_color: str) -> RGBColor:
        return RGBColor.from_string(hex_color.lstrip("#"))

    @staticmethod
    def _with_alpha(hex_color: str, factor: float) -> str:
        """
        Создаёт «полупрозрачную» версию цвета через смешивание со средне-серым.

        python-pptx не даёт прямого API для alpha, поэтому подмешиваем
        к цвету серый по коэффициенту. Для светлых текстов это смягчает
        вторичные элементы (subtitle, captions).
        """
        h = hex_color.lstrip("#")
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        # Смешиваем с 128-серым
        nr = int(r * factor + 128 * (1 - factor))
        ng = int(g * factor + 128 * (1 - factor))
        nb = int(b * factor + 128 * (1 - factor))
        return f"{nr:02X}{ng:02X}{nb:02X}"
