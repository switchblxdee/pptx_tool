"""
Сборка .pptx из PresentationSpec.

Каждый лэйаут рендерится отдельным методом. Графики — нативные
объекты PowerPoint (через python-pptx), не картинки: их можно
редактировать в самом PowerPoint после генерации.

Архитектурно: builder знает только про spec и про pptx-API.
LLM сюда не достаёт. Это делает рендер детерминированным и тестируемым.
"""
from __future__ import annotations

from pathlib import Path
from typing import Tuple

from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
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
    PresentationStyle,
    QuoteSlide,
    SectionHeaderSlide,
    Slide,
    TitleSlide,
    TwoColumnSlide,
)

# Размеры 16:9
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

# Соответствие типов графиков
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
        """Главная точка входа — собирает и сохраняет."""
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        for slide_spec in self.spec.slides:
            self._render_slide(slide_spec)

        self.prs.save(output_path)
        return output_path

    # ----------------------------------------------------------------------- #
    # Диспетчер слайдов
    # ----------------------------------------------------------------------- #

    def _render_slide(self, s: Slide) -> None:
        """Маршрутизация по типу слайда (Pydantic-дискриминация)."""
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
    # Рендеры по типам
    # ----------------------------------------------------------------------- #

    def _render_title(self, slide, s: TitleSlide) -> None:
        # Темный фон обложки = primary цвет палитры
        self._set_background(slide, self.style.palette.primary)

        # Декоративный акцентный прямоугольник слева (визуальный мотив)
        accent_bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.35), SLIDE_HEIGHT
        )
        self._fill_solid(accent_bar, self.style.palette.accent)
        accent_bar.line.fill.background()

        # Заголовок
        title_box = slide.shapes.add_textbox(
            Inches(1.0), Inches(2.6), Inches(11.3), Inches(2.0)
        )
        self._set_text(
            title_box.text_frame, s.title,
            font=self.style.typography.heading_font,
            size=54, bold=True, color=self.style.palette.text_light,
        )

        if s.subtitle:
            sub_box = slide.shapes.add_textbox(
                Inches(1.0), Inches(4.5), Inches(11.3), Inches(0.8)
            )
            self._set_text(
                sub_box.text_frame, s.subtitle,
                font=self.style.typography.body_font,
                size=22, color=self.style.palette.text_light,
            )

        # Автор внизу
        if self.spec.author:
            author_box = slide.shapes.add_textbox(
                Inches(1.0), Inches(6.6), Inches(11.3), Inches(0.5)
            )
            self._set_text(
                author_box.text_frame, self.spec.author,
                font=self.style.typography.body_font,
                size=14, color=self.style.palette.accent,
            )

    def _render_section(self, slide, s: SectionHeaderSlide) -> None:
        self._set_background(slide, self.style.palette.secondary)

        title_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(2.5), Inches(11.7), Inches(2.0)
        )
        self._set_text(
            title_box.text_frame, s.title,
            font=self.style.typography.heading_font,
            size=44, bold=True, color=self.style.palette.text_light,
        )

        if s.description:
            desc_box = slide.shapes.add_textbox(
                Inches(0.8), Inches(4.3), Inches(11.7), Inches(2.0)
            )
            self._set_text(
                desc_box.text_frame, s.description,
                font=self.style.typography.body_font,
                size=20, color=self.style.palette.text_light,
            )

    def _render_bullets(self, slide, s: BulletsSlide) -> None:
        self._set_background(slide, self.style.palette.background)
        self._add_title(slide, s.title)

        body_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(1.7), Inches(11.7), Inches(5.5)
        )
        tf = body_box.text_frame
        tf.word_wrap = True
        tf.margin_left = Emu(0)

        for i, bullet in enumerate(s.bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = f"●  {bullet}"
            p.space_after = Pt(14)
            for run in p.runs:
                run.font.name = self.style.typography.body_font
                run.font.size = Pt(20)
                run.font.color.rgb = self._rgb(self.style.palette.text_dark)

    def _render_two_column(self, slide, s: TwoColumnSlide) -> None:
        self._set_background(slide, self.style.palette.background)
        self._add_title(slide, s.title)

        col_width = Inches(5.8)
        col_top = Inches(1.7)
        col_height = Inches(5.4)
        gap = Inches(0.4)

        left = (Inches(0.7), col_top, col_width, col_height)
        right = (Inches(0.7) + col_width + gap, col_top, col_width, col_height)

        self._render_column(slide, left, s.left_heading, s.left_bullets)
        self._render_column(slide, right, s.right_heading, s.right_bullets)

    def _render_column(
        self, slide, rect: Tuple[Inches, Inches, Inches, Inches],
        heading: str, bullets: list,
    ) -> None:
        left, top, width, height = rect

        # Заголовок колонки на фоне акцента
        head_box = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.6)
        )
        self._fill_solid(head_box, self.style.palette.primary)
        head_box.line.fill.background()
        self._set_text(
            head_box.text_frame, heading,
            font=self.style.typography.heading_font,
            size=18, bold=True, color=self.style.palette.text_light,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE,
            margin_left=Inches(0.2),
        )

        body_box = slide.shapes.add_textbox(
            left, top + Inches(0.7), width, height - Inches(0.7)
        )
        tf = body_box.text_frame
        tf.word_wrap = True
        for i, bullet in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = f"●  {bullet}"
            p.space_after = Pt(10)
            for run in p.runs:
                run.font.name = self.style.typography.body_font
                run.font.size = Pt(16)
                run.font.color.rgb = self._rgb(self.style.palette.text_dark)

    def _render_chart(self, slide, s: ChartSlide) -> None:
        self._set_background(slide, self.style.palette.background)
        self._add_title(slide, s.title)

        self._add_chart(
            slide, s.chart,
            left=Inches(0.7), top=Inches(1.6),
            width=Inches(11.9), height=Inches(5.5),
        )

    def _render_chart_with_text(self, slide, s: ChartWithTextSlide) -> None:
        self._set_background(slide, self.style.palette.background)
        self._add_title(slide, s.title)

        # График слева (60% ширины)
        self._add_chart(
            slide, s.chart,
            left=Inches(0.5), top=Inches(1.6),
            width=Inches(7.5), height=Inches(5.5),
        )

        # Инсайты справа
        insight_box = slide.shapes.add_textbox(
            Inches(8.3), Inches(1.6), Inches(4.4), Inches(5.5)
        )
        tf = insight_box.text_frame
        tf.word_wrap = True

        # Подзаголовок "Ключевые выводы"
        p0 = tf.paragraphs[0]
        p0.text = "Ключевые выводы"
        for run in p0.runs:
            run.font.name = self.style.typography.heading_font
            run.font.size = Pt(18)
            run.font.bold = True
            run.font.color.rgb = self._rgb(self.style.palette.accent)
        p0.space_after = Pt(14)

        for insight in s.insights:
            p = tf.add_paragraph()
            p.text = f"▸  {insight}"
            p.space_after = Pt(10)
            for run in p.runs:
                run.font.name = self.style.typography.body_font
                run.font.size = Pt(14)
                run.font.color.rgb = self._rgb(self.style.palette.text_dark)

    def _render_kpi(self, slide, s: KPISlide) -> None:
        self._set_background(slide, self.style.palette.background)
        self._add_title(slide, s.title)

        n = len(s.kpis)
        # Раскладываем карточки в ряд
        total_width = Inches(12.0)
        gap = Inches(0.3)
        card_width = Emu(int((total_width - gap * (n - 1)) / n))
        start_left = Inches(0.65)
        top = Inches(2.4)
        card_height = Inches(3.0)

        for i, kpi in enumerate(s.kpis):
            left = start_left + (card_width + gap) * i
            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, left, top, card_width, card_height
            )
            self._fill_solid(card, self.style.palette.primary)
            card.line.fill.background()
            card.adjustments[0] = 0.06  # лёгкое скругление

            # Значение (большое)
            val_box = slide.shapes.add_textbox(
                left, top + Inches(0.5), card_width, Inches(1.6)
            )
            self._set_text(
                val_box.text_frame, kpi.value,
                font=self.style.typography.heading_font,
                size=54, bold=True, color=self.style.palette.text_light,
                align=PP_ALIGN.CENTER,
            )

            # Подпись
            lbl_box = slide.shapes.add_textbox(
                left + Inches(0.15), top + Inches(2.1),
                card_width - Inches(0.3), Inches(0.8),
            )
            self._set_text(
                lbl_box.text_frame, kpi.label,
                font=self.style.typography.body_font,
                size=14, color=self.style.palette.text_light,
                align=PP_ALIGN.CENTER,
            )

    def _render_quote(self, slide, s: QuoteSlide) -> None:
        self._set_background(slide, self.style.palette.primary)

        # Большие кавычки как декоративный элемент
        quote_mark = slide.shapes.add_textbox(
            Inches(0.8), Inches(0.8), Inches(2.0), Inches(2.0)
        )
        self._set_text(
            quote_mark.text_frame, "\u201C",
            font="Georgia", size=140, bold=True,
            color=self.style.palette.accent,
        )

        # Текст цитаты
        quote_box = slide.shapes.add_textbox(
            Inches(1.5), Inches(2.5), Inches(10.5), Inches(3.0)
        )
        self._set_text(
            quote_box.text_frame, s.quote,
            font=self.style.typography.heading_font,
            size=28, color=self.style.palette.text_light,
        )

        if s.attribution:
            attr_box = slide.shapes.add_textbox(
                Inches(1.5), Inches(5.8), Inches(10.5), Inches(0.6)
            )
            self._set_text(
                attr_box.text_frame, f"— {s.attribution}",
                font=self.style.typography.body_font,
                size=16, color=self.style.palette.accent,
            )

    def _render_closing(self, slide, s: ClosingSlide) -> None:
        self._set_background(slide, self.style.palette.primary)

        title_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(2.8), Inches(11.7), Inches(1.5)
        )
        self._set_text(
            title_box.text_frame, s.title,
            font=self.style.typography.heading_font,
            size=54, bold=True, color=self.style.palette.text_light,
            align=PP_ALIGN.CENTER,
        )

        if s.subtitle:
            sub_box = slide.shapes.add_textbox(
                Inches(0.8), Inches(4.5), Inches(11.7), Inches(1.0)
            )
            self._set_text(
                sub_box.text_frame, s.subtitle,
                font=self.style.typography.body_font,
                size=22, color=self.style.palette.accent,
                align=PP_ALIGN.CENTER,
            )

    # ----------------------------------------------------------------------- #
    # Графики
    # ----------------------------------------------------------------------- #

    def _add_chart(
        self, slide, chart_data: ChartData,
        left: Inches, top: Inches, width: Inches, height: Inches,
    ) -> None:
        """Создаёт нативный pptx-чарт нужного типа."""
        xl_type = _CHART_TYPE_MAP[chart_data.chart_type]

        if chart_data.chart_type == ChartType.SCATTER:
            data = self._build_scatter_data(chart_data)
        else:
            data = self._build_category_data(chart_data)

        graphic_frame = slide.shapes.add_chart(xl_type, left, top, width, height, data)
        chart = graphic_frame.chart

        self._style_chart(chart, chart_data)

    @staticmethod
    def _build_category_data(cd: ChartData) -> CategoryChartData:
        data = CategoryChartData()
        data.categories = cd.categories
        for series in cd.series:
            data.add_series(series.name, series.values)
        return data

    @staticmethod
    def _build_scatter_data(cd: ChartData) -> XyChartData:
        """
        Для scatter: первая серия — X, вторая+ — Y.
        Если серия одна, используем индексы как X.
        """
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
        """Применяет стиль палитры к графику."""
        chart.has_title = True
        chart.chart_title.text_frame.text = cd.title
        for run in chart.chart_title.text_frame.paragraphs[0].runs:
            run.font.name = self.style.typography.heading_font
            run.font.size = Pt(18)
            run.font.bold = True
            run.font.color.rgb = self._rgb(self.style.palette.text_dark)

        # Легенда — снизу, кроме pie (там встроенные подписи)
        if cd.chart_type == ChartType.PIE:
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.RIGHT
            chart.legend.include_in_layout = False
            # Подписи долей на pie
            plot = chart.plots[0]
            plot.has_data_labels = True
            plot.data_labels.show_percentage = True
            plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        else:
            chart.has_legend = len(cd.series) > 1
            if chart.has_legend:
                chart.legend.position = XL_LEGEND_POSITION.BOTTOM
                chart.legend.include_in_layout = False

        # Раскраска серий цветами палитры
        palette_colors = [
            self.style.palette.primary,
            self.style.palette.accent,
            self.style.palette.secondary,
        ]

        if cd.chart_type == ChartType.PIE:
            # У pie одна серия, но раскрашивать надо каждую точку отдельно
            self._color_pie_points(chart, palette_colors)
        elif cd.chart_type == ChartType.SCATTER:
            # У scatter цвет задаётся через маркер, а не fill серии
            for idx, series in enumerate(chart.series):
                color = palette_colors[idx % len(palette_colors)]
                marker = series.marker
                marker.format.fill.solid()
                marker.format.fill.fore_color.rgb = self._rgb(color)
                marker.format.line.color.rgb = self._rgb(color)
                marker.size = 10
        else:
            for idx, series in enumerate(chart.series):
                color = palette_colors[idx % len(palette_colors)]
                fill = series.format.fill
                fill.solid()
                fill.fore_color.rgb = self._rgb(color)

                # Для линий красим саму линию (а не только маркеры)
                if cd.chart_type == ChartType.LINE:
                    line = series.format.line
                    line.color.rgb = self._rgb(color)
                    line.width = Pt(2.5)

    def _color_pie_points(self, chart, palette_colors: list) -> None:
        """
        Pie-chart требует раскрашивать каждую data point отдельно.
        python-pptx не даёт прямого API для этого, используем lxml.
        """
        from copy import deepcopy
        from pptx.oxml.ns import qn

        # Расширенная палитра, чтобы хватило на много секторов
        extended = palette_colors + [
            self.style.palette.text_dark,
            self.style.palette.secondary,
            self.style.palette.accent,
        ]

        plot = chart.plots[0]
        series_xml = plot.series[0]._element  # CT_PieSer

        # Удаляем существующие dPt, если есть
        for dpt in series_xml.findall(qn("c:dPt")):
            series_xml.remove(dpt)

        # Найдём, после какого элемента вставлять dPt (по спеке — после tx, spPr, explosion)
        # Безопаснее всего — вставлять после последнего из этих элементов
        insert_after_tags = ["c:idx", "c:order", "c:tx", "c:spPr", "c:explosion"]
        insert_after_qnames = [qn(t) for t in insert_after_tags]

        # Найдём индекс точки, куда вставлять
        children = list(series_xml)
        insert_idx = 0
        for i, child in enumerate(children):
            if child.tag in insert_after_qnames:
                insert_idx = i + 1

        # Сколько секторов? Берём из cat
        cat = series_xml.find(".//" + qn("c:cat"))
        if cat is None:
            return
        pt_count_elem = cat.find(".//" + qn("c:ptCount"))
        n_points = int(pt_count_elem.get("val")) if pt_count_elem is not None else 0

        from lxml import etree
        nsmap = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                 "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

        for i in range(n_points):
            color = extended[i % len(extended)].lstrip("#")
            dpt_xml = f"""<c:dPt xmlns:c="{nsmap['c']}" xmlns:a="{nsmap['a']}">
                <c:idx val="{i}"/>
                <c:bubble3D val="0"/>
                <c:spPr>
                    <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
                    <a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>
                </c:spPr>
            </c:dPt>"""
            dpt = etree.fromstring(dpt_xml)
            series_xml.insert(insert_idx + i, dpt)

    # ----------------------------------------------------------------------- #
    # Утилиты стиля
    # ----------------------------------------------------------------------- #

    def _add_title(self, slide, text: str) -> None:
        """Стандартный заголовок контентного слайда."""
        title_box = slide.shapes.add_textbox(
            Inches(0.7), Inches(0.4), Inches(12.0), Inches(0.9)
        )
        self._set_text(
            title_box.text_frame, text,
            font=self.style.typography.heading_font,
            size=32, bold=True, color=self.style.palette.text_dark,
        )

    def _set_background(self, slide, hex_color: str) -> None:
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = self._rgb(hex_color)

    def _fill_solid(self, shape, hex_color: str) -> None:
        shape.fill.solid()
        shape.fill.fore_color.rgb = self._rgb(hex_color)

    @staticmethod
    def _rgb(hex_color: str) -> RGBColor:
        return RGBColor.from_string(hex_color.lstrip("#"))

    def _set_text(
        self, text_frame, text: str,
        font: str, size: int,
        color: str,
        bold: bool = False,
        align: PP_ALIGN = PP_ALIGN.LEFT,
        anchor: MSO_ANCHOR = MSO_ANCHOR.TOP,
        margin_left=None,
    ) -> None:
        """Единая точка форматирования текста."""
        text_frame.word_wrap = True
        text_frame.vertical_anchor = anchor
        if margin_left is not None:
            text_frame.margin_left = margin_left

        text_frame.text = text
        p = text_frame.paragraphs[0]
        p.alignment = align
        for run in p.runs:
            run.font.name = font
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.color.rgb = self._rgb(color)
