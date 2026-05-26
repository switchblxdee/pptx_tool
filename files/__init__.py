"""
pptx_generator — LangChain-инструмент для генерации .pptx через GigaChat.

Главный публичный API:
    GeneratePresentationTool — LangChain Tool
    PresentationBuilder      — низкоуровневая сборка из спеки
    ExcelReader              — чтение и профилирование xlsx
    PresentationSpec         — Pydantic-схема презентации
"""
from .builder import PresentationBuilder
from .excel_reader import DataContext, ExcelReader
from .schemas import (
    BulletsSlide,
    ChartData,
    ChartSeries,
    ChartSlide,
    ChartType,
    ChartWithTextSlide,
    ClosingSlide,
    ColorPalette,
    KPIItem,
    KPISlide,
    PresentationSpec,
    PresentationStyle,
    QuoteSlide,
    SectionHeaderSlide,
    SlideLayout,
    TitleSlide,
    TwoColumnSlide,
    Typography,
)
from .tool import GeneratePresentationInput, GeneratePresentationTool

__all__ = [
    # Tool
    "GeneratePresentationTool",
    "GeneratePresentationInput",
    # Builder
    "PresentationBuilder",
    # Reader
    "ExcelReader",
    "DataContext",
    # Schemas
    "PresentationSpec",
    "PresentationStyle",
    "ColorPalette",
    "Typography",
    "ChartType",
    "ChartData",
    "ChartSeries",
    "SlideLayout",
    "TitleSlide",
    "SectionHeaderSlide",
    "BulletsSlide",
    "TwoColumnSlide",
    "ChartSlide",
    "ChartWithTextSlide",
    "KPISlide",
    "KPIItem",
    "QuoteSlide",
    "ClosingSlide",
]
__version__ = "0.1.0"
