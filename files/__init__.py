"""
pptx_generator — LangChain-инструмент для генерации дайджеста .pptx через GigaChat.

Главный публичный API:
    GenerateDigestTool — LangChain Tool
    DigestBuilder      — низкоуровневая сборка из спеки
    ExcelReader        — чтение и профилирование xlsx
    DigestSpec         — Pydantic-схема дайджеста
"""
from .builder import DigestBuilder
from .excel_reader import DataContext, ExcelReader
from .json_repair import repair_json
from .schemas import (
    AttentionItem,
    AttentionSlide,
    ClosingSlide,
    ColorPalette,
    CoverSlide,
    DigestMeta,
    DigestSpec,
    DigestStyle,
    ExecutiveSummarySlide,
    KPICard,
    PatternItem,
    PatternsSlide,
    SummaryPoint,
    TopicItem,
    TopicSlide,
    Typography,
)
from .tool import (
    GenerateDigestInput,
    GenerateDigestTool,
    # backward-compat aliases
    GeneratePresentationInput,
    GeneratePresentationTool,
)

__all__ = [
    # Tool
    "GenerateDigestTool",
    "GenerateDigestInput",
    # Backward compatibility
    "GeneratePresentationTool",
    "GeneratePresentationInput",
    # Builder
    "DigestBuilder",
    # Reader
    "ExcelReader",
    "DataContext",
    # Schemas
    "DigestSpec",
    "DigestStyle",
    "DigestMeta",
    "ColorPalette",
    "Typography",
    "CoverSlide",
    "TopicSlide",
    "TopicItem",
    "KPICard",
    "ExecutiveSummarySlide",
    "SummaryPoint",
    "PatternsSlide",
    "PatternItem",
    "AttentionSlide",
    "AttentionItem",
    "ClosingSlide",
    # JSON repair (полезно для отладки)
    "repair_json",
]
__version__ = "0.2.0"
