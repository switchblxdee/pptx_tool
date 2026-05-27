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
from .schemas import (
    ColorPalette,
    CoverSlide,
    DigestMeta,
    DigestSpec,
    DigestStyle,
    KPICard,
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
]
__version__ = "0.2.0"
