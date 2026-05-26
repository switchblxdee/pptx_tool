"""
Чтение и предварительный анализ входного xlsx.

Ожидаем два листа:
- "Сырье" — табличные данные (любые: числа, текст, мешанина)
- "Суммаризация" — текст с описанием проблем

Модуль возвращает структурированный DataContext, который потом
передаётся в LLM в текстовом виде. LLM не видит сырой xlsx —
он видит уже подготовленный, очищенный, компактный summary.

Это критично: подавать сырые 10к строк в контекст бессмысленно
и дорого. Лучше дать LLM качественную сводку + сэмплы.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd


RAW_SHEET = "Сырье"
SUMMARY_SHEET = "Суммаризация"

# Эвристика: если колонка числовая и в ней разумное число уникальных значений —
# она годится для агрегаций/графиков.
MAX_SAMPLE_ROWS = 15           # сколько строк показать LLM как сэмпл
MAX_UNIQUE_FOR_CATEGORY = 50   # больше — уже не категориальная переменная


@dataclass
class ColumnProfile:
    """Профиль одной колонки — то, что увидит LLM."""
    name: str
    dtype: str
    non_null_count: int
    null_count: int
    unique_count: int
    sample_values: List[Any] = field(default_factory=list)
    # Только для числовых
    min_value: Optional[float] = None
    max_value: Optional[float] = None
    mean_value: Optional[float] = None
    # Топ-значения для категориальных
    top_values: Dict[str, int] = field(default_factory=dict)

    def to_prompt_text(self) -> str:
        lines = [f"  • {self.name} ({self.dtype})"]
        lines.append(f"      не пусто: {self.non_null_count}, пусто: {self.null_count}, уникальных: {self.unique_count}")

        if self.min_value is not None:
            lines.append(
                f"      min={self.min_value:.2f}, max={self.max_value:.2f}, mean={self.mean_value:.2f}"
            )
        if self.top_values:
            top_str = ", ".join(f"{k!r}={v}" for k, v in list(self.top_values.items())[:5])
            lines.append(f"      топ: {top_str}")
        if self.sample_values:
            sample_str = ", ".join(repr(v) for v in self.sample_values[:5])
            lines.append(f"      примеры: {sample_str}")
        return "\n".join(lines)


@dataclass
class DataContext:
    """Полный контекст данных для LLM."""
    source_path: str
    row_count: int
    column_count: int
    columns: List[ColumnProfile]
    sample_rows: List[Dict[str, Any]]
    summary_text: str

    def to_prompt_text(self) -> str:
        """Сериализация в компактный текст для системного промпта."""
        parts = [
            f"=== ИСТОЧНИК: {self.source_path} ===",
            f"Строк: {self.row_count}, колонок: {self.column_count}",
            "",
            "=== СТРУКТУРА ДАННЫХ (лист 'Сырье') ===",
        ]
        for col in self.columns:
            parts.append(col.to_prompt_text())

        parts.append("")
        parts.append("=== СЭМПЛ ДАННЫХ (первые строки) ===")
        for i, row in enumerate(self.sample_rows[:MAX_SAMPLE_ROWS], 1):
            row_repr = ", ".join(f"{k}={v!r}" for k, v in row.items() if v is not None)
            parts.append(f"  [{i}] {row_repr}")

        parts.append("")
        parts.append("=== СУММАРИЗАЦИЯ ПРОБЛЕМ (лист 'Суммаризация') ===")
        parts.append(self.summary_text.strip())
        return "\n".join(parts)


# --------------------------------------------------------------------------- #
# Чтение и профилирование
# --------------------------------------------------------------------------- #

class ExcelReader:
    """Читает xlsx и собирает DataContext."""

    def __init__(self, path: str | Path):
        self.path = Path(path)
        if not self.path.exists():
            raise FileNotFoundError(f"Файл не найден: {self.path}")
        if self.path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError(f"Ожидается xlsx-файл, получено: {self.path.suffix}")

    def read(self) -> DataContext:
        """Главная точка входа."""
        xls = pd.ExcelFile(self.path)
        available = set(xls.sheet_names)

        missing = {RAW_SHEET, SUMMARY_SHEET} - available
        if missing:
            raise ValueError(
                f"В файле {self.path.name} отсутствуют листы: {missing}. "
                f"Найдены: {sorted(available)}"
            )

        raw_df = pd.read_excel(xls, sheet_name=RAW_SHEET)
        summary_df = pd.read_excel(xls, sheet_name=SUMMARY_SHEET, header=None)

        return DataContext(
            source_path=str(self.path),
            row_count=len(raw_df),
            column_count=len(raw_df.columns),
            columns=[self._profile_column(raw_df[c]) for c in raw_df.columns],
            sample_rows=self._extract_samples(raw_df),
            summary_text=self._extract_summary(summary_df),
        )

    # --- private ------------------------------------------------------------

    @staticmethod
    def _profile_column(series: pd.Series) -> ColumnProfile:
        """Собирает профиль колонки в зависимости от её типа."""
        non_null = series.dropna()
        profile = ColumnProfile(
            name=str(series.name),
            dtype=str(series.dtype),
            non_null_count=int(non_null.count()),
            null_count=int(series.isna().sum()),
            unique_count=int(non_null.nunique()),
        )

        if pd.api.types.is_numeric_dtype(series) and not non_null.empty:
            profile.min_value = float(non_null.min())
            profile.max_value = float(non_null.max())
            profile.mean_value = float(non_null.mean())

        # Сэмпл уникальных значений
        unique_vals = non_null.unique()[:5]
        profile.sample_values = [_to_jsonable(v) for v in unique_vals]

        # Топ-значения для категориальных
        if profile.unique_count <= MAX_UNIQUE_FOR_CATEGORY and profile.unique_count > 0:
            counts = non_null.value_counts().head(7)
            profile.top_values = {str(k): int(v) for k, v in counts.items()}

        return profile

    @staticmethod
    def _extract_samples(df: pd.DataFrame) -> List[Dict[str, Any]]:
        """Берёт первые строки и приводит значения к JSON-совместимым."""
        sample = df.head(MAX_SAMPLE_ROWS).fillna("")
        return [
            {str(k): _to_jsonable(v) for k, v in row.items()}
            for row in sample.to_dict(orient="records")
        ]

    @staticmethod
    def _extract_summary(df: pd.DataFrame) -> str:
        """
        Лист 'Суммаризация' может быть в любом формате: одна ячейка,
        несколько строк, табличка. Склеиваем всё непустое.
        """
        chunks: List[str] = []
        for _, row in df.iterrows():
            for cell in row:
                if pd.notna(cell):
                    text = str(cell).strip()
                    if text:
                        chunks.append(text)
        return "\n".join(chunks)


def _to_jsonable(value: Any) -> Any:
    """Преобразует numpy/pandas-типы в нативные Python."""
    if pd.isna(value):
        return None
    if hasattr(value, "item"):  # numpy scalars
        return value.item()
    if isinstance(value, (pd.Timestamp,)):
        return value.isoformat()
    return value
