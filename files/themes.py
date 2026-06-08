"""
Готовые палитры-пресеты и детекция темы из промпта пользователя.

Проблема: LLM плохо слушает указания про цвета («сделай тёмное оформление»)
и копирует палитру из few-shot примера. Решение: детектируем намерение
пользователя по ключевым словам и подставляем готовую палитру программно.

Так стиль из запроса применяется ГАРАНТИРОВАННО, а LLM отвечает только
за контент.

Каждый пресет — это полный набор из 8 цветов под схему ColorPalette.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional


@dataclass
class ThemePreset:
    """Готовая палитра (8 цветов) + ключевые слова для детекции."""
    name: str
    keywords: List[str]
    palette: Dict[str, str]


# --------------------------------------------------------------------------- #
# Пресеты палитр
# --------------------------------------------------------------------------- #

THEME_PRESETS: List[ThemePreset] = [
    ThemePreset(
        name="dark",
        keywords=[
            "тёмн", "темн", "dark", "чёрн", "черн", "ночн", "тёмное оформление",
            "dark mode", "тёмная тема", "темная тема", "графит",
        ],
        palette={
            "gradient_start": "1A1D29",   # почти чёрный сине-серый
            "gradient_end": "2D3142",     # тёмно-синий графит
            "card_bg": "252836",          # тёмная карточка
            "kpi_bg": "2F3343",           # чуть светлее для KPI
            "text_dark": "F5F6FA",        # СВЕТЛЫЙ текст (на тёмном фоне!)
            "text_muted": "9BA1B0",       # приглушённый светло-серый
            "accent": "5B8DEF",           # яркий синий акцент
            "badge": "FF6B6B",            # коралловый бейдж
        },
    ),
    ThemePreset(
        name="dark_emerald",
        keywords=["тёмно-зелён", "темно-зелен", "изумруд", "emerald dark", "тёмный зелёный"],
        palette={
            "gradient_start": "0F1B14",
            "gradient_end": "1A2E22",
            "card_bg": "16271D",
            "kpi_bg": "1E3329",
            "text_dark": "EAF5EE",
            "text_muted": "8FA89A",
            "accent": "2ECC71",
            "badge": "F39C12",
        },
    ),
    ThemePreset(
        name="navy_corporate",
        keywords=[
            "корпоратив", "corporate", "деловой", "официальн", "совет директоров",
            "navy", "тёмно-син", "темно-син", "премиум", "premium",
        ],
        palette={
            "gradient_start": "0B1F3A",
            "gradient_end": "1B3A5C",
            "card_bg": "F0F4F8",          # светлые карточки на тёмном фоне
            "kpi_bg": "FFFFFF",
            "text_dark": "0B1F3A",        # тёмный текст для светлых карточек
            "text_muted": "5A6B7D",
            "accent": "C9A227",           # золотой акцент
            "badge": "E07A5F",
        },
    ),
    ThemePreset(
        name="vibrant",
        keywords=[
            "яркий", "яркое", "энергичн", "vibrant", "сочн", "стартап",
            "startup", "молодёжн", "молодежн", "креатив",
        ],
        palette={
            "gradient_start": "6C5CE7",   # фиолетовый
            "gradient_end": "FD79A8",     # розовый
            "card_bg": "FFF0F6",
            "kpi_bg": "FFFFFF",
            "text_dark": "2D3436",
            "text_muted": "636E72",
            "accent": "00B894",           # бирюзовый акцент
            "badge": "FDCB6E",            # жёлтый
        },
    ),
    ThemePreset(
        name="ocean",
        keywords=["океан", "ocean", "морск", "синий", "blue", "голуб", "вода"],
        palette={
            "gradient_start": "2980B9",
            "gradient_end": "6DD5FA",
            "card_bg": "EAF6FB",
            "kpi_bg": "FFFFFF",
            "text_dark": "0A3D62",
            "text_muted": "5B7A8C",
            "accent": "0984E3",
            "badge": "FF7675",
        },
    ),
    ThemePreset(
        name="warm",
        keywords=["тёпл", "тепл", "warm", "оранж", "orange", "осен", "закат", "sunset"],
        palette={
            "gradient_start": "F8B500",
            "gradient_end": "FCEEB5",
            "card_bg": "FFF6E5",
            "kpi_bg": "FFFFFF",
            "text_dark": "5C3A21",
            "text_muted": "977C5E",
            "accent": "E17055",
            "badge": "D63031",
        },
    ),
    ThemePreset(
        name="minimal_mono",
        keywords=[
            "минимал", "minimal", "монохром", "mono", "ч/б", "чёрно-бел",
            "черно-бел", "строг", "сдержан", "лаконичн",
        ],
        palette={
            "gradient_start": "F5F5F5",
            "gradient_end": "E0E0E0",
            "card_bg": "FFFFFF",
            "kpi_bg": "FAFAFA",
            "text_dark": "1A1A1A",
            "text_muted": "757575",
            "accent": "424242",
            "badge": "E53935",
        },
    ),
    # Дефолт — бирюзово-персиковый, как на референсе «Голос IT»
    ThemePreset(
        name="default_teal_peach",
        keywords=["бирюзов", "персик", "пастель", "teal", "peach", "голос it"],
        palette={
            "gradient_start": "A8D8D5",
            "gradient_end": "F5D5BA",
            "card_bg": "FFE8EA",
            "kpi_bg": "FFFFFF",
            "text_dark": "1A1A2E",
            "text_muted": "6B7280",
            "accent": "2C5F5D",
            "badge": "F97316",
        },
    ),
]

DEFAULT_PRESET = THEME_PRESETS[-1]  # бирюзово-персиковый


def detect_theme(style_prompt: str) -> Optional[ThemePreset]:
    """
    Определяет пресет палитры по ключевым словам в промпте.

    Возвращает первый подходящий пресет или None, если ничего не совпало
    (тогда вызывающая сторона решает — дефолт или доверить LLM).

    Порядок важен: более специфичные пресеты (dark_emerald) идут раньше
    общих (dark), чтобы «тёмно-зелёный» не поймался как просто «тёмный».
    """
    if not style_prompt:
        return None

    text = style_prompt.lower()

    # Сначала специфичные многословные, потом общие — за счёт порядка в списке.
    # dark_emerald и navy_corporate идут до общего dark по длине ключей.
    # Сортируем кандидатов по длине самого длинного совпавшего ключа.
    best: Optional[ThemePreset] = None
    best_key_len = 0

    for preset in THEME_PRESETS:
        for kw in preset.keywords:
            if kw in text and len(kw) > best_key_len:
                best = preset
                best_key_len = len(kw)

    return best


def resolve_palette(style_prompt: str, fallback_to_default: bool = True) -> Optional[Dict[str, str]]:
    """
    Возвращает словарь палитры для заданного промпта.

    :param fallback_to_default: если True и ничего не найдено — вернуть дефолт.
        Если False — вернуть None (доверить палитру LLM).
    """
    preset = detect_theme(style_prompt)
    if preset:
        return dict(preset.palette)
    if fallback_to_default:
        return dict(DEFAULT_PRESET.palette)
    return None


def available_themes() -> List[str]:
    """Список имён доступных тем (для документации/отладки)."""
    return [p.name for p in THEME_PRESETS]
