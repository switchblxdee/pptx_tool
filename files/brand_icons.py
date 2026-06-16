"""
brand_icons.py — векторные иконки из фирменного шаблона SberF1.

Иконки в шаблоне нарисованы кастомной геометрией (freeform-пути),
сгруппированы по одной на значок. Мы извлекли часть из них в
assets/icons_sberf1.json и умеем вставлять любую на слайд: задаём позицию,
размер и цвет (иконки перекрашиваются под палитру/контраст).

Формат assets/icons_sberf1.json:
    {
      "icons": { "i73": {"xml": "<p:grpSp ...>", "aspect": 1.0}, ... },
      "hints": { "warning": "i73", "clock": "i66", ... }
    }

Полный набор (540 иконок) можно достать скриптом extract_icons.py — он
рендерит пронумерованный «контактный лист», по которому выбираешь нужные id.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Optional

from lxml import etree

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _q(tag: str) -> str:
    pfx, name = tag.split(":")
    return f"{{{_A if pfx == 'a' else _P}}}{name}"


_LIB = None


def _load() -> dict:
    global _LIB
    if _LIB is None:
        p = Path(__file__).resolve().parent / "assets" / "icons_sberf1.json"
        try:
            _LIB = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            _LIB = {"icons": {}, "hints": {}}
    return _LIB


def has_brand_icon(hint: Optional[str]) -> bool:
    if not hint:
        return False
    return hint.lower().strip() in _load().get("hints", {})


def available_brand_hints() -> list[str]:
    return sorted(_load().get("hints", {}).keys())


def _recolor(group, hex_color: str) -> None:
    """Перекрашивает все заливки (и контуры) иконки в один цвет."""
    hexc = hex_color.lstrip("#").upper()
    for sf in group.iter(_q("a:solidFill")):
        for ch in list(sf):
            sf.remove(ch)
        etree.SubElement(sf, _q("a:srgbClr")).set("val", hexc)


def stamp_icon(slide, icon_id: str, left, top, size, hex_color: str) -> bool:
    """Вставляет иконку по её id (например 'i73') на слайд.

    left/top/size — в EMU. Иконка масштабируется в квадрат size×size
    (с учётом исходной пропорции) и перекрашивается в hex_color.
    """
    item = _load().get("icons", {}).get(icon_id)
    if not item:
        return False
    g = etree.fromstring(item["xml"])
    xfrm = g.find(_q("p:grpSpPr") + "/" + _q("a:xfrm"))
    if xfrm is None:
        return False
    off = xfrm.find(_q("a:off"))
    ext = xfrm.find(_q("a:ext"))
    off.set("x", str(int(left)))
    off.set("y", str(int(top)))
    ext.set("cx", str(int(size)))
    ext.set("cy", str(int(size * item.get("aspect", 1.0))))
    _recolor(g, hex_color)
    slide.shapes._spTree.append(g)
    return True


def draw_brand_icon(slide, hint: str, left, top, size, color: str) -> bool:
    """Рисует иконку по смысловому хинту (warning, clock, signal...).

    Возвращает True, если иконка найдена и вставлена; иначе False
    (вызывающий код может откатиться на встроенные иконки icons.py).
    """
    key = _load().get("hints", {}).get((hint or "").lower().strip())
    if not key:
        return False
    try:
        return stamp_icon(slide, key, left, top, size, color)
    except Exception:
        return False
