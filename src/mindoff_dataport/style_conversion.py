from typing import Optional

import xml.etree.ElementTree as ET

from openpyxl.styles import Color
from openpyxl.styles.borders import Side

from .schema import BorderSide

# §1. Constants & Exceptions

# Default Office theme base colors in openpyxl/Excel index order:
# lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink.
DEFAULT_THEME_COLORS: list[str] = [
    "FFFFFFFF",
    "FF000000",
    "FFEEECE1",
    "FF1F497D",
    "FF4F81BD",
    "FFC0504D",
    "FF9BBB59",
    "FF8064A2",
    "FF4BACC6",
    "FFF79646",
    "FF0000FF",
    "FF800080",
]
_THEME_ORDER = (
    "lt1",
    "dk1",
    "lt2",
    "dk2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
)
_DRAWINGML_NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

# openpyxl indexed color table (legacy palette)
_INDEXED_COLORS = [
    "FF000000",
    "FFFFFFFF",
    "FFFF0000",
    "FF00FF00",
    "FF0000FF",
    "FFFFFF00",
    "FFFF00FF",
    "FF00FFFF",
    "FF000000",
    "FFFFFFFF",
    "FFFF0000",
    "FF00FF00",
    "FF0000FF",
    "FFFFFF00",
    "FFFF00FF",
    "FF00FFFF",
    "FF800000",
    "FF008000",
    "FF000080",
    "FF808000",
    "FF800080",
    "FF008080",
    "FFC0C0C0",
    "FF808080",
    "FF9999FF",
    "FF993366",
    "FFFFFFCC",
    "FFCCFFFF",
    "FF660066",
    "FFFF8080",
    "FF0066CC",
    "FFCCCCFF",
    "FF000080",
    "FFFF00FF",
    "FFFFFF00",
    "FF00FFFF",
    "FF800080",
    "FF800000",
    "FF008080",
    "FF0000FF",
    "FF00CCFF",
    "FFCCFFFF",
    "FFCCFFCC",
    "FFFFFF99",
    "FF99CCFF",
    "FFFF99CC",
    "FFCC99FF",
    "FFFFCC99",
    "FF3366FF",
    "FF33CCCC",
    "FF99CC00",
    "FFFFCC00",
    "FFFF9900",
    "FFFF6600",
    "FF666699",
    "FF969696",
    "FF003366",
    "FF339966",
    "FF003300",
    "FF333300",
    "FF993300",
    "FF993366",
    "FF333399",
    "FF333333",
]

# §2. Classes and Sub Classes

# §3. Private Helper Functions


def normalize_color(color: Optional[Color]) -> Optional[str]:
    """Convert openpyxl Color to ARGB hex string, or None."""
    if color is None:
        return None

    color_type = getattr(color, "type", None)

    if color_type == "rgb":
        rgb = color.rgb
        if rgb in ("00000000", "000000"):
            return None
        return rgb if len(rgb) == 8 else f"FF{rgb}"

    if color_type == "theme":
        tint = getattr(color, "tint", 0.0)
        theme_idx = getattr(color, "theme", 0)
        return f"theme:{theme_idx}:{tint}"

    if color_type == "indexed":
        idx = getattr(color, "indexed", 0)
        if 0 <= idx < len(_INDEXED_COLORS):
            return _INDEXED_COLORS[idx]
        return None

    return None


def extract_theme_colors(theme_xml: bytes | str | None) -> list[str]:
    if not theme_xml:
        return list(DEFAULT_THEME_COLORS)
    try:
        root = ET.fromstring(theme_xml)
    except ET.ParseError:
        return list(DEFAULT_THEME_COLORS)

    color_scheme = root.find(".//a:clrScheme", _DRAWINGML_NS)
    if color_scheme is None:
        return list(DEFAULT_THEME_COLORS)

    result: list[str] = []
    for key in _THEME_ORDER:
        item = color_scheme.find(f"a:{key}", _DRAWINGML_NS)
        color = _theme_color_value(item)
        if color is None:
            return list(DEFAULT_THEME_COLORS)
        result.append(color)
    return result


def _theme_color_value(item: ET.Element | None) -> str | None:
    if item is None:
        return None
    srgb = item.find("a:srgbClr", _DRAWINGML_NS)
    if srgb is not None and srgb.get("val"):
        return _rgb_to_argb(srgb.get("val"))
    system = item.find("a:sysClr", _DRAWINGML_NS)
    if system is not None and system.get("lastClr"):
        return _rgb_to_argb(system.get("lastClr"))
    return None


def _rgb_to_argb(value: str | None) -> str | None:
    if not value:
        return None
    raw = value.strip().upper()
    if len(raw) == 8:
        return raw
    if len(raw) == 6:
        return f"FF{raw}"
    return None


def resolve_theme_color(
    value: str | None,
    theme_colors: list[str] | None = None,
) -> str | None:
    if not value:
        return None
    if not value.startswith("theme:"):
        return value

    parts = value.split(":")
    if len(parts) != 3:
        return None
    try:
        idx, tint = int(parts[1]), float(parts[2])
    except ValueError:
        return None

    palette = theme_colors or DEFAULT_THEME_COLORS
    if idx < 0 or idx >= len(palette):
        return None
    base = palette[idx]
    raw = base[-6:]
    if len(raw) != 6:
        return None
    try:
        rgb = tuple(int(raw[offset : offset + 2], 16) for offset in (0, 2, 4))
    except ValueError:
        return None
    r, g, b = _apply_tint(rgb, tint)
    return f"FF{r:02X}{g:02X}{b:02X}"


def _apply_tint(rgb: tuple[int, int, int], tint: float) -> tuple[int, int, int]:
    r, g, b = rgb
    if tint > 0:
        return (
            int(r + (255 - r) * tint),
            int(g + (255 - g) * tint),
            int(b + (255 - b) * tint),
        )
    if tint < 0:
        return (int(r * (1 + tint)), int(g * (1 + tint)), int(b * (1 + tint)))
    return r, g, b


def argb_to_color(argb: Optional[str]) -> Optional[Color]:
    """Convert ARGB hex string (or theme:index:tint) to openpyxl Color."""
    if argb is None:
        return None

    if argb.startswith("theme:"):
        parts = argb.split(":")
        theme_idx = int(parts[1])
        tint = float(parts[2])
        return Color(theme=theme_idx, tint=tint)

    return Color(rgb=argb)


def border_side_to_dict(side: Optional[Side]) -> BorderSide:
    """Convert openpyxl Side to a BorderSide dict."""
    if side is None:
        return {"style": None, "color": None}
    return {
        "style": side.border_style,
        "color": normalize_color(side.color),
    }


def dict_to_border_side(d: BorderSide) -> Side:
    """Convert BorderSide dict to openpyxl Side."""
    color = argb_to_color(d["color"])
    if color is not None:
        return Side(border_style=d["style"], color=color)
    return Side(border_style=d["style"])


# §4. Public Functions

# §5. Entrypoints
