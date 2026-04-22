from typing import Optional

from openpyxl.styles import Color
from openpyxl.styles.borders import Side

from .schema import BorderSide

# §1 Constants & Exceptions

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

# §2 Classes and Sub Classes

# §3 Private Helper Functions


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


# §4 Public Functions

# §5 Entrypoints
