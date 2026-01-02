"""
Style mapping module

Maps draw.io styles such as dash patterns, arrows, and corner radius to PowerPoint format
"""
from typing import Optional, Tuple
from ..config import DASH_PATTERN_MAP, ARROW_TYPE_MAP, rounded_to_arc_size


def map_dash_pattern(drawio_dash: Optional[str]) -> Optional[str]:
    """
    Map dash pattern
    
    Args:
        drawio_dash: draw.io dash pattern name
    
    Returns:
        DrawingML prstDash value, or None
    """
    if not drawio_dash:
        return None
    
    return DASH_PATTERN_MAP.get(drawio_dash.lower(), None)


def map_arrow_type(drawio_arrow: Optional[str]) -> Optional[Tuple[str, str, str]]:
    """
    Map arrow type
    
    Args:
        drawio_arrow: draw.io arrow type name
    
    Returns:
        (type, w, len) tuple, or None
    """
    if not drawio_arrow or drawio_arrow.lower() == "none":
        return None
    
    return ARROW_TYPE_MAP.get(drawio_arrow.lower())


def map_arrow_size_px_to_pptx(size_px: Optional[float]) -> Optional[str]:
    """
    Map draw.io marker size (startSize/endSize) in px to PowerPoint's discrete arrow size.

    PowerPoint supports only 'sm' / 'med' / 'lg' for line-end width/length.
    draw.io exposes a numeric marker size (often 6 by default). We approximate by thresholds.
    """
    if size_px is None:
        return None
    try:
        v = float(size_px)
    except Exception:
        return None
    if v <= 6.5:
        return "sm"
    if v <= 10.5:
        return "med"
    return "lg"


def map_arrow_type_with_size(drawio_arrow: Optional[str], size_px: Optional[float]) -> Optional[Tuple[str, str, str]]:
    """
    Map arrow type, optionally overriding (w, len) based on draw.io startSize/endSize.
    """
    base = map_arrow_type(drawio_arrow)
    if not base:
        return None
    override = map_arrow_size_px_to_pptx(size_px)
    if not override:
        return base
    arrow_type, _w, _len = base
    return (arrow_type, override, override)


def map_corner_radius(rounded: bool, width: float, height: float) -> Optional[int]:
    """
    Map corner radius
    
    Args:
        rounded: Whether corner radius is enabled
        width: Shape width (px)
        height: Shape height (px)
    
    Returns:
        arcSize value (0-100000), or None
    """
    return rounded_to_arc_size(rounded, width, height)
