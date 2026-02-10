"""
Configuration and policy module

Manages settings such as approximation tables, scale policies, margins, DPI, font replacements for conversion
"""
from dataclasses import dataclass
from typing import Dict, Optional, Tuple


# EMU conversion constants (conversion from screen pixels to EMU)
EMU_PER_PX = 9525
PT_PER_PX = 0.75  # Assumption: for 96 DPI

# Default skew for parallelogram (MSO_SHAPE.PARALLELOGRAM) (approximately 0.0-0.5)
# Used when converting draw.io → PPTX to match shape appearance with connection point calculation.
PARALLELOGRAM_SKEW: float = 0.2

# Default inner padding (px) for swimlane/container shapes when draw.io does not specify spacing.
# Applied to header (title) area so box content is not flush to edges.
SWIMLANE_DEFAULT_PADDING_PX: float = 4.0


@dataclass
class ConversionConfig:
    """Conversion configuration"""
    # Scale policy: 'fit-to-width', 'fit-to-height', 'contain', 'cover', 'none'
    scale_policy: str = 'none'
    
    # Margin (px)
    margin_x: float = 0.0
    margin_y: float = 0.0
    
    # DPI setting
    dpi: float = 96.0
    
    # Font replacement map
    font_replacements: Dict[str, str] = None
    
    # Whether to warn about unsupported effects
    warn_unsupported_effects: bool = True
    
    # Tolerance for coordinate errors (px)
    coordinate_tolerance: float = 0.1
    
    def __post_init__(self):
        if self.font_replacements is None:
            self.font_replacements = {}


# Dash pattern mapping (draw.io → DrawingML prstDash)
DASH_PATTERN_MAP: Dict[str, str] = {
    'dashed': 'dash',
    'dash': 'dash',
    'dotted': 'dot',
    'dot': 'dot',
    'dashDot': 'dashDot',
    'dashdot': 'dashDot',
    'dashDotDot': 'lgDashDotDot',
    'dashdotdot': 'lgDashDotDot',
    'longDash': 'lgDash',
    'longdash': 'lgDash',
    'longDashDot': 'lgDashDot',
    'longdashdot': 'lgDashDot',
    'solid': 'solid',
    'none': 'solid',
}


# Arrow type mapping (draw.io → PowerPoint)
ARROW_TYPE_MAP: Dict[str, Tuple[str, str, str]] = {
    # (type, w, len) format. stealth = block arrow
    'classic': ('stealth', 'med', 'med'),
    'block': ('stealth', 'med', 'med'),
    'classicthin': ('stealth', 'sm', 'sm'),
    'blockthin': ('stealth', 'sm', 'sm'),
    # PowerPoint supports an "arrow" line-end type which corresponds to an open arrow head.
    'open': ('arrow', 'med', 'med'),
    'openthin': ('arrow', 'sm', 'sm'),
    'oval': ('oval', 'med', 'med'),
    'diamond': ('diamond', 'med', 'med'),
    'diamondthin': ('diamond', 'sm', 'sm'),
    'cross': ('stealth', 'med', 'med'),
    'crossthin': ('stealth', 'sm', 'sm'),
    'dash': ('triangle', 'sm', 'sm'),
    'dashthin': ('triangle', 'sm', 'sm'),
    'line': ('triangle', 'sm', 'sm'),
    'linethin': ('triangle', 'sm', 'sm'),
    'none': None,
}


# Corner radius approximation (rounded → arcSize)
# arcSize is a value from 0-100000, where 100000 is a complete circle
def rounded_to_arc_size(rounded: bool, width: float, height: float) -> Optional[int]:
    """
    Calculate arcSize from rounded flag and size
    
    Args:
        rounded: Whether corner radius is enabled
        width: Shape width (px)
        height: Shape height (px)
    
    Returns:
        arcSize value (0-100000), or None (no corner radius)
    """
    if not rounded:
        return None
    
    # Assume default corner radius (approximately 10% of size)
    radius = min(width, height) * 0.1
    
    # arcSize = (radius / min(width, height)) * 100000
    if min(width, height) > 0:
        arc_size = int((radius / min(width, height)) * 100000)
        return min(arc_size, 100000)
    
    return 50000  # Default value


# Default configuration instance
default_config = ConversionConfig()
