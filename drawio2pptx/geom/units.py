"""
Unit conversion module

Provides conversions between px/pt/EMU with precision management
"""
from pptx.util import Emu, Pt
from typing import Union


# Conversion constants
EMU_PER_PX = 9525  # Conversion from screen pixels to EMU (assuming 96 DPI)
PT_PER_PX = 0.75   # For 96 DPI: 1px = 0.75pt
EMU_PER_PT = 12700  # 1pt = 12700 EMU

# Font size scale factor
# Scale for converting draw.io font size (points) to PowerPoint font size
# Considering coordinate transformation, font size also needs to be scaled
# Use same value as PT_PER_PX (for 96 DPI: 1px = 0.75pt)
FONT_SIZE_SCALE = PT_PER_PX  # 0.75


def px_to_emu(px: float) -> Emu:
    """
    Convert pixels to EMU
    
    Args:
        px: Pixel value
    
    Returns:
        EMU object
    """
    return Emu(int(px * EMU_PER_PX))


def px_to_pt(px: float) -> Pt:
    """
    Convert pixels to points
    
    Args:
        px: Pixel value
    
    Returns:
        Pt object
    """
    return Pt(px * PT_PER_PX)


def pt_to_emu(pt: float) -> Emu:
    """
    Convert points to EMU
    
    Args:
        pt: Point value
    
    Returns:
        EMU object
    """
    return Emu(int(pt * EMU_PER_PT))


def emu_to_px(emu: Union[Emu, int]) -> float:
    """
    Convert EMU to pixels
    
    Args:
        emu: EMU value (Emu object or int)
    
    Returns:
        Pixel value (float)
    """
    if isinstance(emu, Emu):
        emu_value = emu
    else:
        emu_value = int(emu)
    return emu_value / EMU_PER_PX


def emu_to_pt(emu: Union[Emu, int]) -> float:
    """
    Convert EMU to points
    
    Args:
        emu: EMU value (Emu object or int)
    
    Returns:
        Point value (float)
    """
    if isinstance(emu, Emu):
        emu_value = emu
    else:
        emu_value = int(emu)
    return emu_value / EMU_PER_PT


def pt_to_px(pt: float) -> float:
    """
    Convert points to pixels
    
    Args:
        pt: Point value
    
    Returns:
        Pixel value (float)
    """
    return pt / PT_PER_PX


def scale_font_size_for_pptx(font_size_drawio: float) -> float:
    """
    Scale convert draw.io font size for PowerPoint
    
    Since draw.io and PowerPoint have different coordinate systems, font size also needs to be scaled.
    Considering coordinate transformation, font size also needs to be scaled similarly.
    
    Args:
        font_size_drawio: draw.io font size (draw.io units; often behaves like px)
    
    Returns:
        PowerPoint font size (points)
    """
    return font_size_drawio * FONT_SIZE_SCALE
