"""
Font validation and replacement module

Provides font validation and replacement
"""
from typing import Optional, Dict
from .config import ConversionConfig, default_config

# draw.io's default font family
# In draw.io, Helvetica is used when font is not specified
# Falls back to Arial if Helvetica is not available
DRAWIO_DEFAULT_FONT_FAMILY = "Helvetica"


def validate_font(font_family: Optional[str]) -> bool:
    """
    Validate if font is available
    
    Args:
        font_family: Font family name
    
    Returns:
        True if available
    """
    # Simple implementation: always return True
    # In actual implementation, check system font list
    return True


def replace_font(font_family: Optional[str], config: Optional[ConversionConfig] = None) -> Optional[str]:
    """
    Replace font (based on configuration)
    
    Args:
        font_family: Original font family name
        config: ConversionConfig instance (uses default_config if None)
    
    Returns:
        Replaced font family name
    """
    if not font_family:
        return None
    
    # Get replacement map from configuration
    config_to_use = config or default_config
    replacements = config_to_use.font_replacements
    
    if font_family in replacements:
        return replacements[font_family]
    
    return font_family
