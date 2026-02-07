"""Test module for font validation and replacement"""

import pytest
from drawio2pptx.fonts import validate_font, replace_font, DRAWIO_DEFAULT_FONT_FAMILY
from drawio2pptx.config import ConversionConfig, default_config


def test_validate_font():
    """Test font validation"""
    # Simple implementation always returns True
    assert validate_font("Arial") is True
    assert validate_font("Helvetica") is True
    assert validate_font(None) is True
    assert validate_font("") is True


def test_replace_font_none():
    """Test font replacement with None"""
    result = replace_font(None)
    assert result is None


def test_replace_font_no_replacement():
    """Test font replacement when no replacement is configured"""
    # Clear any existing replacements
    original_replacements = default_config.font_replacements.copy()
    default_config.font_replacements.clear()
    
    try:
        result = replace_font("Arial")
        assert result == "Arial"
    finally:
        default_config.font_replacements = original_replacements


def test_replace_font_with_replacement():
    """Test font replacement when replacement is configured"""
    # Set up replacement
    original_replacements = default_config.font_replacements.copy()
    default_config.font_replacements = {'Arial': 'Helvetica'}
    
    try:
        result = replace_font("Arial")
        assert result == "Helvetica"
        
        # Font not in replacement map should return as-is
        result2 = replace_font("Times New Roman")
        assert result2 == "Times New Roman"
    finally:
        default_config.font_replacements = original_replacements


def test_drawio_default_font_family():
    """Test DRAWIO_DEFAULT_FONT_FAMILY constant"""
    assert DRAWIO_DEFAULT_FONT_FAMILY == "Helvetica"


def test_replace_font_with_custom_config():
    """Test font replacement with custom config"""
    custom_config = ConversionConfig(font_replacements={'Arial': 'Times New Roman'})
    
    # Should use custom config
    result = replace_font("Arial", config=custom_config)
    assert result == "Times New Roman"
    
    # Should not affect default config
    result2 = replace_font("Arial")
    assert result2 == "Arial"  # default_config has no replacement for Arial
