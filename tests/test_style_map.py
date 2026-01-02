"""Test module for style mapping"""

import pytest
from drawio2pptx.mapping.style_map import (
    map_dash_pattern,
    map_arrow_type,
    map_arrow_size_px_to_pptx,
    map_arrow_type_with_size,
    map_corner_radius,
)


def test_map_dash_pattern_none():
    """Test map_dash_pattern with None"""
    result = map_dash_pattern(None)
    assert result is None


def test_map_dash_pattern_valid():
    """Test map_dash_pattern with valid patterns"""
    assert map_dash_pattern("dashed") == "dash"
    assert map_dash_pattern("dotted") == "dot"
    assert map_dash_pattern("dashDot") == "dashDot"
    assert map_dash_pattern("solid") == "solid"


def test_map_dash_pattern_case_insensitive():
    """Test map_dash_pattern is case insensitive"""
    assert map_dash_pattern("DASHED") == "dash"
    assert map_dash_pattern("Dotted") == "dot"


def test_map_dash_pattern_invalid():
    """Test map_dash_pattern with invalid pattern"""
    result = map_dash_pattern("invalid_pattern")
    assert result is None


def test_map_arrow_type_none():
    """Test map_arrow_type with None"""
    result = map_arrow_type(None)
    assert result is None


def test_map_arrow_type_none_string():
    """Test map_arrow_type with 'none' string"""
    result = map_arrow_type("none")
    assert result is None


def test_map_arrow_type_valid():
    """Test map_arrow_type with valid arrow types"""
    result = map_arrow_type("classic")
    assert result == ("triangle", "med", "med")
    
    result = map_arrow_type("oval")
    assert result == ("oval", "med", "med")
    
    result = map_arrow_type("diamond")
    assert result == ("diamond", "med", "med")


def test_map_arrow_type_case_insensitive():
    """Test map_arrow_type is case insensitive"""
    result = map_arrow_type("CLASSIC")
    assert result == ("triangle", "med", "med")


def test_map_arrow_type_invalid():
    """Test map_arrow_type with invalid arrow type"""
    result = map_arrow_type("invalid_arrow")
    assert result is None


def test_map_arrow_size_px_to_pptx_none():
    """Test map_arrow_size_px_to_pptx with None"""
    result = map_arrow_size_px_to_pptx(None)
    assert result is None


def test_map_arrow_size_px_to_pptx_small():
    """Test map_arrow_size_px_to_pptx with small size"""
    result = map_arrow_size_px_to_pptx(6.0)
    assert result == "sm"
    
    result = map_arrow_size_px_to_pptx(6.5)
    assert result == "sm"


def test_map_arrow_size_px_to_pptx_medium():
    """Test map_arrow_size_px_to_pptx with medium size"""
    result = map_arrow_size_px_to_pptx(7.0)
    assert result == "med"
    
    result = map_arrow_size_px_to_pptx(10.5)
    assert result == "med"


def test_map_arrow_size_px_to_pptx_large():
    """Test map_arrow_size_px_to_pptx with large size"""
    result = map_arrow_size_px_to_pptx(11.0)
    assert result == "lg"
    
    result = map_arrow_size_px_to_pptx(20.0)
    assert result == "lg"


def test_map_arrow_size_px_to_pptx_invalid():
    """Test map_arrow_size_px_to_pptx with invalid input"""
    result = map_arrow_size_px_to_pptx("invalid")
    assert result is None


def test_map_arrow_type_with_size_none_arrow():
    """Test map_arrow_type_with_size with None arrow"""
    result = map_arrow_type_with_size(None, 6.0)
    assert result is None


def test_map_arrow_type_with_size_no_size():
    """Test map_arrow_type_with_size without size override"""
    result = map_arrow_type_with_size("classic", None)
    assert result == ("triangle", "med", "med")


def test_map_arrow_type_with_size_with_override():
    """Test map_arrow_type_with_size with size override"""
    result = map_arrow_type_with_size("classic", 6.0)
    assert result == ("triangle", "sm", "sm")
    
    result = map_arrow_type_with_size("oval", 8.0)
    assert result == ("oval", "med", "med")
    
    result = map_arrow_type_with_size("diamond", 15.0)
    assert result == ("diamond", "lg", "lg")


def test_map_corner_radius_not_rounded():
    """Test map_corner_radius when not rounded"""
    result = map_corner_radius(False, 100.0, 50.0)
    assert result is None


def test_map_corner_radius_rounded():
    """Test map_corner_radius when rounded"""
    result = map_corner_radius(True, 100.0, 50.0)
    assert result is not None
    assert 0 <= result <= 100000

