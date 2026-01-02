"""Test module for configuration"""

import pytest
from drawio2pptx.config import (
    ConversionConfig,
    rounded_to_arc_size,
    default_config,
    DASH_PATTERN_MAP,
    ARROW_TYPE_MAP,
)


def test_conversion_config_default():
    """Test default ConversionConfig"""
    config = ConversionConfig()
    assert config.scale_policy == 'none'
    assert config.margin_x == 0.0
    assert config.margin_y == 0.0
    assert config.dpi == 96.0
    assert config.font_replacements == {}
    assert config.warn_unsupported_effects is True
    assert config.coordinate_tolerance == 0.1


def test_conversion_config_custom():
    """Test custom ConversionConfig"""
    config = ConversionConfig(
        scale_policy='fit-to-width',
        margin_x=10.0,
        margin_y=20.0,
        dpi=120.0,
        font_replacements={'Arial': 'Helvetica'},
        warn_unsupported_effects=False,
        coordinate_tolerance=0.5
    )
    assert config.scale_policy == 'fit-to-width'
    assert config.margin_x == 10.0
    assert config.margin_y == 20.0
    assert config.dpi == 120.0
    assert config.font_replacements == {'Arial': 'Helvetica'}
    assert config.warn_unsupported_effects is False
    assert config.coordinate_tolerance == 0.5


def test_rounded_to_arc_size_not_rounded():
    """Test rounded_to_arc_size when rounded is False"""
    result = rounded_to_arc_size(False, 100.0, 50.0)
    assert result is None


def test_rounded_to_arc_size_rounded():
    """Test rounded_to_arc_size when rounded is True"""
    result = rounded_to_arc_size(True, 100.0, 50.0)
    assert result is not None
    assert 0 <= result <= 100000
    # Should be approximately 10% of min dimension
    expected = int((5.0 / 50.0) * 100000)  # radius = 5.0 (10% of 50.0)
    assert result == expected


def test_rounded_to_arc_size_zero_size():
    """Test rounded_to_arc_size with zero size"""
    result = rounded_to_arc_size(True, 0.0, 0.0)
    assert result == 50000  # Default value


def test_rounded_to_arc_size_square():
    """Test rounded_to_arc_size with square shape"""
    result = rounded_to_arc_size(True, 100.0, 100.0)
    assert result is not None
    assert 0 <= result <= 100000
    # Should be approximately 10% of size
    expected = int((10.0 / 100.0) * 100000)
    assert result == expected


def test_dash_pattern_map():
    """Test DASH_PATTERN_MAP contains expected mappings"""
    assert DASH_PATTERN_MAP['dashed'] == 'dash'
    assert DASH_PATTERN_MAP['dotted'] == 'dot'
    assert DASH_PATTERN_MAP['dashDot'] == 'dashDot'
    assert DASH_PATTERN_MAP['solid'] == 'solid'


def test_arrow_type_map():
    """Test ARROW_TYPE_MAP contains expected mappings"""
    assert ARROW_TYPE_MAP['classic'] == ('triangle', 'med', 'med')
    assert ARROW_TYPE_MAP['oval'] == ('oval', 'med', 'med')
    assert ARROW_TYPE_MAP['none'] is None


def test_default_config():
    """Test default_config instance"""
    assert isinstance(default_config, ConversionConfig)
    assert default_config.scale_policy == 'none'

