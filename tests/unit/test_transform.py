"""Test module for geometry transformations"""

import pytest
from drawio2pptx.geom.transform import apply_transform
from drawio2pptx.model.intermediate import Transform


def test_apply_transform_no_transform():
    """Test apply_transform with no transformation"""
    transform = Transform()
    x, y = apply_transform(10.0, 20.0, transform)
    assert x == 10.0
    assert y == 20.0


def test_apply_transform_translation():
    """Test apply_transform with translation only"""
    transform = Transform(translate_x=5.0, translate_y=10.0)
    x, y = apply_transform(10.0, 20.0, transform)
    assert x == 15.0
    assert y == 30.0


def test_apply_transform_scale():
    """Test apply_transform with scale only"""
    transform = Transform(scale_x=2.0, scale_y=3.0)
    x, y = apply_transform(10.0, 20.0, transform)
    assert x == 20.0
    assert y == 60.0


def test_apply_transform_rotation():
    """Test apply_transform with rotation only"""
    transform = Transform(rotation=90.0)
    x, y = apply_transform(10.0, 0.0, transform)
    # 90 degree rotation: (10, 0) -> (0, 10)
    assert abs(x - 0.0) < 0.0001
    assert abs(y - 10.0) < 0.0001


def test_apply_transform_flip_horizontal():
    """Test apply_transform with horizontal flip"""
    transform = Transform(flip_h=True)
    x, y = apply_transform(10.0, 20.0, transform)
    assert x == -10.0
    assert y == 20.0


def test_apply_transform_flip_vertical():
    """Test apply_transform with vertical flip"""
    transform = Transform(flip_v=True)
    x, y = apply_transform(10.0, 20.0, transform)
    assert x == 10.0
    assert y == -20.0


def test_apply_transform_with_origin():
    """Test apply_transform with custom origin"""
    transform = Transform(scale_x=2.0, scale_y=2.0)
    x, y = apply_transform(10.0, 20.0, transform, origin_x=5.0, origin_y=10.0)
    # Scale relative to origin: (10-5)*2+5 = 15, (20-10)*2+10 = 30
    assert x == 15.0
    assert y == 30.0


def test_apply_transform_combined():
    """Test apply_transform with combined transformations"""
    transform = Transform(
        translate_x=5.0,
        translate_y=10.0,
        scale_x=2.0,
        scale_y=2.0,
        rotation=180.0
    )
    x, y = apply_transform(10.0, 20.0, transform)
    # Complex transformation, just verify it doesn't crash
    assert isinstance(x, float)
    assert isinstance(y, float)
