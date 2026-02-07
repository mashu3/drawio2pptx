"""Test for boundary calculation of parallelogram connection points (exitX/exitY, entryX/entryY)"""

import pytest

from drawio2pptx.io.drawio_loader import DrawIOLoader
from drawio2pptx.model.intermediate import ShapeElement
from drawio2pptx.config import PARALLELOGRAM_SKEW


def _approx(a: float, b: float, tol: float = 1e-6) -> bool:
    return abs(a - b) <= tol


def test_parallelogram_right_edge_midpoint_is_on_edge():
    loader = DrawIOLoader()
    shape = ShapeElement(id="s1", x=100.0, y=200.0, w=120.0, h=50.0, shape_type="parallelogram")

    x, y = loader._calculate_boundary_point(shape, rel_x=1.0, rel_y=0.5, offset_x=0.0, offset_y=0.0)

    offset = float(PARALLELOGRAM_SKEW) * shape.h
    expected_x = shape.x + shape.w - offset / 2.0
    expected_y = shape.y + shape.h / 2.0

    assert _approx(x, expected_x)
    assert _approx(y, expected_y)


def test_parallelogram_left_edge_midpoint_is_on_edge():
    loader = DrawIOLoader()
    shape = ShapeElement(id="s1", x=10.0, y=20.0, w=200.0, h=80.0, shape_type="parallelogram")

    x, y = loader._calculate_boundary_point(shape, rel_x=0.0, rel_y=0.5, offset_x=0.0, offset_y=0.0)

    offset = float(PARALLELOGRAM_SKEW) * shape.h
    expected_x = shape.x + offset / 2.0
    expected_y = shape.y + shape.h / 2.0

    assert _approx(x, expected_x)
    assert _approx(y, expected_y)


def test_parallelogram_corners_match_rel_coordinates():
    loader = DrawIOLoader()
    shape = ShapeElement(id="s1", x=0.0, y=0.0, w=100.0, h=40.0, shape_type="parallelogram")

    offset = float(PARALLELOGRAM_SKEW) * shape.h

    # top-left (rel_y=0 on left edge)
    x, y = loader._calculate_boundary_point(shape, rel_x=0.0, rel_y=0.0, offset_x=0.0, offset_y=0.0)
    assert _approx(x, shape.x + offset)
    assert _approx(y, shape.y)

    # bottom-left
    x, y = loader._calculate_boundary_point(shape, rel_x=0.0, rel_y=1.0, offset_x=0.0, offset_y=0.0)
    assert _approx(x, shape.x)
    assert _approx(y, shape.y + shape.h)

    # top-right
    x, y = loader._calculate_boundary_point(shape, rel_x=1.0, rel_y=0.0, offset_x=0.0, offset_y=0.0)
    assert _approx(x, shape.x + shape.w)
    assert _approx(y, shape.y)

    # bottom-right
    x, y = loader._calculate_boundary_point(shape, rel_x=1.0, rel_y=1.0, offset_x=0.0, offset_y=0.0)
    assert _approx(x, shape.x + shape.w - offset)
    assert _approx(y, shape.y + shape.h)
