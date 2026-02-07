"""Test module for unit conversions"""

import pytest
from pptx.util import Emu, Pt
from drawio2pptx.geom.units import (
    px_to_emu,
    px_to_pt,
    pt_to_emu,
    emu_to_px,
    emu_to_pt,
    pt_to_px,
    scale_font_size_for_pptx,
    EMU_PER_PX,
    PT_PER_PX,
    EMU_PER_PT,
    FONT_SIZE_SCALE,
)


def test_px_to_emu():
    """Test pixel to EMU conversion"""
    result = px_to_emu(100.0)
    assert isinstance(result, Emu)
    assert result == Emu(100 * EMU_PER_PX)


def test_px_to_pt():
    """Test pixel to point conversion"""
    result = px_to_pt(100.0)
    assert isinstance(result, Pt)
    assert abs(result.pt - (100.0 * PT_PER_PX)) < 0.001


def test_pt_to_emu():
    """Test point to EMU conversion"""
    result = pt_to_emu(12.0)
    assert isinstance(result, Emu)
    assert result == Emu(int(12.0 * EMU_PER_PT))


def test_emu_to_px_emu_object():
    """Test EMU to pixel conversion with Emu object"""
    emu = Emu(9525)  # 1 pixel
    result = emu_to_px(emu)
    assert abs(result - 1.0) < 0.001


def test_emu_to_px_int():
    """Test EMU to pixel conversion with int"""
    result = emu_to_px(9525)  # 1 pixel
    assert abs(result - 1.0) < 0.001


def test_emu_to_pt_emu_object():
    """Test EMU to point conversion with Emu object"""
    emu = Emu(12700)  # 1 point
    result = emu_to_pt(emu)
    assert abs(result - 1.0) < 0.001


def test_emu_to_pt_int():
    """Test EMU to point conversion with int"""
    result = emu_to_pt(12700)  # 1 point
    assert abs(result - 1.0) < 0.001


def test_pt_to_px():
    """Test point to pixel conversion"""
    result = pt_to_px(12.0)
    expected = 12.0 / PT_PER_PX
    assert abs(result - expected) < 0.001


def test_scale_font_size_for_pptx():
    """Test font size scaling for PowerPoint"""
    result = scale_font_size_for_pptx(12.0)
    expected = 12.0 * FONT_SIZE_SCALE
    assert abs(result - expected) < 0.001


def test_conversion_constants():
    """Test conversion constants"""
    assert EMU_PER_PX == 9525
    assert PT_PER_PX == 0.75
    assert EMU_PER_PT == 12700
    assert FONT_SIZE_SCALE == 0.75


def test_round_trip_conversions():
    """Test round-trip conversions"""
    # px -> emu -> px
    px_value = 100.0
    emu = px_to_emu(px_value)
    px_back = emu_to_px(emu)
    assert abs(px_back - px_value) < 0.1
    
    # pt -> emu -> pt
    pt_value = 12.0
    emu = pt_to_emu(pt_value)
    pt_back = emu_to_pt(emu)
    assert abs(pt_back - pt_value) < 0.001
    
    # px -> pt -> px
    px_value = 100.0
    pt = px_to_pt(px_value)
    px_back = pt_to_px(pt.pt)
    assert abs(px_back - px_value) < 0.1
