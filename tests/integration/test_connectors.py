"""
Integration tests: connectors (arrows, routing, z-order, markers).
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

import pytest

from drawio2pptx.io.drawio_loader import DrawIOLoader
from drawio2pptx.io.pptx_writer import PPTXWriter
from drawio2pptx.model.intermediate import ConnectorElement, ShapeElement


def _approx(a: float, b: float, tol: float = 1e-6) -> bool:
    return abs(a - b) <= tol


# ---- Orthogonal connector endpoints (ellipse ↔ parallelogram) ----
def test_connector_ellipse_side_goes_down_from_ellipse(sample_dir: Path):
    """
    In sample.drawio, the orthogonal connector between parallelogram and ellipse
    attaches to the bottom of the ellipse (segment adjacent to ellipse is vertical).
    """
    sample_path = sample_dir / "sample.drawio"
    loader = DrawIOLoader()
    diagrams = loader.load_file(sample_path)
    elements = loader.extract_elements(diagrams[0])

    shapes = {e.id: e for e in elements if isinstance(e, ShapeElement)}
    connectors = [e for e in elements if isinstance(e, ConnectorElement)]

    edge_id = "GStdcLXKth4fSFfuQepI-11"
    conn = next(c for c in connectors if c.id == edge_id)

    target = shapes[conn.target_id]
    assert "ellipse" in (target.shape_type or "").lower()

    end_x, end_y = conn.points[-1]
    assert _approx(end_y, target.y + target.h)

    prev_x, prev_y = conn.points[-2]
    assert _approx(prev_x, end_x)
    assert prev_y > end_y


# ---- Orthogonal connector (smiley → hexagon) ----
def test_connector_smiley_to_hexagon_respects_entry_side_left(sample_dir: Path):
    """
    In sample.drawio, the smiley→hexagon connector has entryX=0,entryY=0.5,
    so the segment adjacent to the hexagon is horizontal (approaching from the left).
    """
    sample_path = sample_dir / "sample.drawio"
    loader = DrawIOLoader()
    diagrams = loader.load_file(sample_path)
    elements = loader.extract_elements(diagrams[0])

    shapes = {e.id: e for e in elements if isinstance(e, ShapeElement)}
    connectors = [e for e in elements if isinstance(e, ConnectorElement)]

    edge_id = "fYnb-Lad83hC8_SQXFiI-27"
    conn = next(c for c in connectors if c.id == edge_id)

    target = shapes[conn.target_id]
    assert "hexagon" in (target.shape_type or "").lower()
    assert conn.edge_style == "orthogonal"
    assert len(conn.points) >= 3

    end_x, end_y = conn.points[-1]
    prev_x, prev_y = conn.points[-2]
    assert _approx(prev_y, end_y)
    assert prev_x < end_x


# ---- Default end arrow (when omitted) ----
def test_loader_applies_default_end_arrow_when_omitted(sample_dir: Path):
    """When draw.io omits endArrow, the loader applies classic as default."""
    sample_path = sample_dir / "sample.drawio"
    loader = DrawIOLoader()
    diagrams = loader.load_file(sample_path)
    elements = loader.extract_elements(diagrams[0])

    edge_id = "GStdcLXKth4fSFfuQepI-11"
    conn = next(e for e in elements if isinstance(e, ConnectorElement) and e.id == edge_id)
    assert (conn.style.arrow_end or "").lower() == "classic"


def test_generated_pptx_contains_triangle_arrow_type_for_default_end_arrow(
    sample_dir: Path, tmp_path: Path
):
    """Default end arrow (classic) is emitted as triangle in the generated slide XML."""
    sample_path = sample_dir / "sample.drawio"
    out_path = tmp_path / "sample_default_arrow.pptx"

    loader = DrawIOLoader()
    diagrams = loader.load_file(sample_path)
    page_size = loader.extract_page_size(diagrams[0])
    writer = PPTXWriter()
    prs, blank_layout = writer.create_presentation(page_size)
    elements = loader.extract_elements(diagrams[0])
    writer.add_slide(prs, blank_layout, elements)
    prs.save(out_path)

    with zipfile.ZipFile(out_path) as z:
        slide_xml = z.read("ppt/slides/slide1.xml").decode("utf-8", errors="ignore")

    assert 'type="triangle"' in slide_xml
    assert '<a:tailEnd' in slide_xml and 'type="triangle"' in slide_xml


# ---- Arrow size from endSize ----
def test_generated_pptx_respects_endSize_for_filled_oval_arrow(
    sample_dir: Path, tmp_path: Path
):
    """sample.drawio endArrow=oval; endFill=1; endSize=6 → PPTX w/len=\"sm\"."""
    sample_path = sample_dir / "sample.drawio"
    out_path = tmp_path / "sample_endSize_oval.pptx"

    loader = DrawIOLoader()
    diagrams = loader.load_file(sample_path)
    page_size = loader.extract_page_size(diagrams[0])
    writer = PPTXWriter()
    prs, blank_layout = writer.create_presentation(page_size)
    elements = loader.extract_elements(diagrams[0])
    writer.add_slide(prs, blank_layout, elements)
    prs.save(out_path)

    with zipfile.ZipFile(out_path) as z:
        slide_xml = z.read("ppt/slides/slide1.xml").decode("utf-8", errors="ignore")

    m = re.search(r'<a:tailEnd[^>]*type="oval"[^>]*/>', slide_xml)
    assert m, "Expected an a:tailEnd element with type='oval' in slide1.xml"
    frag = m.group(0)
    assert 'w="sm"' in frag
    assert 'len="sm"' in frag


# ---- Open oval marker (startFill=0) ----
def test_generated_pptx_emulates_open_oval_marker_when_startFill_is_zero(
    sample_dir: Path, tmp_path: Path
):
    """Open circle (startArrow=oval, startFill=0) is emulated as oval outline in PPTX."""
    sample_path = sample_dir / "sample.drawio"
    out_path = tmp_path / "sample_open_oval_marker.pptx"

    loader = DrawIOLoader()
    diagrams = loader.load_file(sample_path)
    page_size = loader.extract_page_size(diagrams[0])
    writer = PPTXWriter()
    prs, blank_layout = writer.create_presentation(page_size)
    elements = loader.extract_elements(diagrams[0])
    writer.add_slide(prs, blank_layout, elements)
    prs.save(out_path)

    with zipfile.ZipFile(out_path) as z:
        slide_xml = z.read("ppt/slides/slide1.xml").decode("utf-8", errors="ignore")

    assert "drawio2pptx:marker:open-oval:GStdcLXKth4fSFfuQepI-8:start" in slide_xml
    assert slide_xml.count('type="oval"') == 1

    marker_name = "drawio2pptx:marker:open-oval:GStdcLXKth4fSFfuQepI-8:start"
    assert re.search(
        r'name="' + re.escape(marker_name) + r'".{0,2000}?<a:noFill/>',
        slide_xml,
    ), "Expected the open-oval marker shape to have <a:noFill/>"

    m = re.search(
        r'name="' + re.escape(marker_name) + r'"[\s\S]*?<a:ext cx="(\d+)" cy="(\d+)"',
        slide_xml,
    )
    assert m, "Expected to find open-oval marker shape extents in slide XML"
    cx, cy = int(m.group(1)), int(m.group(2))
    assert abs(cx - 69056) <= 20
    assert abs(cy - 69056) <= 20


# ---- z-order: connectors behind node shapes ----
def test_connector_is_behind_target_shape_in_sample_drawio(
    sample_dir: Path, tmp_path: Path
):
    """In sample.drawio, the target shape of edge ...-7 is drawn in front of the connector."""
    sample_path = sample_dir / "sample.drawio"
    out_path = tmp_path / "sample_zorder.pptx"

    loader = DrawIOLoader()
    diagrams = loader.load_file(sample_path)
    page_size = loader.extract_page_size(diagrams[0])
    writer = PPTXWriter()
    prs, blank_layout = writer.create_presentation(page_size)
    elements = loader.extract_elements(diagrams[0])
    writer.add_slide(prs, blank_layout, elements)
    prs.save(out_path)

    with zipfile.ZipFile(out_path) as z:
        slide_xml = z.read("ppt/slides/slide1.xml").decode("utf-8", errors="ignore")

    target_shape_name = 'name="drawio2pptx:shape:GStdcLXKth4fSFfuQepI-4"'
    assert target_shape_name in slide_xml

    seg_name_prefix = 'name="drawio2pptx:connector:GStdcLXKth4fSFfuQepI-7:seg:'
    seg_positions = [m.start() for m in re.finditer(re.escape(seg_name_prefix), slide_xml)]
    assert seg_positions
    assert max(seg_positions) < slide_xml.index(target_shape_name)
