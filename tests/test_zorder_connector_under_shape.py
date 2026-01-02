"""Regression test for z-order: connectors should not cover node shapes."""

from __future__ import annotations

from pathlib import Path
import re
import zipfile

from drawio2pptx.io.drawio_loader import DrawIOLoader
from drawio2pptx.io.pptx_writer import PPTXWriter


def test_connector_is_behind_target_shape_in_sample_drawio(tmp_path: Path) -> None:
    """
    In sample.drawio, edge "GStdcLXKth4fSFfuQepI-7" ends with a filled oval marker.

    The expected stacking is that the target square shape "GStdcLXKth4fSFfuQepI-4" is above the
    connector (including its line-end marker), matching the draw.io mxCell order.
    """
    root = Path(__file__).resolve().parents[1]  # drawio2pptx/
    sample_path = root / "sample" / "sample.drawio"
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

    # Connector "...-7" is orthogonal in the sample, so it is emitted as multiple segments.
    seg_name_prefix = 'name="drawio2pptx:connector:GStdcLXKth4fSFfuQepI-7:seg:'
    seg_positions = [m.start() for m in re.finditer(re.escape(seg_name_prefix), slide_xml)]
    assert seg_positions, "Expected connector segments to be named for z-order testing"

    # In PowerPoint slide XML, later shapes are drawn on top.
    assert max(seg_positions) < slide_xml.index(target_shape_name)


