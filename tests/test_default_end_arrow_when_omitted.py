"""Regression tests for default end arrow when draw.io omits endArrow from style."""

from __future__ import annotations

from pathlib import Path
import zipfile

from drawio2pptx.io.drawio_loader import DrawIOLoader
from drawio2pptx.io.pptx_writer import PPTXWriter
from drawio2pptx.model.intermediate import ConnectorElement


def test_loader_applies_default_end_arrow_when_omitted():
    root = Path(__file__).resolve().parents[1]  # drawio2pptx/
    sample_path = root / "sample" / "sample.drawio"

    loader = DrawIOLoader()
    diagrams = loader.load_file(sample_path)
    elements = loader.extract_elements(diagrams[0])

    edge_id = "GStdcLXKth4fSFfuQepI-11"  # parallelogram -> ellipse in sample.drawio
    conn = next(e for e in elements if isinstance(e, ConnectorElement) and e.id == edge_id)

    # draw.io can omit endArrow when it is the default arrow.
    assert (conn.style.arrow_end or "").lower() == "classic"


def test_generated_pptx_contains_triangle_arrow_type_for_default_end_arrow(tmp_path: Path):
    """
    The default end arrow ('classic') should map to a triangle line-end type in the produced slide XML.

    This test checks the produced OOXML directly to avoid relying on UI rendering.
    """
    root = Path(__file__).resolve().parents[1]  # drawio2pptx/
    sample_path = root / "sample" / "sample.drawio"
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

    # Before the fix, sample.drawio had no triangle arrows at all; classic defaults add them.
    assert 'type="triangle"' in slide_xml
    # The default arrow is the *end* arrow, so it must appear as a:tailEnd (line end), not a:headEnd.
    assert '<a:tailEnd' in slide_xml and 'type="triangle"' in slide_xml


