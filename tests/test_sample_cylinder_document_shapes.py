"""Regression test for shape mapping: cylinder3/document should map to real PPTX shapes."""

from __future__ import annotations

from pathlib import Path
import re
import zipfile

from drawio2pptx.io.drawio_loader import DrawIOLoader
from drawio2pptx.io.pptx_writer import PPTXWriter


def test_cylinder_and_document_are_emitted_as_preset_shapes(tmp_path: Path) -> None:
    """
    sample.drawio contains:
    - a Cylinder node with style "shape=cylinder3" (id "fYnb-Lad83hC8_SQXFiI-1")
    - a Document node with style "shape=document" (id "fYnb-Lad83hC8_SQXFiI-2")

    They should be emitted as preset geometries rather than falling back to rectangles.
    """
    root = Path(__file__).resolve().parents[1]  # drawio2pptx/
    sample_path = root / "sample" / "sample.drawio"
    out_path = tmp_path / "sample_cylinder_document.pptx"

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

    # Validate Cylinder -> can
    cylinder_name = 'name="drawio2pptx:shape:fYnb-Lad83hC8_SQXFiI-1"'
    assert cylinder_name in slide_xml
    assert re.search(cylinder_name + r"[\s\S]*?prst=\"can\"", slide_xml), "Expected cylinder3 to map to prstGeom can"

    # Validate draw.io labelBackgroundColor -> PPTX highlight (best effort)
    # sample.drawio sets: labelBackgroundColor=light-dark(#ff0000, #FFFFFF)
    assert re.search(
        cylinder_name + r'[\s\S]*?<a:highlight>[\s\S]*?val="FF0000"',
        slide_xml,
    ), "Expected Cylinder text to include highlight color FF0000"

    # Validate Document -> flowChartDocument
    document_name = 'name="drawio2pptx:shape:fYnb-Lad83hC8_SQXFiI-2"'
    assert document_name in slide_xml
    assert re.search(
        document_name + r"[\s\S]*?prst=\"flowChartDocument\"",
        slide_xml,
    ), "Expected document to map to prstGeom flowChartDocument"


