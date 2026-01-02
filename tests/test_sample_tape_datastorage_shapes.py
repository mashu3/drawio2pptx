"""Regression test for shape mapping: tape/dataStorage should map to PPTX preset shapes."""

from __future__ import annotations

from pathlib import Path
import re
import zipfile

from drawio2pptx.io.drawio_loader import DrawIOLoader
from drawio2pptx.io.pptx_writer import PPTXWriter


def test_tape_and_datastorage_are_emitted_as_preset_shapes(tmp_path: Path) -> None:
    """
    sample.drawio contains:
    - a Tape node with style "shape=tape" (id "fYnb-Lad83hC8_SQXFiI-5")
    - a Data Storage node with style "shape=dataStorage" (id "fYnb-Lad83hC8_SQXFiI-6")

    They should be emitted as preset geometries rather than falling back to rectangles.
    """
    root = Path(__file__).resolve().parents[1]  # drawio2pptx/
    sample_path = root / "sample" / "sample.drawio"
    out_path = tmp_path / "sample_tape_datastorage.pptx"

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

    # Validate Tape -> flowChartPunchedTape
    tape_name = 'name="drawio2pptx:shape:fYnb-Lad83hC8_SQXFiI-5"'
    assert tape_name in slide_xml
    assert re.search(
        tape_name + r"[\s\S]*?prst=\"flowChartPunchedTape\"",
        slide_xml,
    ), "Expected tape to map to prstGeom flowChartPunchedTape"

    # Validate Data Storage -> flowChartOnlineStorage (Stored Data)
    ds_name = 'name="drawio2pptx:shape:fYnb-Lad83hC8_SQXFiI-6"'
    assert ds_name in slide_xml
    assert re.search(
        ds_name + r"[\s\S]*?prst=\"flowChartOnlineStorage\"",
        slide_xml,
    ), "Expected dataStorage to map to prstGeom flowChartOnlineStorage"


