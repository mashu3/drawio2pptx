"""
Integration tests: shape mapping (cylinder/document, tape/dataStorage, offPageConnector).
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

import pytest

from drawio2pptx.io.drawio_loader import DrawIOLoader
from drawio2pptx.io.pptx_writer import PPTXWriter


def test_cylinder_and_document_are_emitted_as_preset_shapes(
    sample_dir: Path, tmp_path: Path
) -> None:
    """cylinder3 and document in sample.drawio are emitted as PPTX preset geometries."""
    sample_path = sample_dir / "sample.drawio"
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

    cylinder_name = 'name="drawio2pptx:shape:fYnb-Lad83hC8_SQXFiI-1"'
    assert cylinder_name in slide_xml
    assert re.search(cylinder_name + r"[\s\S]*?prst=\"can\"", slide_xml)

    assert re.search(
        cylinder_name + r'[\s\S]*?<a:highlight>[\s\S]*?val="FF0000"',
        slide_xml,
    )

    document_name = 'name="drawio2pptx:shape:fYnb-Lad83hC8_SQXFiI-2"'
    assert document_name in slide_xml
    assert re.search(
        document_name + r"[\s\S]*?prst=\"flowChartDocument\"",
        slide_xml,
    )


def test_tape_and_datastorage_are_emitted_as_preset_shapes(
    sample_dir: Path, tmp_path: Path
) -> None:
    """tape and dataStorage in sample.drawio are emitted as PPTX preset geometries."""
    sample_path = sample_dir / "sample.drawio"
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

    tape_name = 'name="drawio2pptx:shape:fYnb-Lad83hC8_SQXFiI-5"'
    assert tape_name in slide_xml
    assert re.search(
        tape_name + r"[\s\S]*?prst=\"flowChartPunchedTape\"",
        slide_xml,
    )

    ds_name = 'name="drawio2pptx:shape:fYnb-Lad83hC8_SQXFiI-6"'
    assert ds_name in slide_xml
    assert re.search(
        ds_name + r"[\s\S]*?prst=\"flowChartOnlineStorage\"",
        slide_xml,
    )


def test_offpage_connector_is_emitted_as_preset_shape_and_flipV_is_applied(
    sample_dir: Path, tmp_path: Path
) -> None:
    """offPageConnector in timeline3.drawio maps to flowChartOffpageConnector with flipV applied."""
    sample_path = sample_dir / "timeline3.drawio"
    if not sample_path.exists():
        pytest.skip(f"Sample file not found: {sample_path}")

    out_path = tmp_path / "timeline3_offpage_connector.pptx"

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

    stage_ids = [
        "4ec97bd9e5d20128-5",
        "4ec97bd9e5d20128-6",
        "4ec97bd9e5d20128-7",
        "4ec97bd9e5d20128-8",
    ]

    for sid in stage_ids:
        shape_name = f'name="drawio2pptx:shape:{sid}"'
        assert shape_name in slide_xml
        assert re.search(
            shape_name + r"[\s\S]*?prst=\"flowChartOffpageConnector\"",
            slide_xml,
        ), f"Expected {sid} to map to prstGeom flowChartOffpageConnector"

    for sid in ["4ec97bd9e5d20128-7", "4ec97bd9e5d20128-8"]:
        shape_name = f'name="drawio2pptx:shape:{sid}"'
        assert re.search(
            shape_name + r'[\s\S]*?<a:xfrm[^>]*flipV="1"',
            slide_xml,
        ), f"Expected flipV=1 to be applied for {sid}"

    overlay_map = {
        "4ec97bd9e5d20128-7": "Stage 2",
        "4ec97bd9e5d20128-8": "Stage 4",
    }
    for sid, expected_text in overlay_map.items():
        overlay_name = f'name="drawio2pptx:shape-text-overlay:{sid}"'
        assert overlay_name in slide_xml
        assert re.search(
            overlay_name + r'[\s\S]*?<a:t>' + re.escape(expected_text) + r"</a:t>",
            slide_xml,
        ), f"Expected overlay for {sid} to contain text '{expected_text}'"
        assert re.search(
            overlay_name + r'[\s\S]*?<a:xfrm(?![^>]*flipH)(?![^>]*flipV)[^>]*>',
            slide_xml,
        ), f"Expected overlay for {sid} to be unflipped"
