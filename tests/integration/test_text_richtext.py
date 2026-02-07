"""
Integration tests: rich text and headings (<h1>Heading</h1><p>...</p> preservation).
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

import pytest

from drawio2pptx.io.drawio_loader import DrawIOLoader
from drawio2pptx.io.pptx_writer import PPTXWriter


def test_timeline3_richtext_heading_is_preserved(sample_dir: Path, tmp_path: Path) -> None:
    """<h1>Heading</h1><p>... in timeline3.drawio is preserved as 'Heading' after conversion."""
    sample_path = sample_dir / "timeline3.drawio"
    if not sample_path.exists():
        pytest.skip(f"Sample file not found: {sample_path}")

    out_path = tmp_path / "timeline3_richtext_heading.pptx"

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

    assert re.search(r"<a:t>Heading</a:t>", slide_xml), (
        "Expected rich-text <h1>Heading</h1> to be preserved"
    )

    assert re.search(
        r"<a:t>Heading</a:t>[\s\S]*?</a:p>[\s\S]*?<a:pPr[\s\S]*?<a:spcAft>\s*<a:spcPts val=\"[1-9]\d*\"",
        slide_xml,
    ) or re.search(
        r"<a:pPr[\s\S]*?<a:spcAft>\s*<a:spcPts val=\"[1-9]\d*\"[\s\S]*?<a:t>Heading</a:t>",
        slide_xml,
    ), "Expected heading paragraph to have non-zero spacing after"

    assert re.search(
        r'name="drawio2pptx:shape:4ec97bd9e5d20128-14"[\s\S]*?<a:normAutofit',
        slide_xml,
    ), "Expected overflow=hidden to map to normAutofit (shrink text to fit)"
