"""Regression test for draw.io endSize/startSize affecting PowerPoint arrow size mapping."""

from __future__ import annotations

import re
from pathlib import Path
import zipfile

from drawio2pptx.io.drawio_loader import DrawIOLoader
from drawio2pptx.io.pptx_writer import PPTXWriter


def test_generated_pptx_respects_endSize_for_filled_oval_arrow(tmp_path: Path):
    """
    sample.drawio contains one filled oval marker:
      endArrow=oval; endFill=1; endSize=6

    We map draw.io endSize (px) to PowerPoint's discrete arrow size:
      endSize=6 -> w/len="sm"
    """
    root = Path(__file__).resolve().parents[1]  # drawio2pptx/
    sample_path = root / "sample" / "sample.drawio"
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


