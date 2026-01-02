"""Regression tests for open (unfilled) oval markers specified by draw.io startFill/endFill=0."""

from __future__ import annotations

from pathlib import Path
import zipfile

from drawio2pptx.io.drawio_loader import DrawIOLoader
from drawio2pptx.io.pptx_writer import PPTXWriter


def test_generated_pptx_emulates_open_oval_marker_when_startFill_is_zero(tmp_path: Path):
    """
    draw.io can encode an "open circle" marker with startArrow=oval and startFill=0.

    PowerPoint line-end elements cannot express an unfilled oval, so we emulate it with a tiny
    ellipse outline placed at the connector endpoint, and trim the connector line so it stops at
    the marker boundary (background-color independent).
    """
    root = Path(__file__).resolve().parents[1]  # drawio2pptx/
    sample_path = root / "sample" / "sample.drawio"
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

    # The connector with id "...-8" in sample.drawio has startArrow=oval and startFill=0.
    assert "drawio2pptx:marker:open-oval:GStdcLXKth4fSFfuQepI-8:start" in slide_xml

    # There is only one filled oval arrow in sample.drawio (edge "...-7" uses endArrow=oval;endFill=1).
    # The open-oval marker should not contribute a line-end type="oval".
    assert slide_xml.count('type="oval"') == 1

    # The marker should be unfilled (noFill).
    marker_name = "drawio2pptx:marker:open-oval:GStdcLXKth4fSFfuQepI-8:start"
    import re
    assert re.search(
        r'name="' + re.escape(marker_name) + r'".{0,2000}?<a:noFill/>',
        slide_xml,
    ), "Expected the open-oval marker shape to have <a:noFill/>"

    # Size sanity check (unit conversion verification):
    # In this connector (id "...-8"), startSize is omitted in sample.drawio, so mxGraph default is 6.
    # The overlay marker diameter is chosen as max(base_d, 6 + strokeWidth*1.25).
    # Here strokeWidth is default 1 -> max(6, 7.25) = 7.25 px.
    # At 96 DPI, 1px = 9525 EMU => 7.25px ~= 69056 EMU.
    marker_name = "drawio2pptx:marker:open-oval:GStdcLXKth4fSFfuQepI-8:start"
    m = re.search(
        r'name="' + re.escape(marker_name) + r'"[\s\S]*?<a:ext cx="(\d+)" cy="(\d+)"',
        slide_xml,
    )
    assert m, "Expected to find open-oval marker shape extents in slide XML"
    cx = int(m.group(1))
    cy = int(m.group(2))
    assert abs(cx - 69056) <= 20
    assert abs(cy - 69056) <= 20


