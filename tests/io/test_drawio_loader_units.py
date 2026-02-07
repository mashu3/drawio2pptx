"""
Unit tests for drawio_loader: ColorParser, StyleExtractor, DrawIOLoader (load_file, extract_page_size).
"""
from __future__ import annotations

from pathlib import Path

import pytest
from lxml import etree as ET
from pptx.dml.color import RGBColor

from drawio2pptx.io.drawio_loader import ColorParser, StyleExtractor, DrawIOLoader
from drawio2pptx.logger import ConversionLogger


# ---- ColorParser ----
def test_color_parser_none_empty() -> None:
    assert ColorParser.parse(None) is None
    assert ColorParser.parse("") is None


def test_color_parser_none_keyword() -> None:
    assert ColorParser.parse("none") is None
    assert ColorParser.parse("None") is None


def test_color_parser_light_dark() -> None:
    """light-dark(light_color, dark_color) uses first (light) color."""
    rgb = ColorParser.parse("light-dark(#ff0000, #00ff00)")
    assert rgb is not None
    assert rgb[0] == 255 and rgb[1] == 0 and rgb[2] == 0


def test_color_parser_hex_6() -> None:
    rgb = ColorParser.parse("#FF00FF")
    assert rgb is not None
    assert rgb[0] == 255 and rgb[1] == 0 and rgb[2] == 255


def test_color_parser_hex_3() -> None:
    rgb = ColorParser.parse("#f0f")
    assert rgb is not None
    assert rgb[0] == 255 and rgb[1] == 0 and rgb[2] == 255


def test_color_parser_rgb() -> None:
    rgb = ColorParser.parse("rgb(10, 20, 30)")
    assert rgb is not None
    assert rgb[0] == 10 and rgb[1] == 20 and rgb[2] == 30


def test_color_parser_invalid_returns_none() -> None:
    assert ColorParser.parse("notacolor") is None
    assert ColorParser.parse("#gggggg") is None
    assert ColorParser.parse("rgb(1,2)") is None


# ---- StyleExtractor ----
def _cell(attrib: dict, tag: str = "mxCell") -> ET.Element:
    el = ET.Element(tag)
    for k, v in attrib.items():
        el.set(k, str(v))
    return el


def test_style_extractor_extract_style_value() -> None:
    ext = StyleExtractor()
    assert ext.extract_style_value("fillColor=#ff0000;strokeColor=#00ff00", "fillColor") == "#ff0000"
    assert ext.extract_style_value("fillColor=#ff0000", "strokeColor") is None
    assert ext.extract_style_value("", "fillColor") is None


def test_style_extractor_extract_style_float() -> None:
    ext = StyleExtractor()
    assert ext.extract_style_float("fontSize=12;width=100", "fontSize") == 12.0
    assert ext.extract_style_float("fontSize=12", "fontSize", default=10.0) == 12.0
    assert ext.extract_style_float("x=abc", "x", default=5.0) == 5.0


def test_style_extractor_parse_font_style() -> None:
    ext = StyleExtractor()
    r = ext._parse_font_style("1")  # bold
    assert r["bold"] is True and r["italic"] is False
    r = ext._parse_font_style("2")  # italic
    assert r["italic"] is True
    r = ext._parse_font_style("4")  # underline
    assert r["underline"] is True
    r = ext._parse_font_style("0")
    assert r["bold"] is False and r["italic"] is False


def test_style_extractor_is_text_style() -> None:
    ext = StyleExtractor()
    assert ext.is_text_style("text;html=1") is True
    assert ext.is_text_style("shape=text;fillColor=#fff") is True
    assert ext.is_text_style("ellipse;fillColor=#fff") is False


def test_style_extractor_extract_fill_color_attr() -> None:
    ext = StyleExtractor()
    cell = _cell({"fillColor": "#ff0000"})
    assert ext.extract_fill_color(cell) is not None
    cell = _cell({"fillColor": "default"})
    assert ext.extract_fill_color(cell) == "default"
    cell = _cell({"fillColor": "none"})
    assert ext.extract_fill_color(cell) is None


def test_style_extractor_extract_fill_color_style() -> None:
    ext = StyleExtractor()
    cell = _cell({"style": "fillColor=#00ff00;strokeColor=#0000ff"})
    rgb = ext.extract_fill_color(cell)
    assert rgb is not None
    assert rgb[1] == 255


def test_style_extractor_extract_fill_color_vertex_default() -> None:
    ext = StyleExtractor()
    cell = _cell({"vertex": "1"}, tag="mxCell")
    assert ext.extract_fill_color(cell) == "default"


def test_style_extractor_extract_gradient_color() -> None:
    ext = StyleExtractor()
    cell = _cell({"style": "gradientColor=#ff0000"})
    assert ext.extract_gradient_color(cell) is not None
    cell = _cell({"style": "gradientColor=default"})
    assert ext.extract_gradient_color(cell) == "default"
    cell = _cell({"style": "gradientColor=none"})
    assert ext.extract_gradient_color(cell) is None


def test_style_extractor_extract_gradient_direction() -> None:
    ext = StyleExtractor()
    cell = _cell({"style": "gradientDirection=north"})
    assert ext.extract_gradient_direction(cell) == "north"
    cell = _cell({})
    assert ext.extract_gradient_direction(cell) is None


def test_style_extractor_extract_swimlane_fill_color() -> None:
    ext = StyleExtractor()
    cell = _cell({"style": "swimlaneFillColor=#eeeeee"})
    assert ext.extract_swimlane_fill_color(cell) is not None
    cell = _cell({"style": "swimlaneFillColor=none"})
    assert ext.extract_swimlane_fill_color(cell) is None


def test_style_extractor_extract_stroke_color() -> None:
    ext = StyleExtractor()
    cell = _cell({"style": "strokeColor=#000000"})
    assert ext.extract_stroke_color(cell) is not None
    cell = _cell({})
    assert ext.extract_stroke_color(cell) is None


def test_style_extractor_extract_no_stroke() -> None:
    ext = StyleExtractor()
    cell = _cell({"style": "strokeColor=none"})
    assert ext.extract_no_stroke(cell) is True
    cell = _cell({"strokeColor": "none"})
    assert ext.extract_no_stroke(cell) is True
    cell = _cell({"style": "strokeColor=#000"})
    assert ext.extract_no_stroke(cell) is False


def test_style_extractor_extract_font_color_from_style() -> None:
    ext = StyleExtractor()
    cell = _cell({"style": "fontColor=#123456"})
    assert ext.extract_font_color(cell) is not None


def test_style_extractor_extract_label_background_color() -> None:
    ext = StyleExtractor()
    cell = _cell({"style": "labelBackgroundColor=#ffff00"})
    assert ext.extract_label_background_color(cell) is not None
    cell = _cell({"style": "labelBackgroundColor=#ff0000"})
    assert ext.extract_label_background_color(cell) is not None


def test_style_extractor_extract_shadow() -> None:
    ext = StyleExtractor()
    cell = _cell({"style": "shadow=1"})
    assert ext.extract_shadow(cell, None) is True
    cell = _cell({"style": "shadow=0"})
    assert ext.extract_shadow(cell, None) is False
    root = ET.Element("root")
    root.set("shadow", "1")
    assert ext.extract_shadow(_cell({}), root) is True


def test_style_extractor_extract_shape_type_swimlane() -> None:
    ext = StyleExtractor()
    cell = _cell({"style": "shape=swimlane"})
    assert ext.extract_shape_type(cell) == "swimlane"


def test_style_extractor_extract_shape_type_process_predefined() -> None:
    ext = StyleExtractor()
    cell = _cell({"style": "shape=process;backgroundOutline=1"})
    assert ext.extract_shape_type(cell) == "predefinedprocess"
    cell = _cell({"style": "shape=process;size=10"})
    assert ext.extract_shape_type(cell) == "predefinedprocess"


def test_style_extractor_extract_shape_type_map() -> None:
    ext = StyleExtractor()
    cell = _cell({"style": "shape=ellipse"})
    assert ext.extract_shape_type(cell) == "ellipse"
    cell = _cell({"style": "shape=mxgraph.flowchart.document"})
    assert ext.extract_shape_type(cell) == "document"


def test_style_extractor_extract_shape_type_first_part() -> None:
    ext = StyleExtractor()
    cell = _cell({"style": "ellipse;fillColor=#fff"})
    assert ext.extract_shape_type(cell) == "ellipse"
    cell = _cell({"style": "rhombus;fillColor=#fff"})
    assert ext.extract_shape_type(cell) == "rhombus"


def test_style_extractor_extract_shape_type_default_rectangle() -> None:
    ext = StyleExtractor()
    cell = _cell({})
    assert ext.extract_shape_type(cell) == "rectangle"


# ---- DrawIOLoader load_file ----
def test_loader_load_file_diagram_with_mxgraph(tmp_path: Path) -> None:
    """load_file returns list of mxGraphModel when file has <diagram> with inner XML."""
    drawio = tmp_path / "test.drawio"
    drawio.write_text("""<?xml version="1.0"?>
<mxfile>
  <diagram name="Page-1">
    <mxGraphModel dx="800" dy="600">
      <root><mxCell id="0"/><mxCell id="1" parent="0"/></root>
    </mxGraphModel>
  </diagram>
</mxfile>""", encoding="utf-8")
    loader = DrawIOLoader()
    diagrams = loader.load_file(drawio)
    assert len(diagrams) == 1
    root = diagrams[0]
    assert root.tag.endswith("mxGraphModel") or root.tag == "mxGraphModel"
    assert len(root.findall(".//mxCell")) >= 1


def test_loader_load_file_diagram_empty_inner_uses_fallback(tmp_path: Path) -> None:
    """When diagram has no inner content, loader uses mxGraphModel from root if present."""
    drawio = tmp_path / "test.drawio"
    drawio.write_text("""<?xml version="1.0"?>
<mxfile>
  <diagram name="Page-1"></diagram>
  <mxGraphModel><root><mxCell id="0"/></root></mxGraphModel>
</mxfile>""", encoding="utf-8")
    loader = DrawIOLoader()
    diagrams = loader.load_file(drawio)
    assert len(diagrams) >= 1


# ---- DrawIOLoader extract_page_size ----
def test_loader_extract_page_size() -> None:
    loader = DrawIOLoader()
    root = ET.Element("mxGraphModel")
    root.set("pageWidth", "800")
    root.set("pageHeight", "600")
    w, h = loader.extract_page_size(root)
    assert w == 800.0 and h == 600.0


def test_loader_extract_page_size_with_scale() -> None:
    loader = DrawIOLoader()
    root = ET.Element("mxGraphModel")
    root.set("pageWidth", "400")
    root.set("pageHeight", "300")
    root.set("pageScale", "2")
    w, h = loader.extract_page_size(root)
    assert w == 800.0 and h == 600.0


def test_loader_extract_page_size_missing_returns_none() -> None:
    loader = DrawIOLoader()
    root = ET.Element("mxGraphModel")
    w, h = loader.extract_page_size(root)
    assert w is None and h is None


def test_loader_extract_page_size_invalid_scale_defaults_to_one() -> None:
    loader = DrawIOLoader()
    root = ET.Element("mxGraphModel")
    root.set("pageWidth", "100")
    root.set("pageHeight", "100")
    root.set("pageScale", "x")
    w, h = loader.extract_page_size(root)
    assert w == 100.0 and h == 100.0
