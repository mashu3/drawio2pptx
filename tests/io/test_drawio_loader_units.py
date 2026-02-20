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
from drawio2pptx.model.intermediate import ShapeElement


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


# ---- DrawIOLoader shape extraction helpers (refactored for unit testing) ----
def _mx_cell_with_geometry(
    x: float = 0, y: float = 0, width: float = 100, height: float = 50,
    cell_id: str = "cell1", parent: str = "1", value: str = "", style: str = ""
) -> ET.Element:
    """Build an mxCell with mxGeometry child."""
    cell = ET.Element("mxCell")
    cell.set("id", cell_id)
    cell.set("parent", parent)
    if value:
        cell.set("value", value)
    if style:
        cell.set("style", style)
    geo = ET.SubElement(cell, "mxGeometry")
    geo.set("x", str(x))
    geo.set("y", str(y))
    geo.set("width", str(width))
    geo.set("height", str(height))
    geo.set("as", "geometry")
    return cell


def test_parse_geometry_from_cell() -> None:
    """_parse_geometry_from_cell returns (x, y, w, h) or None."""
    loader = DrawIOLoader()
    cell = _mx_cell_with_geometry(10, 20, 80, 40)
    rect = loader._parse_geometry_from_cell(cell)
    assert rect == (10.0, 20.0, 80.0, 40.0)

    cell_no_geo = ET.Element("mxCell")
    cell_no_geo.set("id", "x")
    assert loader._parse_geometry_from_cell(cell_no_geo) is None

    cell_bad = _mx_cell_with_geometry(0, 0, 100, 50)
    cell_bad.find("mxGeometry").set("width", "abc")
    assert loader._parse_geometry_from_cell(cell_bad) is None


def test_apply_shape_parent_offset_no_parent() -> None:
    """_apply_shape_parent_offset returns (x,y) unchanged when parent is None or 0/1."""
    loader = DrawIOLoader()
    cell = _mx_cell_with_geometry(5, 10, 50, 50)
    root = ET.Element("root")
    assert loader._apply_shape_parent_offset(5.0, 10.0, None, cell, root) == (5.0, 10.0)
    assert loader._apply_shape_parent_offset(5.0, 10.0, "0", cell, root) == (5.0, 10.0)
    assert loader._apply_shape_parent_offset(5.0, 10.0, "1", cell, root) == (5.0, 10.0)


def test_apply_shape_parent_offset_with_parent() -> None:
    """_apply_shape_parent_offset adds parent coordinates when parent exists in mgm_root."""
    loader = DrawIOLoader()
    cell = _mx_cell_with_geometry(10, 20, 60, 40, parent="p1")
    root = ET.fromstring("""<root>
      <mxCell id="p1"><mxGeometry x="100" y="200" width="0" height="0" as="geometry"/></mxCell>
    </root>""")
    x, y = loader._apply_shape_parent_offset(10.0, 20.0, "p1", cell, root)
    assert x == 110.0 and y == 220.0


def test_build_shape_transform_default() -> None:
    """_build_shape_transform returns no rotation and no flip by default."""
    loader = DrawIOLoader()
    t = loader._build_shape_transform("", None)
    assert t.rotation == 0.0 and t.flip_h is False and t.flip_v is False


def test_build_shape_transform_rotation_flip() -> None:
    """_build_shape_transform reads rotation, flipH, flipV from style."""
    loader = DrawIOLoader()
    t = loader._build_shape_transform("rotation=90;flipH=1;flipV=1", "rectangle")
    assert t.rotation == 90.0 and t.flip_h is True and t.flip_v is True


def test_build_shape_transform_arrow_direction() -> None:
    """_build_shape_transform folds arrow direction into rotation for right_arrow."""
    loader = DrawIOLoader()
    t = loader._build_shape_transform("direction=north", "right_arrow")
    assert t.rotation == 270.0 and t.flip_h is False and t.flip_v is False
    t2 = loader._build_shape_transform("direction=south;flipV=1", "right_arrow")
    assert t2.rotation == 270.0  # 90 + 180


def test_extract_word_wrap_from_style() -> None:
    """_extract_word_wrap_from_style: wrap -> True, nowrap/unspecified -> False."""
    loader = DrawIOLoader()
    assert loader._extract_word_wrap_from_style("whiteSpace=wrap") is True
    assert loader._extract_word_wrap_from_style("whiteSpace=nowrap") is False
    assert loader._extract_word_wrap_from_style("whiteSpace=foo") is False
    assert loader._extract_word_wrap_from_style("") is False
    assert loader._extract_word_wrap_from_style("fillColor=#fff") is False


def test_extract_shape_returns_none_when_no_geometry() -> None:
    """_extract_shape returns None when cell has no mxGeometry."""
    loader = DrawIOLoader()
    cell = ET.Element("mxCell")
    cell.set("id", "c1")
    cell.set("parent", "1")
    root = ET.Element("root")
    assert loader._extract_shape(cell, root) is None


def test_extract_shape_returns_shape_with_minimal_cell() -> None:
    """_extract_shape returns ShapeElement with geometry and style from minimal mxCell."""
    loader = DrawIOLoader()
    cell = _mx_cell_with_geometry(20, 30, 100, 60, value="Hello", style="rounded=1;fillColor=#ffffff")
    root = ET.Element("root")
    shape = loader._extract_shape(cell, root)
    assert shape is not None
    assert shape.x == 20.0 and shape.y == 30.0 and shape.w == 100.0 and shape.h == 60.0
    assert shape.shape_type == "rectangle"
    assert shape.style.corner_radius is not None
    assert len(shape.text) >= 1
    assert shape.text[0].runs and shape.text[0].runs[0].text == "Hello"


def test_build_style_for_shape_cell_word_wrap_and_rounded() -> None:
    """_build_style_for_shape_cell respects word_wrap parameter and rounded style."""
    loader = DrawIOLoader()
    cell = _mx_cell_with_geometry(0, 0, 80, 40, style="rounded=1;fillColor=#f0f0f0")
    root = ET.Element("root")
    style = loader._build_style_for_shape_cell(cell, root, 80.0, 40.0, word_wrap=True)
    assert style.word_wrap is True
    assert style.corner_radius is not None
    style_nowrap = loader._build_style_for_shape_cell(cell, root, 80.0, 40.0, word_wrap=False)
    assert style_nowrap.word_wrap is False


# ---- DrawIOLoader connector extraction helpers ----
def test_parse_connector_edge_style() -> None:
    """_parse_connector_edge_style returns (edge_style, is_elbow_edge)."""
    loader = DrawIOLoader()
    edge, elbow = loader._parse_connector_edge_style("")
    assert edge == "straight" and elbow is False
    edge, elbow = loader._parse_connector_edge_style("edgeStyle=orthogonalEdgeStyle")
    assert edge == "orthogonal" and elbow is False
    edge, elbow = loader._parse_connector_edge_style("edgeStyle=elbowEdgeStyle")
    assert edge == "orthogonal" and elbow is True
    edge, elbow = loader._parse_connector_edge_style("edgeStyle=curved")
    assert edge == "curved" and elbow is False


def test_build_connector_style() -> None:
    """_build_connector_style returns Style with stroke, arrows, dash, shadow."""
    loader = DrawIOLoader()
    cell = _cell({"style": "strokeColor=#000000;strokeWidth=2;dashed=1;startArrow=classic;endArrow=block;shadow=1"})
    root = ET.Element("mxGraphModel")
    style = loader._build_connector_style(cell, root, cell.attrib.get("style", ""))
    assert style.stroke_width == 2.0
    assert style.dash == "dashed"
    assert style.arrow_start == "classic"
    assert style.arrow_end == "block"
    assert style.has_shadow is True


def test_build_connector_style_default_end_arrow_from_mgm() -> None:
    """When endArrow is missing and mgm_root has arrows=1, end_arrow becomes classic."""
    loader = DrawIOLoader()
    cell = _cell({"style": "strokeColor=#000000"})
    root = ET.Element("mxGraphModel")
    root.set("arrows", "1")
    style = loader._build_connector_style(cell, root, "")
    assert style.arrow_end == "classic"


def test_parse_connector_geometry_empty() -> None:
    """_parse_connector_geometry returns empty lists when no mxGeometry."""
    loader = DrawIOLoader()
    cell = ET.Element("mxCell")
    cell.set("id", "e1")
    root = ET.Element("root")
    points_raw, source_pt, target_pt, points_for_ports = loader._parse_connector_geometry(cell, root)
    assert points_raw == [] and source_pt is None and target_pt is None and points_for_ports == []


def test_parse_connector_geometry_with_array_points() -> None:
    """_parse_connector_geometry parses Array[@as='points'] waypoints."""
    loader = DrawIOLoader()
    root = ET.fromstring("""<root>
      <mxCell id="e1" parent="1">
        <mxGeometry>
          <Array as="points">
            <mxPoint x="50" y="20"/>
            <mxPoint x="100" y="80"/>
          </Array>
        </mxGeometry>
      </mxCell>
    </root>""")
    cell = root.find(".//mxCell")
    points_raw, source_pt, target_pt, points_for_ports = loader._parse_connector_geometry(cell, root)
    assert len(points_raw) == 2
    assert points_raw[0] == (50.0, 20.0) and points_raw[1] == (100.0, 80.0)
    assert source_pt is None and target_pt is None
    assert len(points_for_ports) == 2


def test_extract_connector_keeps_floating_source_target_points() -> None:
    """Standalone edge with sourcePoint/targetPoint must stay floating (no inferred shape attachment)."""
    loader = DrawIOLoader()
    root = ET.fromstring("""<mxGraphModel arrows="1"><root><mxCell id="0"/><mxCell id="1" parent="0"/></root></mxGraphModel>""")
    cell = ET.Element("mxCell")
    cell.set("id", "e1")
    cell.set("edge", "1")
    cell.set("parent", "1")
    cell.set("style", "edgeStyle=none;strokeWidth=3;")
    geo = ET.SubElement(cell, "mxGeometry")
    ET.SubElement(geo, "mxPoint", {"x": "100", "y": "200", "as": "sourcePoint"})
    ET.SubElement(geo, "mxPoint", {"x": "100", "y": "50", "as": "targetPoint"})
    shapes_dict = {
        "s1": ShapeElement(id="s1", x=90.0, y=190.0, w=20.0, h=20.0),
        "t1": ShapeElement(id="t1", x=90.0, y=40.0, w=20.0, h=20.0),
    }

    connector, labels = loader._extract_connector(cell, root, shapes_dict)

    assert labels == []
    assert connector is not None
    assert connector.source_id is None
    assert connector.target_id is None
    assert connector.points == [(100.0, 200.0), (100.0, 50.0)]


# ---- DrawIOLoader port / orthogonal / boundary helpers ----
def test_infer_port_side() -> None:
    """_infer_port_side returns left/right/top/bottom or None for center."""
    loader = DrawIOLoader()
    assert loader._infer_port_side(0.0, 0.5) == "left"
    assert loader._infer_port_side(1.0, 0.5) == "right"
    assert loader._infer_port_side(0.5, 0.0) == "top"
    assert loader._infer_port_side(0.5, 1.0) == "bottom"
    assert loader._infer_port_side(0.5, 0.5) is None
    assert loader._infer_port_side(None, 0.5) is None


def test_snap_to_grid() -> None:
    """_snap_to_grid snaps value to grid or returns unchanged when grid is None/0."""
    loader = DrawIOLoader()
    assert loader._snap_to_grid(12.3, None) == 12.3
    assert loader._snap_to_grid(12.3, 10.0) == 10.0
    assert loader._snap_to_grid(15.0, 10.0) == 20.0


def test_clamp01() -> None:
    """_clamp01 clamps to [0,1] or returns None."""
    loader = DrawIOLoader()
    assert loader._clamp01(None) is None
    assert loader._clamp01(0.5) == 0.5
    assert loader._clamp01(-0.1) == 0.0
    assert loader._clamp01(1.5) == 1.0


def test_boundary_point_rect() -> None:
    """_calculate_boundary_point for rectangle: edges and offset."""
    loader = DrawIOLoader()
    shape = ShapeElement(id="r1", x=10.0, y=20.0, w=100.0, h=50.0, shape_type="rectangle")
    x, y = loader._calculate_boundary_point(shape, 0.0, 0.5, 0.0, 0.0)
    assert x == 10.0 and y == 45.0
    x, y = loader._calculate_boundary_point(shape, 1.0, 0.5, 0.0, 0.0)
    assert x == 110.0 and y == 45.0
    x, y = loader._calculate_boundary_point(shape, 0.5, 0.0, 0.0, 0.0)
    assert x == 60.0 and y == 20.0
    x, y = loader._calculate_boundary_point(shape, 0.5, 1.0, 2.0, 3.0)
    assert x == 62.0 and y == 73.0


def test_boundary_point_rhombus() -> None:
    """_calculate_boundary_point for rhombus: point on edge from center."""
    loader = DrawIOLoader()
    shape = ShapeElement(id="d1", x=0.0, y=0.0, w=40.0, h=40.0, shape_type="rhombus")
    x, y = loader._calculate_boundary_point(shape, 1.0, 0.5, 0.0, 0.0)
    assert abs(x - 40.0) < 1e-5 and abs(y - 20.0) < 1e-5


def test_boundary_point_ellipse() -> None:
    """_calculate_boundary_point for ellipse: point on ellipse from center."""
    loader = DrawIOLoader()
    shape = ShapeElement(id="e1", x=0.0, y=0.0, w=60.0, h=40.0, shape_type="ellipse")
    x, y = loader._calculate_boundary_point(shape, 1.0, 0.5, 0.0, 0.0)
    assert abs(x - 60.0) < 1e-5 and abs(y - 20.0) < 1e-5


def test_ensure_orthogonal_route_respects_ports() -> None:
    """_ensure_orthogonal_route_respects_ports adds bend when exit/entry directions require it."""
    loader = DrawIOLoader()
    pts = [(0.0, 0.0), (100.0, 100.0)]
    out = loader._ensure_orthogonal_route_respects_ports(pts, 1.0, 0.5, 0.5, 0.0, None)
    assert len(out) == 3
    assert out[0] == (0.0, 0.0) and out[-1] == (100.0, 100.0)


def test_parse_edge_label_start_end() -> None:
    """_parse_edge_label_start_end returns (is_start, is_end) from geometry and align."""
    loader = DrawIOLoader()
    cell = ET.Element("mxCell")
    geo = ET.SubElement(cell, "mxGeometry")
    geo.set("relative", "1")
    geo.set("x", "-0.8")
    start, end = loader._parse_edge_label_start_end(cell, "align=left")
    assert start is True and end is False
    geo.set("x", "0.8")
    start, end = loader._parse_edge_label_start_end(cell, "align=right")
    assert start is False and end is True
    start, end = loader._parse_edge_label_start_end(cell, "")
    assert start is False and end is False


def test_resolve_connector_points_two_shapes() -> None:
    """_resolve_connector_points returns list of points for two shapes with no waypoints."""
    loader = DrawIOLoader()
    source = ShapeElement(id="s1", x=0.0, y=0.0, w=40.0, h=30.0)
    target = ShapeElement(id="t1", x=100.0, y=50.0, w=40.0, h=30.0)
    points = loader._resolve_connector_points(
        source, target, [], [], None, None,
        "", "orthogonal", False, None,
    )
    assert len(points) >= 2
    # Start near source right edge (40), end near target (100–140)
    assert points[0][0] == pytest.approx(40.0, abs=2.0)
    assert 98.0 <= points[-1][0] <= 142.0


# ---- StyleExtractor helpers ----
def test_style_extractor_get_attr_or_style_value() -> None:
    """_get_attr_or_style_value returns attribute or value from style string."""
    ext = StyleExtractor()
    cell = _cell({"fillColor": "#ff0000"})
    assert ext._get_attr_or_style_value(cell, "fillColor") == "#ff0000"
    cell2 = _cell({"style": "fillColor=#00ff00"})
    assert ext._get_attr_or_style_value(cell2, "fillColor") == "#00ff00"
    assert ext._get_attr_or_style_value(cell2, "strokeColor") is None


def test_style_extractor_parse_color_value() -> None:
    """_parse_color_value returns default, None, or parsed RGBColor."""
    ext = StyleExtractor()
    assert ext._parse_color_value("default") == "default"
    assert ext._parse_color_value("none") is None
    assert ext._parse_color_value(None) is None
    rgb = ext._parse_color_value("#0000ff")
    assert rgb is not None and rgb[2] == 255


# ---- DrawIOLoader cell index / polyline / label helpers ----
def test_build_cell_index_and_draw_order() -> None:
    """_build_cell_index_and_draw_order returns cells, index, and z-order."""
    loader = DrawIOLoader()
    root = ET.fromstring("""<mxGraphModel>
      <root><mxCell id="0"/><mxCell id="1" parent="0"/><mxCell id="v1" parent="1" vertex="1"/><mxCell id="e1" parent="1" edge="1"/></root>
    </mxGraphModel>""")
    cells, cell_by_id, doc_order, children, containers, cell_order = loader._build_cell_index_and_draw_order(root)
    assert len(cells) >= 3
    assert "v1" in cell_by_id and "e1" in cell_by_id
    assert "1" in children
    assert "v1" in cell_order and "e1" in cell_order


def test_polyline_segment_lengths() -> None:
    """_polyline_segment_lengths returns total length and per-segment lengths."""
    loader = DrawIOLoader()
    pts = [(0.0, 0.0), (3.0, 0.0), (3.0, 4.0)]
    total, segs = loader._polyline_segment_lengths(pts)
    assert total == 7.0
    assert segs == [3.0, 4.0]
    t2, s2 = loader._polyline_segment_lengths([])
    assert t2 == 0.0 and s2 == []


def test_point_along_polyline() -> None:
    """_point_along_polyline returns base point and segment direction at t_rel."""
    loader = DrawIOLoader()
    pts = [(0.0, 0.0), (10.0, 0.0), (10.0, 10.0)]
    total, segs = loader._polyline_segment_lengths(pts)
    bx, by, sdx, sdy = loader._point_along_polyline(pts, segs, total, 0.0)
    assert (bx, by) == (0.0, 0.0)
    bx, by, sdx, sdy = loader._point_along_polyline(pts, segs, total, 0.5)
    assert abs(bx - 10.0) < 1e-5 and abs(by) < 1e-5
    bx, by, sdx, sdy = loader._point_along_polyline(pts, segs, total, 1.0)
    assert (bx, by) == (10.0, 10.0)


def test_segment_normal_for_label() -> None:
    """_segment_normal_for_label returns unit normal for orthogonal and straight."""
    loader = DrawIOLoader()
    nx, ny = loader._segment_normal_for_label(1.0, 0.0, "orthogonal")
    assert (nx, ny) == (0.0, -1.0)
    nx, ny = loader._segment_normal_for_label(0.0, 1.0, "orthogonal")
    assert (nx, ny) == (1.0, 0.0)
    nx, ny = loader._segment_normal_for_label(3.0, 4.0, "straight")
    assert abs(nx * nx + ny * ny - 1.0) < 1e-9


def test_adjust_label_outside_shape() -> None:
    """_adjust_label_outside_shape moves label right or left when overlapping shape."""
    loader = DrawIOLoader()
    shape = ShapeElement(id="s1", x=100.0, y=100.0, w=50.0, h=30.0)
    # Label top-left inside shape -> push right
    x, y = loader._adjust_label_outside_shape(110.0, 105.0, 20.0, 10.0, shape, push_to_right=True)
    assert x == 155.0  # shape_right + 5
    # Label overlapping from left -> push left
    x, y = loader._adjust_label_outside_shape(80.0, 105.0, 40.0, 10.0, shape, push_to_right=False)
    assert x == 55.0  # shape.x - label_w - 5
