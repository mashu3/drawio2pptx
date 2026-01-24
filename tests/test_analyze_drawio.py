"""
Test module for drawio file analysis functionality
"""
import pytest
from pathlib import Path
from lxml import etree as ET
from lxml import html as lxml_html
from drawio2pptx.io.drawio_loader import ColorParser


def parse_color(color_str):
    """Parse color string and return formatted string"""
    if not color_str:
        return "None"
    
    # Use ColorParser to parse color
    rgb_color = ColorParser.parse(color_str)
    if rgb_color:
        # RGBColor has tuple-like structure, accessible by index
        r, g, b = rgb_color[0], rgb_color[1], rgb_color[2]
        
        # For hexadecimal format, preserve original format if short form
        if color_str.startswith('#'):
            # Use original string for hex format
            return f"RGB({r}, {g}, {b}) / #{color_str[1:]}"
        else:
            # For rgb format, don't include hex
            return f"RGB({r}, {g}, {b})"
    
    return color_str


def extract_style_value(style_str, key):
    """Extract value for specified key from style string"""
    if not style_str:
        return None
    for part in style_str.split(";"):
        if "=" in part:
            k, v = part.split("=", 1)
            if k.strip() == key:
                return v.strip()
    return None


def parse_drawio_file(path):
    """Parse drawio file and return list of diagrams"""
    tree = ET.parse(path)
    root = tree.getroot()
    
    diagrams = []
    for d in root.findall(".//diagram"):
        inner = (d.text or "").strip()
        if not inner:
            mgm = d.find(".//mxGraphModel")
            if mgm is not None:
                diagrams.append(mgm)
                continue
            continue
        
        # Process HTML entities
        if "&lt;" in inner or "&amp;" in inner:
            try:
                wrapped = f"<div>{inner}</div>"
                parsed = lxml_html.fromstring(wrapped)
                inner = parsed.text_content()
            except Exception:
                pass
        
        # Parse XML fragment
        if "<mxGraphModel" in inner or "<root" in inner or "<mxCell" in inner:
            try:
                parsed = ET.fromstring(inner)
                mgm = None
                if parsed.tag.endswith("mxGraphModel") or parsed.tag == "mxGraphModel":
                    mgm = parsed
                else:
                    mgm = parsed.find(".//mxGraphModel")
                    if mgm is None and parsed.tag.endswith("root"):
                        mgm = parsed
                if mgm is not None:
                    diagrams.append(mgm)
                    continue
            except ET.ParseError:
                pass
        
        # Fallback
        mgm_global = root.find(".//mxGraphModel")
        if mgm_global is not None:
            diagrams.append(mgm_global)
        else:
            diagrams.append(root)
    
    if not diagrams:
        mgm_global = root.find(".//mxGraphModel")
        if mgm_global is not None:
            diagrams.append(mgm_global)
        else:
            diagrams.append(root)
    
    return diagrams


def analyze_mxcell(cell):
    """Analyze mxCell element and return dictionary"""
    result = {
        'id': cell.attrib.get('id', 'N/A'),
        'value': cell.attrib.get('value', ''),
        'vertex': cell.attrib.get('vertex', 'N/A'),
        'fill_color_attr': None,
        'stroke_color_attr': None,
        'font_color_attr': None,
        'fill_color_style': None,
        'stroke_color_style': None,
        'font_color_style': None,
        'geometry': None,
        'text_properties': {}
    }
    
    # Get colors from direct attributes
    fill_color_attr = cell.attrib.get("fillColor")
    stroke_color_attr = cell.attrib.get("strokeColor")
    font_color_attr = cell.attrib.get("fontColor")
    
    if fill_color_attr:
        result['fill_color_attr'] = parse_color(fill_color_attr)
    if stroke_color_attr:
        result['stroke_color_attr'] = parse_color(stroke_color_attr)
    if font_color_attr:
        result['font_color_attr'] = parse_color(font_color_attr)
    
    # Get colors from style attribute
    style = cell.attrib.get("style", "")
    if style:
        fill_color_style = extract_style_value(style, "fillColor")
        stroke_color_style = extract_style_value(style, "strokeColor")
        font_color_style = extract_style_value(style, "fontColor")
        
        if fill_color_style:
            result['fill_color_style'] = parse_color(fill_color_style)
        if stroke_color_style:
            result['stroke_color_style'] = parse_color(stroke_color_style)
        if font_color_style:
            result['font_color_style'] = parse_color(font_color_style)
        
        # Text properties
        result['text_properties'] = {
            'fontSize': extract_style_value(style, "fontSize"),
            'fontStyle': extract_style_value(style, "fontStyle"),
            'align': extract_style_value(style, "align"),
            'verticalAlign': extract_style_value(style, "verticalAlign"),
            'spacingTop': extract_style_value(style, "spacingTop"),
            'spacingLeft': extract_style_value(style, "spacingLeft"),
            'spacingBottom': extract_style_value(style, "spacingBottom"),
            'spacingRight': extract_style_value(style, "spacingRight"),
        }
    
    # Geometry information
    geo = cell.find(".//mxGeometry")
    if geo is not None:
        result['geometry'] = {
            'x': geo.attrib.get("x", "0"),
            'y': geo.attrib.get("y", "0"),
            'width': geo.attrib.get("width", "0"),
            'height': geo.attrib.get("height", "0"),
        }
    
    return result


class TestDrawIOAnalysis:
    """Test class for drawio file analysis"""
    
    @pytest.fixture
    def sample_drawio_path(self):
        """Path to sample drawio file"""
        path = Path(__file__).parent.parent / "sample" / "sample.drawio"
        if not path.exists():
            pytest.skip(f"Sample file not found: {path}")
        return path
    
    def test_parse_color_hex(self):
        """Test hexadecimal color parsing"""
        assert parse_color("#FF0000") == "RGB(255, 0, 0) / #FF0000"
        assert parse_color("#00FF00") == "RGB(0, 255, 0) / #00FF00"
        assert parse_color("#0000FF") == "RGB(0, 0, 255) / #0000FF"
        assert parse_color("#F00") == "RGB(255, 0, 0) / #F00"
        assert parse_color("") == "None"
        assert parse_color(None) == "None"
    
    def test_parse_color_rgb(self):
        """Test RGB format color parsing"""
        assert parse_color("rgb(255, 0, 0)") == "RGB(255, 0, 0)"
        assert parse_color("rgb(0, 255, 0)") == "RGB(0, 255, 0)"
        assert parse_color("rgb(0, 0, 255)") == "RGB(0, 0, 255)"
    
    def test_extract_style_value(self):
        """Test value extraction from style string"""
        style = "fillColor=#FF0000;strokeColor=#0000FF;fontSize=12"
        assert extract_style_value(style, "fillColor") == "#FF0000"
        assert extract_style_value(style, "strokeColor") == "#0000FF"
        assert extract_style_value(style, "fontSize") == "12"
        assert extract_style_value(style, "nonexistent") is None
        assert extract_style_value("", "fillColor") is None
    
    def test_parse_drawio_file(self, sample_drawio_path):
        """Test drawio file parsing"""
        diagrams = parse_drawio_file(sample_drawio_path)
        assert len(diagrams) > 0, "At least one diagram must exist"
        
        # Verify each diagram contains mxGraphModel or root element
        for diagram in diagrams:
            assert diagram is not None
    
    def test_analyze_mxcell_structure(self, sample_drawio_path):
        """Test mxCell element structure analysis"""
        diagrams = parse_drawio_file(sample_drawio_path)
        assert len(diagrams) > 0
        
        # Get mxCell elements from first diagram
        cells = diagrams[0].findall(".//mxCell")
        vertex_cells = [c for c in cells if c.attrib.get("vertex") == "1"]
        
        if len(vertex_cells) > 0:
            # Analyze first vertex cell
            result = analyze_mxcell(vertex_cells[0])
            
            # Verify basic structure exists
            assert 'id' in result
            assert 'vertex' in result
            assert 'geometry' in result or result['geometry'] is None
            assert 'text_properties' in result
    
    def test_drawio_file_has_cells(self, sample_drawio_path):
        """Verify drawio file contains mxCell elements"""
        diagrams = parse_drawio_file(sample_drawio_path)
        assert len(diagrams) > 0
        
        total_cells = 0
        vertex_cells = 0
        
        for diagram in diagrams:
            cells = diagram.findall(".//mxCell")
            total_cells += len(cells)
            vertex_cells += len([c for c in cells if c.attrib.get("vertex") == "1"])
        
        assert total_cells > 0, "At least one mxCell element must exist"
    
    def test_drawio_loader_integration(self, sample_drawio_path):
        """Integration test with drawio2pptx DrawIOLoader"""
        from drawio2pptx.io.drawio_loader import DrawIOLoader
        
        loader = DrawIOLoader()
        diagrams = loader.load_file(sample_drawio_path)
        
        assert len(diagrams) > 0, "DrawIOLoader must return at least one diagram"




