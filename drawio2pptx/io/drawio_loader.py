"""
draw.io file loading and parsing module

Provides loading of .drawio/.xml/.mxfile files, page selection, layer extraction,
parsing of mxGraphModel → mxCell (vertex/edge), and style string parsing
"""
import re
from pathlib import Path
from typing import List, Optional, Dict, Any
from lxml import etree as ET
from lxml import html as lxml_html
from pptx.dml.color import RGBColor  # type: ignore[import]

from ..model.intermediate import (
    ShapeElement, ConnectorElement, BaseElement, TextElement, PolygonElement,
    Transform, Style, TextParagraph, TextRun
)
from ..logger import ConversionLogger
from ..fonts import DRAWIO_DEFAULT_FONT_FAMILY
from ..config import PARALLELOGRAM_SKEW, ConversionConfig, default_config


class ColorParser:
    """Convert draw.io color strings to RGBColor"""
    
    @staticmethod
    def parse(color_str: Optional[str]) -> Optional[RGBColor]:
        """
        Convert draw.io color string to RGBColor
        
        Args:
            color_str: Color string (#RRGGBB, #RGB, rgb(r,g,b), light-dark(...), etc.)
        
        Returns:
            RGBColor object, or None
        """
        if not color_str:
            return None
        
        color_str = color_str.strip()
        
        # Process light-dark(color1,color2) format (use light mode color)
        light_dark_match = re.match(r'^light-dark\s*\((.*)\)$', color_str)
        if light_dark_match:
            inner = light_dark_match.group(1)
            # Split by comma (ignore commas inside parentheses)
            parts = []
            depth = 0
            start = 0
            for i, char in enumerate(inner):
                if char == '(':
                    depth += 1
                elif char == ')':
                    depth -= 1
                elif char == ',' and depth == 0:
                    parts.append(inner[start:i].strip())
                    start = i + 1
            parts.append(inner[start:].strip())
            
            if len(parts) >= 1:
                # Use light mode color (first argument)
                light_color = parts[0]
                return ColorParser.parse(light_color)
        
        # Return None if "none"
        if color_str.lower() == "none":
            return None
        
        # Hexadecimal format (#RRGGBB or #RGB)
        hex_match = re.match(r'^#([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$', color_str)
        if hex_match:
            hex_val = hex_match.group(1)
            if len(hex_val) == 3:
                # Expand short form (#RGB)
                r = int(hex_val[0] * 2, 16)
                g = int(hex_val[1] * 2, 16)
                b = int(hex_val[2] * 2, 16)
            else:
                r = int(hex_val[0:2], 16)
                g = int(hex_val[2:4], 16)
                b = int(hex_val[4:6], 16)
            return RGBColor(r, g, b)
        
        # rgb(r, g, b) format
        rgb_match = re.match(r'^rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$', color_str)
        if rgb_match:
            r = int(rgb_match.group(1))
            g = int(rgb_match.group(2))
            b = int(rgb_match.group(3))
            return RGBColor(r, g, b)
        
        return None


class StyleExtractor:
    """Extract style properties from mxCell elements"""
    
    # Mapping dictionary: draw.io shape type -> normalized shape type
    _SHAPE_TYPE_MAP: dict[str, str] = {
        # Basic shapes
        'rect': 'rectangle',
        'rectangle': 'rectangle',
        'square': 'rectangle',
        'ellipse': 'ellipse',
        'circle': 'ellipse',
        'line': 'line',
        # mxgraph.basic shapes
        'mxgraph.basic.pentagon': 'pentagon',
        'mxgraph.basic.octagon2': 'octagon',
        'mxgraph.basic.acute_triangle': 'isosceles_triangle',
        'mxgraph.basic.orthogonal_triangle': 'right_triangle',
        'mxgraph.basic.4_point_star_2': '4_point_star',
        'mxgraph.basic.star': '5_point_star',
        'mxgraph.basic.6_point_star': '6_point_star',
        'mxgraph.basic.8_point_star': '8_point_star',
        'mxgraph.basic.smiley': 'smiley',
        # mxgraph.flowchart shapes
        'mxgraph.flowchart.decision': 'decision',
        'mxgraph.flowchart.data': 'data',
        'mxgraph.flowchart.document': 'document',
        'mxgraph.flowchart.process': 'process',
        'mxgraph.flowchart.predefined_process': 'predefinedprocess',
        'mxgraph.flowchart.paper_tape': 'tape',
        'mxgraph.flowchart.manual_input': 'manualinput',
        'mxgraph.flowchart.extract': 'extract',
        'mxgraph.flowchart.merge_or_storage': 'merge',
        # mxgraph.bpmn shapes (gateways are typically diamond-shaped)
        'mxgraph.bpmn.shape': 'rhombus',
    }
    
    # Font style bit flags: bit position -> attribute name
    _FONT_STYLE_BITS: dict[int, str] = {
        0: 'bold',
        1: 'italic',
        2: 'underline',
    }
    
    def __init__(self, color_parser: Optional[ColorParser] = None, logger: Optional[ConversionLogger] = None):
        """
        Args:
            color_parser: ColorParser instance (creates new one if None)
            logger: ConversionLogger instance (optional)
        """
        self.color_parser = color_parser or ColorParser()
        self.logger = logger
    
    def extract_style_value(self, style_str: str, key: str) -> Optional[str]:
        """Extract value for specified key from style string"""
        if not style_str:
            return None
        for part in style_str.split(";"):
            if "=" in part:
                k, v = part.split("=", 1)
                if k.strip() == key:
                    return v.strip()
        return None

    def is_text_style(self, style_str: str) -> bool:
        """Return True when the cell is a draw.io text shape."""
        if not style_str:
            return False
        parts = [p.strip().lower() for p in style_str.split(";") if p.strip()]
        if parts and parts[0] == "text":
            return True
        shape_type = self.extract_style_value(style_str, "shape")
        return bool(shape_type and shape_type.strip().lower() == "text")
    
    def extract_style_float(self, style_str: str, key: str, default: Optional[float] = None) -> Optional[float]:
        """Extract float value from style string"""
        value_str = self.extract_style_value(style_str, key)
        if value_str:
            try:
                return float(value_str)
            except ValueError:
                pass
        return default
    
    def _parse_font_style(self, font_style_str: Optional[str]) -> dict[str, bool]:
        """
        Parse font style bit flags
        
        Args:
            font_style_str: Font style string (integer as string)
        
        Returns:
            Dictionary with 'bold', 'italic', 'underline' keys
        """
        result = {'bold': False, 'italic': False, 'underline': False}
        if font_style_str:
            try:
                font_style_int = int(font_style_str) if font_style_str.isdigit() else 0
                for bit_pos, attr_name in self._FONT_STYLE_BITS.items():
                    if (font_style_int & (1 << bit_pos)) != 0:
                        result[attr_name] = True
            except (ValueError, TypeError):
                pass
        return result
    
    def extract_fill_color(self, cell: ET.Element) -> Optional[Any]:
        """
        Extract fillColor
        
        Returns:
            RGBColor object (when color is specified), "default" (when default/auto), None (when transparent)
        """
        # Get from direct attribute
        fill_color = cell.attrib.get("fillColor")
        if fill_color:
            fill_color_lower = fill_color.lower().strip()
            if fill_color_lower in ["default", "auto"]:
                return "default"
            elif fill_color_lower == "none":
                return None
            else:
                parsed = self.color_parser.parse(fill_color)
                if parsed:
                    return parsed
        
        # Get from style attribute
        style = cell.attrib.get("style", "")
        if style:
            for part in style.split(";"):
                if "=" in part:
                    key, value = part.split("=", 1)
                    if key.strip() == "fillColor":
                        value_lower = value.strip().lower()
                        if value_lower in ["default", "auto"]:
                            return "default"
                        elif value_lower == "none":
                            return None
                        else:
                            parsed = self.color_parser.parse(value.strip())
                            if parsed:
                                return parsed
        # Text shapes should be transparent unless fillColor is explicit.
        try:
            if self.is_text_style(style):
                return None
        except Exception as e:
            if self.logger:
                self.logger.debug(f"Failed to check text style for fill: {e}")
        # If fillColor is omitted, draw.io uses a default fill for most vertex shapes.
        # Treat missing fillColor as "default" for vertices, but keep edges transparent.
        try:
            if cell.attrib.get("vertex") == "1" and cell.attrib.get("edge") != "1":
                return "default"
        except Exception as e:
            if self.logger:
                self.logger.debug(f"Failed to check vertex attribute: {e}")

        return None

    def extract_gradient_color(self, cell: ET.Element) -> Optional[Any]:
        """
        Extract gradientColor

        Returns:
            RGBColor object (when color is specified), "default" (when default/auto), None (when no gradient/none)
        """
        # Direct attribute (rare, but keep consistent with fillColor)
        grad_color = cell.attrib.get("gradientColor")
        if grad_color:
            grad_color_lower = grad_color.lower().strip()
            if grad_color_lower in ["default", "auto"]:
                return "default"
            if grad_color_lower == "none":
                return None
            parsed = self.color_parser.parse(grad_color)
            if parsed:
                return parsed

        style = cell.attrib.get("style", "")
        if style:
            for part in style.split(";"):
                if "=" in part:
                    key, value = part.split("=", 1)
                    if key.strip() == "gradientColor":
                        value_lower = value.strip().lower()
                        if value_lower in ["default", "auto"]:
                            return "default"
                        if value_lower == "none":
                            return None
                        parsed = self.color_parser.parse(value.strip())
                        if parsed:
                            return parsed
                        return None

        return None

    def extract_gradient_direction(self, cell: ET.Element) -> Optional[str]:
        """Extract gradientDirection (e.g., north/south/east/west)"""
        style = cell.attrib.get("style", "")
        if not style:
            return None
        value = self.extract_style_value(style, "gradientDirection")
        return value.strip() if value else None

    def extract_swimlane_fill_color(self, cell: ET.Element) -> Optional[Any]:
        """
        Extract swimlaneFillColor (body area fill). Header uses fillColor.
        Returns: RGBColor, "default", or None (treated as white).
        """
        style = cell.attrib.get("style", "")
        if not style:
            return None
        value = self.extract_style_value(style, "swimlaneFillColor")
        if not value:
            return None
        value_lower = value.strip().lower()
        if value_lower in ["default", "auto"]:
            return "default"
        if value_lower == "none":
            return None
        parsed = self.color_parser.parse(value.strip())
        return parsed
    
    def extract_stroke_color(self, cell: ET.Element) -> Optional[RGBColor]:
        """Extract strokeColor"""
        stroke_color = cell.attrib.get("strokeColor")
        if stroke_color:
            parsed = self.color_parser.parse(stroke_color)
            if parsed:
                return parsed
        
        style = cell.attrib.get("style", "")
        if style:
            for part in style.split(";"):
                if "=" in part:
                    key, value = part.split("=", 1)
                    if key.strip() == "strokeColor":
                        parsed = self.color_parser.parse(value.strip())
                        if parsed:
                            return parsed
        
        return None

    def extract_no_stroke(self, cell: ET.Element) -> bool:
        """Detect strokeColor=none explicitly set on the cell."""
        stroke_color = (cell.attrib.get("strokeColor") or "").strip().lower()
        if stroke_color == "none":
            return True
        style = cell.attrib.get("style", "")
        if style:
            for part in style.split(";"):
                if "=" in part:
                    key, value = part.split("=", 1)
                    if key.strip() == "strokeColor" and value.strip().lower() == "none":
                        return True
        return False
    
    def extract_font_color(self, cell: ET.Element) -> Optional[RGBColor]:
        """Extract fontColor (also checks inside HTML tags)"""
        font_color = cell.attrib.get("fontColor")
        if font_color:
            parsed = self.color_parser.parse(font_color)
            if parsed:
                return parsed
        
        style = cell.attrib.get("style", "")
        if style:
            for part in style.split(";"):
                if "=" in part:
                    key, value = part.split("=", 1)
                    if key.strip() == "fontColor":
                        parsed = self.color_parser.parse(value.strip())
                        if parsed:
                            return parsed
        
        # Get from style attribute in HTML tags (for Square/Circle support)
        value = cell.attrib.get("value", "")
        if value and "<font" in value:
            try:
                wrapped = f"<div>{value}</div>"
                parsed = lxml_html.fromstring(wrapped)
                font_tags = parsed.findall(".//font")
                for font_tag in font_tags:
                    font_style = font_tag.get("style", "")
                    if font_style and "color:" in font_style:
                        color_match = re.search(r'color:\s*([^;]+)', font_style)
                        if color_match:
                            color_value = color_match.group(1).strip()
                            parsed_color = self.color_parser.parse(color_value)
                            if parsed_color:
                                return parsed_color
            except Exception as e:
                if self.logger:
                    self.logger.debug(f"Failed to parse font color from HTML: {e}")
        
        return None

    def extract_label_background_color(self, cell: ET.Element) -> Optional[RGBColor]:
        """Extract labelBackgroundColor (draw.io label background color; equivalent to highlight)"""
        label_bg = cell.attrib.get("labelBackgroundColor")
        if label_bg:
            parsed = self.color_parser.parse(label_bg)
            if parsed:
                return parsed

        style = cell.attrib.get("style", "")
        if style:
            for part in style.split(";"):
                if "=" in part:
                    key, value = part.split("=", 1)
                    if key.strip() == "labelBackgroundColor":
                        parsed = self.color_parser.parse(value.strip())
                        if parsed:
                            return parsed
                        return None
        return None
    
    def extract_shadow(self, cell: ET.Element, mgm_root: Optional[ET.Element]) -> bool:
        """Extract shadow setting"""
        style = cell.attrib.get("style", "")
        if style:
            for part in style.split(";"):
                if "=" in part:
                    key, value = part.split("=", 1)
                    if key.strip() == "shadow":
                        return value.strip() == "1"
        
        if mgm_root is not None:
            mgm_shadow = mgm_root.attrib.get("shadow")
            if mgm_shadow == "1":
                return True
        
        return False
    
    def extract_shape_type(self, cell: ET.Element) -> str:
        """Extract and normalize shape type"""
        style = cell.attrib.get("style", "")

        # Prefer explicit "shape=..." when present.
        # draw.io sometimes emits e.g. "ellipse;shape=cloud;..." where the first token is generic.
        shape_type = self.extract_style_value(style, "shape")
        if shape_type:
            shape_type = shape_type.lower()
            if shape_type == "swimlane":
                return "swimlane"
            # draw.io flowchart: "Predefined process" is often represented as shape=process with backgroundOutline=1.
            # Map it to a dedicated pseudo-type so we can use PowerPoint's predefined-process shape.
            if shape_type == "process":
                try:
                    bg_outline = self.extract_style_value(style, "backgroundOutline")
                    if (bg_outline or "").strip() == "1":
                        return "predefinedprocess"
                    # Some diagrams.net exports omit backgroundOutline for predefined process,
                    # but keep a non-zero "size" parameter. Treat it as predefined process
                    # to better match the expected appearance in PowerPoint.
                    size_value = self.extract_style_value(style, "size")
                    if size_value is not None:
                        try:
                            if float(size_value) > 0:
                                return "predefinedprocess"
                        except ValueError:
                            pass
                except Exception as e:
                    if self.logger:
                        self.logger.debug(f"Failed to check backgroundOutline: {e}")
            # Use dictionary mapping
            normalized = self._SHAPE_TYPE_MAP.get(shape_type)
            if normalized:
                return normalized
            # Keep rhombus/parallelogram/cloud/trapezoid/etc. as-is.
            return shape_type

        if style:
            parts = style.split(";")
            first_part = (parts[0].strip().lower() if parts else "")
            if first_part == "swimlane":
                return "swimlane"
            # Use dictionary mapping
            normalized = self._SHAPE_TYPE_MAP.get(first_part)
            if normalized:
                return normalized
            # Keep rhombus/process as-is (process not in _SHAPE_TYPE_MAP key; value is mxgraph.flowchart.process)
            if first_part == "rhombus":
                return "rhombus"
            if first_part == "process":
                return "process"

        return "rectangle"


class DrawIOLoader:
    """draw.io file loading and parsing"""
    
    def __init__(self, logger: Optional[ConversionLogger] = None, config: Optional[ConversionConfig] = None):
        """
        Args:
            logger: ConversionLogger instance
            config: ConversionConfig instance (uses default_config if None)
        """
        self.config = config or default_config
        self.logger = logger
        self.color_parser = ColorParser()
        self.style_extractor = StyleExtractor(self.color_parser, logger)
    
    def load_file(self, path: Path) -> List[ET.Element]:
        """
        Load draw.io file and return list of diagrams
        
        Args:
            path: File path
        
        Returns:
            List of mxGraphModel elements (corresponding to each diagram)
        """
        tree = ET.parse(path)
        root = tree.getroot()
        
        diagrams = []
        # Process <diagram> elements
        for d in root.findall(".//diagram"):
            inner = (d.text or "").strip()
            if not inner:
                mgm = d.find(".//mxGraphModel")
                if mgm is not None:
                    diagrams.append(mgm)
                    continue
                continue
            
            # Unescape HTML entities
            if "&lt;" in inner or "&amp;" in inner:
                try:
                    wrapped = f"<div>{inner}</div>"
                    parsed = lxml_html.fromstring(wrapped)
                    inner = parsed.text_content()
                except Exception as e:
                    if self.logger:
                        self.logger.debug(f"Failed to unescape HTML entities: {e}")
            
            # Parse as XML fragment
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
                    diagrams.append(parsed)
                    continue
                except ET.ParseError:
                    pass
            
            # Fallback
            mgm_global = root.find(".//mxGraphModel")
            if mgm_global is not None:
                diagrams.append(mgm_global)
            else:
                diagrams.append(root)
        
        # If <diagram> tag is not present
        if not diagrams:
            mgm_global = root.find(".//mxGraphModel")
            if mgm_global is not None:
                diagrams.append(mgm_global)
            else:
                diagrams.append(root)
        
        return diagrams
    
    def extract_page_size(self, mgm_root: ET.Element) -> tuple:
        """
        Extract page size
        
        Returns:
            (width, height) tuple (px), or (None, None)
        """
        page_width = mgm_root.attrib.get("pageWidth")
        page_height = mgm_root.attrib.get("pageHeight")
        
        if page_width and page_height:
            try:
                width = float(page_width)
                height = float(page_height)
                return (width, height)
            except ValueError:
                pass
        
        return (None, None)
    
    def extract_elements(self, mgm_root: ET.Element) -> List[BaseElement]:
        """
        Extract elements from mxGraphModel and convert to intermediate model
        
        Args:
            mgm_root: mxGraphModel element
        
        Returns:
            List of elements
        """
        elements = []

        # Preserve draw.io stacking order.
        #
        # In mxGraphModel, the relative stacking / draw order is primarily encoded by the order of
        # mxCell nodes under <root>. PowerPoint stacking is determined by the order of shapes in
        # the slide XML (later = topmost), so we map mxCell order to BaseElement.z_index.
        #
        # NOTE: We still extract shapes first to build shapes_dict for connector routing; z_index
        #       sorting will restore the original order afterwards.
        cells = list(mgm_root.findall(".//mxCell"))
        cell_by_id: Dict[str, ET.Element] = {}
        document_order: Dict[str, int] = {}
        children_by_parent: Dict[str, List[str]] = {}

        # Build cell index and parent->children order based on document order.
        for idx, cell in enumerate(cells):
            cid = cell.attrib.get("id")
            if cid is None:
                continue
            if cid not in cell_by_id:
                cell_by_id[cid] = cell
            if cid not in document_order:
                document_order[cid] = idx
            parent_id = cell.attrib.get("parent")
            if parent_id:
                children_by_parent.setdefault(parent_id, []).append(cid)

        # Identify container-like vertices (parents of other cells).
        parent_ids = set()
        for cell in cells:
            parent_id = cell.attrib.get("parent")
            if parent_id and parent_id not in ("0", "1"):
                parent_ids.add(parent_id)
        container_vertex_ids = set()
        for pid in parent_ids:
            pcell = cell_by_id.get(pid)
            if pcell is None:
                continue
            if pcell.attrib.get("vertex") == "1" and pcell.attrib.get("edge") != "1":
                container_vertex_ids.add(pid)

        # Derive draw order by traversing the parent-child hierarchy.
        # In mxGraph, the order within a parent's children list defines z-order,
        # and children are drawn above their parent.
        draw_order_ids: List[str] = []
        visited: set[str] = set()

        def _append_draw_order(parent_id: str) -> None:
            for cid in children_by_parent.get(parent_id, []):
                if cid in visited:
                    continue
                visited.add(cid)
                cell = cell_by_id.get(cid)
                if cell is None:
                    continue
                if cell.attrib.get("vertex") == "1" or cell.attrib.get("edge") == "1":
                    draw_order_ids.append(cid)
                _append_draw_order(cid)

        # Start from the root layer(s). Default layer is id="1" under root id="0".
        if "0" in children_by_parent:
            _append_draw_order("0")
        elif "1" in children_by_parent:
            _append_draw_order("1")

        # Fallback to document order if traversal yielded nothing.
        if not draw_order_ids:
            for cid, _ in sorted(document_order.items(), key=lambda x: x[1]):
                cell = cell_by_id.get(cid)
                if cell is None:
                    continue
                if cell.attrib.get("vertex") == "1" or cell.attrib.get("edge") == "1":
                    draw_order_ids.append(cid)

        cell_order: Dict[str, int] = {cid: idx for idx, cid in enumerate(draw_order_ids)}
        
        # First extract shapes
        shapes_dict = {}
        for cell in cells:
            if cell.attrib.get("vertex") == "1":
                shape = self._extract_shape(cell, mgm_root)
                if shape:
                    try:
                        if shape.id is not None:
                            shape.z_index = cell_order.get(shape.id, 0)
                            # Containers should stay behind their contents and connectors.
                            if shape.id in container_vertex_ids:
                                shape.z_index -= 100000
                    except Exception as e:
                        if self.logger:
                            self.logger.debug(f"Failed to set z_index for shape {shape.id}: {e}")
                    elements.append(shape)
                    if shape.id:
                        shapes_dict[shape.id] = shape
        
        # Then extract edges
        for cell in cells:
            if cell.attrib.get("edge") == "1":
                connector, labels = self._extract_connector(cell, mgm_root, shapes_dict)
                if connector:
                    try:
                        if connector.id is not None:
                            connector.z_index = cell_order.get(connector.id, 0)
                            # Ensure connectors are drawn above their endpoints.
                            # Some draw.io exports place edges before vertices; PPTX would then hide arrowheads.
                            try:
                                src_z = shapes_dict.get(connector.source_id).z_index if connector.source_id in shapes_dict else None
                                tgt_z = shapes_dict.get(connector.target_id).z_index if connector.target_id in shapes_dict else None
                                max_z = max(z for z in (src_z, tgt_z) if z is not None)
                                if connector.z_index <= max_z:
                                    connector.z_index = max_z + 1
                            except Exception:
                                pass
                    except Exception as e:
                        if self.logger:
                            self.logger.debug(f"Failed to set z_index for connector {connector.id}: {e}")
                    elements.append(connector)
                    if labels:
                        for label in labels:
                            # Keep label above the connector line.
                            label.z_index = connector.z_index
                            elements.append(label)
        
        # If the entire diagram is off the page, normalize to the page origin.
        self._maybe_normalize_page_offset(elements, mgm_root)

        # Sort by Z-order
        elements.sort(key=lambda e: e.z_index)
        
        return elements

    def _maybe_normalize_page_offset(self, elements: List[BaseElement], mgm_root: ET.Element) -> None:
        """Shift diagram to page origin when content is fully outside the page."""
        if not elements:
            return

        page_width, page_height = self.extract_page_size(mgm_root)
        if page_width is None or page_height is None:
            return

        bounds = self._get_elements_bounds(elements)
        if bounds is None:
            return

        min_x, min_y, max_x, max_y = bounds
        # Only shift when the entire diagram is outside the page rectangle.
        if max_x < 0 or max_y < 0 or min_x > page_width or min_y > page_height:
            width = max_x - min_x
            height = max_y - min_y
            if width <= page_width:
                offset_x = (page_width - width) / 2.0 - min_x
            else:
                offset_x = -min_x
            if height <= page_height:
                offset_y = (page_height - height) / 2.0 - min_y
            else:
                offset_y = -min_y
            self._apply_translation(elements, offset_x, offset_y)
            if self.logger:
                self.logger.debug(
                    f"Applied page offset normalization: ({offset_x:.2f}, {offset_y:.2f})"
                )

    def _get_elements_bounds(self, elements: List[BaseElement]) -> Optional[tuple[float, float, float, float]]:
        """Compute bounding box for a list of elements."""
        min_x = float("inf")
        min_y = float("inf")
        max_x = float("-inf")
        max_y = float("-inf")

        for element in elements:
            bounds = self._get_element_bounds(element)
            if bounds is None:
                continue
            e_min_x, e_min_y, e_max_x, e_max_y = bounds
            min_x = min(min_x, e_min_x)
            min_y = min(min_y, e_min_y)
            max_x = max(max_x, e_max_x)
            max_y = max(max_y, e_max_y)

        if min_x == float("inf"):
            return None

        return (min_x, min_y, max_x, max_y)

    def _get_element_bounds(self, element: BaseElement) -> Optional[tuple[float, float, float, float]]:
        """Get bounding box for a single element."""
        if isinstance(element, (ConnectorElement, PolygonElement)) and element.points:
            xs = [p[0] for p in element.points]
            ys = [p[1] for p in element.points]
            return (min(xs), min(ys), max(xs), max(ys))

        x0, y0 = element.x, element.y
        x1, y1 = x0 + element.w, y0 + element.h
        return (min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))

    def _apply_translation(self, elements: List[BaseElement], dx: float, dy: float) -> None:
        """Apply translation to element coordinates and point lists."""
        if dx == 0 and dy == 0:
            return

        for element in elements:
            element.x += dx
            element.y += dy
            if isinstance(element, (ConnectorElement, PolygonElement)) and element.points:
                element.points = [(x + dx, y + dy) for x, y in element.points]
    
    def _get_parent_coordinates(self, parent_id: str, mgm_root: ET.Element) -> tuple[float, float]:
        """
        Get parent element's coordinates recursively
        
        Args:
            parent_id: Parent element ID
            mgm_root: mxGraphModel root element
        
        Returns:
            (parent_x, parent_y) tuple (accumulated coordinates from all ancestors)
        """
        if not parent_id or parent_id in ("0", "1"):
            # Root elements (0 or 1) have no coordinates
            return (0.0, 0.0)
        
        # Find parent cell
        parent_cell = None
        for cell in mgm_root.findall(".//mxCell"):
            if cell.attrib.get("id") == parent_id:
                parent_cell = cell
                break
        
        if parent_cell is None:
            return (0.0, 0.0)
        
        # Get parent's geometry
        parent_geo = parent_cell.find(".//mxGeometry")
        if parent_geo is None:
            return (0.0, 0.0)
        
        try:
            parent_x = float(parent_geo.attrib.get("x", "0") or 0)
            parent_y = float(parent_geo.attrib.get("y", "0") or 0)
        except ValueError:
            return (0.0, 0.0)
        
        # Recursively get grandparent coordinates
        grandparent_id = parent_cell.attrib.get("parent")
        if grandparent_id:
            grandparent_x, grandparent_y = self._get_parent_coordinates(grandparent_id, mgm_root)
            parent_x += grandparent_x
            parent_y += grandparent_y
        
        return (parent_x, parent_y)
    
    def _extract_shape(self, cell: ET.Element, mgm_root: ET.Element) -> Optional[ShapeElement]:
        """Extract shape"""
        geo = cell.find(".//mxGeometry")
        if geo is None:
            return None
        
        try:
            x = float(geo.attrib.get("x", "0") or 0)
            y = float(geo.attrib.get("y", "0") or 0)
            w = float(geo.attrib.get("width", "0") or 0)
            h = float(geo.attrib.get("height", "0") or 0)
        except ValueError:
            return None
        
        # Add parent coordinates if parent exists
        parent_id = cell.attrib.get("parent")
        if parent_id and parent_id not in ("0", "1"):
            # Handle swimlane header offsets only when the child coordinates
            # are relative to the content area (not including header).
            try:
                parent_cell = None
                for pcell in mgm_root.findall(".//mxCell"):
                    if pcell.attrib.get("id") == parent_id:
                        parent_cell = pcell
                        break
                if parent_cell is not None:
                    parent_style = parent_cell.attrib.get("style", "") or ""
                    if parent_style and "swimlane" in parent_style:
                        start_size = self.style_extractor.extract_style_float(parent_style, "startSize", 0.0) or 0.0
                        horizontal_str = self.style_extractor.extract_style_value(parent_style, "horizontal")
                        if horizontal_str is None:
                            # draw.io default for swimlane is horizontal=1 (header on top).
                            is_horizontal = True
                        else:
                            is_horizontal = horizontal_str.strip() == "1"
                        eps = 1e-6
                        if is_horizontal:
                            # Header on top: only shift when child y is inside header coords.
                            if y < start_size - eps:
                                y += start_size
                        else:
                            # Header on left: only shift when child x is inside header coords.
                            if x < start_size - eps:
                                x += start_size
            except Exception as e:
                if self.logger:
                    self.logger.debug(f"Failed to adjust swimlane child offset: {e}")

            parent_x, parent_y = self._get_parent_coordinates(parent_id, mgm_root)
            x += parent_x
            y += parent_y
        
        # Basic information
        shape_id = cell.attrib.get("id")
        text_raw = cell.attrib.get("value", "") or ""
        
        # Extract style
        fill_color = self.style_extractor.extract_fill_color(cell)
        gradient_color = self.style_extractor.extract_gradient_color(cell)
        gradient_direction = self.style_extractor.extract_gradient_direction(cell)
        stroke_color = self.style_extractor.extract_stroke_color(cell)
        font_color = self.style_extractor.extract_font_color(cell)
        label_bg_color = self.style_extractor.extract_label_background_color(cell)
        has_shadow = self.style_extractor.extract_shadow(cell, mgm_root)

        style_str = cell.attrib.get("style", "")

        # Extract BPMN symbol attribute before shape_type normalization.
        original_shape = self.style_extractor.extract_style_value(style_str, "shape")
        bpmn_symbol = None
        if original_shape and "mxgraph.bpmn.shape" in original_shape.lower():
            bpmn_symbol = self.style_extractor.extract_style_value(style_str, "symbol")

        shape_type = self.style_extractor.extract_shape_type(cell)
        stroke_width = self.style_extractor.extract_style_float(style_str, "strokeWidth", 1.0)
        is_text_style = self.style_extractor.is_text_style(style_str)
        no_stroke = is_text_style or self.style_extractor.extract_no_stroke(cell)
        
        # Extract whiteSpace (text wrapping) setting
        # draw.io: whiteSpace=wrap (default) enables text wrapping, whiteSpace=nowrap disables it
        white_space = self.style_extractor.extract_style_value(style_str, "whiteSpace")
        word_wrap = True  # Default is wrap (draw.io's default)
        if white_space and white_space.lower() == "nowrap":
            word_wrap = False
        # For draw.io text shapes, default behavior is no wrapping unless explicitly set.
        if white_space is None and is_text_style:
            word_wrap = False
        
        # Extract rounded attribute (corner radius)
        # draw.io: rounded=1 enables corner radius, rounded=0 disables it
        corner_radius = None
        rounded_str = self.style_extractor.extract_style_value(style_str, "rounded")
        if rounded_str and rounded_str.strip() == "1":
            # Calculate default corner radius (approximately 10% of min dimension)
            if min(w, h) > 0:
                corner_radius = min(w, h) * 0.1
        
        # Extract step size (draw.io style: size) for step shapes
        step_size = None
        try:
            if shape_type and shape_type.lower() == "step":
                step_size = self.style_extractor.extract_style_float(style_str, "size")
        except Exception as e:
            if self.logger:
                self.logger.debug(f"Failed to extract step size: {e}")

        # Extract text
        text_paragraphs = self._extract_text(text_raw, font_color, style_str)
        
        # Create style
        style = Style(
            fill=fill_color,
            gradient_color=gradient_color,
            gradient_direction=gradient_direction,
            stroke=stroke_color,
            stroke_width=stroke_width,
            opacity=1.0,
            corner_radius=corner_radius,
            label_background_color=label_bg_color,
            has_shadow=has_shadow,
            word_wrap=word_wrap,
            no_stroke=no_stroke,
            bpmn_symbol=bpmn_symbol,
            step_size_px=step_size,
        )

        # Swimlane/container metadata (used for header layout in PPTX)
        try:
            if shape_type and shape_type.lower() == "swimlane":
                style.is_swimlane = True
                start_size = self.style_extractor.extract_style_float(style_str, "startSize", 0.0)
                style.swimlane_start_size = float(start_size or 0.0)
                horizontal_str = self.style_extractor.extract_style_value(style_str, "horizontal")
                if horizontal_str is None:
                    # draw.io default for swimlane is horizontal=1 (header on top).
                    style.swimlane_horizontal = True
                else:
                    style.swimlane_horizontal = horizontal_str.strip() == "1"
                swimlane_line_str = self.style_extractor.extract_style_value(style_str, "swimlaneLine")
                if swimlane_line_str is None:
                    style.swimlane_line = True
                else:
                    style.swimlane_line = swimlane_line_str.strip() != "0"
                style.swimlane_fill_color = self.style_extractor.extract_swimlane_fill_color(cell)
        except Exception as e:
            if self.logger:
                self.logger.debug(f"Failed to extract swimlane metadata: {e}")
        
        # Create ShapeElement
        shape = ShapeElement(
            id=shape_id,
            x=x,
            y=y,
            w=w,
            h=h,
            shape_type=shape_type,
            text=text_paragraphs,
            style=style,
            z_index=0  # TODO: extract z-index
        )
        
        return shape
    
    def _extract_connector(self, cell: ET.Element, mgm_root: ET.Element, shapes_dict: Dict[str, ShapeElement]) -> tuple[Optional[ConnectorElement], List[TextElement]]:
        """Extract connector and its labels (if present)"""
        source_id = cell.attrib.get("source")
        target_id = cell.attrib.get("target")
        
        source_shape = shapes_dict.get(source_id) if source_id else None
        target_shape = shapes_dict.get(target_id) if target_id else None
        
        connector_id = cell.attrib.get("id")
        style_str = cell.attrib.get("style", "")
        
        # Extract style
        stroke_color = self.style_extractor.extract_stroke_color(cell)
        stroke_width = self.style_extractor.extract_style_float(style_str, "strokeWidth", 1.0)
        has_shadow = self.style_extractor.extract_shadow(cell, mgm_root)
        font_color = self.style_extractor.extract_font_color(cell)
        label_bg_color = self.style_extractor.extract_label_background_color(cell)
        
        # Edge style
        edge_style = "straight"
        edge_style_str = self.style_extractor.extract_style_value(style_str, "edgeStyle")
        edge_style_lower = edge_style_str.lower() if edge_style_str else ""
        is_elbow_edge = False
        if edge_style_str:
            if "orthogonal" in edge_style_lower or "elbow" in edge_style_lower:
                edge_style = "orthogonal"
            elif "curved" in edge_style_lower:
                edge_style = "curved"
        if "elbow" in edge_style_lower:
            is_elbow_edge = True
        
        # Arrow settings
        start_arrow = self.style_extractor.extract_style_value(style_str, "startArrow")
        end_arrow = self.style_extractor.extract_style_value(style_str, "endArrow")
        start_fill = self.style_extractor.extract_style_value(style_str, "startFill") != "0"
        end_fill = self.style_extractor.extract_style_value(style_str, "endFill") != "0"
        start_size = self.style_extractor.extract_style_float(style_str, "startSize", None)
        end_size = self.style_extractor.extract_style_float(style_str, "endSize", None)
        
        # Dash pattern
        # draw.io uses dashed=1 for dashed lines, or dashed=<pattern> for specific patterns
        dash_pattern = None
        dashed_value = self.style_extractor.extract_style_value(style_str, "dashed")
        if dashed_value:
            if dashed_value == "1" or dashed_value == "true":
                # Default dashed pattern
                dash_pattern = "dashed"
            else:
                # Specific dash pattern (e.g., "dotted", "dashDot", etc.)
                dash_pattern = dashed_value

        # draw.io can omit default arrow styles from the edge `style` string.
        # In diagrams.net, a newly created connector typically has an arrow at the target side by default.
        # When the mxGraphModel sets arrows="1", treat missing endArrow as the default "classic".
        if end_arrow is None:
            try:
                if mgm_root is not None and mgm_root.attrib.get("arrows") == "1":
                    end_arrow = "classic"
            except Exception as e:
                if self.logger:
                    self.logger.debug(f"Failed to check arrows attribute: {e}")
        
        # Extract points (execute first for automatic port determination)
        geo = cell.find(".//mxGeometry")
        points_raw = []
        points_for_ports = []
        points_raw_offset_flags = []
        source_point = None
        target_point = None
        if geo is not None:
            # Find Array[@as="points"] (drawio's waypoints are stored here)
            array_elem = geo.find('./Array[@as="points"]')
            if array_elem is not None:
                # Find mxPoint as direct child of Array
                for point_elem in array_elem.findall("./mxPoint"):
                    px = float(point_elem.attrib.get("x", "0") or 0)
                    py = float(point_elem.attrib.get("y", "0") or 0)
                    points_raw.append((px, py))
                    points_for_ports.append((px, py))
                    points_raw_offset_flags.append(False)
            else:
                # If Array is not present, find mxPoint as direct child of mxGeometry.
                # Ignore label offsets and source/target hint points; they are not waypoints.
                for point_elem in geo.findall("./mxPoint"):
                    role = (point_elem.attrib.get("as") or "").strip()
                    if role == "sourcePoint":
                        px = float(point_elem.attrib.get("x", "0") or 0)
                        py = float(point_elem.attrib.get("y", "0") or 0)
                        source_point = (px, py)
                        continue
                    if role == "targetPoint":
                        px = float(point_elem.attrib.get("x", "0") or 0)
                        py = float(point_elem.attrib.get("y", "0") or 0)
                        target_point = (px, py)
                        continue
                    if role == "offset":
                        continue
                    px = float(point_elem.attrib.get("x", "0") or 0)
                    py = float(point_elem.attrib.get("y", "0") or 0)
                    points_raw.append((px, py))
                    points_raw_offset_flags.append(False)
                    # Only use waypoint points for port inference.
                    points_for_ports.append((px, py))

        # Waypoints are stored relative to the connector's parent. Align them to the
        # same absolute coordinate space as shapes.
        parent_id = cell.attrib.get("parent")
        if parent_id and parent_id not in ("0", "1") and (points_raw or source_point or target_point):
            parent_x, parent_y = self._get_parent_coordinates(parent_id, mgm_root)
            points_for_ports = [(px + parent_x, py + parent_y) for px, py in points_for_ports]
            if points_raw_offset_flags and len(points_raw_offset_flags) == len(points_raw):
                points_raw = [
                    (px + parent_x, py + parent_y) if not is_offset else (px, py)
                    for (px, py), is_offset in zip(points_raw, points_raw_offset_flags)
                ]
            else:
                points_raw = [(px + parent_x, py + parent_y) for px, py in points_raw]
            if source_point:
                source_point = (source_point[0] + parent_x, source_point[1] + parent_y)
            if target_point:
                target_point = (target_point[0] + parent_x, target_point[1] + parent_y)

        # Floating edges (no source/target shapes) are stored with sourcePoint/targetPoint.
        # Preserve their geometry and arrow styles even without bound shapes.
        if not source_shape or not target_shape:
            points: List[tuple] = []
            if source_point and target_point:
                points = [source_point]
                if points_raw:
                    points.extend(points_raw)
                points.append(target_point)
            elif points_raw and len(points_raw) >= 2:
                points = list(points_raw)
            else:
                return None, []

            style = Style(
                stroke=stroke_color,
                stroke_width=stroke_width,
                dash=dash_pattern,
                arrow_start=start_arrow,
                arrow_end=end_arrow,
                arrow_start_fill=start_fill,
                arrow_end_fill=end_fill,
                arrow_start_size_px=start_size,
                arrow_end_size_px=end_size,
                has_shadow=has_shadow
            )

            connector = ConnectorElement(
                id=connector_id,
                source_id=source_id,
                target_id=target_id,
                points=points,
                edge_style=edge_style,
                style=style
            )

            # Extract label text (edge value)
            labels: List[TextElement] = []
            label_element = self._extract_connector_label(cell, connector, font_color, label_bg_color, style_str)
            if label_element is not None:
                labels.append(label_element)

            # Edge labels can also be stored as child mxCell nodes (edgeLabel style).
            if connector_id and mgm_root is not None:
                try:
                    child_cells = mgm_root.findall(f".//mxCell[@parent='{connector_id}']")
                except Exception:
                    child_cells = []
                for label_cell in child_cells:
                    if label_cell.attrib.get("vertex") != "1":
                        continue
                    style_val = label_cell.attrib.get("style", "") or ""
                    is_edge_label = "edgeLabel" in style_val
                    if not is_edge_label and label_cell.attrib.get("connectable") != "0":
                        continue
                    child_label = self._extract_edge_label_cell(
                        label_cell=label_cell,
                        connector=connector,
                        default_font_color=font_color,
                        default_label_bg_color=label_bg_color,
                        source_shape=source_shape,
                        target_shape=target_shape,
                    )
                    if child_label is not None:
                        labels.append(child_label)

            return connector, labels

        def _infer_port_side(rel_x: Optional[float], rel_y: Optional[float]) -> Optional[str]:
            """
            Infer which side a port points to.

            This is used to decide whether exitX/exitY (or entryX/entryY) contradict
            the orthogonal route implied by waypoints.

            Returns: 'left' | 'right' | 'top' | 'bottom' | None (ambiguous)
            """
            if rel_x is None or rel_y is None:
                return None
            dx = abs(rel_x - 0.5)
            dy = abs(rel_y - 0.5)
            if dx < 1e-9 and dy < 1e-9:
                return None
            if dx >= dy:
                return "right" if rel_x >= 0.5 else "left"
            return "bottom" if rel_y >= 0.5 else "top"

        def _snap_to_grid(val: float, grid_size: Optional[float]) -> float:
            """Snap a coordinate to draw.io grid."""
            if not grid_size:
                return val
            try:
                if grid_size <= 0:
                    return val
                return round(val / grid_size) * grid_size
            except Exception:
                return val

        def _ensure_orthogonal_route_respects_ports(
            pts: List[tuple],
            exit_x_: Optional[float],
            exit_y_: Optional[float],
            entry_x_: Optional[float],
            entry_y_: Optional[float],
            grid_size_: Optional[float],
        ) -> List[tuple]:
            """
            Ensure an orthogonal connector polyline respects the exit/entry side direction.

            When draw.io stores no waypoints for an orthogonal edge, we still must honor port sides:
            - If entry is on left/right, the segment adjacent to the target must be horizontal.
            - If entry is on top/bottom, the segment adjacent to the target must be vertical.

            A single-bend L-shape cannot satisfy both ends when both ends require the same orientation
            (e.g., exit=right and entry=left => both horizontal). In that case, generate 2 bends:
            - H-V-H or V-H-V with a snapped midpoint.
            """
            if not pts or len(pts) != 2:
                return pts
            (sx, sy), (tx, ty) = pts
            if sx == tx or sy == ty:
                return pts

            exit_side = _infer_port_side(exit_x_, exit_y_)
            entry_side = _infer_port_side(entry_x_, entry_y_)

            def _dir_for_side(side: Optional[str]) -> Optional[str]:
                if side in ("left", "right"):
                    return "h"
                if side in ("top", "bottom"):
                    return "v"
                return None

            start_dir = _dir_for_side(exit_side)
            end_dir = _dir_for_side(entry_side)

            # Fallback to legacy heuristic when ambiguous.
            if start_dir is None or end_dir is None:
                dx_ = abs(tx - sx)
                dy_ = abs(ty - sy)
                if dx_ >= dy_:
                    return [(sx, sy), (tx, sy), (tx, ty)]
                return [(sx, sy), (sx, ty), (tx, ty)]

            # Different orientations can be satisfied with a single bend (2 segments).
            if start_dir == "h" and end_dir == "v":
                return [(sx, sy), (tx, sy), (tx, ty)]
            if start_dir == "v" and end_dir == "h":
                return [(sx, sy), (sx, ty), (tx, ty)]

            # Same orientation at both ends requires 2 bends (3 segments).
            if start_dir == "h" and end_dir == "h":
                x_mid = _snap_to_grid((sx + tx) / 2.0, grid_size_)
                # Avoid zero-length segments after snapping.
                if x_mid == sx:
                    x_mid = _snap_to_grid(sx + (grid_size_ or 10.0), grid_size_)
                if x_mid == tx:
                    x_mid = _snap_to_grid(tx - (grid_size_ or 10.0), grid_size_)
                return [(sx, sy), (x_mid, sy), (x_mid, ty), (tx, ty)]

            if start_dir == "v" and end_dir == "v":
                y_mid = _snap_to_grid((sy + ty) / 2.0, grid_size_)
                if y_mid == sy:
                    y_mid = _snap_to_grid(sy + (grid_size_ or 10.0), grid_size_)
                if y_mid == ty:
                    y_mid = _snap_to_grid(ty - (grid_size_ or 10.0), grid_size_)
                return [(sx, sy), (sx, y_mid), (tx, y_mid), (tx, ty)]

            return pts

        def _build_default_orthogonal_points(
            source_shape_: ShapeElement,
            target_shape_: ShapeElement,
            exit_dx_: float,
            exit_dy_: float,
            entry_dx_: float,
            entry_dy_: float,
        ) -> List[tuple]:
            """
            Build a reasonable default orthogonal polyline when draw.io doesn't store waypoints.

            We evaluate two 1-bend patterns:
              - HV: horizontal then vertical (mid = (target_x, source_y))
              - VH: vertical then horizontal (mid = (source_x, target_y))

            The chosen pattern also determines which sides to use at each end so the
            segment adjacent to the target can go vertically when needed (e.g. "downward from ellipse").
            """
            sx_c = source_shape_.x + source_shape_.w / 2.0
            sy_c = source_shape_.y + source_shape_.h / 2.0
            tx_c = target_shape_.x + target_shape_.w / 2.0
            ty_c = target_shape_.y + target_shape_.h / 2.0

            # If centers are aligned horizontally/vertically, use straight ports.
            # This avoids ambiguous HV/VH tie-breaking that can land on the wrong side.
            dx_c = tx_c - sx_c
            dy_c = ty_c - sy_c
            eps_center = 1e-6
            if abs(dy_c) <= eps_center:
                # Horizontal alignment: exit right/left, entry left/right.
                exit_x = 1.0 if dx_c >= 0 else 0.0
                exit_y = 0.5
                entry_x = 0.0 if dx_c >= 0 else 1.0
                entry_y = 0.5
                p1 = self._calculate_boundary_point(source_shape_, exit_x, exit_y, exit_dx_, exit_dy_)
                p2 = self._calculate_boundary_point(target_shape_, entry_x, entry_y, entry_dx_, entry_dy_)
                return [p1, p2]

            if abs(dx_c) <= eps_center:
                # Vertical alignment: exit bottom/top, entry top/bottom.
                exit_x = 0.5
                exit_y = 1.0 if dy_c >= 0 else 0.0
                entry_x = 0.5
                entry_y = 0.0 if dy_c >= 0 else 1.0
                p1 = self._calculate_boundary_point(source_shape_, exit_x, exit_y, exit_dx_, exit_dy_)
                p2 = self._calculate_boundary_point(target_shape_, entry_x, entry_y, entry_dx_, entry_dy_)
                return [p1, p2]

            # Candidate A: HV (exit left/right, entry top/bottom)
            exit_a_x = 1.0 if tx_c >= sx_c else 0.0
            exit_a_y = 0.5
            entry_a_x = 0.5
            entry_a_y = 1.0 if sy_c >= ty_c else 0.0  # comes from below -> bottom, from above -> top

            p1a = self._calculate_boundary_point(source_shape_, exit_a_x, exit_a_y, exit_dx_, exit_dy_)
            p2a = self._calculate_boundary_point(target_shape_, entry_a_x, entry_a_y, entry_dx_, entry_dy_)
            if p1a[0] == p2a[0] or p1a[1] == p2a[1]:
                points_a = [p1a, p2a]
                len_a = abs(p2a[0] - p1a[0]) + abs(p2a[1] - p1a[1])
            else:
                mid_a = (p2a[0], p1a[1])
                points_a = [p1a, mid_a, p2a]
                len_a = abs(mid_a[0] - p1a[0]) + abs(mid_a[1] - p1a[1]) + abs(p2a[0] - mid_a[0]) + abs(p2a[1] - mid_a[1])

            # Candidate B: VH (exit top/bottom, entry left/right)
            exit_b_x = 0.5
            exit_b_y = 1.0 if ty_c >= sy_c else 0.0  # towards target
            entry_b_x = 1.0 if sx_c >= tx_c else 0.0  # comes from right -> right, from left -> left
            entry_b_y = 0.5

            p1b = self._calculate_boundary_point(source_shape_, exit_b_x, exit_b_y, exit_dx_, exit_dy_)
            p2b = self._calculate_boundary_point(target_shape_, entry_b_x, entry_b_y, entry_dx_, entry_dy_)
            if p1b[0] == p2b[0] or p1b[1] == p2b[1]:
                points_b = [p1b, p2b]
                len_b = abs(p2b[0] - p1b[0]) + abs(p2b[1] - p1b[1])
            else:
                mid_b = (p1b[0], p2b[1])
                points_b = [p1b, mid_b, p2b]
                len_b = abs(mid_b[0] - p1b[0]) + abs(mid_b[1] - p1b[1]) + abs(p2b[0] - mid_b[0]) + abs(p2b[1] - mid_b[1])

            # Prefer the route that exits along the dominant axis between centers.
            # This keeps left/right connections leaving from the side when the target is
            # mostly horizontal from the source, which matches draw.io's default routing.
            if abs(dx_c) > abs(dy_c):
                return points_a
            if abs(dy_c) > abs(dx_c):
                return points_b
            return points_a if len_a <= len_b else points_b

        # Calculate connection points (automatically infer ports if not specified)
        exit_x_val = self.style_extractor.extract_style_float(style_str, "exitX")
        exit_y_val = self.style_extractor.extract_style_float(style_str, "exitY")
        entry_x_val = self.style_extractor.extract_style_float(style_str, "entryX")
        entry_y_val = self.style_extractor.extract_style_float(style_str, "entryY")
        exit_dx = self.style_extractor.extract_style_float(style_str, "exitDx", 0.0)
        exit_dy = self.style_extractor.extract_style_float(style_str, "exitDy", 0.0)
        entry_dx = self.style_extractor.extract_style_float(style_str, "entryDx", 0.0)
        entry_dy = self.style_extractor.extract_style_float(style_str, "entryDy", 0.0)
        grid_size = None
        try:
            if mgm_root is not None:
                grid_size = float(mgm_root.attrib.get("gridSize", "0") or 0) or None
        except Exception:
            grid_size = None

        def _clamp01(val: Optional[float]) -> Optional[float]:
            if val is None:
                return None
            return max(0.0, min(1.0, val))

        hint_exit_x = None
        hint_exit_y = None
        hint_entry_x = None
        hint_entry_y = None
        if edge_style == "orthogonal" and not points_raw and source_point and target_point:
            sxp, syp = source_point
            txp, typ = target_point
            align_tol = 1.0
            dx_hint = abs(sxp - txp)
            dy_hint = abs(syp - typ)
            if not (dx_hint <= align_tol and dy_hint <= align_tol):
                if dy_hint <= align_tol:
                    y_hint = (syp + typ) / 2.0
                    rel_exit_y = _clamp01((y_hint - source_shape.y) / source_shape.h if source_shape.h else 0.5)
                    rel_entry_y = _clamp01((y_hint - target_shape.y) / target_shape.h if target_shape.h else 0.5)
                    if (target_shape.x + target_shape.w / 2.0) >= (source_shape.x + source_shape.w / 2.0):
                        hint_exit_x, hint_entry_x = 1.0, 0.0
                    else:
                        hint_exit_x, hint_entry_x = 0.0, 1.0
                    hint_exit_y, hint_entry_y = rel_exit_y, rel_entry_y
                elif dx_hint <= align_tol:
                    x_hint = (sxp + txp) / 2.0
                    rel_exit_x = _clamp01((x_hint - source_shape.x) / source_shape.w if source_shape.w else 0.5)
                    rel_entry_x = _clamp01((x_hint - target_shape.x) / target_shape.w if target_shape.w else 0.5)
                    if (target_shape.y + target_shape.h / 2.0) >= (source_shape.y + source_shape.h / 2.0):
                        hint_exit_y, hint_entry_y = 1.0, 0.0
                    else:
                        hint_exit_y, hint_entry_y = 0.0, 1.0
                    hint_exit_x, hint_entry_x = rel_exit_x, rel_entry_x

        # If draw.io doesn't store waypoints and explicit ports are also missing,
        # build a better default orthogonal polyline by evaluating HV/VH patterns.
        if edge_style == "orthogonal" and not points_raw and exit_x_val is None and exit_y_val is None and entry_x_val is None and entry_y_val is None and not is_elbow_edge:
            points = _build_default_orthogonal_points(source_shape, target_shape, exit_dx, exit_dy, entry_dx, entry_dy)
            source_x, source_y = points[0]
            target_x, target_y = points[-1]
        else:
            if edge_style == "orthogonal":
                auto_exit_x, auto_exit_y, auto_entry_x, auto_entry_y = self._auto_determine_ports(
                    source_shape, target_shape, points_for_ports
                )
            else:
                auto_exit_x, auto_exit_y, auto_entry_x, auto_entry_y = self._auto_determine_ports(
                    source_shape, target_shape, points_for_ports
                )

            if exit_x_val is None and hint_exit_x is not None:
                exit_x_val = hint_exit_x
            if exit_y_val is None and hint_exit_y is not None:
                exit_y_val = hint_exit_y
            if entry_x_val is None and hint_entry_x is not None:
                entry_x_val = hint_entry_x
            if entry_y_val is None and hint_entry_y is not None:
                entry_y_val = hint_entry_y

            exit_x = exit_x_val if exit_x_val is not None else auto_exit_x
            exit_y = exit_y_val if exit_y_val is not None else auto_exit_y
            entry_x = entry_x_val if entry_x_val is not None else auto_entry_x
            entry_y = entry_y_val if entry_y_val is not None else auto_entry_y

            # For orthogonal edges with waypoints, prefer the side implied by the first/last waypoint
            # when the declared port contradicts it.
            if edge_style == "orthogonal" and points_for_ports:
                declared_exit_side = _infer_port_side(exit_x, exit_y)
                implied_exit_side = _infer_port_side(auto_exit_x, auto_exit_y)
                if implied_exit_side and declared_exit_side != implied_exit_side:
                    exit_x, exit_y = auto_exit_x, auto_exit_y

                declared_entry_side = _infer_port_side(entry_x, entry_y)
                implied_entry_side = _infer_port_side(auto_entry_x, auto_entry_y)
                if implied_entry_side and declared_entry_side != implied_entry_side:
                    entry_x, entry_y = auto_entry_x, auto_entry_y

            # For orthogonal edges with waypoints, project the first/last waypoint onto the
            # source/target boundary so the connector starts (ends) exactly above/below or
            # beside the bend, not at the center of the edge. This preserves "straight down
            # then right" instead of "diagonal from center to waypoint".
            if edge_style == "orthogonal" and points_for_ports:
                first_pt = points_for_ports[0]
                last_pt = points_for_ports[-1]
                exit_side = _infer_port_side(exit_x, exit_y)
                entry_side = _infer_port_side(entry_x, entry_y)
                if exit_side and source_shape.w and source_shape.h:
                    if exit_side == "bottom":
                        exit_x = _clamp01((first_pt[0] - source_shape.x) / source_shape.w)
                        exit_y = 1.0
                    elif exit_side == "top":
                        exit_x = _clamp01((first_pt[0] - source_shape.x) / source_shape.w)
                        exit_y = 0.0
                    elif exit_side == "right":
                        exit_x = 1.0
                        exit_y = _clamp01((first_pt[1] - source_shape.y) / source_shape.h)
                    elif exit_side == "left":
                        exit_x = 0.0
                        exit_y = _clamp01((first_pt[1] - source_shape.y) / source_shape.h)
                if entry_side and target_shape.w and target_shape.h:
                    if entry_side == "bottom":
                        entry_x = _clamp01((last_pt[0] - target_shape.x) / target_shape.w)
                        entry_y = 1.0
                    elif entry_side == "top":
                        entry_x = _clamp01((last_pt[0] - target_shape.x) / target_shape.w)
                        entry_y = 0.0
                    elif entry_side == "right":
                        entry_x = 1.0
                        entry_y = _clamp01((last_pt[1] - target_shape.y) / target_shape.h)
                    elif entry_side == "left":
                        entry_x = 0.0
                        entry_y = _clamp01((last_pt[1] - target_shape.y) / target_shape.h)

            # Calculate connection points
            source_x, source_y = self._calculate_boundary_point(
                source_shape, exit_x, exit_y, exit_dx, exit_dy
            )
            target_x, target_y = self._calculate_boundary_point(
                target_shape, entry_x, entry_y, entry_dx, entry_dy
            )
        
        # Build points
        if 'points' not in locals():
            points = []
            # Filter out (0.0, 0.0) waypoints that come from as="offset" mxPoint elements
            # These are label offsets, not actual waypoints, but we need to maintain
            # backward compatibility for endpoint calculation
            filtered_points_raw = [p for p in points_raw if p != (0.0, 0.0)]
            
            if not filtered_points_raw:
                points = [(source_x, source_y), (target_x, target_y)]
            else:
                points = list(filtered_points_raw)
                # drawio's <Array as="points"> contains only waypoints
                # Start and end points are not included, so they need to be added before and after
                # Insert start point at the beginning
                points.insert(0, (source_x, source_y))
                # Append end point at the end
                points.append((target_x, target_y))
        
        # For orthogonalEdgeStyle, add one intermediate point to create an L-shape (2 segments)
        # even if there are no intermediate points
        if edge_style == "orthogonal" and len(points) == 2:
            # For connectors without stored waypoints, ensure the route respects the declared/auto ports.
            # This is crucial when entry/exit are on left/right: the last segment must be horizontal.
            try:
                points = _ensure_orthogonal_route_respects_ports(
                    points,
                    exit_x if "exit_x" in locals() else None,
                    exit_y if "exit_y" in locals() else None,
                    entry_x if "entry_x" in locals() else None,
                    entry_y if "entry_y" in locals() else None,
                    grid_size,
                )
            except Exception as e:
                # Keep the original points on any unexpected failure.
                if self.logger:
                    self.logger.debug(f"Failed to ensure orthogonal route respects ports: {e}")
        
        # Create style
        style = Style(
            stroke=stroke_color,
            stroke_width=stroke_width,
            dash=dash_pattern,
            arrow_start=start_arrow,
            arrow_end=end_arrow,
            arrow_start_fill=start_fill,
            arrow_end_fill=end_fill,
            arrow_start_size_px=start_size,
            arrow_end_size_px=end_size,
            has_shadow=has_shadow
        )
        
        # Create ConnectorElement
        connector = ConnectorElement(
            id=connector_id,
            source_id=source_id,
            target_id=target_id,
            points=points,
            edge_style=edge_style,
            style=style,
            z_index=0  # TODO: extract z-index
        )

        # Extract label text (edge value)
        labels: List[TextElement] = []

        label_element = self._extract_connector_label(cell, connector, font_color, label_bg_color, style_str)
        if label_element is not None:
            labels.append(label_element)

        # Edge labels can also be stored as child mxCell nodes (edgeLabel style).
        if connector_id and mgm_root is not None:
            try:
                child_cells = mgm_root.findall(f".//mxCell[@parent='{connector_id}']")
            except Exception:
                child_cells = []
            for label_cell in child_cells:
                if label_cell.attrib.get("vertex") != "1":
                    continue
                style_val = label_cell.attrib.get("style", "") or ""
                is_edge_label = "edgeLabel" in style_val
                if not is_edge_label and label_cell.attrib.get("connectable") != "0":
                    continue
                child_label = self._extract_edge_label_cell(
                    label_cell=label_cell,
                    connector=connector,
                    default_font_color=font_color,
                    default_label_bg_color=label_bg_color,
                    source_shape=source_shape,
                    target_shape=target_shape,
                )
                if child_label is not None:
                    labels.append(child_label)

        return connector, labels

    def _extract_connector_label(
        self,
        cell: ET.Element,
        connector: ConnectorElement,
        font_color: Optional[RGBColor],
        label_bg_color: Optional[RGBColor],
        style_str: str,
    ) -> Optional[TextElement]:
        """Extract edge label as a standalone text element."""
        text_raw = cell.attrib.get("value", "") or ""
        if not text_raw.strip():
            return None

        text_paragraphs = self._extract_text(text_raw, font_color, style_str)
        if not text_paragraphs:
            return None

        label_x, label_y = self._calculate_connector_label_position(cell, connector.points, connector.edge_style, style_str)
        label_w, label_h = self._estimate_text_box_size(text_paragraphs)

        return TextElement(
            x=label_x - label_w / 2.0,
            y=label_y - label_h / 2.0,
            w=label_w,
            h=label_h,
            text=text_paragraphs,
            style=Style(
                label_background_color=label_bg_color,
                word_wrap=False,
            ),
        )

    def _extract_edge_label_cell(
        self,
        label_cell: ET.Element,
        connector: ConnectorElement,
        default_font_color: Optional[RGBColor],
        default_label_bg_color: Optional[RGBColor],
        source_shape: Optional[ShapeElement] = None,
        target_shape: Optional[ShapeElement] = None,
    ) -> Optional[TextElement]:
        """Extract edge label from a child mxCell (edgeLabel style)."""
        text_raw = label_cell.attrib.get("value", "") or ""
        if not text_raw.strip():
            return None

        label_style_str = label_cell.attrib.get("style", "") or ""
        label_font_color = self.style_extractor.extract_font_color(label_cell) or default_font_color
        label_bg_color = self.style_extractor.extract_label_background_color(label_cell) or default_label_bg_color

        text_paragraphs = self._extract_text(text_raw, label_font_color, label_style_str)
        if not text_paragraphs:
            return None

        label_x, label_y = self._calculate_connector_label_position(
            label_cell, connector.points, connector.edge_style, label_style_str
        )
        label_w, label_h = self._estimate_text_box_size(text_paragraphs)

        # Determine alignment for positioning
        align = self.style_extractor.extract_style_value(label_style_str, "align")
        vertical_align = self.style_extractor.extract_style_value(label_style_str, "verticalAlign")
        
        # Get geometry to determine if this is a start/end label
        geo = label_cell.find(".//mxGeometry")
        is_start_label = False
        is_end_label = False
        if geo is not None and (geo.attrib.get("relative") or "").strip() == "1":
            try:
                rel_x = float(geo.attrib.get("x", "0") or 0)
                if align == "left" and rel_x <= -0.5:
                    is_start_label = True
                elif align == "right" and rel_x >= 0.5:
                    is_end_label = True
            except ValueError:
                pass

        # Adjust position based on alignment
        if is_start_label:
            # Start label: align to left
            text_x = label_x
        elif is_end_label:
            # End label: align to right
            text_x = label_x - label_w
        else:
            # Center label: center align
            text_x = label_x - label_w / 2.0

        # Adjust vertical position based on verticalAlign
        if vertical_align == "bottom":
            # Place above the line (label's bottom edge aligns with the line)
            text_y = label_y - label_h
        elif vertical_align == "top":
            # Place above the line (label's top edge aligns with the line)
            text_y = label_y - label_h
        else:
            # Center vertically
            text_y = label_y - label_h / 2.0

        # Adjust position to ensure label is outside containers
        if is_start_label and source_shape:
            # Ensure start label is outside source container
            # Check if label overlaps with source shape
            label_right = text_x + label_w
            label_bottom = text_y + label_h
            source_right = source_shape.x + source_shape.w
            source_bottom = source_shape.y + source_shape.h
            
            # If label is inside source shape, move it outside
            if (source_shape.x <= text_x <= source_right and 
                source_shape.y <= text_y <= source_bottom):
                # Move label to the right of the source shape
                text_x = source_right + 5.0
            elif (source_shape.x <= label_right <= source_right and 
                  source_shape.y <= label_bottom <= source_bottom):
                # Move label to the right of the source shape
                text_x = source_right + 5.0
            elif (text_x <= source_shape.x <= label_right and 
                  text_y <= source_shape.y <= label_bottom):
                # Move label below the source shape
                text_y = source_bottom + 5.0
                
        elif is_end_label and target_shape:
            # Ensure end label is outside target container
            # Check if label overlaps with target shape
            label_right = text_x + label_w
            label_bottom = text_y + label_h
            target_right = target_shape.x + target_shape.w
            target_bottom = target_shape.y + target_shape.h
            
            # If label is inside target shape, move it outside
            if (target_shape.x <= text_x <= target_right and 
                target_shape.y <= text_y <= target_bottom):
                # Move label to the left of the target shape
                text_x = target_shape.x - label_w - 5.0
            elif (target_shape.x <= label_right <= target_right and 
                  target_shape.y <= label_bottom <= target_bottom):
                # Move label to the left of the target shape
                text_x = target_shape.x - label_w - 5.0
            elif (text_x <= target_shape.x <= label_right and 
                  text_y <= target_shape.y <= label_bottom):
                # Move label below the target shape
                text_y = target_bottom + 5.0

        return TextElement(
            x=text_x,
            y=text_y,
            w=label_w,
            h=label_h,
            text=text_paragraphs,
            style=Style(
                label_background_color=label_bg_color,
                word_wrap=False,
            ),
        )

    def _calculate_connector_label_position(
        self,
        cell: ET.Element,
        points: List[tuple],
        edge_style: str,
        style_str: str = "",
    ) -> tuple:
        """Return label anchor position (px) for a connector."""
        if not points or len(points) < 2:
            return (0.0, 0.0)

        # Midpoint along the polyline length.
        def _dist(a: tuple, b: tuple) -> float:
            dx = b[0] - a[0]
            dy = b[1] - a[1]
            return (dx * dx + dy * dy) ** 0.5

        total_len = 0.0
        seg_lengths: List[float] = []
        for i in range(len(points) - 1):
            seg_len = _dist(points[i], points[i + 1])
            seg_lengths.append(seg_len)
            total_len += seg_len

        # Determine if this is a start or end label
        align = self.style_extractor.extract_style_value(style_str, "align") if style_str else None
        vertical_align = self.style_extractor.extract_style_value(style_str, "verticalAlign") if style_str else None
        geo = cell.find(".//mxGeometry")
        is_start_label = False
        is_end_label = False
        if geo is not None and (geo.attrib.get("relative") or "").strip() == "1":
            try:
                rel_x = float(geo.attrib.get("x", "0") or 0)
                if align == "left" and rel_x <= -0.5:
                    is_start_label = True
                elif align == "right" and rel_x >= 0.5:
                    is_end_label = True
            except ValueError:
                pass

        if total_len <= 1e-6:
            sx, sy = points[0]
            tx, ty = points[-1]
            if is_start_label:
                base_x, base_y = sx, sy
            elif is_end_label:
                base_x, base_y = tx, ty
            else:
                base_x, base_y = (sx + tx) / 2.0, (sy + ty) / 2.0
            seg_dx, seg_dy = (tx - sx), (ty - sy)
        else:
            rel_pos, rel_offset, abs_offset = self._extract_edge_label_geometry(cell)
            # Override rel_pos for start/end labels
            if is_start_label:
                rel_pos = 0.0
            elif is_end_label:
                rel_pos = 1.0
            t_rel = min(max(rel_pos, 0.0), 1.0)
            target_len = total_len * t_rel
            acc = 0.0
            base_x, base_y = points[0]
            seg_dx, seg_dy = (points[1][0] - points[0][0]), (points[1][1] - points[0][1])
            for i in range(len(points) - 1):
                seg_len = seg_lengths[i]
                if acc + seg_len >= target_len:
                    t = (target_len - acc) / max(seg_len, 1e-6)
                    base_x = points[i][0] + (points[i + 1][0] - points[i][0]) * t
                    base_y = points[i][1] + (points[i + 1][1] - points[i][1]) * t
                    seg_dx = points[i + 1][0] - points[i][0]
                    seg_dy = points[i + 1][1] - points[i][1]
                    break
                acc += seg_len

        # Apply relative offset (perpendicular to local segment) + absolute offset (screen space)
        rel_pos, rel_offset, abs_offset = self._extract_edge_label_geometry(cell)
        rel_x, rel_y = rel_offset
        abs_x, abs_y = abs_offset

        # Normal direction
        if edge_style == "orthogonal" and abs(seg_dx) + abs(seg_dy) > 1e-6:
            # For orthogonal segments, draw.io's label offset is relative to the segment direction.
            # Use the right-hand normal of the segment so positive rel_y moves to the "right" side.
            if abs(seg_dx) >= abs(seg_dy):
                # Horizontal segment: right-hand normal is up when moving right, down when moving left.
                n_x, n_y = (0.0, -1.0) if seg_dx >= 0 else (0.0, 1.0)
            else:
                # Vertical segment: right-hand normal is right when moving down, left when moving up.
                n_x, n_y = (1.0, 0.0) if seg_dy >= 0 else (-1.0, 0.0)
        else:
            seg_len = (seg_dx * seg_dx + seg_dy * seg_dy) ** 0.5
            if seg_len <= 1e-6:
                n_x, n_y = 0.0, 1.0
            else:
                # Counter-clockwise normal.
                n_x = -seg_dy / seg_len
                n_y = seg_dx / seg_len

        # For verticalAlign=bottom, the label should be above the line
        # In draw.io, positive rel_y means upward, so we keep the normal direction
        offset_multiplier = 1.0

        pos_x = base_x + n_x * rel_y * offset_multiplier + abs_x
        pos_y = base_y + n_y * rel_y * offset_multiplier + abs_y

        # If geometry has non-relative x/y, treat them as absolute offsets in screen space.
        if not self._edge_label_is_relative(cell):
            pos_x += rel_x
            pos_y += rel_y

        return (pos_x, pos_y)

    def _edge_label_is_relative(self, cell: ET.Element) -> bool:
        geo = cell.find(".//mxGeometry")
        if geo is None:
            return False
        return (geo.attrib.get("relative") or "").strip() == "1"

    def _extract_edge_label_geometry(self, cell: ET.Element) -> tuple:
        """Extract edge label geometry (relative position, relative offset, absolute offset)."""
        geo = cell.find(".//mxGeometry")
        rel_pos = 0.5
        rel_x = 0.0
        rel_y = 0.0
        abs_x = 0.0
        abs_y = 0.0
        has_rel_x = False

        if geo is not None:
            if "x" in geo.attrib:
                try:
                    rel_x = float(geo.attrib.get("x", "0") or 0)
                    has_rel_x = True
                except ValueError:
                    rel_x = 0.0
            if "y" in geo.attrib:
                try:
                    rel_y = float(geo.attrib.get("y", "0") or 0)
                except ValueError:
                    rel_y = 0.0

            # For relative geometries, x is a position along the edge.
            # mxGraph uses -1..1 where 0 is center, -1 is source, +1 is target.
            # Some exports may store values outside that range; fall back to treating
            # them as offsets from the center.
            if (geo.attrib.get("relative") or "").strip() == "1":
                if has_rel_x:
                    if -1.0 <= rel_x <= 1.0:
                        rel_pos = 0.5 + (rel_x / 2.0)
                    else:
                        rel_pos = 0.5 + rel_x
                else:
                    rel_pos = 0.5

            offset_point = geo.find('./mxPoint[@as="offset"]')
            if offset_point is not None:
                try:
                    if offset_point.attrib.get("x") is not None:
                        abs_x = float(offset_point.attrib.get("x") or 0)
                except ValueError:
                    pass
                try:
                    if offset_point.attrib.get("y") is not None:
                        abs_y = float(offset_point.attrib.get("y") or 0)
                except ValueError:
                    pass

        return (rel_pos, (rel_x, rel_y), (abs_x, abs_y))

    def _estimate_text_box_size(self, paragraphs: List[TextParagraph]) -> tuple:
        """Estimate text box size (px) from text content and font size."""
        if not paragraphs:
            return (10.0, 10.0)

        lines: List[str] = []
        font_size = None
        for para in paragraphs:
            text = "".join(run.text for run in para.runs if run.text)
            if text:
                lines.extend(text.splitlines() or [text])
            if font_size is None:
                for run in para.runs:
                    if run.font_size:
                        font_size = run.font_size
                        break

        if not lines:
            lines = [""]
        if font_size is None:
            font_size = 12.0

        max_len = max(len(line) for line in lines)
        avg_char_px = float(font_size) * 0.6
        padding = float(font_size) * 0.6
        width = max(max_len * avg_char_px + padding, float(font_size) * 1.5)
        height = max(len(lines) * float(font_size) * 1.4, float(font_size) * 1.2)
        return (width, height)
    
    def _calculate_boundary_point(self, shape: ShapeElement, rel_x: float, rel_y: float, offset_x: float, offset_y: float) -> tuple:
        """
        Calculate connection point on shape boundary
        
        rel_x, rel_y: Relative coordinates (0.0-1.0)
        - 0.0 = left/top edge
        - 0.5 = center
        - 1.0 = right/bottom edge
        
        Returns: (x, y) tuple (absolute coordinates)
        """
        # Correction based on shape type
        shape_type = shape.shape_type.lower() if shape.shape_type else ""

        # Base point from rel_x/rel_y (offset is applied only once at the end)
        base_x = shape.x + shape.w * rel_x
        base_y = shape.y + shape.h * rel_y

        # Apply correction for special shapes
        # Parallelogram and flowchart Data (input/output) both have slanted left/right edges.
        # Use the same edge interpolation so connectors attach on the actual slant, not the bounding box.
        if "parallelogram" in shape_type or "data" in shape_type:
            # For parallelograms (and Data), rel_x/rel_y often refers to "position on edge" (e.g., exitX=1, exitY=0.5 is the midpoint of the right edge).
            # If processed with nearest neighbor projection, it may deviate from the midpoint (especially on left/right edges), resulting in appearing "inside the edge".
            # Therefore, for boundary specifications (rel_x/rel_y near 0/1), use linear interpolation on the corresponding edge to ensure it is always on the edge.
            #
            # Furthermore, PPTX's parallelogram tends to have adjust values that affect "horizontal offset amount based on height",
            # and if estimated based on width, the diagonal edge becomes too slanted and connection points tend to go inside. Here, we calculate the offset based on height.
            skew = float(PARALLELOGRAM_SKEW)
            skew = max(0.0, min(skew, 0.49))

            x0, y0, w, h = shape.x, shape.y, shape.w, shape.h
            # Horizontal offset amount (px)
            offset = skew * h
            # Suppress if offset is too large (would cause edge reversal) (less than half of w)
            offset = max(0.0, min(offset, w * 0.49))

            # Default parallelogram tilted to the right (equivalent to PPT's parallelogram)
            tl = (x0 + offset, y0)       # top-left
            tr = (x0 + w, y0)            # top-right
            br = (x0 + w - offset, y0 + h)  # bottom-right
            bl = (x0, y0 + h)            # bottom-left

            def _lerp(a: tuple, b: tuple, t: float) -> tuple:
                return (a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t)

            def _closest_point_on_segment(px: float, py: float, ax: float, ay: float, bx: float, by: float) -> tuple:
                abx = bx - ax
                aby = by - ay
                apx = px - ax
                apy = py - ay
                denom = abx * abx + aby * aby
                if denom <= 1e-12:
                    return ax, ay
                t = (apx * abx + apy * aby) / denom
                t = 0.0 if t < 0.0 else 1.0 if t > 1.0 else t
                return ax + t * abx, ay + t * aby

            eps = 1e-9
            if rel_x <= 0.0 + eps:
                # Left edge (top-left -> bottom-left)
                x, y = _lerp(tl, bl, rel_y)
            elif rel_x >= 1.0 - eps:
                # Right edge (top-right -> bottom-right)
                x, y = _lerp(tr, br, rel_y)
            elif rel_y <= 0.0 + eps:
                # Top edge (top-left -> top-right)
                x, y = _lerp(tl, tr, rel_x)
            elif rel_y >= 1.0 - eps:
                # Bottom edge (bottom-left -> bottom-right)
                x, y = _lerp(bl, br, rel_x)
            else:
                # For internal specification, project to nearest edge
                poly = [tl, tr, br, bl]
                best_x, best_y = poly[0]
                best_d2 = float("inf")
                for i in range(len(poly)):
                    ax, ay = poly[i]
                    bx, by = poly[(i + 1) % len(poly)]
                    cx, cy = _closest_point_on_segment(base_x, base_y, ax, ay, bx, by)
                    dx = base_x - cx
                    dy = base_y - cy
                    d2 = dx * dx + dy * dy
                    if d2 < best_d2:
                        best_d2 = d2
                        best_x, best_y = cx, cy
                x, y = best_x, best_y

        elif "rhombus" in shape_type:
            # Correction for rhombus
            cx = shape.x + shape.w / 2
            cy = shape.y + shape.h / 2
            dx = base_x - cx
            dy = base_y - cy
            
            if shape.w > 0:
                denom = (abs(dx) / (shape.w / 2)) + (abs(dy) / (shape.h / 2))
                if denom > 0:
                    t = 1.0 / denom
                    x = cx + t * dx
                    y = cy + t * dy
                else:
                    x = base_x
                    y = base_y
            else:
                x = base_x
                y = base_y
        elif "ellipse" in shape_type or "circle" in shape_type:
            # Correction for ellipse
            cx = shape.x + shape.w / 2
            cy = shape.y + shape.h / 2
            dx = base_x - cx
            dy = base_y - cy
            
            if shape.w > 0 and shape.h > 0:
                denom_sq = (dx / (shape.w / 2))**2 + (dy / (shape.h / 2))**2
                if denom_sq > 0:
                    t = 1.0 / (denom_sq ** 0.5)
                    x = cx + t * dx
                    y = cy + t * dy
                else:
                    x = base_x
                    y = base_y
            else:
                x = base_x
                y = base_y
        else:
            # Normal rectangle (follow legacy implementation)
            if rel_x <= 0.0:
                # Left edge
                x = shape.x
                y = base_y
            elif rel_x >= 1.0:
                # Right edge
                x = shape.x + shape.w
                y = base_y
            elif rel_y <= 0.0:
                # Top edge
                x = base_x
                y = shape.y
            elif rel_y >= 1.0:
                # Bottom edge
                x = base_x
                y = shape.y + shape.h
            else:
                # For intermediate points, select nearest edge
                dist_left = abs(base_x - shape.x)
                dist_right = abs(base_x - (shape.x + shape.w))
                dist_top = abs(base_y - shape.y)
                dist_bottom = abs(base_y - (shape.y + shape.h))
                
                min_dist = min(dist_left, dist_right, dist_top, dist_bottom)
                
                if min_dist == dist_left:
                    x = shape.x
                    y = base_y
                elif min_dist == dist_right:
                    x = shape.x + shape.w
                    y = base_y
                elif min_dist == dist_top:
                    x = base_x
                    y = shape.y
                else:  # dist_bottom
                    x = base_x
                    y = shape.y + shape.h

        # Add offset only once (exitDx/exitDy, entryDx/entryDy)
        x += offset_x
        y += offset_y

        return (x, y)

    def _auto_determine_ports(self, source_shape: ShapeElement, target_shape: ShapeElement, points: List[tuple] = None) -> tuple:
        """
        Determine default ports when exit/entry are not specified for orthogonalEdgeStyle
        
        If points (waypoints) are specified, determine ports based on direction to first/last point.
        
        Returns:
            (exit_x, exit_y, entry_x, entry_y) relative coordinates
        """
        sx = source_shape.x + source_shape.w / 2
        sy = source_shape.y + source_shape.h / 2
        tx = target_shape.x + target_shape.w / 2
        ty = target_shape.y + target_shape.h / 2
        
        eps = 1e-6

        # --- Exit Port Determination ---
        # If points exist, look at direction to first point
        if points:
            first_point = points[0]
            dx = first_point[0] - sx
            dy = first_point[1] - sy
        else:
            # Otherwise, direction to target center
            dx = tx - sx
            dy = ty - sy

        dx_abs = abs(dx)
        dy_abs = abs(dy)
        
        if dx_abs >= dy_abs:
            # Exit horizontally
            if dx > eps:
                exit_x = 1.0  # Right
            elif dx < -eps:
                exit_x = 0.0  # Left
            else:
                exit_x = 0.5
            exit_y = 0.5
        else:
            # Exit vertically
            if dy > eps:
                exit_y = 1.0  # Bottom
            elif dy < -eps:
                exit_y = 0.0  # Top
            else:
                exit_y = 0.5
            exit_x = 0.5

        # --- Entry Port Determination ---
        # If points exist, look at direction from last point
        if points:
            last_point = points[-1]
            # Direction of last point from target center (entering from there)
            dx_entry = last_point[0] - tx
            dy_entry = last_point[1] - ty
        else:
            # Otherwise, direction from Source center
            dx_entry = sx - tx
            dy_entry = sy - ty
            
        dx_entry_abs = abs(dx_entry)
        dy_entry_abs = abs(dy_entry)
        
        if dx_entry_abs >= dy_entry_abs:
            # Enter from horizontal direction (i.e., left/right edges)
            if dx_entry > eps:
                entry_x = 1.0 # Enter from right edge
            elif dx_entry < -eps:
                entry_x = 0.0 # Enter from left edge
            else:
                entry_x = 0.5
            entry_y = 0.5
        else:
            # Enter from vertical direction (top/bottom edges)
            if dy_entry > eps:
                entry_y = 1.0 # Enter from bottom edge
            elif dy_entry < -eps:
                entry_y = 0.0 # Enter from top edge
            else:
                entry_y = 0.5
            entry_x = 0.5

        return exit_x, exit_y, entry_x, entry_y
    
    def _extract_text(self, text_raw: str, font_color: Optional[RGBColor], style_str: str) -> List[TextParagraph]:
        """Extract text and convert to paragraph list"""
        # Decode HTML entities
        import html as html_module
        if "&lt;" in text_raw or "&gt;" in text_raw or "&amp;" in text_raw:
            text_raw = html_module.unescape(text_raw)
        
        # If HTML tags exist, extract from HTML
        if "<" in text_raw and ">" in text_raw:
            try:
                wrapped = f"<div>{text_raw}</div>"
                parsed = lxml_html.fromstring(wrapped)
                # Extract paragraphs from HTML
                paragraphs = self._extract_text_from_html(parsed, font_color, style_str)
                if paragraphs:
                    return paragraphs
            except Exception as e:
                if self.logger:
                    self.logger.debug(f"Failed to parse HTML text: {e}")
        
        # Process as plain text if HTML is not present or parsing fails
        plain_text = text_raw
        if "<" in plain_text and ">" in plain_text:
            try:
                wrapped = f"<div>{plain_text}</div>"
                parsed = lxml_html.fromstring(wrapped)
                plain_text = parsed.text_content()
            except Exception as e:
                if self.logger:
                    self.logger.debug(f"Failed to extract text content from HTML: {e}")
        
        if not plain_text:
            return []
        
        # Extract text properties from style
        fontSize = self.style_extractor.extract_style_float(style_str, "fontSize")
        fontFamily = self.style_extractor.extract_style_value(style_str, "fontFamily")
        # Treat empty string as None
        if fontFamily == "":
            fontFamily = None
        # Use draw.io's default font if font is not specified
        if fontFamily is None:
            fontFamily = DRAWIO_DEFAULT_FONT_FAMILY
        fontStyle_str = self.style_extractor.extract_style_value(style_str, "fontStyle")
        font_style_flags = self.style_extractor._parse_font_style(fontStyle_str)
        bold = font_style_flags['bold']
        italic = font_style_flags['italic']
        underline = font_style_flags['underline']
        
        align = self.style_extractor.extract_style_value(style_str, "align")
        vertical_align = self.style_extractor.extract_style_value(style_str, "verticalAlign")
        spacing_top = self.style_extractor.extract_style_float(style_str, "spacingTop")
        spacing_left = self.style_extractor.extract_style_float(style_str, "spacingLeft")
        spacing_bottom = self.style_extractor.extract_style_float(style_str, "spacingBottom")
        spacing_right = self.style_extractor.extract_style_float(style_str, "spacingRight")
        
        # Create paragraph
        paragraph = TextParagraph(
            runs=[TextRun(
                text=plain_text,
                font_family=fontFamily,
                font_size=fontSize,
                font_color=font_color,
                bold=bold,
                italic=italic,
                underline=underline
            )],
            align=align.lower() if align else None,
            vertical_align=vertical_align.lower() if vertical_align else None,
            spacing_top=spacing_top,
            spacing_left=spacing_left,
            spacing_bottom=spacing_bottom,
            spacing_right=spacing_right
        )
        
        return [paragraph]
    
    def _extract_text_from_html(self, root_elem, default_font_color: Optional[RGBColor], style_str: str) -> List[TextParagraph]:
        """Extract text paragraphs from HTML element"""
        from ..mapping.text_map import html_to_paragraphs
        
        # Convert HTML back to string
        html_text = lxml_html.tostring(root_elem, encoding='unicode', method='html')
        # Remove <div> tags
        if html_text.startswith('<div>') and html_text.endswith('</div>'):
            html_text = html_text[5:-6]
        
        # Extract default font information
        default_font_size = self.style_extractor.extract_style_float(style_str, "fontSize")
        default_font_family = self.style_extractor.extract_style_value(style_str, "fontFamily")
        # Treat empty string as None
        if default_font_family == "":
            default_font_family = None
        # Use draw.io's default font if font is not specified
        if default_font_family is None:
            default_font_family = DRAWIO_DEFAULT_FONT_FAMILY
        default_font_style = self.style_extractor.extract_style_value(style_str, "fontStyle")
        default_font_style_flags = self.style_extractor._parse_font_style(default_font_style)
        # Note: bold/italic/underline are extracted from HTML, so defaults are not used
        
        # Extract paragraphs from HTML (using text_map)
        paragraphs = html_to_paragraphs(html_text, default_font_color,
                                       default_font_family, default_font_size)
        
        # Apply default font information to each run
        for para in paragraphs:
            for run in para.runs:
                # Also treat empty string as None
                if not run.font_family:
                    run.font_family = default_font_family
                # If still None, use draw.io's default font
                if run.font_family is None:
                    run.font_family = DRAWIO_DEFAULT_FONT_FAMILY
                if run.font_size is None:
                    run.font_size = default_font_size
                if run.font_color is None:
                    run.font_color = default_font_color
                # Apply default bold/italic/underline from style attribute if not set in HTML
                # If bold/italic/underline are not extracted from HTML, apply fontStyle from style attribute as default
                if not run.bold and default_font_style_flags['bold']:
                    run.bold = True
                if not run.italic and default_font_style_flags['italic']:
                    run.italic = True
                if not run.underline and default_font_style_flags['underline']:
                    run.underline = True
        
        # Set paragraph alignment information
        # Mapping: style key -> (para attribute, extractor method, transform function)
        para_attr_map = {
            'align': ('align', self.style_extractor.extract_style_value, lambda v: v.lower() if v else None),
            'verticalAlign': ('vertical_align', self.style_extractor.extract_style_value, lambda v: v.lower() if v else None),
            'spacingTop': ('spacing_top', self.style_extractor.extract_style_float, lambda v: v),
            'spacingLeft': ('spacing_left', self.style_extractor.extract_style_float, lambda v: v),
            'spacingBottom': ('spacing_bottom', self.style_extractor.extract_style_float, lambda v: v),
            'spacingRight': ('spacing_right', self.style_extractor.extract_style_float, lambda v: v),
        }
        
        extracted_values = {}
        for style_key, (attr_name, extractor, transform) in para_attr_map.items():
            value = extractor(style_str, style_key)
            extracted_values[attr_name] = transform(value)
        
        for para in paragraphs:
            for attr_name, value in extracted_values.items():
                if getattr(para, attr_name) is None:
                    setattr(para, attr_name, value)
        
        return paragraphs
