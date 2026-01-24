"""
Text mapping module

HTML fragments → paragraph splitting, inline formatting → run splitting, bullet list processing
"""
import re
from typing import List, Optional
from lxml import html as lxml_html
from ..model.intermediate import TextParagraph, TextRun
from pptx.dml.color import RGBColor
from ..io.drawio_loader import ColorParser
from ..logger import get_logger


def html_to_paragraphs(html_text: str, default_font_color: RGBColor = None,
                       default_font_family: Optional[str] = None,
                       default_font_size: Optional[float] = None) -> List[TextParagraph]:
    """
    Convert HTML fragment to paragraph list
    
    Args:
        html_text: HTML text
        default_font_color: Default font color
        default_font_family: Default font family
        default_font_size: Default font size
    
    Returns:
        List of paragraphs
    """
    if not html_text:
        return []
    
    try:
        # Parse HTML
        wrapped = f"<div>{html_text}</div>"
        parsed = lxml_html.fromstring(wrapped)
        
        paragraphs = []
        
        # First, split by <p> tags
        p_tags = parsed.findall('.//p')
        if p_tags:
            for p_elem in p_tags:
                runs = _extract_runs_from_element(p_elem, default_font_color,
                                                default_font_family, default_font_size)
                if runs:
                    paragraphs.append(TextParagraph(runs=runs))
        
        # If <p> tags are not present, extract runs directly from root element
        if not paragraphs:
            runs = _extract_runs_from_element(parsed, default_font_color,
                                            default_font_family, default_font_size)
            if runs:
                paragraphs.append(TextParagraph(runs=runs))
            else:
                # Fallback: process as plain text
                plain_text = parsed.text_content()
                if plain_text:
                    runs = [TextRun(text=plain_text, font_color=default_font_color,
                                   font_family=default_font_family, font_size=default_font_size)]
                    paragraphs.append(TextParagraph(runs=runs))
        
        return paragraphs
    except Exception as e:
        # Process as plain text if parsing fails
        logger = get_logger()
        logger.debug(f"Failed to parse HTML text, falling back to plain text: {e}")
        if html_text:
            runs = [TextRun(text=html_text, font_color=default_font_color,
                           font_family=default_font_family, font_size=default_font_size)]
            return [TextParagraph(runs=runs)]
        return []


def _extract_runs_from_element(elem, default_font_color: RGBColor = None,
                                default_font_family: Optional[str] = None,
                                default_font_size: Optional[float] = None,
                                parent_font_family: Optional[str] = None,
                                parent_font_size: Optional[float] = None,
                                parent_font_color: Optional[RGBColor] = None,
                                parent_bold: bool = False,
                                parent_italic: bool = False,
                                parent_underline: bool = False) -> List[TextRun]:
    """Extract runs from element (includes font information, inherits parent element's style)"""
    runs = []
    
    # Apply default values to parent style
    effective_parent_font_family = parent_font_family or default_font_family
    effective_parent_font_size = parent_font_size if parent_font_size is not None else default_font_size
    effective_parent_font_color = parent_font_color or default_font_color
    
    # Get style from element tag (add to parent style)
    elem_bold = parent_bold
    elem_italic = parent_italic
    elem_underline = parent_underline
    if elem.tag in ['b', 'strong']:
        elem_bold = True
    if elem.tag in ['i', 'em']:
        elem_italic = True
    if elem.tag == 'u':
        elem_underline = True
    
    # Element text
    if elem.text:
        run = _create_run_from_element(elem, elem.text, default_font_color,
                                       effective_parent_font_family, effective_parent_font_size, effective_parent_font_color,
                                       elem_bold, elem_italic, elem_underline,
                                       default_font_family)
        runs.append(run)
        # Update current element's style as parent style
        # Treat empty string as None
        current_font_family = run.font_family if run.font_family else effective_parent_font_family
        current_font_size = run.font_size if run.font_size is not None else effective_parent_font_size
        current_font_color = run.font_color or effective_parent_font_color
        current_bold = run.bold or elem_bold
        current_italic = run.italic or elem_italic
        current_underline = run.underline or elem_underline
    else:
        current_font_family = effective_parent_font_family
        current_font_size = effective_parent_font_size
        current_font_color = effective_parent_font_color
        current_bold = elem_bold
        current_italic = elem_italic
        current_underline = elem_underline
    
    # Process child elements
    for child in elem:
        child_runs = _extract_runs_from_element(child, default_font_color,
                                                default_font_family, default_font_size,
                                                current_font_family, current_font_size, current_font_color,
                                                current_bold, current_italic, current_underline)
        runs.extend(child_runs)
        
        # Tail text (inherit parent element's style)
        if child.tail:
            run = _create_run_from_element(elem, child.tail, default_font_color,
                                          current_font_family, current_font_size, current_font_color,
                                          current_bold, current_italic, current_underline,
                                          default_font_family)
            runs.append(run)
    
    return runs


def _create_run_from_element(elem, text: str, default_font_color: RGBColor = None,
                             parent_font_family: Optional[str] = None,
                             parent_font_size: Optional[float] = None,
                             parent_font_color: Optional[RGBColor] = None,
                             parent_bold: bool = False,
                             parent_italic: bool = False,
                             parent_underline: bool = False,
                             default_font_family: Optional[str] = None) -> TextRun:
    """Create TextRun from element (extract font information, inherit parent element's style)"""
    # Use parent element's style as default (use default if parent is None or empty string)
    # Treat empty string as None
    effective_parent_font_family = parent_font_family if parent_font_family else None
    effective_default_font_family = default_font_family if default_font_family else None
    font_family = effective_parent_font_family or effective_default_font_family
    font_size = parent_font_size
    font_color = parent_font_color or default_font_color
    bold = parent_bold
    italic = parent_italic
    underline = parent_underline
    
    # Extract font information from <font> tag
    if elem.tag == 'font':
        # face attribute (font family)
        face_value = elem.get('face')
        font_family = face_value if face_value else None
        # size attribute (font size, relative value 1-7)
        size_attr = elem.get('size')
        if size_attr:
            try:
                size_int = int(size_attr)
                # Convert relative size to pt (1=8pt, 2=10pt, 3=12pt, 4=14pt, 5=18pt, 6=24pt, 7=36pt)
                size_map = {1: 8, 2: 10, 3: 12, 4: 14, 5: 18, 6: 24, 7: 36}
                font_size = size_map.get(size_int, 12)
            except ValueError:
                # Invalid size attribute, keep default
                pass
        # color attribute
        color_attr = elem.get('color')
        if color_attr:
            font_color = ColorParser.parse(color_attr)
    
    # Extract font information from style attribute
    style_attr = elem.get('style', '')
    if style_attr:
        # font-family
        font_family_match = re.search(r'font-family:\s*([^;]+)', style_attr)
        if font_family_match:
            extracted_font = font_family_match.group(1).strip().strip('"\'')
            font_family = extracted_font if extracted_font else None
        
        # font-size
        font_size_match = re.search(r'font-size:\s*([^;]+)', style_attr)
        if font_size_match:
            size_str = font_size_match.group(1).strip()
            font_size = _parse_font_size(size_str)
        
        # color
        color_match = re.search(r'color:\s*([^;]+)', style_attr)
        if color_match:
            color_value = color_match.group(1).strip()
            parsed_color = ColorParser.parse(color_value)
            if parsed_color:
                font_color = parsed_color
        
        # font-weight (bold)
        font_weight_match = re.search(r'font-weight:\s*([^;]+)', style_attr)
        if font_weight_match:
            weight = font_weight_match.group(1).strip().lower()
            bold = weight in ['bold', 'bolder', '700', '800', '900']
        
        # font-style (italic)
        font_style_match = re.search(r'font-style:\s*([^;]+)', style_attr)
        if font_style_match:
            style = font_style_match.group(1).strip().lower()
            italic = style == 'italic' or style == 'oblique'
        
        # text-decoration (underline)
        text_decoration_match = re.search(r'text-decoration:\s*([^;]+)', style_attr)
        if text_decoration_match:
            decoration = text_decoration_match.group(1).strip().lower()
            underline = 'underline' in decoration
    
    # <b>, <strong> tags
    if elem.tag in ['b', 'strong']:
        bold = True
    
    # <i>, <em> tags
    if elem.tag in ['i', 'em']:
        italic = True
    
    # <u> tag
    if elem.tag == 'u':
        underline = True
    
    # If font family is not set (None or empty string), inherit parent's font family
    if not font_family:
        font_family = effective_parent_font_family or effective_default_font_family
    
    # If still None, use draw.io's default font (final fallback)
    if not font_family:
        from ..fonts import DRAWIO_DEFAULT_FONT_FAMILY
        font_family = DRAWIO_DEFAULT_FONT_FAMILY
    
    return TextRun(
        text=text,
        font_family=font_family,
        font_size=font_size,
        font_color=font_color,
        bold=bold,
        italic=italic,
        underline=underline
    )


def _parse_font_size(size_str: str) -> Optional[float]:
    """Convert font size string to pt"""
    if not size_str:
        return None
    
    size_str = size_str.strip()
    
    # pt unit
    pt_match = re.match(r'^(\d+(?:\.\d+)?)\s*pt$', size_str, re.IGNORECASE)
    if pt_match:
        return float(pt_match.group(1))
    
    # px unit (assuming 96 DPI)
    px_match = re.match(r'^(\d+(?:\.\d+)?)\s*px$', size_str, re.IGNORECASE)
    if px_match:
        px_value = float(px_match.group(1))
        return px_value * 0.75  # px to pt (96 DPI)
    
    # em unit (based on 12pt)
    em_match = re.match(r'^(\d+(?:\.\d+)?)\s*em$', size_str, re.IGNORECASE)
    if em_match:
        em_value = float(em_match.group(1))
        return 12.0 * em_value
    
    # Numeric only (treat as pt)
    try:
        return float(size_str)
    except ValueError:
        # Invalid numeric format, return None
        pass
    
    return None


def _extract_runs_from_text(text: str, default_font_color: RGBColor = None) -> List[TextRun]:
    """Extract runs from text (simple version)"""
    if not text:
        return []
    return [TextRun(text=text, font_color=default_font_color)]


def plain_text_to_paragraphs(text: str, default_font_color: RGBColor = None) -> List[TextParagraph]:
    """
    Convert plain text to paragraph list (split by newline)
    
    Args:
        text: Plain text
        default_font_color: Default font color
    
    Returns:
        List of paragraphs
    """
    if not text:
        return []
    
    # Split by newline
    lines = text.split('\n')
    paragraphs = []
    
    for line in lines:
        if line.strip():  # Skip empty lines
            runs = [TextRun(text=line, font_color=default_font_color)]
            paragraphs.append(TextParagraph(runs=runs))
    
    # Return one empty paragraph if empty
    if not paragraphs:
        paragraphs.append(TextParagraph(runs=[TextRun(text="", font_color=default_font_color)]))
    
    return paragraphs
