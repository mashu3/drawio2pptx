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

        # Prefer preserving block-level order for common draw.io rich-text:
        # e.g. "<h1>Heading</h1><p>Paragraph</p>".
        #
        # Previous behavior extracted only <p> tags (all descendants), which dropped headings entirely.
        # HTML default relative sizes (roughly): h1=2em, h2=1.5em, h3≈1.17em, h4=1em, h5≈0.83em, h6≈0.67em.
        # draw.io uses HTML fragments in labels; matching these ratios makes the PPTX output closer to the editor view.
        heading_scale = {
            "h1": 2.00,
            "h2": 1.50,
            "h3": 1.17,
            "h4": 1.00,
            "h5": 0.83,
            "h6": 0.67,
        }

        block_tags = {"p", "div", "br", "h1", "h2", "h3", "h4", "h5", "h6", "li"}

        def _scaled_heading_size(tag: str) -> Optional[float]:
            if default_font_size is None:
                return None
            try:
                scale = heading_scale.get(tag, 1.0)
                return float(default_font_size) * float(scale)
            except Exception:
                return default_font_size

        def _heading_space_after_pt(tag: str) -> Optional[float]:
            """
            Best-effort paragraph spacing after headings to mimic HTML default margins.
            Use a fraction of the heading font size so it scales with the label.
            """
            size = _scaled_heading_size(tag)
            if size is None:
                return None
            try:
                # Approximate: h1..h6 typically have a noticeable bottom margin in HTML.
                # Use ~0.6em so it reads as a "gap" under the heading in PPT as well.
                return max(3.0, float(size) * 0.60)
            except Exception:
                return None

        # Collect paragraphs from direct children in order when possible.
        for child in parsed:
            tag = (getattr(child, "tag", "") or "").lower()
            if tag not in block_tags:
                continue
            if tag == "br":
                # Explicit line break: add an empty paragraph (PowerPoint will keep spacing).
                paragraphs.append(TextParagraph(runs=[TextRun(text="")]))
                continue

            # Headings: make them bold and (best-effort) larger.
            if tag in heading_scale:
                runs = _extract_runs_from_element(
                    child,
                    default_font_color,
                    default_font_family,
                    _scaled_heading_size(tag),
                    parent_bold=True,
                )
            else:
                runs = _extract_runs_from_element(child, default_font_color,
                                                  default_font_family, default_font_size)
            if any((r.text or "").strip() for r in runs):
                if tag in heading_scale:
                    paragraphs.append(
                        TextParagraph(
                            runs=runs,
                            space_after_pt=_heading_space_after_pt(tag),
                        )
                    )
                else:
                    paragraphs.append(TextParagraph(runs=runs))

        # Fallback: if we didn't find any direct-child block paragraphs, split by <p> tags (descendants).
        if not paragraphs:
            p_tags = parsed.findall('.//p')
            if p_tags:
                for p_elem in p_tags:
                    runs = _extract_runs_from_element(p_elem, default_font_color,
                                                    default_font_family, default_font_size)
                    if any((r.text or "").strip() for r in runs):
                        paragraphs.append(TextParagraph(runs=runs))

        # If <p> tags are not present, treat top-level <div> / <br> as line breaks.
        # draw.io often encodes newlines as "<div>...</div>" segments inside a label.
        if not paragraphs:
            has_top_level_breaks = any(
                (getattr(child, "tag", "") or "").lower() in ("div", "br")
                for child in parsed
            )
            if has_top_level_breaks:
                current_runs: List[TextRun] = []

                def _flush_current_runs():
                    nonlocal current_runs
                    if any((r.text or "").strip() for r in current_runs):
                        paragraphs.append(TextParagraph(runs=current_runs))
                    current_runs = []

                # Leading text before any child tags
                if parsed.text and parsed.text.strip():
                    current_runs.append(
                        TextRun(
                            text=parsed.text,
                            font_color=default_font_color,
                            font_family=default_font_family,
                            font_size=default_font_size,
                        )
                    )

                for child in parsed:
                    tag = (getattr(child, "tag", "") or "").lower()
                    if tag == "br":
                        _flush_current_runs()
                        continue

                    if tag == "div":
                        _flush_current_runs()
                        div_runs = _extract_runs_from_element(child, default_font_color,
                                                             default_font_family, default_font_size)
                        if any((r.text or "").strip() for r in div_runs):
                            paragraphs.append(TextParagraph(runs=div_runs))
                        # Tail text after a block should start a new line
                        if child.tail and child.tail.strip():
                            current_runs.append(
                                TextRun(
                                    text=child.tail,
                                    font_color=default_font_color,
                                    font_family=default_font_family,
                                    font_size=default_font_size,
                                )
                            )
                        continue

                    # Inline element: append into current line
                    inline_runs = _extract_runs_from_element(child, default_font_color,
                                                             default_font_family, default_font_size)
                    current_runs.extend(inline_runs)
                    if child.tail and child.tail.strip():
                        current_runs.append(
                            TextRun(
                                text=child.tail,
                                font_color=default_font_color,
                                font_family=default_font_family,
                                font_size=default_font_size,
                            )
                        )

                _flush_current_runs()
        
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
            # Base size for relative units (em): prefer current inherited size.
            base = parent_font_size if parent_font_size is not None else default_font_size
            font_size = _parse_font_size(size_str, base_size=base)
        
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


def _parse_font_size(size_str: str, base_size: Optional[float] = None) -> Optional[float]:
    """
    Convert CSS font-size into the converter's internal font-size unit.

    Important:
        This project stores TextRun.font_size in the same unit as draw.io's style 'fontSize'
        (best-effort treated as "draw.io units"), and later converts it to PowerPoint points
        via scale_font_size_for_pptx(). Therefore, this function must NOT convert px->pt here,
        otherwise font sizes get scaled twice.
    """
    if not size_str:
        return None
    
    size_str = size_str.strip()
    
    # pt unit -> convert to draw.io units (assume 96 DPI: 1px = 0.75pt)
    pt_match = re.match(r'^(\d+(?:\.\d+)?)\s*pt$', size_str, re.IGNORECASE)
    if pt_match:
        pt_value = float(pt_match.group(1))
        # pt -> px-equivalent
        return pt_value / 0.75
    
    # px unit
    px_match = re.match(r'^(\d+(?:\.\d+)?)\s*px$', size_str, re.IGNORECASE)
    if px_match:
        return float(px_match.group(1))
    
    # em unit (relative to current font size)
    em_match = re.match(r'^(\d+(?:\.\d+)?)\s*em$', size_str, re.IGNORECASE)
    if em_match:
        em_value = float(em_match.group(1))
        effective_base = base_size if base_size is not None else 12.0
        return float(effective_base) * em_value
    
    # Numeric only (treat as draw.io units)
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
