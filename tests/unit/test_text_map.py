"""Test module for text mapping"""

import pytest
from pptx.dml.color import RGBColor
from drawio2pptx.mapping.text_map import (
    html_to_paragraphs,
    plain_text_to_paragraphs,
    _parse_font_size,
    _extract_runs_from_text,
)
from drawio2pptx.model.intermediate import TextParagraph, TextRun


def test_html_to_paragraphs_empty():
    """Test html_to_paragraphs with empty string"""
    result = html_to_paragraphs("")
    assert result == []


def test_html_to_paragraphs_none():
    """Test html_to_paragraphs with None"""
    result = html_to_paragraphs(None)
    assert result == []


def test_html_to_paragraphs_plain_text():
    """Test html_to_paragraphs with plain text"""
    result = html_to_paragraphs("Hello World")
    assert len(result) == 1
    assert len(result[0].runs) == 1
    assert result[0].runs[0].text == "Hello World"


def test_html_to_paragraphs_with_defaults():
    """Test html_to_paragraphs with default values"""
    default_color = RGBColor(255, 0, 0)
    result = html_to_paragraphs("Hello", default_font_color=default_color,
                               default_font_family="Arial", default_font_size=12.0)
    assert len(result) == 1
    assert len(result[0].runs) == 1
    assert result[0].runs[0].font_color == default_color
    assert result[0].runs[0].font_family == "Arial"
    assert result[0].runs[0].font_size == 12.0


def test_html_to_paragraphs_with_p_tags():
    """Test html_to_paragraphs with <p> tags"""
    html = "<p>Paragraph 1</p><p>Paragraph 2</p>"
    result = html_to_paragraphs(html)
    assert len(result) == 2
    assert result[0].runs[0].text == "Paragraph 1"
    assert result[1].runs[0].text == "Paragraph 2"


def test_html_to_paragraphs_with_bold():
    """Test html_to_paragraphs with <b> tag"""
    html = "Hello <b>World</b>"
    result = html_to_paragraphs(html)
    assert len(result) == 1
    assert len(result[0].runs) >= 2
    # Find bold run
    bold_run = next((r for r in result[0].runs if r.bold), None)
    assert bold_run is not None
    assert "World" in bold_run.text


def test_html_to_paragraphs_with_italic():
    """Test html_to_paragraphs with <i> tag"""
    html = "Hello <i>World</i>"
    result = html_to_paragraphs(html)
    assert len(result) == 1
    # Find italic run
    italic_run = next((r for r in result[0].runs if r.italic), None)
    assert italic_run is not None
    assert "World" in italic_run.text


def test_html_to_paragraphs_with_underline():
    """Test html_to_paragraphs with <u> tag"""
    html = "Hello <u>World</u>"
    result = html_to_paragraphs(html)
    assert len(result) == 1
    # Find underline run
    underline_run = next((r for r in result[0].runs if r.underline), None)
    assert underline_run is not None
    assert "World" in underline_run.text


def test_html_to_paragraphs_with_strong():
    """Test html_to_paragraphs with <strong> tag"""
    html = "Hello <strong>World</strong>"
    result = html_to_paragraphs(html)
    assert len(result) == 1
    # Find bold run (strong is bold)
    bold_run = next((r for r in result[0].runs if r.bold), None)
    assert bold_run is not None


def test_html_to_paragraphs_with_em():
    """Test html_to_paragraphs with <em> tag"""
    html = "Hello <em>World</em>"
    result = html_to_paragraphs(html)
    assert len(result) == 1
    # Find italic run (em is italic)
    italic_run = next((r for r in result[0].runs if r.italic), None)
    assert italic_run is not None


def test_html_to_paragraphs_with_font_tag():
    """Test html_to_paragraphs with <font> tag"""
    html = '<font color="#FF0000">Red text</font>'
    result = html_to_paragraphs(html)
    assert len(result) == 1
    # Find run with color
    colored_run = next((r for r in result[0].runs if r.font_color), None)
    assert colored_run is not None


def test_html_to_paragraphs_malformed_html():
    """Test html_to_paragraphs with malformed HTML (should fallback to plain text)"""
    html = "<div><p>Unclosed tag"
    result = html_to_paragraphs(html)
    # Should still return something (fallback to plain text)
    assert len(result) >= 1


def test_html_to_paragraphs_multiple_formats():
    """Test html_to_paragraphs with multiple formatting tags"""
    html = "Normal <b>Bold</b> <i>Italic</i> <u>Underline</u> text"
    result = html_to_paragraphs(html)
    assert len(result) == 1
    assert len(result[0].runs) >= 4  # At least 4 runs (normal, bold, italic, underline)


def test_html_to_paragraphs_css_font_size_px_is_not_pre_scaled() -> None:
    """CSS font-size in px should be kept in draw.io units (writer applies scaling later)."""
    html = '<span style="font-size: 20px">X</span>'
    result = html_to_paragraphs(html, default_font_size=12.0)
    assert len(result) == 1
    assert len(result[0].runs) >= 1
    assert result[0].runs[0].font_size == 20.0


# ---- Headings (h1-h6), br, block order ----
def test_html_to_paragraphs_h1_heading_with_default_font_size() -> None:
    """h1 produces bold run with scaled font size and space_after_pt."""
    html = "<h1>Title</h1>"
    result = html_to_paragraphs(html, default_font_size=12.0)
    assert len(result) == 1
    assert result[0].runs[0].text == "Title"
    assert result[0].runs[0].bold is True
    assert result[0].runs[0].font_size == 24.0  # 12 * 2.0
    assert result[0].space_after_pt is not None
    assert result[0].space_after_pt >= 3.0


def test_html_to_paragraphs_h2_h6_scaled_sizes() -> None:
    """h2-h6 use heading scale factors."""
    for tag, expected_scale in [("h2", 1.5), ("h3", 1.17), ("h4", 1.0), ("h5", 0.83), ("h6", 0.67)]:
        html = f"<{tag}>Heading</{tag}>"
        result = html_to_paragraphs(html, default_font_size=10.0)
        assert len(result) == 1
        assert result[0].runs[0].font_size == pytest.approx(10.0 * expected_scale, rel=0.01)


def test_html_to_paragraphs_br_adds_empty_paragraph() -> None:
    """Explicit <br> adds an empty paragraph."""
    html = "<p>Line 1</p><br/><p>Line 2</p>"
    result = html_to_paragraphs(html)
    assert len(result) >= 2
    empty_paras = [p for p in result if len(p.runs) == 1 and p.runs[0].text == ""]
    assert len(empty_paras) >= 1


def test_html_to_paragraphs_fallback_findall_p() -> None:
    """When no direct block children, fallback to .//p."""
    html = "<div><span><p>Nested p</p></span></div>"
    result = html_to_paragraphs(html)
    assert len(result) == 1
    assert result[0].runs[0].text == "Nested p"


def test_html_to_paragraphs_top_level_div_br_leading_text() -> None:
    """Top-level div/br path: leading text before first child."""
    html = "<div>Leading</div><div>Second</div>"
    result = html_to_paragraphs(html)
    assert len(result) >= 2
    texts = [p.runs[0].text for p in result if p.runs]
    assert "Leading" in texts
    assert "Second" in texts


def test_html_to_paragraphs_div_with_tail() -> None:
    """Empty div with tail hits the top-level div/br path and child.tail."""
    html = "<div></div> tail"
    result = html_to_paragraphs(html)
    assert len(result) >= 1
    all_text = "".join(r.text for p in result for r in p.runs)
    assert "tail" in all_text


def test_html_to_paragraphs_root_fallback_and_plain_text_fallback() -> None:
    """When no block structure, extract from root; else plain text."""
    html = "Just text no tags"
    result = html_to_paragraphs(html)
    assert len(result) == 1
    assert result[0].runs[0].text == "Just text no tags"


def test_html_to_paragraphs_exception_fallback_to_plain_text() -> None:
    """On parse exception, fallback to plain text."""
    from unittest.mock import patch
    html = "Fallback content"
    with patch("drawio2pptx.mapping.text_map.lxml_html.fromstring", side_effect=ValueError("parse error")):
        result = html_to_paragraphs(html, default_font_size=12.0)
    assert len(result) == 1
    assert len(result[0].runs) == 1
    assert result[0].runs[0].text == "Fallback content"


# ---- _create_run_from_element: font tag, style attr ----
def test_html_to_paragraphs_font_tag_face_and_size() -> None:
    """<font face="Arial" size="4"> uses face and size (4 -> 14pt)."""
    html = '<font face="Arial" size="4">Sized</font>'
    result = html_to_paragraphs(html)
    assert len(result) == 1
    r = result[0].runs[0]
    assert r.font_family == "Arial"
    assert r.font_size == 14


def test_html_to_paragraphs_font_tag_size_invalid_keeps_default() -> None:
    """<font size="invalid"> keeps default size."""
    html = '<font size="x">X</font>'
    result = html_to_paragraphs(html, default_font_size=12.0)
    assert len(result) == 1
    assert result[0].runs[0].font_size == 12.0


def test_html_to_paragraphs_style_font_family() -> None:
    """style='font-family: Tahoma' is applied."""
    html = '<span style="font-family: Tahoma">T</span>'
    result = html_to_paragraphs(html)
    assert len(result) == 1
    assert result[0].runs[0].font_family == "Tahoma"


def test_html_to_paragraphs_style_font_weight_bold() -> None:
    """style='font-weight: bold' sets bold."""
    html = '<span style="font-weight: bold">B</span>'
    result = html_to_paragraphs(html)
    assert len(result) == 1
    assert result[0].runs[0].bold is True


def test_html_to_paragraphs_style_font_style_italic() -> None:
    """style='font-style: italic' sets italic."""
    html = '<span style="font-style: italic">I</span>'
    result = html_to_paragraphs(html)
    assert len(result) == 1
    assert result[0].runs[0].italic is True


def test_html_to_paragraphs_style_text_decoration_underline() -> None:
    """style='text-decoration: underline' sets underline."""
    html = '<span style="text-decoration: underline">U</span>'
    result = html_to_paragraphs(html)
    assert len(result) == 1
    assert result[0].runs[0].underline is True


def test_html_to_paragraphs_no_font_family_uses_default() -> None:
    """When no font family set, uses draw.io default (Helvetica)."""
    html = "Plain"
    result = html_to_paragraphs(html)
    assert len(result) == 1
    assert result[0].runs[0].font_family == "Helvetica"


# ---- _parse_font_size ----
def test_parse_font_size_empty_returns_none() -> None:
    assert _parse_font_size("") is None
    assert _parse_font_size(None) is None


def test_parse_font_size_pt() -> None:
    """pt unit converted to px-equivalent (pt/0.75)."""
    assert _parse_font_size("12pt") == pytest.approx(16.0, rel=0.01)
    assert _parse_font_size("9pt") == pytest.approx(12.0, rel=0.01)


def test_parse_font_size_px() -> None:
    assert _parse_font_size("20px") == 20.0
    assert _parse_font_size("14px") == 14.0


def test_parse_font_size_em() -> None:
    """em is relative to base_size."""
    assert _parse_font_size("2em", base_size=10.0) == 20.0
    assert _parse_font_size("1.5em", base_size=12.0) == 18.0
    assert _parse_font_size("1em") == 12.0  # default base 12.0


def test_parse_font_size_numeric() -> None:
    assert _parse_font_size("14") == 14.0
    assert _parse_font_size("12.5") == 12.5


def test_parse_font_size_invalid_returns_none() -> None:
    assert _parse_font_size("big") is None
    assert _parse_font_size("abc12") is None


# ---- plain_text_to_paragraphs, _extract_runs_from_text ----
def test_plain_text_to_paragraphs_empty() -> None:
    assert plain_text_to_paragraphs("") == []
    assert plain_text_to_paragraphs(None) == []


def test_plain_text_to_paragraphs_splits_newlines() -> None:
    result = plain_text_to_paragraphs("Line1\nLine2\nLine3")
    assert len(result) == 3
    assert result[0].runs[0].text == "Line1"
    assert result[1].runs[0].text == "Line2"
    assert result[2].runs[0].text == "Line3"


def test_plain_text_to_paragraphs_skips_empty_lines() -> None:
    result = plain_text_to_paragraphs("A\n\nB")
    assert len(result) == 2
    assert result[0].runs[0].text == "A"
    assert result[1].runs[0].text == "B"


def test_plain_text_to_paragraphs_all_empty_returns_one_empty_paragraph() -> None:
    result = plain_text_to_paragraphs("\n\n")
    assert len(result) == 1
    assert result[0].runs[0].text == ""


def test_plain_text_to_paragraphs_default_font_color() -> None:
    red = RGBColor(255, 0, 0)
    result = plain_text_to_paragraphs("Hi", default_font_color=red)
    assert len(result) == 1
    assert result[0].runs[0].font_color == red


def test_extract_runs_from_text_empty() -> None:
    assert _extract_runs_from_text("") == []
    assert _extract_runs_from_text(None) == []


def test_extract_runs_from_text_simple() -> None:
    runs = _extract_runs_from_text("Hello")
    assert len(runs) == 1
    assert runs[0].text == "Hello"
    assert runs[0].font_color is None


def test_extract_runs_from_text_with_color() -> None:
    red = RGBColor(255, 0, 0)
    runs = _extract_runs_from_text("Hi", default_font_color=red)
    assert len(runs) == 1
    assert runs[0].font_color == red


# ---- More edge cases for coverage ----
def test_html_to_paragraphs_h1_without_default_font_size() -> None:
    """h1 with default_font_size=None: _scaled_heading_size and _heading_space_after_pt return None."""
    html = "<h1>Title</h1>"
    result = html_to_paragraphs(html, default_font_size=None)
    assert len(result) == 1
    assert result[0].runs[0].text == "Title"
    assert result[0].space_after_pt is None


def test_html_to_paragraphs_fallback_findall_p_no_direct_block() -> None:
    """No direct block children; paragraphs come only from .//p (covers 118-122)."""
    html = "<span><p>Nested</p></span>"
    result = html_to_paragraphs(html)
    assert len(result) == 1
    assert result[0].runs[0].text == "Nested"


def test_html_to_paragraphs_leading_text_before_child() -> None:
    """Leading text before first child in top-level div path (154-155)."""
    html = "Leading<div></div>"
    result = html_to_paragraphs(html)
    assert len(result) >= 1
    all_text = "".join(r.text for p in result for r in p.runs)
    assert "Leading" in all_text


def test_html_to_paragraphs_br_in_top_level_breaks_path() -> None:
    """Br in the inner loop flushes current runs (162)."""
    html = "<div>A</div><br/><div>B</div>"
    result = html_to_paragraphs(html)
    assert len(result) >= 2
    texts = [p.runs[0].text if p.runs else "" for p in result]
    assert "A" in texts
    assert "B" in texts


def test_html_to_paragraphs_inline_element_with_tail() -> None:
    """Inline element with tail (176-180)."""
    html = "<div></div><span>X</span> y"
    result = html_to_paragraphs(html)
    all_text = "".join(r.text for p in result for r in p.runs)
    assert "X" in all_text
    assert "y" in all_text


def test_html_to_paragraphs_exception_empty_html_returns_empty_list() -> None:
    """Exception path with empty html_text returns [] (214)."""
    from unittest.mock import patch
    with patch("drawio2pptx.mapping.text_map.lxml_html.fromstring", side_effect=ValueError()):
        result = html_to_paragraphs("")
    assert result == []


def test_html_to_paragraphs_style_font_family_empty_treated_as_none() -> None:
    """Empty font-family in style yields None (334-335)."""
    html = '<span style="font-family: ">X</span>'
    result = html_to_paragraphs(html)
    assert len(result) == 1
    # Empty or whitespace-only extracted_font -> None, then fallback to default
    assert result[0].runs[0].font_family is not None


def test_html_to_paragraphs_font_tag_size_map_1_and_7() -> None:
    """<font size="1"> and size="7"> use size_map (348-351)."""
    html1 = '<font size="1">S1</font>'
    result1 = html_to_paragraphs(html1)
    assert result1[0].runs[0].font_size == 8
    html7 = '<font size="7">S7</font>'
    result7 = html_to_paragraphs(html7)
    assert result7[0].runs[0].font_size == 36
