"""Test module for text mapping"""

import pytest
from pptx.dml.color import RGBColor
from drawio2pptx.mapping.text_map import html_to_paragraphs
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

