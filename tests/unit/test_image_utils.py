"""Test module for image utilities"""

import pytest
import base64
from drawio2pptx.media.image_utils import extract_data_uri_image, svg_to_png


def test_extract_data_uri_image_base64():
    """Test extracting image from base64 data URI"""
    # Create a simple test image data
    test_data = b"test image data"
    encoded = base64.b64encode(test_data).decode('utf-8')
    data_uri = f"data:image/png;base64,{encoded}"
    
    result = extract_data_uri_image(data_uri)
    assert result == test_data


def test_extract_data_uri_image_url_encoded():
    """Test extracting image from URL-encoded data URI"""
    test_data = "test image data"
    data_uri = f"data:image/png,{test_data}"
    
    result = extract_data_uri_image(data_uri)
    assert result == test_data.encode('utf-8')


def test_extract_data_uri_image_invalid():
    """Test extract_data_uri_image with invalid input"""
    assert extract_data_uri_image(None) is None
    assert extract_data_uri_image("") is None
    assert extract_data_uri_image("not a data uri") is None
    assert extract_data_uri_image("data:") is None


def test_extract_data_uri_image_malformed():
    """Test extract_data_uri_image with malformed data URI"""
    # Missing comma
    result = extract_data_uri_image("data:image/png;base64")
    assert result is None


def test_svg_to_png():
    """Test svg_to_png function (not implemented yet)"""
    result = svg_to_png("<svg></svg>")
    assert result is None


