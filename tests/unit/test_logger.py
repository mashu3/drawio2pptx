"""Test module for logger functionality"""

import pytest
from drawio2pptx.logger import ConversionLogger, ConversionWarning, get_logger


def test_conversion_warning():
    """Test ConversionWarning dataclass"""
    warning = ConversionWarning(
        element_id="test-id",
        warning_type="unsupported_effect",
        message="Test warning",
        details={"key": "value"}
    )
    assert warning.element_id == "test-id"
    assert warning.warning_type == "unsupported_effect"
    assert warning.message == "Test warning"
    assert warning.details == {"key": "value"}


def test_conversion_logger_init():
    """Test ConversionLogger initialization"""
    logger = ConversionLogger(warn_unsupported=True)
    assert logger.warn_unsupported is True
    assert logger.warnings == []
    assert logger.logger is not None


def test_conversion_logger_warn_unsupported_effect():
    """Test warn_unsupported_effect method"""
    logger = ConversionLogger(warn_unsupported=True)
    logger.warn_unsupported_effect("elem1", "rotation", {"angle": 45})
    
    assert len(logger.warnings) == 1
    warning = logger.warnings[0]
    assert warning.element_id == "elem1"
    assert warning.warning_type == "unsupported_effect"
    assert "rotation" in warning.message
    assert warning.details == {"angle": 45}


def test_conversion_logger_warn_unsupported_effect_disabled():
    """Test warn_unsupported_effect when disabled"""
    logger = ConversionLogger(warn_unsupported=False)
    logger.warn_unsupported_effect("elem1", "rotation")
    
    assert len(logger.warnings) == 0


def test_conversion_logger_warn_coordinate_error():
    """Test warn_coordinate_error method"""
    logger = ConversionLogger()
    logger.warn_coordinate_error("elem1", (10.0, 20.0), (10.5, 20.5), 0.1)
    
    assert len(logger.warnings) == 1
    warning = logger.warnings[0]
    assert warning.element_id == "elem1"
    assert warning.warning_type == "coordinate_error"
    assert warning.details['expected'] == (10.0, 20.0)
    assert warning.details['actual'] == (10.5, 20.5)
    assert warning.details['tolerance'] == 0.1


def test_conversion_logger_warn_font_missing():
    """Test warn_font_missing method"""
    logger = ConversionLogger()
    logger.warn_font_missing("elem1", "Arial", "Helvetica")
    
    assert len(logger.warnings) == 1
    warning = logger.warnings[0]
    assert warning.element_id == "elem1"
    assert warning.warning_type == "font_missing"
    assert "Arial" in warning.message
    assert "Helvetica" in warning.message
    assert warning.details['font_family'] == "Arial"
    assert warning.details['replacement'] == "Helvetica"


def test_conversion_logger_warn_font_missing_no_replacement():
    """Test warn_font_missing without replacement"""
    logger = ConversionLogger()
    logger.warn_font_missing("elem1", "Arial")
    
    assert len(logger.warnings) == 1
    warning = logger.warnings[0]
    assert warning.details['replacement'] is None


def test_conversion_logger_logging_methods():
    """Test logging methods (info, debug, error)"""
    logger = ConversionLogger()
    
    # These should not raise exceptions
    logger.info("Test info message")
    logger.debug("Test debug message")
    logger.error("Test error message")


def test_conversion_logger_get_warnings():
    """Test get_warnings method"""
    logger = ConversionLogger()
    logger.warn_unsupported_effect("elem1", "test")
    
    warnings = logger.get_warnings()
    assert len(warnings) == 1
    assert warnings[0].element_id == "elem1"


def test_conversion_logger_clear_warnings():
    """Test clear_warnings method"""
    logger = ConversionLogger()
    logger.warn_unsupported_effect("elem1", "test")
    assert len(logger.warnings) == 1
    
    logger.clear_warnings()
    assert len(logger.warnings) == 0


def test_get_logger():
    """Test get_logger function"""
    logger = get_logger()
    assert isinstance(logger, ConversionLogger)


def test_conversion_logger_with_config():
    """Test ConversionLogger with config"""
    from drawio2pptx.config import ConversionConfig
    
    config = ConversionConfig(warn_unsupported_effects=False, coordinate_tolerance=0.5)
    logger = ConversionLogger(config=config)
    
    assert logger.warn_unsupported is False
    assert logger.config.coordinate_tolerance == 0.5
    
    # Test that coordinate error uses config tolerance
    logger.warn_coordinate_error("elem1", (10.0, 20.0), (10.5, 20.5))
    assert len(logger.warnings) == 1
    assert logger.warnings[0].details['tolerance'] == 0.5


def test_conversion_logger_warn_coordinate_error_without_tolerance():
    """Test warn_coordinate_error without explicit tolerance (uses config)"""
    from drawio2pptx.config import ConversionConfig
    
    config = ConversionConfig(coordinate_tolerance=0.2)
    logger = ConversionLogger(config=config)
    logger.warn_coordinate_error("elem1", (10.0, 20.0), (10.5, 20.5))
    
    assert len(logger.warnings) == 1
    assert logger.warnings[0].details['tolerance'] == 0.2
