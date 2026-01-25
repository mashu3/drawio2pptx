"""
Logging and QA module

Provides detection and warnings for unsupported effects, coordinate error checking, and log output for automated testing
"""
import logging
from typing import List, Dict, Any, Optional
from dataclasses import dataclass, field
from .config import ConversionConfig, default_config


@dataclass
class ConversionWarning:
    """Warning during conversion"""
    element_id: Optional[str]
    warning_type: str  # 'unsupported_effect', 'coordinate_error', 'font_missing', etc.
    message: str
    details: Dict[str, Any] = field(default_factory=dict)


class ConversionLogger:
    """Logger for conversion process"""
    
    def __init__(self, warn_unsupported: bool = None, config: Optional[ConversionConfig] = None):
        """
        Args:
            warn_unsupported: Whether to warn about unsupported effects (deprecated: use config instead)
            config: ConversionConfig instance (uses default_config if None)
        """
        self.config = config or default_config
        
        # Use config.warn_unsupported_effects if warn_unsupported is not explicitly provided
        if warn_unsupported is None:
            self.warn_unsupported = self.config.warn_unsupported_effects
        else:
            self.warn_unsupported = warn_unsupported
        
        self.warnings: List[ConversionWarning] = []
        self.logger = logging.getLogger('drawio2pptx')
        
        # Logger configuration (default)
        if not self.logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            )
            handler.setFormatter(formatter)
            self.logger.addHandler(handler)
            self.logger.setLevel(logging.INFO)
    
    def warn_unsupported_effect(self, element_id: Optional[str], effect_type: str, details: Dict[str, Any] = None):
        """Record warning for unsupported effect"""
        if not self.warn_unsupported:
            return
        
        message = f"Unsupported effect: {effect_type}"
        warning = ConversionWarning(
            element_id=element_id,
            warning_type='unsupported_effect',
            message=message,
            details=details or {}
        )
        self.warnings.append(warning)
        self.logger.warning(f"[{element_id}] {message}")
    
    def warn_coordinate_error(self, element_id: Optional[str], expected: tuple, actual: tuple, tolerance: Optional[float] = None):
        """Record warning for coordinate error"""
        # Use config.coordinate_tolerance if tolerance is not provided
        if tolerance is None:
            tolerance = self.config.coordinate_tolerance
        
        message = f"Coordinate error: expected {expected}, got {actual} (tolerance: {tolerance})"
        warning = ConversionWarning(
            element_id=element_id,
            warning_type='coordinate_error',
            message=message,
            details={
                'expected': expected,
                'actual': actual,
                'tolerance': tolerance
            }
        )
        self.warnings.append(warning)
        self.logger.warning(f"[{element_id}] {message}")
    
    def warn_font_missing(self, element_id: Optional[str], font_family: str, replacement: str = None):
        """Record warning for missing font"""
        message = f"Font not found: {font_family}"
        if replacement:
            message += f" (replaced with {replacement})"
        
        warning = ConversionWarning(
            element_id=element_id,
            warning_type='font_missing',
            message=message,
            details={
                'font_family': font_family,
                'replacement': replacement
            }
        )
        self.warnings.append(warning)
        self.logger.warning(f"[{element_id}] {message}")
    
    def info(self, message: str):
        """Info log"""
        self.logger.info(message)
    
    def debug(self, message: str):
        """Debug log"""
        self.logger.debug(message)
    
    def error(self, message: str):
        """Error log"""
        self.logger.error(message)
    
    def get_warnings(self) -> List[ConversionWarning]:
        """Get warning list"""
        return self.warnings
    
    def clear_warnings(self):
        """Clear warning list"""
        self.warnings.clear()


# Global logger instance
_default_logger = ConversionLogger()


def get_logger() -> ConversionLogger:
    """Get default logger"""
    return _default_logger
