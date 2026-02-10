"""
Image processing module

SVG → PNG rasterization, dataURI → /ppt/media/* expansion, image rotation, scaling, and cropping
"""
from typing import Optional, Tuple
from pathlib import Path
import re
from ..config import default_config


def svg_to_png(svg_data: str, dpi: float = None) -> Optional[bytes]:
    """
    Rasterize SVG to PNG using resvg
    
    Args:
        svg_data: SVG data (string)
        dpi: DPI setting (uses default setting if None, defaults to 96 DPI)
    
    Returns:
        PNG data (bytes), or None
    """
    try:
        from resvg import render, usvg
        import affine
        
        if dpi is None:
            dpi = default_config.get('svg_dpi', 96.0) if hasattr(default_config, 'get') else 96.0
        
        # Set up resvg
        db = usvg.FontDatabase.default()
        db.load_system_fonts()
        
        options = usvg.Options.default()
        
        # Parse SVG
        tree = usvg.Tree.from_str(svg_data, options, db)
        
        # Calculate scale factor for DPI
        scale = dpi / 96.0
        
        # Create affine transformation matrix
        # Apply uniform scaling based on DPI only (for general-purpose use)
        transform = affine.Affine.scale(scale, scale)
        transform_tuple = transform[0:6]
        
        # Render to PNG
        png_data = render(tree, transform_tuple)
        return bytes(png_data)
    except ImportError:
        # Explicitly fail if library is not available
        raise
    except Exception:
        return None


def svg_bytes_to_png(svg_bytes: bytes, target_width: Optional[int] = None, target_height: Optional[int] = None) -> Optional[bytes]:
    """
    Convert SVG bytes to PNG bytes using resvg (for PowerPoint conversion)
    
    This function renders SVG with identity matrix, allowing PowerPoint to handle scaling.
    This preserves aspect ratio and prevents size mismatches.
    
    Args:
        svg_bytes: SVG image data as bytes
        target_width: Target width in pixels (optional, currently unused but kept for API compatibility)
        target_height: Target height in pixels (optional, currently unused but kept for API compatibility)
    
    Returns:
        PNG image data as bytes, or None if conversion fails
    
    Raises:
        ImportError: If resvg or affine libraries are not available
    """
    try:
        from resvg import render, usvg
        import affine
        
        # Convert bytes to string
        svg_str = svg_bytes.decode('utf-8')
        
        # Set up resvg
        db = usvg.FontDatabase.default()
        db.load_system_fonts()
        
        options = usvg.Options.default()
        
        # Parse SVG
        tree = usvg.Tree.from_str(svg_str, options, db)
        
        # Render SVG as-is (with identity matrix) and let PowerPoint handle scaling
        # This preserves aspect ratio and prevents size mismatches
        transform = affine.Affine.identity()
        transform_tuple = transform[0:6]
        
        # Render to PNG
        png_data = render(tree, transform_tuple)
        return bytes(png_data)
    except ImportError:
        # Explicitly fail if library is not available
        raise
    except Exception:
        return None


def extract_svg_dimensions(svg_bytes: bytes) -> Tuple[Optional[float], Optional[float]]:
    """
    Extract width and height from SVG bytes
    
    Args:
        svg_bytes: SVG image data as bytes
    
    Returns:
        Tuple of (width, height) in pixels, or (None, None) if not found
    """
    try:
        svg_str = svg_bytes.decode('utf-8')
        svg_width = None
        svg_height = None
        
        # Try to get width/height from SVG element attributes
        width_match = re.search(r'width=["\']([^"\']+)["\']', svg_str)
        height_match = re.search(r'height=["\']([^"\']+)["\']', svg_str)
        
        if width_match:
            width_str = width_match.group(1)
            # Remove 'px' if present and convert to float
            width_str = width_str.replace('px', '').strip()
            try:
                svg_width = float(width_str)
            except ValueError:
                pass
        
        if height_match:
            height_str = height_match.group(1)
            height_str = height_str.replace('px', '').strip()
            try:
                svg_height = float(height_str)
            except ValueError:
                pass
        
        # If not found, try viewBox
        if svg_width is None or svg_height is None:
            viewbox_match = re.search(r'viewBox=["\']([^"\']+)["\']', svg_str)
            if viewbox_match:
                viewbox_values = viewbox_match.group(1).split()
                if len(viewbox_values) >= 4:
                    try:
                        svg_width = float(viewbox_values[2])
                        svg_height = float(viewbox_values[3])
                    except (ValueError, IndexError):
                        pass
        
        return svg_width, svg_height
    except Exception:
        return None, None


def extract_data_uri_image(data_uri: str) -> Optional[bytes]:
    """
    Extract image data from data URI
    
    Args:
        data_uri: data URI string (data:image/png;base64,... or data:image/svg+xml,...)
    
    Returns:
        Image data (bytes), or None
    """
    if not data_uri or not data_uri.startswith('data:'):
        return None
    
    try:
        # data URI format: data:[<mediatype>][;base64],<data>
        header, data = data_uri.split(',', 1)
        
        if 'base64' in header:
            import base64
            return base64.b64decode(data)
        else:
            # For SVG data URIs, the data may be URL-encoded and/or base64-encoded
            # Try URL decode first
            import urllib.parse
            import base64
            
            # URL decode
            decoded_url = urllib.parse.unquote(data)
            
            # Check if it's base64-encoded (common for SVG in data URIs)
            # Base64 strings typically contain only A-Z, a-z, 0-9, +, /, = and are length multiples of 4
            try:
                # Try base64 decode if it looks like base64
                if len(decoded_url) % 4 == 0 and all(c in 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=' for c in decoded_url):
                    decoded_base64 = base64.b64decode(decoded_url)
                    # Check if result is valid SVG
                    if b'<svg' in decoded_base64[:1000]:
                        return decoded_base64
            except Exception:
                pass
            
            # If base64 decode failed or doesn't look like base64, treat as plain text
            if isinstance(decoded_url, str):
                return decoded_url.encode('utf-8')
            return decoded_url
    except Exception:
        return None


def save_image_to_media(image_data: bytes, output_dir: Path, filename: str = None) -> Optional[str]:
    """
    Save image data to /ppt/media/ directory
    
    Args:
        image_data: Image data (bytes)
        output_dir: Output directory
        filename: Filename (auto-generated if None)
    
    Returns:
        Relative path (/ppt/media/filename), or None
    """
    # TODO: Implementation needed
    # Save image to PowerPoint media directory and add relationship
    return None
