"""
Image processing module

SVG → PNG rasterization (cairosvg default, resvg optional), image extraction from data URIs, and DPI calculation.
CairoSVG is LGPL; used as library only (no modification).
"""
from typing import Optional, Tuple
import re
import io
import urllib.request
from ..config import default_config


def _svg_to_png_cairosvg(svg_data: str, dpi: float, output_width: Optional[int] = None, output_height: Optional[int] = None) -> Optional[bytes]:
    """
    Rasterize SVG to PNG using cairosvg (LGPL; use as library only, no modification).
    
    Args:
        svg_data: SVG data as string
        dpi: DPI setting (affects how SVG units are interpreted)
        output_width: Output width in pixels (scaled by DPI if provided)
        output_height: Output height in pixels (scaled by DPI if provided)
    """
    import cairosvg
    # Scale output dimensions by DPI ratio (96 DPI is baseline)
    # This ensures high-resolution output
    scale = dpi / 96.0
    kwargs = {'dpi': dpi}
    
    if output_width is not None:
        kwargs['output_width'] = int(output_width * scale)
    if output_height is not None:
        kwargs['output_height'] = int(output_height * scale)
    
    # cairosvg.svg2png(bytestring=..., dpi=..., output_width=..., output_height=...) returns bytes when write_to is omitted
    out = cairosvg.svg2png(bytestring=svg_data.encode('utf-8'), **kwargs)
    return bytes(out) if out else None


def _svg_to_png_resvg(svg_data: str, dpi: float) -> Optional[bytes]:
    """
    Rasterize SVG to PNG using resvg.
    resvg's render function scales content using transform matrix, but
    output size is determined by SVG width/height attributes.
    To output at 2x resolution, scale SVG size to 2x before rendering.
    """
    from resvg import render, usvg
    import affine

    scale = dpi / 96.0

    if scale != 1.0:
        def scale_svg_size(match, scale_factor):
            """Scale SVG width/height attribute values"""
            attr_name = match.group(1)
            value = match.group(2)
            num_match = re.search(r'(\d+(?:\.\d+)?)', value)
            if num_match:
                num = float(num_match.group(1))
                new_num = num * scale_factor
                unit = value.replace(num_match.group(1), '').strip()
                return f'{attr_name}="{new_num}{unit}"'
            return match.group(0)

        svg_data = re.sub(
            r'(width)=["\']([^"\']+)["\']',
            lambda m: scale_svg_size(m, scale),
            svg_data
        )
        svg_data = re.sub(
            r'(height)=["\']([^"\']+)["\']',
            lambda m: scale_svg_size(m, scale),
            svg_data
        )

        viewbox_match = re.search(r'viewBox=["\']([^"\']+)["\']', svg_data)
        if viewbox_match:
            viewbox_values = viewbox_match.group(1).split()
            if len(viewbox_values) >= 4:
                try:
                    viewbox_x = float(viewbox_values[0])
                    viewbox_y = float(viewbox_values[1])
                    viewbox_width = float(viewbox_values[2])
                    viewbox_height = float(viewbox_values[3])
                    scaled_width = viewbox_width * scale
                    scaled_height = viewbox_height * scale
                    new_viewbox = f'{viewbox_x} {viewbox_y} {scaled_width} {scaled_height}'
                    svg_data = re.sub(
                        r'viewBox=["\'][^"\']+["\']',
                        f'viewBox="{new_viewbox}"',
                        svg_data,
                        count=1
                    )
                    svg_data = re.sub(
                        r'width=["\'][^"\']+["\']',
                        f'width="{scaled_width}"',
                        svg_data,
                        count=1
                    )
                    svg_data = re.sub(
                        r'height=["\'][^"\']+["\']',
                        f'height="{scaled_height}"',
                        svg_data,
                        count=1
                    )
                    if not re.search(r'width=["\']', svg_data):
                        svg_data = re.sub(
                            r'(<svg[^>]*?)>',
                            lambda m: f'{m.group(1)} width="{scaled_width}">',
                            svg_data,
                            count=1
                        )
                    if not re.search(r'height=["\']', svg_data):
                        svg_data = re.sub(
                            r'(<svg[^>]*?)>',
                            lambda m: f'{m.group(1)} height="{scaled_height}">',
                            svg_data,
                            count=1
                        )
                except (ValueError, IndexError):
                    pass

    db = usvg.FontDatabase.default()
    db.load_system_fonts()
    options = usvg.Options.default()
    tree = usvg.Tree.from_str(svg_data, options, db)
    transform = affine.Affine.scale(scale, scale)
    transform_tuple = transform[0:6]
    png_data = render(tree, transform_tuple)
    return bytes(png_data)


def svg_to_png(svg_data: str, dpi: float = None, output_width: Optional[int] = None, output_height: Optional[int] = None) -> Optional[bytes]:
    """
    Rasterize SVG to PNG using the configured backend (default: cairosvg).

    Backends:
        - cairosvg (default): LGPL, used as library only.
        - resvg: set config.svg_backend = 'resvg' and install resvg, affine.

    Args:
        svg_data: SVG data (string)
        dpi: DPI setting (uses default_config.dpi if None, defaults to 192 DPI)
        output_width: Output width in pixels (optional, scaled by DPI for high resolution)
        output_height: Output height in pixels (optional, scaled by DPI for high resolution)

    Returns:
        PNG data (bytes), or None on conversion failure.

    Raises:
        ImportError: When the selected backend is not installed.
    """
    if dpi is None:
        dpi = default_config.dpi if hasattr(default_config, 'dpi') else 192.0

    backend = getattr(default_config, 'svg_backend', 'cairosvg')
    try:
        if backend == 'resvg':
            return _svg_to_png_resvg(svg_data, dpi)
        else:
            return _svg_to_png_cairosvg(svg_data, dpi, output_width, output_height)
    except ImportError as e:
        if backend == 'resvg':
            raise ImportError(
                "SVG backend is set to 'resvg' but resvg or affine is not installed. "
                "Install with: pip install resvg affine"
            ) from e
        raise ImportError(
            "SVG backend is 'cairosvg' (default) but cairosvg is not installed. "
            "Install with: pip install cairosvg. "
            "Alternatively use resvg: pip install resvg affine and set config.svg_backend = 'resvg'."
        ) from e
    except Exception:
        return None


def svg_bytes_to_png(svg_bytes: bytes, target_width: Optional[int] = None, target_height: Optional[int] = None, dpi: Optional[float] = None) -> Optional[bytes]:
    """
    Convert SVG bytes to PNG bytes (for PowerPoint conversion).

    Uses the configured SVG backend (default: cairosvg; optional: resvg).
    Renders SVG with specified DPI; higher DPI gives higher resolution PNG.
    For cairosvg, output_width and output_height are scaled by DPI for high resolution.

    Args:
        svg_bytes: SVG image data as bytes
        target_width: Target width in pixels (optional, scaled by DPI for high resolution)
        target_height: Target height in pixels (optional, scaled by DPI for high resolution)
        dpi: DPI for rendering (uses default_config.dpi if None, defaults to 192 DPI)

    Returns:
        PNG image data as bytes, or None if conversion fails

    Raises:
        ImportError: If the selected SVG backend (cairosvg or resvg) is not installed
    """
    try:
        # Convert bytes to string
        svg_str = svg_bytes.decode('utf-8')
        
        # Use svg_to_png which handles DPI scaling correctly
        if dpi is None:
            dpi = default_config.dpi if hasattr(default_config, 'dpi') else 192.0
        
        return svg_to_png(svg_str, dpi=dpi, output_width=target_width, output_height=target_height)
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
        svg_str = svg_bytes.decode('utf-8') if isinstance(svg_bytes, bytes) else svg_bytes
        svg_width = None
        svg_height = None
        
        # Try to get size from viewBox first (preferred)
        viewbox_match = re.search(r'viewBox=["\']([^"\']+)["\']', svg_str)
        if viewbox_match:
            viewbox_values = viewbox_match.group(1).split()
            if len(viewbox_values) >= 4:
                try:
                    svg_width = float(viewbox_values[2])
                    svg_height = float(viewbox_values[3])
                except (ValueError, IndexError):
                    pass
        
        # If viewBox not found, try width/height attributes
        if svg_width is None or svg_height is None:
            width_match = re.search(r'width=["\']([^"\']+)["\']', svg_str)
            height_match = re.search(r'height=["\']([^"\']+)["\']', svg_str)
            
            if width_match:
                width_str = width_match.group(1).replace('px', '').strip()
                try:
                    svg_width = float(width_str)
                except ValueError:
                    pass
            
            if height_match:
                height_str = height_match.group(1).replace('px', '').strip()
                try:
                    svg_height = float(height_str)
                except ValueError:
                    pass
        
        return svg_width, svg_height
    except Exception:
        return None, None


def calculate_optimal_dpi(svg_bytes: bytes, base_dpi: float = None) -> float:
    """
    Calculate optimal DPI for SVG to PNG conversion
    
    Ensures minimum short edge of 100px for better quality.
    
    Args:
        svg_bytes: SVG image data as bytes
        base_dpi: Base DPI setting (uses default_config.dpi if None, defaults to 192 DPI)
    
    Returns:
        Optimal DPI value (at least base_dpi, higher if needed for 100px short edge)
    """
    if base_dpi is None:
        base_dpi = default_config.dpi if hasattr(default_config, 'dpi') else 192.0
    
    # Extract SVG dimensions
    svg_width, svg_height = extract_svg_dimensions(svg_bytes)
    
    if svg_width is None or svg_height is None or svg_width <= 0 or svg_height <= 0:
        # Use base DPI if size cannot be determined
        return base_dpi
    
    # Calculate short edge
    short_edge = min(svg_width, svg_height)
    
    # Calculate DPI needed for short edge to be at least 100px
    # At 96 DPI, short_edge units = short_edge pixels
    # We need: short_edge * (dpi / 96) >= 100
    # So: dpi >= 100 * 96 / short_edge
    min_dpi_for_100px = (100.0 * 96.0) / short_edge
    
    # Use the higher of: base DPI or minimum for 100px
    return max(base_dpi, min_dpi_for_100px)


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
            # Many SVG data URIs are base64-encoded even without ;base64 in the header
            import urllib.parse
            import base64
            
            # First, try base64 decode directly (common case for SVG data URIs)
            try:
                # Base64 strings typically contain only A-Z, a-z, 0-9, +, /, = and are length multiples of 4
                # Check if it looks like base64 (at least first 50 chars)
                if len(data) >= 4:
                    # Check if first part looks like base64
                    sample = data[:min(100, len(data))]
                    if all(c in 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=' for c in sample):
                        # Try base64 decode (without validate to handle padding issues)
                        decoded_base64 = base64.b64decode(data)
                        # Check if result is valid SVG
                        if b'<svg' in decoded_base64[:1000] or b'<?xml' in decoded_base64[:1000]:
                            return decoded_base64
            except Exception as e:
                # If base64 decode fails, continue to URL decode
                pass
            
            # If base64 decode failed, try URL decode first, then base64
            decoded_url = urllib.parse.unquote(data)
            
            # Check if URL-decoded data is base64-encoded
            try:
                if len(decoded_url) >= 4:
                    sample = decoded_url[:min(100, len(decoded_url))]
                    if all(c in 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=' for c in sample):
                        decoded_base64 = base64.b64decode(decoded_url)
                        if b'<svg' in decoded_base64[:1000] or b'<?xml' in decoded_base64[:1000]:
                            return decoded_base64
            except Exception:
                pass
            
            # If base64 decode failed or doesn't look like base64, treat as plain text
            if isinstance(decoded_url, str):
                # Check if it's already valid SVG text
                if decoded_url.strip().startswith('<svg') or decoded_url.strip().startswith('<?xml'):
                    return decoded_url.encode('utf-8')
                return decoded_url.encode('utf-8')
            return decoded_url if isinstance(decoded_url, bytes) else decoded_url.encode('utf-8')
    except Exception:
        return None


def load_image_bytes(data_uri: Optional[str] = None, file_path: Optional[str] = None) -> Optional[bytes]:
    """
    Load image bytes from data URI, HTTP(S) URL, or local file path.
    """
    if data_uri:
        return extract_data_uri_image(data_uri)

    if not file_path:
        return None

    try:
        if file_path.startswith(("http://", "https://")):
            req = urllib.request.Request(
                file_path,
                headers={"User-Agent": "drawio2pptx/1.0"},
            )
            with urllib.request.urlopen(req, timeout=10) as response:
                return response.read()

        with open(file_path, "rb") as f:
            return f.read()
    except Exception:
        return None


def is_svg_image(image_bytes: bytes, data_uri: Optional[str] = None, file_path: Optional[str] = None) -> bool:
    """
    Detect whether image bytes/source represent an SVG image.
    """
    if file_path and file_path.lower().endswith(".svg"):
        return True
    if data_uri and "svg" in data_uri.lower():
        return True
    if image_bytes.startswith(b"<svg") or image_bytes.startswith(b"<?xml"):
        return b"<svg" in image_bytes[:1000]
    return False


def trim_transparent_padding(image_bytes: bytes) -> bytes:
    """
    Trim transparent outer padding from a raster image.
    Returns original bytes if trimming is not possible.
    """
    try:
        from PIL import Image

        img = Image.open(io.BytesIO(image_bytes)).convert("RGBA")
        bbox = img.getchannel("A").getbbox()
        if not bbox:
            return image_bytes

        full_bbox = (0, 0, img.width, img.height)
        if bbox == full_bbox:
            return image_bytes

        cropped = img.crop(bbox)
        out = io.BytesIO()
        cropped.save(out, format="PNG")
        return out.getvalue()
    except Exception:
        return image_bytes


def get_image_size(image_bytes: bytes) -> Tuple[Optional[int], Optional[int]]:
    """
    Get raster image dimensions (width, height) in pixels.
    """
    try:
        from PIL import Image

        img = Image.open(io.BytesIO(image_bytes))
        return int(img.width), int(img.height)
    except Exception:
        return None, None


def prepare_image_for_pptx(
    *,
    data_uri: Optional[str] = None,
    file_path: Optional[str] = None,
    shape_type: Optional[str] = None,
    target_width_px: Optional[int] = None,
    target_height_px: Optional[int] = None,
    base_dpi: float = 192.0,
) -> Tuple[Optional[bytes], Optional[int], Optional[int], bool]:
    """
    End-to-end image preparation for PPTX placement.

    Steps:
      1) Load bytes from data URI / URL / local file.
      2) Convert SVG -> PNG with high-resolution settings.
      3) Trim transparent outer padding for AWS icons.
      4) Return final bytes and pixel dimensions.
    """
    image_bytes = load_image_bytes(data_uri=data_uri, file_path=file_path)
    if not image_bytes:
        return None, None, None, False

    svg = is_svg_image(image_bytes, data_uri=data_uri, file_path=file_path)
    if svg:
        dpi = calculate_optimal_dpi(image_bytes, base_dpi=base_dpi)
        image_bytes = svg_bytes_to_png(
            image_bytes,
            target_width=target_width_px,
            target_height=target_height_px,
            dpi=dpi,
        )
        if not image_bytes:
            return None, None, None, True

    if shape_type and shape_type.startswith("mxgraph.aws4"):
        image_bytes = trim_transparent_padding(image_bytes)

    w, h = get_image_size(image_bytes)
    return image_bytes, w, h, svg


