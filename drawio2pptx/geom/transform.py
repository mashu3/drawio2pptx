"""
Geometry transformation module

Provides pre-baking of group inheritance (rotation, scale, translation), bounding box calculation, etc.
"""
import math
from typing import Tuple, List, Optional
from ..model.intermediate import Transform, BaseElement


def apply_transform(
    x: float, y: float,
    transform: Transform,
    origin_x: float = 0.0, origin_y: float = 0.0
) -> Tuple[float, float]:
    """
    Apply transformation to coordinates (rotation, scale, flip, translation)
    
    Args:
        x, y: Coordinates before transformation
        transform: Transformation information
        origin_x, origin_y: Transformation origin (center of rotation/scale)
    
    Returns:
        Transformed coordinates (x, y)
    """
    # Move relative to origin
    x -= origin_x
    y -= origin_y
    
    # Flip
    if transform.flip_h:
        x = -x
    if transform.flip_v:
        y = -y
    
    # Scale
    x *= transform.scale_x
    y *= transform.scale_y
    
    # Rotate
    if transform.rotation != 0.0:
        angle_rad = math.radians(transform.rotation)
        cos_a = math.cos(angle_rad)
        sin_a = math.sin(angle_rad)
        x_new = x * cos_a - y * sin_a
        y_new = x * sin_a + y * cos_a
        x, y = x_new, y_new
    
    # Return to origin
    x += origin_x
    y += origin_y
    
    # Translate
    x += transform.translate_x
    y += transform.translate_y
    
    return (x, y)


def apply_group_transform(
    element: BaseElement,
    group_transform: Transform,
    group_x: float, group_y: float
) -> BaseElement:
    """
    Pre-bake (inherit) group's transformation to element
    
    Apply group's transformation to elements within the group and convert to absolute coordinates
    
    Args:
        element: Element to transform
        group_transform: Group's transformation information
        group_x, group_y: Group position
    
    Returns:
        Transformed element (new instance)
    """
    # Convert element's relative coordinates to group's absolute coordinates
    rel_x = element.x
    rel_y = element.y
    
    # Apply group's transformation
    abs_x, abs_y = apply_transform(
        rel_x, rel_y,
        group_transform,
        origin_x=group_x, origin_y=group_y
    )
    
    # Create new element (copy)
    import copy
    new_element = copy.deepcopy(element)
    new_element.x = abs_x
    new_element.y = abs_y
    
    # Also compose element's own transformation with group's transformation
    new_element.transform.rotation += group_transform.rotation
    new_element.transform.scale_x *= group_transform.scale_x
    new_element.transform.scale_y *= group_transform.scale_y
    new_element.transform.flip_h = new_element.transform.flip_h != group_transform.flip_h
    new_element.transform.flip_v = new_element.transform.flip_v != group_transform.flip_v
    new_element.transform.translate_x += group_transform.translate_x
    new_element.transform.translate_y += group_transform.translate_y
    
    return new_element


def calculate_bounding_box(
    points: List[Tuple[float, float]],
    transform: Optional[Transform] = None
) -> Tuple[float, float, float, float]:
    """
    Calculate bounding box from list of points (including rotation)
    
    Args:
        points: List of points [(x, y), ...]
        transform: Transformation to apply (optional)
    
    Returns:
        (min_x, min_y, width, height)
    """
    if not points:
        return (0.0, 0.0, 0.0, 0.0)
    
    # Apply transformation
    if transform:
        transformed_points = [
            apply_transform(x, y, transform)
            for x, y in points
        ]
    else:
        transformed_points = points
    
    # Calculate bounding box
    xs = [p[0] for p in transformed_points]
    ys = [p[1] for p in transformed_points]
    
    min_x = min(xs)
    max_x = max(xs)
    min_y = min(ys)
    max_y = max(ys)
    
    width = max_x - min_x
    height = max_y - min_y
    
    return (min_x, min_y, width, height)


def calculate_rotated_bounding_box(
    x: float, y: float, w: float, h: float,
    rotation: float
) -> Tuple[float, float, float, float]:
    """
    Calculate bounding box of rotated rectangle
    
    Args:
        x, y: Top-left coordinates of rectangle
        w, h: Width and height of rectangle
        rotation: Rotation angle (degrees)
    
    Returns:
        (min_x, min_y, width, height)
    """
    if rotation == 0.0:
        return (x, y, w, h)
    
    # Rectangle's 4 vertices
    corners = [
        (x, y),
        (x + w, y),
        (x + w, y + h),
        (x, y + h),
    ]
    
    # Center point
    center_x = x + w / 2
    center_y = y + h / 2
    
    # Rotation transformation
    transform = Transform(rotation=rotation)
    rotated_corners = [
        apply_transform(cx, cy, transform, origin_x=center_x, origin_y=center_y)
        for cx, cy in corners
    ]
    
    return calculate_bounding_box(rotated_corners)


def split_polyline_to_segments(
    points: List[Tuple[float, float]]
) -> List[Tuple[Tuple[float, float], Tuple[float, float]]]:
    """
    Split polyline into straight segments
    
    Args:
        points: List of points [(x, y), ...]
    
    Returns:
        List of segments [((x1, y1), (x2, y2)), ...]
    """
    if len(points) < 2:
        return []
    
    segments = []
    for i in range(len(points) - 1):
        segments.append((points[i], points[i + 1]))
    
    return segments


def catmull_rom_to_bezier(
    p0: Tuple[float, float],
    p1: Tuple[float, float],
    p2: Tuple[float, float],
    p3: Tuple[float, float],
    num_segments: int = 10
) -> List[Tuple[float, float]]:
    """
    Convert Catmull-Rom spline to Bezier curve (approximation)
    
    Args:
        p0, p1, p2, p3: Control points
        num_segments: Number of segments
    
    Returns:
        List of Bezier control points
    """
    # Simple implementation: split Catmull-Rom into multiple Bezier segments
    # Use Catmull-Rom formula for more accurate implementation
    points = []
    
    for i in range(num_segments + 1):
        t = i / num_segments
        # Catmull-Rom interpolation (simple version)
        # Use more accurate formula in actual implementation
        x = (1 - t) ** 3 * p1[0] + 3 * (1 - t) ** 2 * t * p1[0] + \
            3 * (1 - t) * t ** 2 * p2[0] + t ** 3 * p2[0]
        y = (1 - t) ** 3 * p1[1] + 3 * (1 - t) ** 2 * t * p1[1] + \
            3 * (1 - t) * t ** 2 * p2[1] + t ** 3 * p2[1]
        points.append((x, y))
    
    return points
