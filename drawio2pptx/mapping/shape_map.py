"""
Shape mapping module

Conversion logic from intermediate model to PowerPoint format (custGeom/connector, etc.)
"""
from typing import Optional
from pptx.enum.shapes import MSO_SHAPE

# Mapping dictionary: draw.io shape type -> PowerPoint MSO_SHAPE
# Multiple draw.io shape types can map to the same PowerPoint shape
_SHAPE_TYPE_MAP: dict[str, MSO_SHAPE] = {
    # Basic shapes
    'ellipse': MSO_SHAPE.OVAL,
    'circle': MSO_SHAPE.OVAL,
    'rect': MSO_SHAPE.RECTANGLE,
    'rectangle': MSO_SHAPE.RECTANGLE,
    'square': MSO_SHAPE.RECTANGLE,
    'rhombus': MSO_SHAPE.DIAMOND,
    'parallelogram': MSO_SHAPE.PARALLELOGRAM,
    'cloud': MSO_SHAPE.CLOUD,
    'trapezoid': MSO_SHAPE.TRAPEZOID,
    'step': MSO_SHAPE.CHEVRON,
    # 3D shapes
    'cylinder3': MSO_SHAPE.CAN,
    'cylinder': MSO_SHAPE.CAN,
    # Flowchart shapes
    'document': MSO_SHAPE.FLOWCHART_DOCUMENT,
    'tape': MSO_SHAPE.FLOWCHART_PUNCHED_TAPE,
    'data': MSO_SHAPE.FLOWCHART_DATA,
    'datastorage': MSO_SHAPE.FLOWCHART_STORED_DATA,
    'data_storage': MSO_SHAPE.FLOWCHART_STORED_DATA,
    'data-storage': MSO_SHAPE.FLOWCHART_STORED_DATA,
    'internalstorage': MSO_SHAPE.FLOWCHART_INTERNAL_STORAGE,
    'internal_storage': MSO_SHAPE.FLOWCHART_INTERNAL_STORAGE,
    'internal-storage': MSO_SHAPE.FLOWCHART_INTERNAL_STORAGE,
    'process': MSO_SHAPE.FLOWCHART_PROCESS,
    'predefinedprocess': MSO_SHAPE.FLOWCHART_PREDEFINED_PROCESS,
    'predefined_process': MSO_SHAPE.FLOWCHART_PREDEFINED_PROCESS,
    'predefined-process': MSO_SHAPE.FLOWCHART_PREDEFINED_PROCESS,
    'decision': MSO_SHAPE.FLOWCHART_DECISION,
    'manualinput': MSO_SHAPE.FLOWCHART_MANUAL_INPUT,
    'manual_input': MSO_SHAPE.FLOWCHART_MANUAL_INPUT,
    'extract': MSO_SHAPE.FLOWCHART_EXTRACT,
    'merge': MSO_SHAPE.FLOWCHART_MERGE,
    'merge_or_storage': MSO_SHAPE.FLOWCHART_MERGE,
    # Off-page connector (draw.io: offPageConnector; PowerPoint: Flowchart Off-page connector)
    'offpageconnector': MSO_SHAPE.FLOWCHART_OFFPAGE_CONNECTOR,
    # Polygon shapes
    'hexagon': MSO_SHAPE.HEXAGON,
    'pentagon': MSO_SHAPE.REGULAR_PENTAGON,
    'octagon': MSO_SHAPE.OCTAGON,
    # Triangle shapes
    'isosceles_triangle': MSO_SHAPE.ISOSCELES_TRIANGLE,
    'right_triangle': MSO_SHAPE.RIGHT_TRIANGLE,
    # Star shapes
    '4_point_star': MSO_SHAPE.STAR_4_POINT,
    '5_point_star': MSO_SHAPE.STAR_5_POINT,
    '6_point_star': MSO_SHAPE.STAR_6_POINT,
    '8_point_star': MSO_SHAPE.STAR_8_POINT,
    # Smiley face
    'smiley': MSO_SHAPE.SMILEY_FACE,
    # Arrows (vertex shapes)
    'right_arrow': MSO_SHAPE.RIGHT_ARROW,
    'notched_right_arrow': MSO_SHAPE.NOTCHED_RIGHT_ARROW,
}


def map_shape_type_to_pptx(shape_type: str) -> MSO_SHAPE:
    """
    Map draw.io shape type to PowerPoint shape type
    
    Args:
        shape_type: draw.io shape type ('rectangle', 'ellipse', etc.)
    
    Returns:
        MSO_SHAPE enumeration value
    """
    shape_type_lower = (shape_type or "").lower()
    return _SHAPE_TYPE_MAP.get(shape_type_lower, MSO_SHAPE.RECTANGLE)
