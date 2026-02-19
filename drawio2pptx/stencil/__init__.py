"""
External stencil support for resolving draw.io shapes that reference
icons by name (e.g. mxgraph.aws*.*) without embedded image data.
"""
from .aws_icons import (
    get_aws_icon_data_uri,
    get_aws_icon_image_data,
    is_aws_shape_type,
    resolve_aws_group_metadata,
)

__all__ = [
    "get_aws_icon_data_uri",
    "get_aws_icon_image_data",
    "is_aws_shape_type",
    "resolve_aws_group_metadata",
]
