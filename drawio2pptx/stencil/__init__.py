"""
External stencil support for resolving draw.io shapes that reference
icons by name (e.g. mxgraph.aws4.*) without embedded image data.
"""
from .aws_icons import get_aws_icon_data_uri, get_aws_icon_image_data

__all__ = ["get_aws_icon_data_uri", "get_aws_icon_image_data"]
