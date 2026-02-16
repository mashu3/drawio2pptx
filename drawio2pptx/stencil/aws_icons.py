"""
AWS Architecture Icons stencil support.

Resolves draw.io shapes like shape=mxgraph.aws4.lambda_function (without embedded
image data) by mapping to external icon sources:
- MKAbuMattar/aws-icons (npm, CDN) — official AWS Architecture Icons, updated (e.g. 07/31/2025).
  See https://github.com/MKAbuMattar/aws-icons
- weibeld/aws-icons-svg (raw GitHub) — fallback for specific icons.
"""
import re
from typing import Optional, Dict, List

# Fallback: MKAbuMattar/aws-icons via jsDelivr (official AWS icons, npm package)
# https://github.com/MKAbuMattar/aws-icons — architecture-service/ and resource/
_AWS_ICON_CDN_BASE = "https://cdn.jsdelivr.net/npm/aws-icons@latest/icons"
_ARCH = _AWS_ICON_CDN_BASE + "/architecture-service"
_RES = _AWS_ICON_CDN_BASE + "/resource"

# Fallback for specific icons: weibeld/aws-icons-svg (raw GitHub)
# Some icons (CloudWatch Event, SQS Queue) are more accurate in weibeld
_AWS_ICON_SVG_BASE = (
    "https://raw.githubusercontent.com/weibeld/aws-icons-svg/main/"
    "q1-2022/Resource-Icons_01312022"
)

_AWS_ICON_URL_BY_KEY: Dict[str, str] = {
    # Lambda service icon (architecture-service) vs Lambda Function resource icon (resource)
    "lambda": f"{_ARCH}/AWSLambda.svg",
    "lambda_function": f"{_RES}/AWSLambdaLambdaFunction.svg",
    "cloudwatch": f"{_ARCH}/AmazonCloudWatch.svg",
    "amazon_cloudwatch": f"{_ARCH}/AmazonCloudWatch.svg",
    "sns": f"{_ARCH}/AmazonSimpleNotificationService.svg",
    "amazon_sns": f"{_ARCH}/AmazonSimpleNotificationService.svg",
    "dynamodb": f"{_ARCH}/AmazonDynamoDB.svg",
    "amazon_dynamodb": f"{_ARCH}/AmazonDynamoDB.svg",
    # CloudWatch Event, SQS Queue, and S3 Bucket: use weibeld (more accurate icons)
    "queue": f"{_AWS_ICON_SVG_BASE}/Res_Application-Integration/Res_48_Light/Res_Amazon-Simple-Queue-Service_Queue_48_Light.svg",
    "event_time_based": f"{_AWS_ICON_SVG_BASE}/Res_Management-Governance/Res_48_Light/Res_Amazon-CloudWatch_Event-Time-Based_48_Light.svg",
    "event_event_based": f"{_AWS_ICON_SVG_BASE}/Res_Management-Governance/Res_48_Light/Res_Amazon-CloudWatch_Event-Event-Based_48_Light.svg",
    "bucket_with_objects": f"{_AWS_ICON_SVG_BASE}/Res_Storage/Res_48_Light/Res_Amazon-Simple-Storage-Service_Bucket-With-Objects_48_Light.svg",
    # Other icons: MKAbuMattar CDN
    "topic": f"{_RES}/AmazonSimpleNotificationServiceTopic.svg",
    "template": f"{_RES}/AWSCloudFormationTemplate.svg",
    "role": f"{_RES}/AWSIdentityAccessManagementRole.svg",
}

def _shape_type_to_lookup_keys(shape_type: str, res_icon: Optional[str]) -> List[str]:
    """
    Map draw.io shape type and optional resIcon to stencil lookup keys to try.
    shape_type is normalized lower (e.g. mxgraph.aws4.lambda_function).
    res_icon is from style resIcon= (e.g. mxgraph.aws4.cloudwatch).
    """
    keys: List[str] = []
    if not shape_type or "mxgraph.aws4" not in shape_type:
        return keys

    # Direct shape: mxgraph.aws4.lambda_function -> try "lambda_function"
    if "resourceicon" not in shape_type:
        suffix = shape_type.split(".")[-1] if "." in shape_type else shape_type
        if suffix:
            keys.append(suffix)

    # resourceIcon: use resIcon suffix, and common aliases
    if res_icon:
        icon_suffix = res_icon.split(".")[-1] if "." in res_icon else res_icon
        if icon_suffix:
            keys.append(icon_suffix)
        # Some stencil labels are "Amazon X" so try "amazon_<icon>"
        if icon_suffix:
            keys.append(f"amazon_{icon_suffix}")
    return keys


def get_aws_icon_data_uri(
    shape_type: str,
    style_str: Optional[str] = None,
    url: Optional[str] = None,
) -> Optional[str]:
    """
    Resolve an AWS stencil shape to an image data URI (legacy; prefer get_aws_icon_image_data).
    """
    img = get_aws_icon_image_data(shape_type, style_str, url)
    if img and img.data_uri:
        return img.data_uri
    return None


def get_aws_icon_image_data(
    shape_type: str,
    style_str: Optional[str] = None,
    url: Optional[str] = None,
):
    """
    Resolve an AWS stencil shape to ImageData (file_path URL from icon sources).

    Args:
        shape_type: Normalized shape type (e.g. mxgraph.aws4.lambda_function).
        style_str: Raw style string (to read resIcon for resourceIcon shapes).
        url: Unused (kept for API compatibility).

    Returns:
        ImageData with file_path set if the icon is found, else None.
    """
    from ..model.intermediate import ImageData

    if not shape_type or "mxgraph.aws4" not in shape_type.lower():
        return None

    res_icon = None
    if style_str:
        for part in style_str.split(";"):
            if "=" in part:
                k, v = part.split("=", 1)
                if k.strip() == "resIcon":
                    res_icon = v.strip()
                    break

    keys = _shape_type_to_lookup_keys(shape_type.lower(), res_icon)
    if not keys:
        return None

    # Look up icon URL from mapping dictionary
    for k in keys:
        if k in _AWS_ICON_URL_BY_KEY:
            return ImageData(file_path=_AWS_ICON_URL_BY_KEY[k])
    return None
