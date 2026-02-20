"""
AWS Architecture Icons stencil support.

Resolves draw.io shapes like shape=mxgraph.aws4.lambda_function (without embedded
image data) by mapping to external icon sources:
- MKAbuMattar/aws-icons (npm, CDN) — official AWS Architecture Icons, updated (e.g. 07/31/2025).
  See https://github.com/MKAbuMattar/aws-icons
- weibeld/aws-icons-svg (raw GitHub) — fallback for specific icons.
"""
import re
import base64
from typing import Optional, Dict, List, Set, cast

_AWS_SHAPE_PREFIX_RE = re.compile(r"^mxgraph\.aws\d*(?:\.|$)")


def is_aws_shape_type(shape_type: Optional[str]) -> bool:
    """Return True for draw.io AWS stencil prefixes: aws, aws2, aws3, aws4."""
    if not shape_type:
        return False
    return bool(_AWS_SHAPE_PREFIX_RE.match(shape_type.lower()))

# Fallback: MKAbuMattar/aws-icons via jsDelivr (official AWS icons, npm package)
# https://github.com/MKAbuMattar/aws-icons — architecture-service/ and resource/
_AWS_ICON_CDN_BASE = "https://cdn.jsdelivr.net/npm/aws-icons@latest/icons"
_ARCH = _AWS_ICON_CDN_BASE + "/architecture-service"
_RES = _AWS_ICON_CDN_BASE + "/resource"
_GROUP = _AWS_ICON_CDN_BASE + "/architecture-group"
_CATEGORY = _AWS_ICON_CDN_BASE + "/category"

# Fallback for specific icons: weibeld/aws-icons-svg (raw GitHub)
# Some icons (CloudWatch Event, SQS Queue) are more accurate in weibeld
_AWS_ICON_SVG_BASE = (
    "https://raw.githubusercontent.com/weibeld/aws-icons-svg/main/"
    "q1-2022/Resource-Icons_01312022"
)
_AWS4_STENCIL_XML_URL = (
    "https://raw.githubusercontent.com/jgraph/drawio/dev/src/main/webapp/stencils/aws4.xml"
)

def _url_spec(value: str) -> tuple[str, str]:
    return ("url", value)


def _aws4_spec(
    shape_name: str,
    background_hex: str,
    foreground_hex: str,
    canvas_w: float,
    canvas_h: float,
) -> tuple[str, str, str, str, float, float]:
    return ("aws4xml", shape_name, background_hex, foreground_hex, canvas_w, canvas_h)


def _group_icon_spec(
    ref: str,
    *,
    padding_ratio: float = 0.18,
    padding_color_mode: str = "stroke",
    match_label: Optional[str] = None,
    match_fill: Optional[str] = None,
) -> Dict[str, object]:
    """Spec for an aws4 group overlay icon entry.

    match_label / match_fill are optional conditions for variant lists.
    Entries without either condition act as the unconditional default.
    """
    entry: Dict[str, object] = {
        "spec": _url_spec(ref),
        "padding_ratio": float(padding_ratio),
        "padding_color_mode": padding_color_mode,
    }
    if match_label is not None:
        entry["match_label"] = match_label
    if match_fill is not None:
        entry["match_fill"] = match_fill
    return entry

# Unified mapping for known draw.io aws4 keys.
# Value format:
# - ("url", "<url-or-data-uri>")
# - ("aws4xml", "<shape_name>", "<bg_hex>", "<fg_hex>", canvas_w, canvas_h)
_AWS4_ICON_SPEC_BY_KEY: Dict[str, tuple] = {
    # Lambda service icon (architecture-service) vs Lambda Function resource icon (resource)
    "lambda": _url_spec(f"{_ARCH}/AWSLambda.svg"),
    "lambda_function": _url_spec(f"{_RES}/AWSLambdaLambdaFunction.svg"),
    "cloudwatch": _url_spec(f"{_ARCH}/AmazonCloudWatch.svg"),
    "amazon_cloudwatch": _url_spec(f"{_ARCH}/AmazonCloudWatch.svg"),
    "sns": _url_spec(f"{_ARCH}/AmazonSimpleNotificationService.svg"),
    "amazon_sns": _url_spec(f"{_ARCH}/AmazonSimpleNotificationService.svg"),
    "dynamodb": _url_spec(f"{_ARCH}/AmazonDynamoDB.svg"),
    "amazon_dynamodb": _url_spec(f"{_ARCH}/AmazonDynamoDB.svg"),
    # CloudWatch Event, SQS Queue, and S3 Bucket: use weibeld (more accurate icons)
    "queue": _url_spec(f"{_AWS_ICON_SVG_BASE}/Res_Application-Integration/Res_48_Light/Res_Amazon-Simple-Queue-Service_Queue_48_Light.svg"),
    "event_time_based": _url_spec(f"{_AWS_ICON_SVG_BASE}/Res_Management-Governance/Res_48_Light/Res_Amazon-CloudWatch_Event-Time-Based_48_Light.svg"),
    "event_event_based": _url_spec(f"{_AWS_ICON_SVG_BASE}/Res_Management-Governance/Res_48_Light/Res_Amazon-CloudWatch_Event-Event-Based_48_Light.svg"),
    "bucket_with_objects": _url_spec(f"{_AWS_ICON_SVG_BASE}/Res_Storage/Res_48_Light/Res_Amazon-Simple-Storage-Service_Bucket-With-Objects_48_Light.svg"),
    # Other icons: MKAbuMattar CDN
    "topic": _url_spec(f"{_RES}/AmazonSimpleNotificationServiceTopic.svg"),
    "template": _url_spec(f"{_RES}/AWSCloudFormationTemplate.svg"),
    "role": _url_spec(f"{_RES}/AWSIdentityAccessManagementRole.svg"),
    # AWS Config, GuardDuty, CloudTrail
    "config": _url_spec(f"{_ARCH}/AWSConfig.svg"),
    "aws_config": _url_spec(f"{_ARCH}/AWSConfig.svg"),
    "guardduty": _url_spec(f"{_ARCH}/AmazonGuardDuty.svg"),
    "amazon_guardduty": _url_spec(f"{_ARCH}/AmazonGuardDuty.svg"),
    "cloudtrail": _url_spec(f"{_ARCH}/AWSCloudTrail.svg"),
    "aws_cloudtrail": _url_spec(f"{_ARCH}/AWSCloudTrail.svg"),
    # Email and notification icons: use weibeld (correct colors, no fixed colors)
    "email": _url_spec(f"{_AWS_ICON_SVG_BASE}/Res_Business-Applications/Res_48_Light/Res_Amazon-Simple-Email-Service_Email_48_Light.svg"),
    "email_notification": _url_spec(f"{_AWS_ICON_SVG_BASE}/Res_Application-Integration/Res_48_Light/Res_Amazon-Simple-Notification-Service_Email-Notification_48_Light.svg"),
    # Rule icons
    "rule_2": _url_spec(f"{_RES}/AmazonCloudWatchRule.svg"),
    "rule": _url_spec(f"{_RES}/AmazonCloudWatchRule.svg"),
    # AWS / General Resources
    "resicon_marketplace": _url_spec(f"{_ARCH}/AWSMarketplace.svg"),
    "marketplace": _url_spec(f"{_ARCH}/AWSMarketplaceDark.svg"),
    "resicon_all_products": _aws4_spec("all products", "#232F3E", "#FFFFFF", 68.0, 68.0),
    "all_products": _aws4_spec("all products", "#FFFFFF", "#232F3D", 68.0, 68.0),
    "resicon_general": _aws4_spec("general", "#232F3E", "#FFFFFF", 64.0, 64.0),
    "general": _aws4_spec("general", "#FFFFFF", "#232F3D", 64.0, 64.0),
    "alert": _url_spec(f"{_RES}/Alert.svg"),
    "authenticated_user": _url_spec(f"{_RES}/AuthenticatedUser.svg"),
    "management_console2": _url_spec(f"{_RES}/AWSManagementConsole.svg"),
    "camera2": _url_spec(f"{_RES}/Camera.svg"),
    "chat": _url_spec(f"{_RES}/Chat.svg"),
    "client": _url_spec(f"{_RES}/Client.svg"),
    "cold_storage": _url_spec(f"{_RES}/ColdStorage.svg"),
    "credentials": _url_spec(f"{_RES}/Credentials.svg"),
    "corporate_data_center": _url_spec(f"{_RES}/Officebuilding.svg"),
    "data_stream": _url_spec(f"{_RES}/DataStream.svg"),
    "data_table": _url_spec(f"{_RES}/DataTable.svg"),
    "disk": _url_spec(f"{_RES}/Disk.svg"),
    "document": _url_spec(f"{_RES}/Document.svg"),
    "resicon_documents": _aws4_spec("documents", "#232F3E", "#FFFFFF", 64.0, 64.0),
    "documents": _aws4_spec("documents", "#FFFFFF", "#232F3D", 64.0, 64.0),
    "resicon_documents2": _aws4_spec("documents2", "#232F3E", "#FFFFFF", 64.0, 64.0),
    "documents2": _aws4_spec("documents2", "#FFFFFF", "#232F3D", 64.0, 64.0),
    "resicon_documents3": _url_spec(f"{_RES}/Documents.svg"),
    "documents3": _url_spec(f"{_RES}/Documents.svg"),
    "email_2": _url_spec(f"{_RES}/Email.svg"),
    "forums": _url_spec(f"{_RES}/Forums.svg"),
    "gear": _url_spec(f"{_RES}/Gear.svg"),
    "generic_application": _url_spec(f"{_RES}/GenericApplication.svg"),
    "generic_database": _url_spec(f"{_RES}/Database.svg"),
    "generic_firewall": _url_spec(f"{_RES}/Firewall.svg"),
    "git_repository": _url_spec(f"{_RES}/GitRepository.svg"),
    "globe": _url_spec(f"{_RES}/Globe.svg"),
    "folder": _url_spec(f"{_RES}/Folder.svg"),
    "folders": _url_spec(f"{_RES}/Folders.svg"),
    "internet": _url_spec(f"{_RES}/Internet.svg"),
    "internet_alt1": _url_spec(f"{_RES}/Internetalt1.svg"),
    "resicon_internet_alt2": _aws4_spec("internet alt2", "#232F3E", "#FFFFFF", 64.0, 64.0),
    "internet_alt2": _aws4_spec("internet alt2", "#FFFFFF", "#232F3D", 64.0, 64.0),
    "internet_alt22": _url_spec(f"{_RES}/Internetalt2.svg"),
    "json_script": _url_spec(f"{_RES}/JSONScript.svg"),
    "logs": _url_spec(f"{_RES}/Logs.svg"),
    "magnifying_glass_2": _url_spec(f"{_RES}/MagnifyingGlass.svg"),
    "metrics": _url_spec(f"{_RES}/Metrics.svg"),
    "mobile_client": _url_spec(f"{_RES}/Mobileclient.svg"),
    "multimedia": _url_spec(f"{_RES}/Multimedia.svg"),
    "office_building": _url_spec(f"{_RES}/Officebuilding.svg"),
    "programming_language": _url_spec(f"{_RES}/ProgrammingLanguage.svg"),
    "question": _url_spec(f"{_RES}/Question.svg"),
    "recover": _url_spec(f"{_RES}/Recover.svg"),
    "saml_token": _url_spec(f"{_RES}/SAMLtoken.svg"),
    "ssl_padlock": _url_spec(f"{_RES}/SSLpadlock.svg"),
    "tape_storage": _url_spec(f"{_RES}/Tapestorage.svg"),
    "traditional_server": _url_spec(f"{_RES}/Server.svg"),
    "user": _url_spec(f"{_RES}/User.svg"),
    "users": _url_spec(f"{_RES}/Users.svg"),
    "servers": _url_spec(f"{_RES}/Servers.svg"),
    "external_toolkit": _url_spec(f"{_RES}/Toolkit.svg"),
    "external_sdk": _url_spec(f"{_RES}/SDK.svg"),
    "shield2": _url_spec(f"{_RES}/Shield.svg"),
    "source_code": _url_spec(f"{_RES}/SourceCode.svg"),
    # AWS / Analytics category icons
    "analytics": _url_spec(f"{_CATEGORY}/Analytics.svg"),
    "athena": _url_spec(f"{_ARCH}/AmazonAthena.svg"),
    "amazon_athena": _url_spec(f"{_ARCH}/AmazonAthena.svg"),
    "datazone": _url_spec(f"{_ARCH}/AmazonDataZone.svg"),
    "amazon_datazone": _url_spec(f"{_ARCH}/AmazonDataZone.svg"),
    "cloudsearch2": _url_spec(f"{_ARCH}/AmazonCloudSearch.svg"),
    "amazon_cloudsearch": _url_spec(f"{_ARCH}/AmazonCloudSearch.svg"),
    "elasticsearch_service": _url_spec(f"{_ARCH}/AmazonOpenSearchService.svg"),
    "amazon_elasticsearch_service": _url_spec(f"{_ARCH}/AmazonOpenSearchService.svg"),
    "opensearch_service": _url_spec(f"{_ARCH}/AmazonOpenSearchService.svg"),
    "emr": _url_spec(f"{_ARCH}/AmazonEMR.svg"),
    "amazon_emr": _url_spec(f"{_ARCH}/AmazonEMR.svg"),
    "finspace": _url_spec(f"{_ARCH}/AmazonFinSpace.svg"),
    "amazon_finspace": _url_spec(f"{_ARCH}/AmazonFinSpace.svg"),
    "kinesis": _url_spec(f"{_ARCH}/AmazonKinesis.svg"),
    "amazon_kinesis": _url_spec(f"{_ARCH}/AmazonKinesis.svg"),
    # Legacy Kinesis Data Analytics/Firehose names in draw.io map to modern AWS icon names.
    "kinesis_data_analytics": _url_spec(f"{_ARCH}/AmazonManagedServiceforApacheFlink.svg"),
    "amazon_kinesis_data_analytics": _url_spec(f"{_ARCH}/AmazonManagedServiceforApacheFlink.svg"),
    "kinesis_data_firehose": _url_spec(f"{_ARCH}/AmazonDataFirehose.svg"),
    "amazon_kinesis_data_firehose": _url_spec(f"{_ARCH}/AmazonDataFirehose.svg"),
    "kinesis_data_streams": _url_spec(f"{_ARCH}/AmazonKinesisDataStreams.svg"),
    "amazon_kinesis_data_streams": _url_spec(f"{_ARCH}/AmazonKinesisDataStreams.svg"),
    "kinesis_video_streams": _url_spec(f"{_ARCH}/AmazonKinesisVideoStreams.svg"),
    "amazon_kinesis_video_streams": _url_spec(f"{_ARCH}/AmazonKinesisVideoStreams.svg"),
    "managed_service_for_apache_flink": _url_spec(f"{_ARCH}/AmazonManagedServiceforApacheFlink.svg"),
    "amazon_managed_service_for_apache_flink": _url_spec(f"{_ARCH}/AmazonManagedServiceforApacheFlink.svg"),
    "quicksight": _url_spec(f"{_ARCH}/AmazonQuickSight.svg"),
    "amazon_quicksight": _url_spec(f"{_ARCH}/AmazonQuickSight.svg"),
    "clean_rooms": _url_spec(f"{_ARCH}/AWSCleanRooms.svg"),
    "amazon_clean_rooms": _url_spec(f"{_ARCH}/AWSCleanRooms.svg"),
    "redshift": _url_spec(f"{_ARCH}/AmazonRedshift.svg"),
    "amazon_redshift": _url_spec(f"{_ARCH}/AmazonRedshift.svg"),
    "sagemaker_2": _url_spec(f"{_ARCH}/AmazonSageMaker.svg"),
    "sagemaker": _url_spec(f"{_ARCH}/AmazonSageMaker.svg"),
    "amazon_sagemaker": _url_spec(f"{_ARCH}/AmazonSageMaker.svg"),
    "data_pipeline": _url_spec(f"{_ARCH}/AWSDataPipeline.svg"),
    "aws_data_pipeline": _url_spec(f"{_ARCH}/AWSDataPipeline.svg"),
    "entity_resolution": _url_spec(f"{_ARCH}/AWSEntityResolution.svg"),
    "aws_entity_resolution": _url_spec(f"{_ARCH}/AWSEntityResolution.svg"),
    "managed_streaming_for_kafka": _url_spec(f"{_ARCH}/AmazonManagedStreamingforApacheKafka.svg"),
    "amazon_managed_streaming_for_kafka": _url_spec(f"{_ARCH}/AmazonManagedStreamingforApacheKafka.svg"),
    "glue": _url_spec(f"{_ARCH}/AWSGlue.svg"),
    "aws_glue": _url_spec(f"{_ARCH}/AWSGlue.svg"),
    # Keep legacy rendering for resourceIcon+resIcon=glue_databrew.
    # This matches existing diagrams that expect the old solid purple icon tile.
    "resicon_glue_databrew": _url_spec(f"{_ARCH}/AWSGlueDataBrew.svg"),
    # For direct shape=mxgraph.aws4.glue_databrew, use aws4 stencil rendering.
    # AWSGlueDataBrew has no dedicated *Dark.svg on aws-icons CDN.
    "glue_databrew": _aws4_spec("glue databrew", "#FFFFFF", "#8C4FFF", 56.0, 56.0),
    "aws_glue_databrew": _aws4_spec("glue databrew", "#FFFFFF", "#8C4FFF", 56.0, 56.0),
    "glue_elastic_views": _url_spec(f"{_ARCH}/AWSGlueElasticViews.svg"),
    "aws_glue_elastic_views": _url_spec(f"{_ARCH}/AWSGlueElasticViews.svg"),
    "lake_formation": _url_spec(f"{_ARCH}/AWSLakeFormation.svg"),
    "aws_lake_formation": _url_spec(f"{_ARCH}/AWSLakeFormation.svg"),
    "data_exchange": _url_spec(f"{_ARCH}/AWSDataExchange.svg"),
    "aws_data_exchange": _url_spec(f"{_ARCH}/AWSDataExchange.svg"),
    # resourceIcon + resIcon=sql_workbench should be the inverted tile style.
    "resicon_sql_workbench": _aws4_spec("sql workbench", "#8C4FFF", "#FFFFFF", 74.0, 74.0),
    "sql_workbench": _url_spec(f"{_RES}/AmazonRedshiftQueryEditorv20.svg"),
    "amazon_redshift_query_editor": _url_spec(f"{_RES}/AmazonRedshiftQueryEditorv20.svg"),
    # Analytics resource icons (specific resource types)
    "athena_data_source_connectors": _url_spec(f"{_RES}/AmazonAthenaDataSourceConnectors.svg"),
    "search_documents": _url_spec(f"{_RES}/AmazonCloudSearchSearchDocuments.svg"),
    "datazone_business_data_catalog": _url_spec(f"{_RES}/AmazonDataZoneBusinessDataCatalog.svg"),
    "datazone_data_portal": _url_spec(f"{_RES}/AmazonDataZoneDataPortal.svg"),
    "datazone_data_projects": _url_spec(f"{_RES}/AmazonDataZoneDataProjects.svg"),
    "cluster": _url_spec(f"{_RES}/AmazonEMRHDFSCluster.svg"),
    "msk_amazon_msk_connect": _url_spec(f"{_RES}/AmazonMSKAmazonMSKConnect.svg"),
    "opensearch_service_cluster_administrator_node": _url_spec(f"{_RES}/AmazonOpenSearchServiceClusterAdministratorNode.svg"),
    "opensearch_service_data_node": _url_spec(f"{_RES}/AmazonOpenSearchServiceDataNode.svg"),
    "opensearch_service_index": _url_spec(f"{_RES}/AmazonOpenSearchServiceIndex.svg"),
    "opensearch_observability": _url_spec(f"{_RES}/AmazonOpenSearchServiceObservability.svg"),
    "opensearch_dashboards": _url_spec(f"{_RES}/AmazonOpenSearchServiceOpenSearchDashboards.svg"),
    "opensearch_ingestion": _url_spec(f"{_RES}/AmazonOpenSearchServiceOpenSearchIngestion.svg"),
    "opensearch_service_traces": _url_spec(f"{_RES}/AmazonOpenSearchServiceTraces.svg"),
    "opensearch_service_ultrawarm_node": _url_spec(f"{_RES}/AmazonOpenSearchServiceUltraWarmNode.svg"),
    "quicksight_paginated_reports": _url_spec(f"{_RES}/AmazonQuicksightPaginatedReports.svg"),
    "redshift_auto_copy": _url_spec(f"{_RES}/AmazonRedshiftAutocopy.svg"),
    "redshift_data_sharing_governance": _url_spec(f"{_RES}/AmazonRedshiftDataSharingGovernance.svg"),
    "data_lake_resource_icon": _url_spec(f"{_RES}/AWSLakeFormationDataLake.svg"),
    "emr_engine": _url_spec(f"{_RES}/AmazonEMREMREngine.svg"),
    # MapR variants are not available as separate static SVG assets on aws-icons CDN.
    # Render dedicated aws4 stencil shapes so m3/m5/m7 can be reproduced distinctly.
    "emr_engine_mapr_m3": _aws4_spec("emr engine mapr m3", "none", "#8C4FFF", 78.109, 59.258),
    "emr_engine_mapr_m5": _aws4_spec("emr engine mapr m5", "none", "#8C4FFF", 78.109, 59.258),
    "emr_engine_mapr_m7": _aws4_spec("emr engine mapr m7", "none", "#8C4FFF", 78.109, 59.258),
    "hdfs_cluster": _url_spec(f"{_RES}/AmazonEMRCluster.svg"),
    "dense_compute_node": _url_spec(f"{_RES}/AmazonRedshiftDenseComputeNode.svg"),
    "dense_storage_node": _url_spec(f"{_RES}/AmazonRedshiftDenseStorageNode.svg"),
    "redshift_ra3": _url_spec(f"{_RES}/AmazonRedshiftRA3.svg"),
    "redshift_streaming_ingestion": _url_spec(f"{_RES}/AmazonRedshiftStreamingIngestion.svg"),
    "data_exchange_for_apis": _url_spec(f"{_RES}/AWSDataExchangeforAPIs.svg"),
    "aws_glue_for_ray": _url_spec(f"{_RES}/AWSGlueAWSGlueforRay.svg"),
    "glue_crawlers": _url_spec(f"{_RES}/AWSGlueCrawler.svg"),
    "glue_data_catalog": _url_spec(f"{_RES}/AWSGlueDataCatalog.svg"),
    "aws_glue_data_quality": _url_spec(f"{_RES}/AWSGlueDataQuality.svg"),
    "redshift_ml": _url_spec(f"{_RES}/AmazonRedshiftML.svg"),
    "redshift_query_editor_v20_light": _url_spec(f"{_RES}/AmazonRedshiftQueryEditorv20.svg"),
}

_AWS4_GROUP_CONFIG: Dict[str, object] = {
    "shape_types": {"mxgraph.aws4.group", "mxgraph.aws4.groupcenter"},
    "icons": {
    # AWS group/container icon overlays (draw.io style key: grIcon=mxgraph.aws4.group_*)
    # group_aws_cloud_alt is the "AWS" text variant in draw.io.
    "group_aws_cloud_alt": _group_icon_spec(f"{_GROUP}/AWSCloudlogo.svg"),
    "group_aws_cloud": _group_icon_spec(f"{_GROUP}/AWSCloud.svg"),
    "group_region": _group_icon_spec(f"{_GROUP}/Region.svg"),
    "group_auto_scaling_group": _group_icon_spec(
        f"{_GROUP}/AutoScalinggroup.svg",
        padding_color_mode="icon",
    ),
    "group_vpc2": _group_icon_spec(f"{_GROUP}/VirtualprivatecloudVPC.svg"),
    # draw.io uses group_security_group for both public/private subnets.
    # Listed in priority order; last entry is the unconditional default.
    "group_security_group": [
        _group_icon_spec(f"{_GROUP}/Publicsubnet.svg", match_label="public", match_fill="#f2f6e8"),
        _group_icon_spec(f"{_GROUP}/Privatesubnet.svg"),  # default
    ],
    "group_on_premise": _group_icon_spec(f"{_GROUP}/Servercontents.svg"),
    "group_corporate_data_center": _group_icon_spec(f"{_GROUP}/Corporatedatacenter.svg"),
    "group_elastic_beanstalk": _group_icon_spec(
        f"{_ARCH}/AWSElasticBeanstalk.svg",
        padding_color_mode="icon",
    ),
    "group_ec2_instance_contents": _group_icon_spec(
        f"{_GROUP}/EC2instancecontents.svg",
        padding_color_mode="icon",
    ),
    "group_spot_fleet": _group_icon_spec(
        f"{_GROUP}/SpotFleet.svg",
        padding_color_mode="icon",
    ),
    "group_aws_step_functions_workflow": _group_icon_spec(
        f"{_ARCH}/AWSStepFunctions.svg",
        padding_color_mode="icon",
    ),
    "group_account": _group_icon_spec(
        f"{_GROUP}/AWSAccount.svg",
        padding_color_mode="icon",
    ),
    "group_iot_greengrass_deployment": _group_icon_spec(f"{_GROUP}/AWSIoTGreengrassDeployment.svg"),
    "group_iot_greengrass": _group_icon_spec(f"{_ARCH}/AWSIoTGreengrass.svg"),
    },
}

_AWS4_GROUP_SHAPE_TYPES = cast(Set[str], _AWS4_GROUP_CONFIG["shape_types"])
_AWS4_GROUP_ICONS = cast(Dict[str, object], _AWS4_GROUP_CONFIG["icons"])


def _image_data_from_ref(ref: str):
    """Build ImageData from URL/data URI."""
    from ..model.intermediate import ImageData

    return ImageData(data_uri=ref) if ref.startswith("data:image/") else ImageData(file_path=ref)


def resolve_aws_group_metadata(
    shape_type: Optional[str],
    style_str: Optional[str] = None,
    label_text: Optional[str] = None,
):
    """
    Resolve AWS group/container metadata in one pass.

    Returns dict:
      - apply_text_padding: bool
      - group_icon_key: Optional[str]
      - group_icon_image_data: Optional[ImageData]
      - group_icon_padding_ratio: Optional[float]
      - group_icon_padding_color_mode: Optional[str] ("stroke"|"icon")
    """
    shape_type_lower = (shape_type or "").strip().lower()
    vertical_align = (_get_style_value(style_str, "verticalAlign") or "").strip().lower()
    apply_text_padding = (shape_type_lower in _AWS4_GROUP_SHAPE_TYPES) or (vertical_align == "top")

    group_key = None
    icon_cfg = None
    if shape_type_lower in _AWS4_GROUP_SHAPE_TYPES:
        group_key = ((_get_style_value(style_str, "grIcon") or "").split(".")[-1].strip().lower() or None)
        if group_key:
            entry = _AWS4_GROUP_ICONS.get(group_key)
            if isinstance(entry, list):
                label = (label_text or "").lower()
                fill = (_get_style_value(style_str, "fillColor") or "").lower()
                icon_cfg = next(
                    (e for e in entry if
                        not (e.get("match_label") or e.get("match_fill"))
                        or (e.get("match_label") and e["match_label"] in label)
                        or (e.get("match_fill") and e["match_fill"] == fill)
                    ),
                    entry[-1],
                )
            else:
                icon_cfg = entry

    icon_data = None
    padding_ratio = None
    padding_color_mode = None
    if icon_cfg:
        _, ref = icon_cfg["spec"]
        icon_data = _image_data_from_ref(ref)
        pr = icon_cfg.get("padding_ratio")
        padding_ratio = float(pr) if pr is not None else None
        padding_color_mode = str(icon_cfg.get("padding_color_mode", "stroke"))

    return {
        "apply_text_padding": bool(apply_text_padding),
        "group_icon_key": group_key,
        "group_icon_image_data": icon_data,
        "group_icon_padding_ratio": padding_ratio,
        "group_icon_padding_color_mode": padding_color_mode,
    }


def _shape_type_to_lookup_keys(shape_type: str, res_icon: Optional[str]) -> List[str]:
    """
    Map draw.io shape type and optional resIcon to stencil lookup keys to try.
    shape_type is normalized lower (e.g. mxgraph.aws4.lambda_function).
    res_icon is from style resIcon= (e.g. mxgraph.aws4.cloudwatch).
    """
    keys: List[str] = []
    if not is_aws_shape_type(shape_type):
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
            # Allow context-specific mapping for resourceIcon variants.
            if "resourceicon" in shape_type:
                keys.append(f"resicon_{icon_suffix}")
            keys.append(icon_suffix)
        # Some stencil labels are "Amazon X" so try "amazon_<icon>"
        if icon_suffix:
            keys.append(f"amazon_{icon_suffix}")

    return keys


def _get_style_value(style_str: Optional[str], key: str) -> Optional[str]:
    """Extract a single style value from draw.io style string."""
    if not style_str:
        return None
    for part in style_str.split(";"):
        if "=" in part:
            k, v = part.split("=", 1)
            if k.strip() == key:
                value = v.strip()
                return value or None
    return None


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

    if not is_aws_shape_type(shape_type):
        return None

    shape_type_lower = shape_type.lower()
    # aws4 group/groupCenter should be rendered as container + small overlay icon.
    # Do not resolve them as full-size shape images here.
    if shape_type_lower in _AWS4_GROUP_SHAPE_TYPES:
        return None

    res_icon = _get_style_value(style_str, "resIcon")
    keys = _shape_type_to_lookup_keys(shape_type_lower, res_icon)
    if not keys:
        return None

    # 1) Look up icon spec from unified mapping table
    for k in keys:
        if k in _AWS4_ICON_SPEC_BY_KEY:
            spec = _AWS4_ICON_SPEC_BY_KEY[k]
            if spec[0] == "aws4xml":
                _, shape_name, bg_hex, fg_hex, canvas_w, canvas_h = spec
                data_uri = _build_shape_data_uri_from_aws4(
                    shape_name=shape_name,
                    background_hex=bg_hex,
                    foreground_hex=fg_hex,
                    canvas_w=canvas_w,
                    canvas_h=canvas_h,
                )
                if data_uri:
                    return ImageData(data_uri=data_uri)
                continue
            _, ref = spec
            return _image_data_from_ref(ref)

    # 2) Dynamic fallback for AWS Illustrations (e.g. mxgraph.aws4.illustration_users).
    # These stencils are present in aws4.xml but not consistently available as static SVG files.
    shape_suffix = shape_type.split(".")[-1] if "." in shape_type else shape_type
    if shape_suffix.startswith("illustration_"):
        shape_name = shape_suffix.replace("_", " ")
        fg = _get_style_value(style_str, "fillColor") or "#879196"
        data_uri = _build_shape_data_uri_from_aws4(
            shape_name=shape_name,
            background_hex="none",
            foreground_hex=fg,
            canvas_w=100.0,
            canvas_h=100.0,
        )
        if data_uri:
            return ImageData(data_uri=data_uri)

    return None


def _build_shape_data_uri_from_aws4(
    *,
    shape_name: str,
    background_hex: str,
    foreground_hex: str,
    canvas_w: float,
    canvas_h: float,
) -> Optional[str]:
    """
    Build a data URI by fetching an aws4.xml shape and drawing it on a canvas.
    """
    spec = _fetch_shape_spec_from_aws4(shape_name)
    if not spec:
        return None
    path_d, shape_w, shape_h = spec

    # Prevent clipping when a stencil shape is larger than requested canvas.
    effective_canvas_w = max(canvas_w, shape_w)
    effective_canvas_h = max(canvas_h, shape_h)

    offset_x = (effective_canvas_w - shape_w) / 2.0
    offset_y = (effective_canvas_h - shape_h) / 2.0

    svg = (
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {effective_canvas_w:.3f} {effective_canvas_h:.3f}">'
        f'<rect x="0" y="0" width="{effective_canvas_w:.3f}" height="{effective_canvas_h:.3f}" fill="{background_hex}"/>'
        f'<path d="{path_d}" transform="translate({offset_x:.3f} {offset_y:.3f})" fill="{foreground_hex}" fill-rule="evenodd"/>'
        "</svg>"
    )
    encoded = base64.b64encode(svg.encode("utf-8")).decode("ascii")
    return f"data:image/svg+xml;base64,{encoded}"


def _fetch_shape_spec_from_aws4(shape_name: str) -> Optional[tuple[str, float, float]]:
    """
    Fetch and parse shape path data from draw.io official aws4.xml.
    """
    from ..media.image_utils import load_image_bytes

    xml_bytes = load_image_bytes(file_path=_AWS4_STENCIL_XML_URL)
    if not xml_bytes:
        return None
    xml_text = xml_bytes.decode("utf-8", errors="ignore")

    shape_match = re.search(
        rf'(<shape[^>]*name="{re.escape(shape_name)}"[^>]*>)(.*?)</shape>',
        xml_text,
        re.IGNORECASE | re.DOTALL,
    )
    if not shape_match:
        return None

    shape_tag = shape_match.group(1)
    shape_body = shape_match.group(2)
    w_match = re.search(r'\bw="([^"]+)"', shape_tag)
    h_match = re.search(r'\bh="([^"]+)"', shape_tag)
    if not w_match or not h_match:
        return None
    try:
        shape_w = float(w_match.group(1))
        shape_h = float(h_match.group(1))
    except ValueError:
        return None

    path_match = re.search(
        r"<path>(.*?)</path>", shape_body, re.IGNORECASE | re.DOTALL
    )
    if not path_match:
        return None

    commands: List[str] = []
    for tag, attrs in re.findall(
        r"<(move|line|curve|arc|close)\b([^>]*)/?>", path_match.group(1), re.IGNORECASE
    ):
        tag = tag.lower()
        kv = dict(re.findall(r'(x1|y1|x2|y2|x3|y3|x|y)="([^"]+)"', attrs))
        if tag == "move" and "x" in kv and "y" in kv:
            commands.append(f'M {kv["x"]} {kv["y"]}')
        elif tag == "line" and "x" in kv and "y" in kv:
            commands.append(f'L {kv["x"]} {kv["y"]}')
        elif tag == "curve" and all(
            k in kv for k in ("x1", "y1", "x2", "y2", "x3", "y3")
        ):
            commands.append(
                f'C {kv["x1"]} {kv["y1"]} {kv["x2"]} {kv["y2"]} {kv["x3"]} {kv["y3"]}'
            )
        elif tag == "arc":
            arc_kv = dict(
                re.findall(
                    r'(rx|ry|x-axis-rotation|large-arc-flag|sweep-flag|x|y)="([^"]+)"',
                    attrs,
                )
            )
            if all(
                k in arc_kv
                for k in (
                    "rx",
                    "ry",
                    "x-axis-rotation",
                    "large-arc-flag",
                    "sweep-flag",
                    "x",
                    "y",
                )
            ):
                commands.append(
                    f'A {arc_kv["rx"]} {arc_kv["ry"]} {arc_kv["x-axis-rotation"]} '
                    f'{arc_kv["large-arc-flag"]} {arc_kv["sweep-flag"]} {arc_kv["x"]} {arc_kv["y"]}'
                )
        elif tag == "close":
            commands.append("Z")

    if not commands:
        return None
    return " ".join(commands), shape_w, shape_h
