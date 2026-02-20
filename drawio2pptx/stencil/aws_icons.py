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

# Legacy alias mapping for known aws4 icon specs.
# New primary lookup path uses draw.io internal keys:
#   (shape=mxgraph.aws4.*) or
#   (shape=mxgraph.aws4.resourceIcon, resIcon=mxgraph.aws4.*)
# Value format:
# - ("url", "<url-or-data-uri>")
# - ("aws4xml", "<shape_name>", "<bg_hex>", "<fg_hex>", canvas_w, canvas_h)
# Draw.io-native AWS icon mapping.
# Key format: (shape_type, res_icon)
#   - direct shape: ("mxgraph.aws4.<shape>", None)
#   - resourceIcon: ("mxgraph.aws4.resourceicon", "mxgraph.aws4.<resIcon>")
# Value format: _url_spec(...) or _aws4_spec(...)
_AWS4_ICON_SPEC_BY_DRAWIO_KEY: Dict[tuple[str, Optional[str]], tuple] = {
    ("mxgraph.aws4.lambda", None): _url_spec(f"{_ARCH}/AWSLambda.svg"),
    ("mxgraph.aws4.lambda_function", None): _url_spec(f"{_RES}/AWSLambdaLambdaFunction.svg"),
    ("mxgraph.aws4.cloudwatch", None): _url_spec(f"{_ARCH}/AmazonCloudWatch.svg"),
    ("mxgraph.aws4.amazon_cloudwatch", None): _url_spec(f"{_ARCH}/AmazonCloudWatch.svg"),
    ("mxgraph.aws4.sns", None): _url_spec(f"{_ARCH}/AmazonSimpleNotificationService.svg"),
    ("mxgraph.aws4.amazon_sns", None): _url_spec(f"{_ARCH}/AmazonSimpleNotificationService.svg"),
    ("mxgraph.aws4.dynamodb", None): _url_spec(f"{_ARCH}/AmazonDynamoDB.svg"),
    ("mxgraph.aws4.amazon_dynamodb", None): _url_spec(f"{_ARCH}/AmazonDynamoDB.svg"),
    ("mxgraph.aws4.queue", None): _url_spec(f"{_AWS_ICON_SVG_BASE}/Res_Application-Integration/Res_48_Light/Res_Amazon-Simple-Queue-Service_Queue_48_Light.svg"),
    ("mxgraph.aws4.event_time_based", None): _url_spec(f"{_AWS_ICON_SVG_BASE}/Res_Management-Governance/Res_48_Light/Res_Amazon-CloudWatch_Event-Time-Based_48_Light.svg"),
    ("mxgraph.aws4.event_event_based", None): _url_spec(f"{_AWS_ICON_SVG_BASE}/Res_Management-Governance/Res_48_Light/Res_Amazon-CloudWatch_Event-Event-Based_48_Light.svg"),
    ("mxgraph.aws4.bucket_with_objects", None): _url_spec(f"{_AWS_ICON_SVG_BASE}/Res_Storage/Res_48_Light/Res_Amazon-Simple-Storage-Service_Bucket-With-Objects_48_Light.svg"),
    ("mxgraph.aws4.topic", None): _url_spec(f"{_RES}/AmazonSimpleNotificationServiceTopic.svg"),
    ("mxgraph.aws4.template", None): _url_spec(f"{_RES}/AWSCloudFormationTemplate.svg"),
    ("mxgraph.aws4.role", None): _url_spec(f"{_RES}/AWSIdentityAccessManagementRole.svg"),
    ("mxgraph.aws4.config", None): _url_spec(f"{_ARCH}/AWSConfig.svg"),
    ("mxgraph.aws4.aws_config", None): _url_spec(f"{_ARCH}/AWSConfig.svg"),
    ("mxgraph.aws4.guardduty", None): _url_spec(f"{_ARCH}/AmazonGuardDuty.svg"),
    ("mxgraph.aws4.amazon_guardduty", None): _url_spec(f"{_ARCH}/AmazonGuardDuty.svg"),
    ("mxgraph.aws4.cloudtrail", None): _url_spec(f"{_ARCH}/AWSCloudTrail.svg"),
    ("mxgraph.aws4.aws_cloudtrail", None): _url_spec(f"{_ARCH}/AWSCloudTrail.svg"),
    ("mxgraph.aws4.email", None): _url_spec(f"{_AWS_ICON_SVG_BASE}/Res_Business-Applications/Res_48_Light/Res_Amazon-Simple-Email-Service_Email_48_Light.svg"),
    ("mxgraph.aws4.email_notification", None): _url_spec(f"{_AWS_ICON_SVG_BASE}/Res_Application-Integration/Res_48_Light/Res_Amazon-Simple-Notification-Service_Email-Notification_48_Light.svg"),
    ("mxgraph.aws4.rule_2", None): _url_spec(f"{_RES}/AmazonCloudWatchRule.svg"),
    ("mxgraph.aws4.rule", None): _url_spec(f"{_RES}/AmazonCloudWatchRule.svg"),
    ("mxgraph.aws4.resourceicon", 'mxgraph.aws4.marketplace'): _url_spec(f"{_ARCH}/AWSMarketplace.svg"),
    ("mxgraph.aws4.marketplace", None): _url_spec(f"{_ARCH}/AWSMarketplaceDark.svg"),
    ("mxgraph.aws4.resourceicon", 'mxgraph.aws4.all_products'): _aws4_spec("all products", "#232F3E", "#FFFFFF", 68.0, 68.0),
    ("mxgraph.aws4.all_products", None): _aws4_spec("all products", "#FFFFFF", "#232F3D", 68.0, 68.0),
    ("mxgraph.aws4.resourceicon", 'mxgraph.aws4.general'): _aws4_spec("general", "#232F3E", "#FFFFFF", 64.0, 64.0),
    ("mxgraph.aws4.general", None): _aws4_spec("general", "#FFFFFF", "#232F3D", 64.0, 64.0),
    ("mxgraph.aws4.alert", None): _url_spec(f"{_RES}/Alert.svg"),
    ("mxgraph.aws4.authenticated_user", None): _url_spec(f"{_RES}/AuthenticatedUser.svg"),
    ("mxgraph.aws4.management_console2", None): _url_spec(f"{_RES}/AWSManagementConsole.svg"),
    ("mxgraph.aws4.camera2", None): _url_spec(f"{_RES}/Camera.svg"),
    ("mxgraph.aws4.chat", None): _url_spec(f"{_RES}/Chat.svg"),
    ("mxgraph.aws4.client", None): _url_spec(f"{_RES}/Client.svg"),
    ("mxgraph.aws4.cold_storage", None): _url_spec(f"{_RES}/ColdStorage.svg"),
    ("mxgraph.aws4.credentials", None): _url_spec(f"{_RES}/Credentials.svg"),
    ("mxgraph.aws4.corporate_data_center", None): _url_spec(f"{_RES}/Officebuilding.svg"),
    ("mxgraph.aws4.data_stream", None): _url_spec(f"{_RES}/DataStream.svg"),
    ("mxgraph.aws4.data_table", None): _url_spec(f"{_RES}/DataTable.svg"),
    ("mxgraph.aws4.disk", None): _url_spec(f"{_RES}/Disk.svg"),
    ("mxgraph.aws4.document", None): _url_spec(f"{_RES}/Document.svg"),
    ("mxgraph.aws4.resourceicon", 'mxgraph.aws4.documents'): _aws4_spec("documents", "#232F3E", "#FFFFFF", 64.0, 64.0),
    ("mxgraph.aws4.documents", None): _aws4_spec("documents", "#FFFFFF", "#232F3D", 64.0, 64.0),
    ("mxgraph.aws4.resourceicon", 'mxgraph.aws4.documents2'): _aws4_spec("documents2", "#232F3E", "#FFFFFF", 64.0, 64.0),
    ("mxgraph.aws4.documents2", None): _aws4_spec("documents2", "#FFFFFF", "#232F3D", 64.0, 64.0),
    ("mxgraph.aws4.resourceicon", 'mxgraph.aws4.documents3'): _url_spec(f"{_RES}/Documents.svg"),
    ("mxgraph.aws4.documents3", None): _url_spec(f"{_RES}/Documents.svg"),
    ("mxgraph.aws4.email_2", None): _url_spec(f"{_RES}/Email.svg"),
    ("mxgraph.aws4.forums", None): _url_spec(f"{_RES}/Forums.svg"),
    ("mxgraph.aws4.gear", None): _url_spec(f"{_RES}/Gear.svg"),
    ("mxgraph.aws4.generic_application", None): _url_spec(f"{_RES}/GenericApplication.svg"),
    ("mxgraph.aws4.generic_database", None): _url_spec(f"{_RES}/Database.svg"),
    ("mxgraph.aws4.generic_firewall", None): _url_spec(f"{_RES}/Firewall.svg"),
    ("mxgraph.aws4.git_repository", None): _url_spec(f"{_RES}/GitRepository.svg"),
    ("mxgraph.aws4.globe", None): _url_spec(f"{_RES}/Globe.svg"),
    ("mxgraph.aws4.folder", None): _url_spec(f"{_RES}/Folder.svg"),
    ("mxgraph.aws4.folders", None): _url_spec(f"{_RES}/Folders.svg"),
    ("mxgraph.aws4.internet", None): _url_spec(f"{_RES}/Internet.svg"),
    ("mxgraph.aws4.internet_alt1", None): _url_spec(f"{_RES}/Internetalt1.svg"),
    ("mxgraph.aws4.resourceicon", 'mxgraph.aws4.internet_alt2'): _aws4_spec("internet alt2", "#232F3E", "#FFFFFF", 64.0, 64.0),
    ("mxgraph.aws4.internet_alt2", None): _aws4_spec("internet alt2", "#FFFFFF", "#232F3D", 64.0, 64.0),
    ("mxgraph.aws4.internet_alt22", None): _url_spec(f"{_RES}/Internetalt2.svg"),
    ("mxgraph.aws4.json_script", None): _url_spec(f"{_RES}/JSONScript.svg"),
    ("mxgraph.aws4.logs", None): _url_spec(f"{_RES}/Logs.svg"),
    ("mxgraph.aws4.magnifying_glass_2", None): _url_spec(f"{_RES}/MagnifyingGlass.svg"),
    ("mxgraph.aws4.metrics", None): _url_spec(f"{_RES}/Metrics.svg"),
    ("mxgraph.aws4.mobile_client", None): _url_spec(f"{_RES}/Mobileclient.svg"),
    ("mxgraph.aws4.multimedia", None): _url_spec(f"{_RES}/Multimedia.svg"),
    ("mxgraph.aws4.office_building", None): _url_spec(f"{_RES}/Officebuilding.svg"),
    ("mxgraph.aws4.programming_language", None): _url_spec(f"{_RES}/ProgrammingLanguage.svg"),
    ("mxgraph.aws4.question", None): _url_spec(f"{_RES}/Question.svg"),
    ("mxgraph.aws4.recover", None): _url_spec(f"{_RES}/Recover.svg"),
    ("mxgraph.aws4.saml_token", None): _url_spec(f"{_RES}/SAMLtoken.svg"),
    ("mxgraph.aws4.ssl_padlock", None): _url_spec(f"{_RES}/SSLpadlock.svg"),
    ("mxgraph.aws4.tape_storage", None): _url_spec(f"{_RES}/Tapestorage.svg"),
    ("mxgraph.aws4.traditional_server", None): _url_spec(f"{_RES}/Server.svg"),
    ("mxgraph.aws4.user", None): _url_spec(f"{_RES}/User.svg"),
    ("mxgraph.aws4.users", None): _url_spec(f"{_RES}/Users.svg"),
    ("mxgraph.aws4.servers", None): _url_spec(f"{_RES}/Servers.svg"),
    ("mxgraph.aws4.external_toolkit", None): _url_spec(f"{_RES}/Toolkit.svg"),
    ("mxgraph.aws4.external_sdk", None): _url_spec(f"{_RES}/SDK.svg"),
    ("mxgraph.aws4.shield2", None): _url_spec(f"{_RES}/Shield.svg"),
    ("mxgraph.aws4.source_code", None): _url_spec(f"{_RES}/SourceCode.svg"),
    # category SVG has a relatively wide frame; use aws4 stencil path to reduce visual padding
    ("mxgraph.aws4.application_integration", None): _aws4_spec("application integration", "#E7157B", "#FFFFFF", 70.0, 70.0),
    # AR & VR category/service are still present in draw.io aws4.xml,
    # but not in current aws-icons static SVG package.
    ("mxgraph.aws4.resourceicon", 'mxgraph.aws4.ar_vr'): _aws4_spec("ar vr", "#BC1356", "#FFFFFF", 70.0, 70.0),
    ("mxgraph.aws4.ar_vr", None): _aws4_spec("ar vr", "#FFFFFF", "#BC1356", 70.0, 70.0),
    ("mxgraph.aws4.resourceicon", 'mxgraph.aws4.sumerian'): _aws4_spec("sumerian", "#BC1356", "#FFFFFF", 64.0, 64.0),
    ("mxgraph.aws4.sumerian", None): _aws4_spec("sumerian", "#FFFFFF", "#BC1356", 64.0, 64.0),
    ("mxgraph.aws4.api_gateway", None): _url_spec(f"{_ARCH}/AmazonAPIGateway.svg"),
    ("mxgraph.aws4.amazon_api_gateway", None): _url_spec(f"{_ARCH}/AmazonAPIGateway.svg"),
    ("mxgraph.aws4.mq", None): _url_spec(f"{_ARCH}/AmazonMQ.svg"),
    ("mxgraph.aws4.amazon_mq", None): _url_spec(f"{_ARCH}/AmazonMQ.svg"),
    ("mxgraph.aws4.sqs", None): _url_spec(f"{_ARCH}/AmazonSimpleQueueService.svg"),
    ("mxgraph.aws4.amazon_sqs", None): _url_spec(f"{_ARCH}/AmazonSimpleQueueService.svg"),
    ("mxgraph.aws4.appsync", None): _url_spec(f"{_ARCH}/AWSAppSync.svg"),
    ("mxgraph.aws4.amazon_appsync", None): _url_spec(f"{_ARCH}/AWSAppSync.svg"),
    ("mxgraph.aws4.b2b_data_interchange", None): _url_spec(f"{_ARCH}/AWSB2BDataInterchange.svg"),
    ("mxgraph.aws4.amazon_b2b_data_interchange", None): _url_spec(f"{_ARCH}/AWSB2BDataInterchange.svg"),
    ("mxgraph.aws4.eventbridge", None): _url_spec(f"{_ARCH}/AmazonEventBridge.svg"),
    ("mxgraph.aws4.amazon_eventbridge", None): _url_spec(f"{_ARCH}/AmazonEventBridge.svg"),
    ("mxgraph.aws4.managed_workflows_for_apache_airflow", None): _url_spec(f"{_ARCH}/AmazonManagedWorkflowsforApacheAirflow.svg"),
    ("mxgraph.aws4.amazon_managed_workflows_for_apache_airflow", None): _url_spec(f"{_ARCH}/AmazonManagedWorkflowsforApacheAirflow.svg"),
    ("mxgraph.aws4.step_functions", None): _url_spec(f"{_ARCH}/AWSStepFunctions.svg"),
    ("mxgraph.aws4.amazon_step_functions", None): _url_spec(f"{_ARCH}/AWSStepFunctions.svg"),
    ("mxgraph.aws4.mobile_application", None): _url_spec(f"{_ARCH}/AWSConsoleMobileApplication.svg"),
    ("mxgraph.aws4.amazon_mobile_application", None): _url_spec(f"{_ARCH}/AWSConsoleMobileApplication.svg"),
    ("mxgraph.aws4.express_workflow", None): _url_spec(f"{_ARCH}/AWSExpressWorkflows.svg"),
    ("mxgraph.aws4.amazon_express_workflow", None): _url_spec(f"{_ARCH}/AWSExpressWorkflows.svg"),
    ("mxgraph.aws4.appflow", None): _url_spec(f"{_ARCH}/AmazonAppFlow.svg"),
    ("mxgraph.aws4.amazon_appflow", None): _url_spec(f"{_ARCH}/AmazonAppFlow.svg"),
    ("mxgraph.aws4.endpoint", None): _url_spec(f"{_RES}/AmazonAPIGatewayEndpoint.svg"),
    ("mxgraph.aws4.event", None): _url_spec(f"{_RES}/AmazonEventBridgeEvent.svg"),
    ("mxgraph.aws4.eventbridge_pipes", None): _url_spec(f"{_RES}/AmazonEventBridgePipes.svg"),
    ("mxgraph.aws4.eventbridge_custom_event_bus_resource", None): _url_spec(f"{_RES}/AmazonEventBridgeCustomEventBus.svg"),
    ("mxgraph.aws4.eventbridge_default_event_bus_resource", None): _url_spec(f"{_RES}/AmazonEventBridgeDefaultEventBus.svg"),
    ("mxgraph.aws4.eventbridge_saas_partner_event_bus_resource", None): _url_spec(f"{_RES}/AmazonEventBridgeSaasPartnerEvent.svg"),
    ("mxgraph.aws4.eventbridge_scheduler", None): _url_spec(f"{_RES}/AmazonEventBridgeScheduler.svg"),
    ("mxgraph.aws4.eventbridge_schema", None): _url_spec(f"{_RES}/AmazonEventBridgeSchema.svg"),
    ("mxgraph.aws4.eventbridge_schema_registry", None): _url_spec(f"{_RES}/AmazonEventBridgeSchemaRegistry.svg"),
    ("mxgraph.aws4.mq_broker", None): _url_spec(f"{_RES}/AmazonMQBroker.svg"),
    ("mxgraph.aws4.event_resource", None): _url_spec(f"{_RES}/AmazonEventBridgeEvent.svg"),
    ("mxgraph.aws4.http_notification", None): _url_spec(f"{_RES}/AmazonSimpleNotificationServiceHTTPNotification.svg"),
    ("mxgraph.aws4.message", None): _url_spec(f"{_RES}/AmazonSimpleQueueServiceMessage.svg"),
    ("mxgraph.aws4.rule_3", None): _url_spec(f"{_RES}/AmazonEventBridgeRule.svg"),
    ("mxgraph.aws4.analytics", None): _url_spec(f"{_CATEGORY}/Analytics.svg"),
    ("mxgraph.aws4.athena", None): _url_spec(f"{_ARCH}/AmazonAthena.svg"),
    ("mxgraph.aws4.amazon_athena", None): _url_spec(f"{_ARCH}/AmazonAthena.svg"),
    ("mxgraph.aws4.datazone", None): _url_spec(f"{_ARCH}/AmazonDataZone.svg"),
    ("mxgraph.aws4.amazon_datazone", None): _url_spec(f"{_ARCH}/AmazonDataZone.svg"),
    ("mxgraph.aws4.cloudsearch2", None): _url_spec(f"{_ARCH}/AmazonCloudSearch.svg"),
    ("mxgraph.aws4.amazon_cloudsearch", None): _url_spec(f"{_ARCH}/AmazonCloudSearch.svg"),
    ("mxgraph.aws4.elasticsearch_service", None): _url_spec(f"{_ARCH}/AmazonOpenSearchService.svg"),
    ("mxgraph.aws4.amazon_elasticsearch_service", None): _url_spec(f"{_ARCH}/AmazonOpenSearchService.svg"),
    ("mxgraph.aws4.opensearch_service", None): _url_spec(f"{_ARCH}/AmazonOpenSearchService.svg"),
    ("mxgraph.aws4.emr", None): _url_spec(f"{_ARCH}/AmazonEMR.svg"),
    ("mxgraph.aws4.amazon_emr", None): _url_spec(f"{_ARCH}/AmazonEMR.svg"),
    ("mxgraph.aws4.finspace", None): _url_spec(f"{_ARCH}/AmazonFinSpace.svg"),
    ("mxgraph.aws4.amazon_finspace", None): _url_spec(f"{_ARCH}/AmazonFinSpace.svg"),
    ("mxgraph.aws4.kinesis", None): _url_spec(f"{_ARCH}/AmazonKinesis.svg"),
    ("mxgraph.aws4.amazon_kinesis", None): _url_spec(f"{_ARCH}/AmazonKinesis.svg"),
    ("mxgraph.aws4.kinesis_data_analytics", None): _url_spec(f"{_ARCH}/AmazonManagedServiceforApacheFlink.svg"),
    ("mxgraph.aws4.amazon_kinesis_data_analytics", None): _url_spec(f"{_ARCH}/AmazonManagedServiceforApacheFlink.svg"),
    ("mxgraph.aws4.kinesis_data_firehose", None): _url_spec(f"{_ARCH}/AmazonDataFirehose.svg"),
    ("mxgraph.aws4.amazon_kinesis_data_firehose", None): _url_spec(f"{_ARCH}/AmazonDataFirehose.svg"),
    ("mxgraph.aws4.kinesis_data_streams", None): _url_spec(f"{_ARCH}/AmazonKinesisDataStreams.svg"),
    ("mxgraph.aws4.amazon_kinesis_data_streams", None): _url_spec(f"{_ARCH}/AmazonKinesisDataStreams.svg"),
    ("mxgraph.aws4.kinesis_video_streams", None): _url_spec(f"{_ARCH}/AmazonKinesisVideoStreams.svg"),
    ("mxgraph.aws4.amazon_kinesis_video_streams", None): _url_spec(f"{_ARCH}/AmazonKinesisVideoStreams.svg"),
    ("mxgraph.aws4.managed_service_for_apache_flink", None): _url_spec(f"{_ARCH}/AmazonManagedServiceforApacheFlink.svg"),
    ("mxgraph.aws4.amazon_managed_service_for_apache_flink", None): _url_spec(f"{_ARCH}/AmazonManagedServiceforApacheFlink.svg"),
    ("mxgraph.aws4.quicksight", None): _url_spec(f"{_ARCH}/AmazonQuickSight.svg"),
    ("mxgraph.aws4.amazon_quicksight", None): _url_spec(f"{_ARCH}/AmazonQuickSight.svg"),
    ("mxgraph.aws4.clean_rooms", None): _url_spec(f"{_ARCH}/AWSCleanRooms.svg"),
    ("mxgraph.aws4.amazon_clean_rooms", None): _url_spec(f"{_ARCH}/AWSCleanRooms.svg"),
    ("mxgraph.aws4.redshift", None): _url_spec(f"{_ARCH}/AmazonRedshift.svg"),
    ("mxgraph.aws4.amazon_redshift", None): _url_spec(f"{_ARCH}/AmazonRedshift.svg"),
    ("mxgraph.aws4.sagemaker_2", None): _url_spec(f"{_ARCH}/AmazonSageMaker.svg"),
    ("mxgraph.aws4.sagemaker", None): _url_spec(f"{_ARCH}/AmazonSageMaker.svg"),
    ("mxgraph.aws4.amazon_sagemaker", None): _url_spec(f"{_ARCH}/AmazonSageMaker.svg"),
    ("mxgraph.aws4.data_pipeline", None): _url_spec(f"{_ARCH}/AWSDataPipeline.svg"),
    ("mxgraph.aws4.aws_data_pipeline", None): _url_spec(f"{_ARCH}/AWSDataPipeline.svg"),
    ("mxgraph.aws4.entity_resolution", None): _url_spec(f"{_ARCH}/AWSEntityResolution.svg"),
    ("mxgraph.aws4.aws_entity_resolution", None): _url_spec(f"{_ARCH}/AWSEntityResolution.svg"),
    ("mxgraph.aws4.managed_streaming_for_kafka", None): _url_spec(f"{_ARCH}/AmazonManagedStreamingforApacheKafka.svg"),
    ("mxgraph.aws4.amazon_managed_streaming_for_kafka", None): _url_spec(f"{_ARCH}/AmazonManagedStreamingforApacheKafka.svg"),
    ("mxgraph.aws4.glue", None): _url_spec(f"{_ARCH}/AWSGlue.svg"),
    ("mxgraph.aws4.aws_glue", None): _url_spec(f"{_ARCH}/AWSGlue.svg"),
    ("mxgraph.aws4.resourceicon", 'mxgraph.aws4.glue_databrew'): _url_spec(f"{_ARCH}/AWSGlueDataBrew.svg"),
    ("mxgraph.aws4.glue_databrew", None): _aws4_spec("glue databrew", "#FFFFFF", "#8C4FFF", 56.0, 56.0),
    ("mxgraph.aws4.aws_glue_databrew", None): _aws4_spec("glue databrew", "#FFFFFF", "#8C4FFF", 56.0, 56.0),
    ("mxgraph.aws4.glue_elastic_views", None): _url_spec(f"{_ARCH}/AWSGlueElasticViews.svg"),
    ("mxgraph.aws4.aws_glue_elastic_views", None): _url_spec(f"{_ARCH}/AWSGlueElasticViews.svg"),
    ("mxgraph.aws4.lake_formation", None): _url_spec(f"{_ARCH}/AWSLakeFormation.svg"),
    ("mxgraph.aws4.aws_lake_formation", None): _url_spec(f"{_ARCH}/AWSLakeFormation.svg"),
    ("mxgraph.aws4.data_exchange", None): _url_spec(f"{_ARCH}/AWSDataExchange.svg"),
    ("mxgraph.aws4.aws_data_exchange", None): _url_spec(f"{_ARCH}/AWSDataExchange.svg"),
    ("mxgraph.aws4.resourceicon", 'mxgraph.aws4.sql_workbench'): _aws4_spec("sql workbench", "#8C4FFF", "#FFFFFF", 74.0, 74.0),
    ("mxgraph.aws4.sql_workbench", None): _url_spec(f"{_RES}/AmazonRedshiftQueryEditorv20.svg"),
    ("mxgraph.aws4.amazon_redshift_query_editor", None): _url_spec(f"{_RES}/AmazonRedshiftQueryEditorv20.svg"),
    ("mxgraph.aws4.athena_data_source_connectors", None): _url_spec(f"{_RES}/AmazonAthenaDataSourceConnectors.svg"),
    ("mxgraph.aws4.search_documents", None): _url_spec(f"{_RES}/AmazonCloudSearchSearchDocuments.svg"),
    ("mxgraph.aws4.datazone_business_data_catalog", None): _url_spec(f"{_RES}/AmazonDataZoneBusinessDataCatalog.svg"),
    ("mxgraph.aws4.datazone_data_portal", None): _url_spec(f"{_RES}/AmazonDataZoneDataPortal.svg"),
    ("mxgraph.aws4.datazone_data_projects", None): _url_spec(f"{_RES}/AmazonDataZoneDataProjects.svg"),
    ("mxgraph.aws4.cluster", None): _url_spec(f"{_RES}/AmazonEMRHDFSCluster.svg"),
    ("mxgraph.aws4.msk_amazon_msk_connect", None): _url_spec(f"{_RES}/AmazonMSKAmazonMSKConnect.svg"),
    ("mxgraph.aws4.opensearch_service_cluster_administrator_node", None): _url_spec(f"{_RES}/AmazonOpenSearchServiceClusterAdministratorNode.svg"),
    ("mxgraph.aws4.opensearch_service_data_node", None): _url_spec(f"{_RES}/AmazonOpenSearchServiceDataNode.svg"),
    ("mxgraph.aws4.opensearch_service_index", None): _url_spec(f"{_RES}/AmazonOpenSearchServiceIndex.svg"),
    ("mxgraph.aws4.opensearch_observability", None): _url_spec(f"{_RES}/AmazonOpenSearchServiceObservability.svg"),
    ("mxgraph.aws4.opensearch_dashboards", None): _url_spec(f"{_RES}/AmazonOpenSearchServiceOpenSearchDashboards.svg"),
    ("mxgraph.aws4.opensearch_ingestion", None): _url_spec(f"{_RES}/AmazonOpenSearchServiceOpenSearchIngestion.svg"),
    ("mxgraph.aws4.opensearch_service_traces", None): _url_spec(f"{_RES}/AmazonOpenSearchServiceTraces.svg"),
    ("mxgraph.aws4.opensearch_service_ultrawarm_node", None): _url_spec(f"{_RES}/AmazonOpenSearchServiceUltraWarmNode.svg"),
    ("mxgraph.aws4.quicksight_paginated_reports", None): _url_spec(f"{_RES}/AmazonQuicksightPaginatedReports.svg"),
    ("mxgraph.aws4.redshift_auto_copy", None): _url_spec(f"{_RES}/AmazonRedshiftAutocopy.svg"),
    ("mxgraph.aws4.redshift_data_sharing_governance", None): _url_spec(f"{_RES}/AmazonRedshiftDataSharingGovernance.svg"),
    ("mxgraph.aws4.data_lake_resource_icon", None): _url_spec(f"{_RES}/AWSLakeFormationDataLake.svg"),
    ("mxgraph.aws4.emr_engine", None): _url_spec(f"{_RES}/AmazonEMREMREngine.svg"),
    ("mxgraph.aws4.emr_engine_mapr_m3", None): _aws4_spec("emr engine mapr m3", "none", "#8C4FFF", 78.109, 59.258),
    ("mxgraph.aws4.emr_engine_mapr_m5", None): _aws4_spec("emr engine mapr m5", "none", "#8C4FFF", 78.109, 59.258),
    ("mxgraph.aws4.emr_engine_mapr_m7", None): _aws4_spec("emr engine mapr m7", "none", "#8C4FFF", 78.109, 59.258),
    ("mxgraph.aws4.hdfs_cluster", None): _url_spec(f"{_RES}/AmazonEMRCluster.svg"),
    ("mxgraph.aws4.dense_compute_node", None): _url_spec(f"{_RES}/AmazonRedshiftDenseComputeNode.svg"),
    ("mxgraph.aws4.dense_storage_node", None): _url_spec(f"{_RES}/AmazonRedshiftDenseStorageNode.svg"),
    ("mxgraph.aws4.redshift_ra3", None): _url_spec(f"{_RES}/AmazonRedshiftRA3.svg"),
    ("mxgraph.aws4.redshift_streaming_ingestion", None): _url_spec(f"{_RES}/AmazonRedshiftStreamingIngestion.svg"),
    ("mxgraph.aws4.data_exchange_for_apis", None): _url_spec(f"{_RES}/AWSDataExchangeforAPIs.svg"),
    ("mxgraph.aws4.aws_glue_for_ray", None): _url_spec(f"{_RES}/AWSGlueAWSGlueforRay.svg"),
    ("mxgraph.aws4.glue_crawlers", None): _url_spec(f"{_RES}/AWSGlueCrawler.svg"),
    ("mxgraph.aws4.glue_data_catalog", None): _url_spec(f"{_RES}/AWSGlueDataCatalog.svg"),
    ("mxgraph.aws4.aws_glue_data_quality", None): _url_spec(f"{_RES}/AWSGlueDataQuality.svg"),
    ("mxgraph.aws4.redshift_ml", None): _url_spec(f"{_RES}/AmazonRedshiftML.svg"),
    ("mxgraph.aws4.redshift_query_editor_v20_light", None): _url_spec(f"{_RES}/AmazonRedshiftQueryEditorv20.svg"),
}


def _normalize_drawio_aws_value(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    norm = value.strip().lower()
    return norm or None

_AWS4_GROUP_CONFIG: Dict[str, object] = {
    "shape_types": {"mxgraph.aws4.group", "mxgraph.aws4.groupcenter"},
    "icons": {
    # AWS group/container icon overlays (draw.io style key: grIcon=mxgraph.aws4.group_*)
    # mxgraph.aws4.group_aws_cloud_alt is the "AWS" text variant in draw.io.
    "mxgraph.aws4.group_aws_cloud_alt": _group_icon_spec(f"{_GROUP}/AWSCloudlogo.svg"),
    "mxgraph.aws4.group_aws_cloud": _group_icon_spec(f"{_GROUP}/AWSCloud.svg"),
    "mxgraph.aws4.group_region": _group_icon_spec(f"{_GROUP}/Region.svg"),
    "mxgraph.aws4.group_auto_scaling_group": _group_icon_spec(
        f"{_GROUP}/AutoScalinggroup.svg",
        padding_color_mode="icon",
    ),
    "mxgraph.aws4.group_vpc2": _group_icon_spec(f"{_GROUP}/VirtualprivatecloudVPC.svg"),
    # draw.io uses group_security_group for both public/private subnets.
    # Listed in priority order; last entry is the unconditional default.
    "mxgraph.aws4.group_security_group": [
        _group_icon_spec(f"{_GROUP}/Publicsubnet.svg", match_label="public", match_fill="#f2f6e8"),
        _group_icon_spec(f"{_GROUP}/Privatesubnet.svg"),  # default
    ],
    "mxgraph.aws4.group_on_premise": _group_icon_spec(f"{_GROUP}/Servercontents.svg"),
    "mxgraph.aws4.group_corporate_data_center": _group_icon_spec(f"{_GROUP}/Corporatedatacenter.svg"),
    "mxgraph.aws4.group_elastic_beanstalk": _group_icon_spec(
        f"{_ARCH}/AWSElasticBeanstalk.svg",
        padding_color_mode="icon",
    ),
    "mxgraph.aws4.group_ec2_instance_contents": _group_icon_spec(
        f"{_GROUP}/EC2instancecontents.svg",
        padding_color_mode="icon",
    ),
    "mxgraph.aws4.group_spot_fleet": _group_icon_spec(
        f"{_GROUP}/SpotFleet.svg",
        padding_color_mode="icon",
    ),
    "mxgraph.aws4.group_aws_step_functions_workflow": _group_icon_spec(
        f"{_ARCH}/AWSStepFunctions.svg",
        padding_color_mode="icon",
    ),
    "mxgraph.aws4.group_account": _group_icon_spec(
        f"{_GROUP}/AWSAccount.svg",
        padding_color_mode="icon",
    ),
    "mxgraph.aws4.group_iot_greengrass_deployment": _group_icon_spec(f"{_GROUP}/AWSIoTGreengrassDeployment.svg"),
    "mxgraph.aws4.group_iot_greengrass": _group_icon_spec(f"{_ARCH}/AWSIoTGreengrass.svg"),
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
        group_key = _normalize_drawio_aws_value(_get_style_value(style_str, "grIcon"))
        if group_key:
            entry = _AWS4_GROUP_ICONS.get(group_key)
            # Compatibility: accept short form "group_*" and expand to full draw.io key.
            if entry is None and "." not in group_key:
                entry = _AWS4_GROUP_ICONS.get(f"mxgraph.aws4.{group_key}")
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


def _drawio_lookup_keys(shape_type: str, res_icon: Optional[str]) -> List[tuple[str, Optional[str]]]:
    """
    Build draw.io-native lookup keys ordered by priority.

    Primary:
      - direct shape key: (shape, None)
      - resourceIcon key: (shape=...resourceIcon, resIcon=...)
    Compatibility fallback:
      - resIcon as direct shape
      - amazon_* alias shapes
    """
    shape_type_norm = _normalize_drawio_aws_value(shape_type)
    res_icon_norm = _normalize_drawio_aws_value(res_icon)
    if not shape_type_norm or not is_aws_shape_type(shape_type_norm):
        return []

    keys: List[tuple[str, Optional[str]]] = []
    is_resource_icon_shape = "resourceicon" in shape_type_norm

    if is_resource_icon_shape and res_icon_norm:
        keys.append((shape_type_norm, res_icon_norm))
    else:
        keys.append((shape_type_norm, None))

    # Fallback: if resIcon exists, try the resIcon shape directly.
    if res_icon_norm:
        keys.append((res_icon_norm, None))
        icon_suffix = res_icon_norm.split(".")[-1] if "." in res_icon_norm else res_icon_norm
        if icon_suffix:
            keys.append((f"mxgraph.aws4.amazon_{icon_suffix}", None))

    # Fallback: try amazon_<shape_suffix> for legacy compatibility.
    shape_suffix = shape_type_norm.split(".")[-1] if "." in shape_type_norm else shape_type_norm
    if shape_suffix:
        keys.append((f"mxgraph.aws4.amazon_{shape_suffix}", None))

    # Keep order while removing duplicates.
    deduped: List[tuple[str, Optional[str]]] = []
    seen = set()
    for k in keys:
        if k in seen:
            continue
        seen.add(k)
        deduped.append(k)
    return deduped


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

    is_resource_icon_shape = "resourceicon" in shape_type_lower
    res_icon = _get_style_value(style_str, "resIcon")
    drawio_keys = _drawio_lookup_keys(shape_type_lower, res_icon)
    if not drawio_keys:
        return None

    # 0) Draw.io-native dictionary lookup (shape or shape+resIcon).
    for k in drawio_keys:
        spec = _AWS4_ICON_SPEC_BY_DRAWIO_KEY.get(k)
        if not spec:
            continue
        if spec[0] == "aws4xml":
            _, shape_name, bg_hex, fg_hex, canvas_w, canvas_h = spec
            effective_bg_hex = bg_hex
            effective_bg_gradient_hex = None
            effective_gradient_direction = None
            if is_resource_icon_shape:
                style_fill = _get_style_value(style_str, "fillColor")
                if style_fill and style_fill.lower() != "none":
                    effective_bg_hex = style_fill
                style_gradient = _get_style_value(style_str, "gradientColor")
                if style_gradient and style_gradient.lower() != "none":
                    effective_bg_gradient_hex = style_gradient
                    effective_gradient_direction = _get_style_value(style_str, "gradientDirection")
            data_uri = _build_shape_data_uri_from_aws4(
                shape_name=shape_name,
                background_hex=effective_bg_hex,
                background_gradient_hex=effective_bg_gradient_hex,
                gradient_direction=effective_gradient_direction,
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
    background_gradient_hex: Optional[str] = None,
    gradient_direction: Optional[str] = None,
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

    bg_fill = background_hex
    defs = ""
    if (
        background_hex.lower() != "none"
        and background_gradient_hex
        and background_gradient_hex.lower() != "none"
    ):
        x1, y1, x2, y2 = _svg_gradient_vector_for_drawio_direction(gradient_direction)
        defs = (
            "<defs>"
            f'<linearGradient id="bgGrad" x1="{x1:.3f}" y1="{y1:.3f}" x2="{x2:.3f}" y2="{y2:.3f}">'
            f'<stop offset="0%" stop-color="{background_hex}"/>'
            f'<stop offset="100%" stop-color="{background_gradient_hex}"/>'
            "</linearGradient>"
            "</defs>"
        )
        bg_fill = "url(#bgGrad)"

    svg = (
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {effective_canvas_w:.3f} {effective_canvas_h:.3f}">'
        f"{defs}"
        f'<rect x="0" y="0" width="{effective_canvas_w:.3f}" height="{effective_canvas_h:.3f}" fill="{bg_fill}"/>'
        f'<path d="{path_d}" transform="translate({offset_x:.3f} {offset_y:.3f})" fill="{foreground_hex}" fill-rule="evenodd"/>'
        "</svg>"
    )
    encoded = base64.b64encode(svg.encode("utf-8")).decode("ascii")
    return f"data:image/svg+xml;base64,{encoded}"


def _svg_gradient_vector_for_drawio_direction(direction: Optional[str]) -> tuple[float, float, float, float]:
    """
    Convert draw.io gradientDirection to SVG linearGradient vector.
    draw.io uses fillColor + gradientColor, where gradientDirection points to the gradientColor side.
    """
    direction_norm = (direction or "").strip().lower()
    if direction_norm == "north":
        return (0.0, 1.0, 0.0, 0.0)
    if direction_norm == "south":
        return (0.0, 0.0, 0.0, 1.0)
    if direction_norm == "east":
        return (0.0, 0.0, 1.0, 0.0)
    if direction_norm == "west":
        return (1.0, 0.0, 0.0, 0.0)
    return (0.0, 0.0, 0.0, 1.0)


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
