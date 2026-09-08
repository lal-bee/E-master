"""P1-07 导入前质量门禁。

消费 P1-02～P1-06 的结果对象，对“是否允许进入后续导入流程”给出
确定性 PASS/FAIL 判定：

- TemplateWorkbookResult 的 TemplateIssue 是前序问题的规范化入口；
- 规则严重级别、blocking 与启用状态全部来自 JSON 配置；
- 只读取、判定与聚合，不重新读取业务 Excel、不修复、不猜测；
- 本模块只产生内存质量结果，不生成报告文件。
"""

from __future__ import annotations

import json
from collections.abc import Mapping
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any

from excel_qc.cleaning import CleaningWorkbookResult
from excel_qc.equipment_type_mapping import EquipmentTypeWorkbookResult
from excel_qc.field_mapping import WorkbookFieldMappingResult
from excel_qc.profiler import WorkbookProfile
from excel_qc.template_generator import (
    TemplateCellIssue,
    TemplateWorkbookResult,
)
from excel_qc.text import is_blank_value


class QualityGateError(Exception):
    """导入前质量门禁领域异常基类。"""


class QualityGateConfigError(QualityGateError):
    """质量门禁配置不合法或映射不完整。"""


class QualityGateDataError(QualityGateError):
    """质量门禁输入结果与结构不一致。"""


class QualitySeverity(str, Enum):
    """质量问题严重级别。"""

    INFO = "INFO"
    WARNING = "WARNING"
    ERROR = "ERROR"
    BLOCKER = "BLOCKER"


class QualityStatus(str, Enum):
    """质量门禁整体/Sheet 状态。"""

    PASS = "PASS"
    FAIL = "FAIL"


class QualityRuleKind(str, Enum):
    """质量规则类型。"""

    REQUIRED_FIELD = "required_field"
    DUPLICATE_KEY = "duplicate_key"
    TEMPLATE_ISSUE = "template_issue"


class QualityIssueSourceStage(str, Enum):
    """质量问题来源阶段。"""

    FIELD_MAPPING = "FIELD_MAPPING"
    CLEANING = "CLEANING"
    EQUIPMENT_TYPE_MAPPING = "EQUIPMENT_TYPE_MAPPING"
    TEMPLATE = "TEMPLATE"
    QUALITY_GATE = "QUALITY_GATE"


@dataclass(frozen=True)
class QualityRule:
    """一条质量门禁规则。

    applies_to 决定规则种类；parameters 存放该种类的最小参数集合：

    - required_field: {"standard_fields": [...]}
    - duplicate_key: {"key_standard_fields": [...],
                      "skip_empty_keys": true}
    - template_issue: {"issue_codes": [...] 和/或
                       "source_stages": [...]}
    """

    rule_code: str
    severity: QualitySeverity
    applies_to: QualityRuleKind
    blocking: bool
    description: str = ""
    enabled: bool = True
    parameters: dict[str, object] = field(default_factory=dict)


@dataclass(frozen=True)
class QualityGateConfig:
    """质量门禁规则配置。"""

    rules: tuple[QualityRule, ...] = ()


@dataclass(frozen=True)
class QualityIssue:
    """一条质量问题，保留完整追溯信息。"""

    rule_code: str
    severity: QualitySeverity
    blocking: bool
    source_stage: QualityIssueSourceStage
    sheet: str
    row_number: int | None
    standard_field: str
    template_column: str
    reason: str
    actual_value: Any = None
    column_number: int | None = None
    column_letter: str | None = None


@dataclass(frozen=True)
class QualityRuleCount:
    """按规则统计的问题数量。"""

    rule_code: str
    count: int = 0


@dataclass(frozen=True)
class QualitySummary:
    """质量门禁统计。"""

    issue_count: int = 0
    info_count: int = 0
    warning_count: int = 0
    error_count: int = 0
    blocker_count: int = 0
    rule_counts: tuple[QualityRuleCount, ...] = ()


@dataclass(frozen=True)
class QualitySheetResult:
    """一个模板 Sheet 的质量门禁结果。"""

    sheet_name: str
    status: QualityStatus
    issues: tuple[QualityIssue, ...]
    summary: QualitySummary


@dataclass(frozen=True)
class QualityWorkbookResult:
    """整个工作簿的质量门禁结果。"""

    source_file: Path
    status: QualityStatus
    sheet_results: tuple[QualitySheetResult, ...]
    issues: tuple[QualityIssue, ...]
    summary: QualitySummary
    can_proceed: bool


_TEMPLATE_SOURCE_STAGES = {
    QualityIssueSourceStage.FIELD_MAPPING.value,
    QualityIssueSourceStage.CLEANING.value,
    QualityIssueSourceStage.EQUIPMENT_TYPE_MAPPING.value,
    QualityIssueSourceStage.TEMPLATE.value,
}


def load_quality_gate_config(path: str | Path) -> QualityGateConfig:
    """从 JSON 文件加载质量门禁规则配置。

    JSON 结构：
    {
      "rules": [
        {
          "rule_code": "REQUIRED_FIELD_EMPTY",
          "description": "必填字段为空",
          "severity": "BLOCKER",
          "applies_to": "required_field",
          "enabled": true,
          "blocking": true,
          "parameters": {"standard_fields": ["设备编码"]}
        }
      ]
    }
    """
    config_path = Path(path)
    if not config_path.exists() or not config_path.is_file():
        raise QualityGateConfigError(
            f"质量门禁配置文件不存在: {config_path}"
        )
    try:
        payload = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise QualityGateConfigError(
            f"质量门禁配置文件无法解析: {config_path}：{exc}"
        ) from exc
    if not isinstance(payload, dict):
        raise QualityGateConfigError(
            "质量门禁配置顶层必须是 JSON 对象"
        )

    raw_rules = payload.get("rules")
    if not isinstance(raw_rules, list) or not raw_rules:
        raise QualityGateConfigError(
            "质量门禁配置缺少非空 rules 列表"
        )

    rules: list[QualityRule] = []
    for index, raw_rule in enumerate(raw_rules, start=1):
        if not isinstance(raw_rule, dict):
            raise QualityGateConfigError(f"rules 第 {index} 项必须是对象")
        rule_code = raw_rule.get("rule_code")
        if not isinstance(rule_code, str) or not rule_code.strip():
            raise QualityGateConfigError(
                f"rules 第 {index} 项 rule_code 必须是非空字符串"
            )

        severity_text = raw_rule.get("severity")
        try:
            severity = QualitySeverity(severity_text)
        except (TypeError, ValueError) as exc:
            raise QualityGateConfigError(
                f"rules 第 {index} 项 severity 必须为 "
                "INFO/WARNING/ERROR/BLOCKER"
            ) from exc

        applies_to_text = raw_rule.get("applies_to")
        try:
            applies_to = QualityRuleKind(applies_to_text)
        except (TypeError, ValueError) as exc:
            raise QualityGateConfigError(
                f"rules 第 {index} 项 applies_to 必须为 "
                "required_field/duplicate_key/template_issue"
            ) from exc

        enabled = raw_rule.get("enabled", True)
        if not isinstance(enabled, bool):
            raise QualityGateConfigError(
                f"rules 第 {index} 项 enabled 必须是布尔值"
            )

        if "blocking" not in raw_rule:
            raise QualityGateConfigError(
                f"rules 第 {index} 项缺少 blocking"
            )
        blocking = raw_rule.get("blocking")
        if not isinstance(blocking, bool):
            raise QualityGateConfigError(
                f"rules 第 {index} 项 blocking 必须是布尔值"
            )

        description = raw_rule.get("description", "")
        if not isinstance(description, str):
            raise QualityGateConfigError(
                f"rules 第 {index} 项 description 必须是字符串"
            )

        parameters = raw_rule.get("parameters")
        if not isinstance(parameters, dict):
            raise QualityGateConfigError(
                f"rules 第 {index} 项 parameters 必须是对象"
            )

        rules.append(
            QualityRule(
                rule_code=rule_code.strip(),
                severity=severity,
                applies_to=applies_to,
                blocking=blocking,
                description=description,
                enabled=enabled,
                parameters=parameters,
            )
        )

    config = QualityGateConfig(tuple(rules))
    _validate_config(config)
    return config


def run_quality_gate(
    profile: WorkbookProfile,
    mapping: WorkbookFieldMappingResult,
    cleaning_result: CleaningWorkbookResult,
    equipment_result: EquipmentTypeWorkbookResult,
    template_result: TemplateWorkbookResult,
    config: QualityGateConfig,
) -> QualityWorkbookResult:
    """对 P1-06 模板结果执行导入前质量门禁。

    - TemplateIssue 按配置映射为 QualityIssue；
    - 执行配置化必填与重复键检查；
    - 不修改任何输入对象、不创建文件。
    """
    _validate_config(config)
    _validate_input_consistency(
        profile=profile,
        mapping=mapping,
        cleaning_result=cleaning_result,
        equipment_result=equipment_result,
        template_result=template_result,
    )

    template_rules = _enabled_rules(config, QualityRuleKind.TEMPLATE_ISSUE)
    required_rules = _enabled_rules(config, QualityRuleKind.REQUIRED_FIELD)
    duplicate_rules = _enabled_rules(config, QualityRuleKind.DUPLICATE_KEY)

    sheet_results: list[QualitySheetResult] = []
    for template_sheet in template_result.sheet_results:
        issues: list[QualityIssue] = []
        issues.extend(
            _convert_template_issues(template_sheet, template_rules)
        )
        for rule in required_rules:
            issues.extend(
                _check_required_fields(template_sheet, rule)
            )
        for rule in duplicate_rules:
            issues.extend(
                _check_duplicate_keys(template_sheet, rule)
            )
        sheet_results.append(
            _sheet_result(template_sheet.sheet_name, tuple(issues))
        )

    issues = tuple(
        issue for result in sheet_results for issue in result.issues
    )
    summary = _summary(issues)
    status = _status_for(issues)
    can_proceed = status is QualityStatus.PASS
    return QualityWorkbookResult(
        source_file=template_result.source_file,
        status=status,
        sheet_results=tuple(sheet_results),
        issues=issues,
        summary=summary,
        can_proceed=can_proceed,
    )


def _validate_config(config: QualityGateConfig) -> None:
    if not isinstance(config, QualityGateConfig):
        raise QualityGateConfigError("config 必须是 QualityGateConfig")
    if not isinstance(config.rules, tuple) or not config.rules:
        raise QualityGateConfigError("质量门禁配置必须包含至少一条规则")
    seen_codes: dict[str, int] = {}
    for index, rule in enumerate(config.rules, start=1):
        if not isinstance(rule, QualityRule):
            raise QualityGateConfigError(
                f"rules 第 {index} 项必须是 QualityRule"
            )
        if not isinstance(rule.rule_code, str) or not rule.rule_code.strip():
            raise QualityGateConfigError(
                f"rules 第 {index} 项 rule_code 必须是非空字符串"
            )
        rule_code = rule.rule_code.strip()
        if rule_code in seen_codes:
            raise QualityGateConfigError(
                f"rule_code 重复: “{rule_code}”（第 {index} 项与第 "
                f"{seen_codes[rule_code]} 项）"
            )
        seen_codes[rule_code] = index
        if not isinstance(rule.severity, QualitySeverity):
            raise QualityGateConfigError(
                f"rules 第 {index} 项 severity 必须是 QualitySeverity"
            )
        if not isinstance(rule.applies_to, QualityRuleKind):
            raise QualityGateConfigError(
                f"rules 第 {index} 项 applies_to 必须是 QualityRuleKind"
            )
        if not isinstance(rule.enabled, bool):
            raise QualityGateConfigError(
                f"rules 第 {index} 项 enabled 必须是布尔值"
            )
        if not isinstance(rule.blocking, bool):
            raise QualityGateConfigError(
                f"rules 第 {index} 项 blocking 必须是布尔值"
            )
        _validate_severity_blocking(rule)
        _validate_rule_parameters(rule)


def _validate_severity_blocking(rule: QualityRule) -> None:
    if rule.severity is QualitySeverity.BLOCKER and not rule.blocking:
        raise QualityGateConfigError(
            f"规则“{rule.rule_code}”severity=BLOCKER 时 blocking 必须为 true"
        )
    if rule.severity is QualitySeverity.ERROR and not rule.blocking:
        raise QualityGateConfigError(
            f"规则“{rule.rule_code}”severity=ERROR 时 blocking 必须为 true"
        )
    if (
        rule.severity in {QualitySeverity.INFO, QualitySeverity.WARNING}
        and rule.blocking
    ):
        raise QualityGateConfigError(
            f"规则“{rule.rule_code}”severity=INFO/WARNING 时 "
            "blocking 必须为 false"
        )


def _validate_rule_parameters(rule: QualityRule) -> None:
    parameters = rule.parameters
    if not isinstance(parameters, Mapping):
        raise QualityGateConfigError(
            f"规则“{rule.rule_code}”parameters 必须是对象"
        )
    if rule.applies_to is QualityRuleKind.REQUIRED_FIELD:
        _validate_field_list_parameter(
            rule,
            parameters,
            key="standard_fields",
            label="standard_fields",
        )
        _reject_unknown_parameters(rule, {"standard_fields"})
        return
    if rule.applies_to is QualityRuleKind.DUPLICATE_KEY:
        _validate_field_list_parameter(
            rule,
            parameters,
            key="key_standard_fields",
            label="key_standard_fields",
        )
        if "skip_empty_keys" in parameters:
            if not isinstance(parameters["skip_empty_keys"], bool):
                raise QualityGateConfigError(
                    f"规则“{rule.rule_code}”的 skip_empty_keys "
                    "必须是布尔值"
                )
        _reject_unknown_parameters(
            rule, {"key_standard_fields", "skip_empty_keys"}
        )
        return
    if rule.applies_to is QualityRuleKind.TEMPLATE_ISSUE:
        _validate_template_issue_parameters(rule, parameters)
        _reject_unknown_parameters(
            rule, {"issue_codes", "source_stages"}
        )
        return
    raise QualityGateConfigError(
        f"规则“{rule.rule_code}”applies_to 不受支持"
    )


def _validate_field_list_parameter(
    rule: QualityRule,
    parameters: Mapping[str, object],
    *,
    key: str,
    label: str,
) -> None:
    values = parameters.get(key)
    if not isinstance(values, (list, tuple)) or not values:
        raise QualityGateConfigError(
            f"规则“{rule.rule_code}”缺少非空 {label}"
        )
    if not all(
        isinstance(item, str) and item.strip() for item in values
    ):
        raise QualityGateConfigError(
            f"规则“{rule.rule_code}”的 {label} 必须是非空字符串列表"
        )
    normalized = [item.strip() for item in values]
    if len(set(normalized)) != len(normalized):
        raise QualityGateConfigError(
            f"规则“{rule.rule_code}”的 {label} 存在重复字段"
        )


def _validate_template_issue_parameters(
    rule: QualityRule,
    parameters: Mapping[str, object],
) -> None:
    has_filter = False
    issue_codes = parameters.get("issue_codes", ())
    if issue_codes:
        if not isinstance(issue_codes, (list, tuple)) or not all(
            isinstance(item, str) and item.strip() for item in issue_codes
        ):
            raise QualityGateConfigError(
                f"规则“{rule.rule_code}”的 issue_codes "
                "必须是非空字符串列表"
            )
        normalized = [item.strip() for item in issue_codes]
        if len(set(normalized)) != len(normalized):
            raise QualityGateConfigError(
                f"规则“{rule.rule_code}”的 issue_codes 存在重复"
            )
        has_filter = True
    source_stages = parameters.get("source_stages", ())
    if source_stages:
        if not isinstance(source_stages, (list, tuple)) or not all(
            isinstance(item, str) and item.strip() for item in source_stages
        ):
            raise QualityGateConfigError(
                f"规则“{rule.rule_code}”的 source_stages "
                "必须是非空字符串列表"
            )
        if not set(source_stages).issubset(_TEMPLATE_SOURCE_STAGES):
            raise QualityGateConfigError(
                f"规则“{rule.rule_code}”的 source_stages 包含非法阶段值"
            )
        has_filter = True
    if not has_filter:
        raise QualityGateConfigError(
            f"规则“{rule.rule_code}”applies_to=template_issue 时 "
            "parameters 必须提供 issue_codes 和/或 source_stages"
        )


def _reject_unknown_parameters(
    rule: QualityRule,
    allowed: set[str],
) -> None:
    unknown = set(rule.parameters) - allowed
    if unknown:
        raise QualityGateConfigError(
            f"规则“{rule.rule_code}”parameters 包含不适用参数: "
            + "、".join(sorted(unknown))
        )


def _validate_input_consistency(
    *,
    profile: WorkbookProfile,
    mapping: WorkbookFieldMappingResult,
    cleaning_result: CleaningWorkbookResult,
    equipment_result: EquipmentTypeWorkbookResult,
    template_result: TemplateWorkbookResult,
) -> None:
    if not isinstance(profile, WorkbookProfile):
        raise QualityGateDataError("profile 必须是 WorkbookProfile")
    if not isinstance(mapping, WorkbookFieldMappingResult):
        raise QualityGateDataError(
            "mapping 必须是 WorkbookFieldMappingResult"
        )
    if not isinstance(cleaning_result, CleaningWorkbookResult):
        raise QualityGateDataError(
            "cleaning_result 必须是 CleaningWorkbookResult"
        )
    if not isinstance(equipment_result, EquipmentTypeWorkbookResult):
        raise QualityGateDataError(
            "equipment_result 必须是 EquipmentTypeWorkbookResult"
        )
    if not isinstance(template_result, TemplateWorkbookResult):
        raise QualityGateDataError(
            "template_result 必须是 TemplateWorkbookResult"
        )

    profile_names = tuple(sheet.sheet_name for sheet in profile.sheets)
    mapping_names = tuple(result.sheet_name for result in mapping.sheet_results)
    cleaning_names = tuple(
        result.sheet_name for result in cleaning_result.sheet_results
    )
    equipment_names = tuple(
        result.sheet_name for result in equipment_result.sheet_results
    )
    if not (
        profile_names == mapping_names == cleaning_names == equipment_names
    ):
        raise QualityGateDataError(
            "profile/mapping/cleaning/equipment 的 Sheet 名称或顺序不一致"
        )

    source_path = Path(profile.path).resolve()
    for label, other in (
        ("mapping", mapping.file_path),
        ("cleaning_result", cleaning_result.source_file),
        ("equipment_result", equipment_result.source_file),
        ("template_result", template_result.source_file),
    ):
        if Path(other).resolve() != source_path:
            raise QualityGateDataError(
                f"{label} 与 profile 不属于同一源文件"
            )

    cleaning_by_sheet = {
        result.sheet_name: result for result in cleaning_result.sheet_results
    }
    equipment_by_sheet = {
        result.sheet_name: result for result in equipment_result.sheet_results
    }
    template_by_sheet = {
        result.sheet_name: result for result in template_result.sheet_results
    }

    if not _is_ordered_subsequence(
        tuple(result.sheet_name for result in template_result.sheet_results),
        profile_names,
    ):
        raise QualityGateDataError(
            "template_result 的 Sheet 顺序与 profile 不一致"
        )

    for sheet_profile in profile.sheets:
        expected_rows = _expected_row_numbers(sheet_profile)
        cleaning_sheet = cleaning_by_sheet[sheet_profile.sheet_name]
        equipment_sheet = equipment_by_sheet[sheet_profile.sheet_name]
        if tuple(
            row.source_row_number for row in cleaning_sheet.rows
        ) != expected_rows:
            raise QualityGateDataError(
                f"Sheet“{sheet_profile.sheet_name}”清洗结果行覆盖范围不一致"
            )
        if tuple(
            audit.row_number for audit in equipment_sheet.audits
        ) != expected_rows:
            raise QualityGateDataError(
                f"Sheet“{sheet_profile.sheet_name}”编码映射结果行覆盖范围不一致"
            )

        template_sheet = template_by_sheet.get(sheet_profile.sheet_name)
        if template_sheet is None:
            continue
        if tuple(
            row.source_row_number for row in template_sheet.rows
        ) != expected_rows:
            raise QualityGateDataError(
                f"Sheet“{sheet_profile.sheet_name}”模板结果行覆盖范围不一致"
            )


def _is_ordered_subsequence(
    values: tuple[str, ...],
    ordered: tuple[str, ...],
) -> bool:
    if len(set(values)) != len(values):
        return False
    position = 0
    for value in values:
        while position < len(ordered) and ordered[position] != value:
            position += 1
        if position >= len(ordered):
            return False
        position += 1
    return True


def _expected_row_numbers(sheet_profile) -> tuple[int, ...]:
    if (
        sheet_profile.data_first_row is None
        or sheet_profile.data_last_row is None
    ):
        return ()
    return tuple(
        range(sheet_profile.data_first_row, sheet_profile.data_last_row + 1)
    )


def _enabled_rules(
    config: QualityGateConfig,
    kind: QualityRuleKind,
) -> tuple[QualityRule, ...]:
    return tuple(
        rule
        for rule in config.rules
        if rule.applies_to is kind and rule.enabled
    )


def _convert_template_issues(
    template_sheet,
    rules: tuple[QualityRule, ...],
) -> list[QualityIssue]:
    issues: list[QualityIssue] = []
    row_by_number = {
        row.source_row_number: row for row in template_sheet.rows
    }
    for template_issue in template_sheet.issues:
        rule = _match_template_issue_rule(template_issue, rules)
        if rule is None:
            raise QualityGateConfigError(
                f"TemplateIssue 未配置映射规则: "
                f"issue_code={template_issue.issue_code}, "
                f"source_stage={template_issue.source_stage.value}, "
                f"sheet={template_issue.sheet}，需补充质量门禁配置"
            )
        trace = None
        actual_value = None
        if template_issue.source_row_number is not None:
            row = row_by_number.get(template_issue.source_row_number)
            if row is not None:
                trace, actual_value = _find_trace_cell(
                    row,
                    standard_field=template_issue.standard_field,
                    template_column=template_issue.template_column,
                )
        issues.append(
            QualityIssue(
                rule_code=rule.rule_code,
                severity=rule.severity,
                blocking=rule.blocking,
                source_stage=QualityIssueSourceStage(
                    template_issue.source_stage.value
                ),
                sheet=template_issue.sheet,
                row_number=template_issue.source_row_number,
                standard_field=template_issue.standard_field,
                template_column=template_issue.template_column,
                reason=template_issue.reason,
                actual_value=actual_value,
                column_number=(
                    trace.source_column_number if trace is not None else None
                ),
                column_letter=(
                    trace.source_column_letter if trace is not None else None
                ),
            )
        )
    return issues


def _match_template_issue_rule(
    issue: TemplateCellIssue,
    rules: tuple[QualityRule, ...],
) -> QualityRule | None:
    for rule in rules:
        parameters = rule.parameters
        issue_codes = parameters.get("issue_codes") or ()
        source_stages = parameters.get("source_stages") or ()
        code_matches = not issue_codes or issue.issue_code in issue_codes
        stage_matches = (
            not source_stages or issue.source_stage.value in source_stages
        )
        if code_matches and stage_matches:
            return rule
    return None


def _find_trace_cell(row, *, standard_field: str, template_column: str):
    for index, trace in enumerate(row.trace):
        if (
            trace.standard_field == standard_field
            and trace.template_column == template_column
        ):
            return trace, row.values[index]
    return None, None


def _check_required_fields(
    template_sheet,
    rule: QualityRule,
) -> list[QualityIssue]:
    standard_fields = rule.parameters["standard_fields"]
    column_by_field: dict[str, int] = {}
    for index, column in enumerate(template_sheet.columns):
        if (
            column.standard_field in standard_fields
            and not column.is_code_column
        ):
            column_by_field[column.standard_field] = index
    for field in standard_fields:
        if field not in column_by_field:
            raise QualityGateConfigError(
                f"规则“{rule.rule_code}”的必填字段“{field}”"
                f"不在模板 Sheet“{template_sheet.sheet_name}”的普通输出列中"
            )

    issues: list[QualityIssue] = []
    for row in template_sheet.rows:
        for field in standard_fields:
            column_index = column_by_field[field]
            actual_value = row.values[column_index]
            if not is_blank_value(actual_value):
                continue
            trace = row.trace[column_index]
            issues.append(
                QualityIssue(
                    rule_code=rule.rule_code,
                    severity=rule.severity,
                    blocking=rule.blocking,
                    source_stage=QualityIssueSourceStage.QUALITY_GATE,
                    sheet=template_sheet.sheet_name,
                    row_number=row.source_row_number,
                    standard_field=field,
                    template_column=trace.template_column,
                    reason=f"必填字段“{field}”为空",
                    actual_value=actual_value,
                    column_number=trace.source_column_number,
                    column_letter=trace.source_column_letter,
                )
            )
    return issues


def _check_duplicate_keys(
    template_sheet,
    rule: QualityRule,
) -> list[QualityIssue]:
    key_fields = rule.parameters["key_standard_fields"]
    column_by_field: dict[str, int] = {}
    for index, column in enumerate(template_sheet.columns):
        if column.standard_field in key_fields:
            column_by_field[column.standard_field] = index
    for field in key_fields:
        if field not in column_by_field:
            raise QualityGateConfigError(
                f"规则“{rule.rule_code}”的重复键字段“{field}”"
                f"不在模板 Sheet“{template_sheet.sheet_name}”的输出列中"
            )

    skip_empty_keys = bool(rule.parameters.get("skip_empty_keys", True))
    seen: dict[tuple[Any, ...], int] = {}
    issues: list[QualityIssue] = []
    for row in template_sheet.rows:
        values = tuple(
            row.values[column_by_field[field]] for field in key_fields
        )
        if skip_empty_keys and any(is_blank_value(value) for value in values):
            continue
        first_row_number = seen.get(values)
        if first_row_number is None:
            seen[values] = row.source_row_number
            continue
        key_label = _key_label(values)
        trace = row.trace[column_by_field[key_fields[0]]]
        issues.append(
            QualityIssue(
                rule_code=rule.rule_code,
                severity=rule.severity,
                blocking=rule.blocking,
                source_stage=QualityIssueSourceStage.QUALITY_GATE,
                sheet=template_sheet.sheet_name,
                row_number=row.source_row_number,
                standard_field=key_fields[0],
                template_column=trace.template_column,
                reason=(
                    f"重复键 {key_label} 与第 {first_row_number} 行重复"
                ),
                actual_value=values,
                column_number=trace.source_column_number,
                column_letter=trace.source_column_letter,
            )
        )
    return issues


def _key_label(values: tuple[Any, ...]) -> str:
    parts = ", ".join(repr(value) for value in values)
    return f"({parts})"


def _sheet_result(
    sheet_name: str,
    issues: tuple[QualityIssue, ...],
) -> QualitySheetResult:
    return QualitySheetResult(
        sheet_name=sheet_name,
        status=_status_for(issues),
        issues=issues,
        summary=_summary(issues),
    )


def _status_for(issues: tuple[QualityIssue, ...]) -> QualityStatus:
    if any(
        issue.severity in {QualitySeverity.ERROR, QualitySeverity.BLOCKER}
        for issue in issues
    ):
        return QualityStatus.FAIL
    return QualityStatus.PASS


def _summary(issues: tuple[QualityIssue, ...]) -> QualitySummary:
    counts = {severity: 0 for severity in QualitySeverity}
    rule_counts: dict[str, int] = {}
    for issue in issues:
        counts[issue.severity] += 1
        rule_counts[issue.rule_code] = (
            rule_counts.get(issue.rule_code, 0) + 1
        )
    return QualitySummary(
        issue_count=len(issues),
        info_count=counts[QualitySeverity.INFO],
        warning_count=counts[QualitySeverity.WARNING],
        error_count=counts[QualitySeverity.ERROR],
        blocker_count=counts[QualitySeverity.BLOCKER],
        rule_counts=tuple(
            QualityRuleCount(rule_code=code, count=count)
            for code, count in sorted(rule_counts.items())
        ),
    )
