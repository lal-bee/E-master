"""P1-07 导入前质量门禁测试。"""

from __future__ import annotations

import json
from dataclasses import replace
from pathlib import Path

import pytest

from excel_qc.cleaning import (
    CleaningKind,
    CleaningWorkbookConfig,
    CleaningWorkbookResult,
    FieldCleaningRule,
    clean_workbook,
)
from excel_qc.equipment_type_mapping import (
    EquipmentTypeCodeConfig,
    EquipmentTypeRule,
    map_equipment_type_codes,
)
from excel_qc.field_mapping import (
    FieldMappingConfig,
    StandardFieldDefinition,
    WorkbookFieldMappingResult,
    map_workbook_fields,
)
from excel_qc.profiler import profile_workbook
from excel_qc.quality_gate import (
    QualityGateConfig,
    QualityGateConfigError,
    QualityGateDataError,
    QualityIssue,
    QualityIssueSourceStage,
    QualityRule,
    QualityRuleKind,
    QualitySeverity,
    QualitySheetResult,
    QualityStatus,
    QualitySummary,
    QualityWorkbookResult,
    load_quality_gate_config,
    run_quality_gate,
)
from excel_qc.template_generator import (
    TemplateColumnConfig,
    TemplateConfig,
    TemplateWorkbookResult,
    build_standard_template,
)
from tests.qc.helpers import build_workbook


_BASIC_FIELDS = (
    StandardFieldDefinition(name="设备名称"),
    StandardFieldDefinition(name="设备类型名称", aliases=("设备类型",)),
    StandardFieldDefinition(name="投用日期"),
    StandardFieldDefinition(name="备注"),
)

_DEFAULT_EQUIPMENT_RULES = (
    EquipmentTypeRule(
        aliases=("除尘器",),
        standard_type_name="除尘器",
        system_code="D-CCQ",
    ),
    EquipmentTypeRule(
        aliases=("空压机",),
        standard_type_name="空压机",
        system_code="D-KYJ",
    ),
)

_ALL_TEMPLATE_ISSUE_CODES = (
    "FIELD_MAPPING_MISSING",
    "FIELD_MAPPING_CONFLICT",
    "CLEANING_INVALID",
    "EQUIPMENT_UNMATCHED",
    "EQUIPMENT_CONFLICT",
    "EQUIPMENT_INVALID",
    "EQUIPMENT_SKIPPED",
    "EQUIPMENT_AUDIT_MISSING",
)


def _equipment_config(
    *,
    target: str = "设备类型名称",
    rules: tuple[EquipmentTypeRule, ...] = _DEFAULT_EQUIPMENT_RULES,
    code_pattern: str | None = None,
) -> EquipmentTypeCodeConfig:
    return EquipmentTypeCodeConfig(
        target_standard_field=target,
        rules=rules,
        code_pattern=code_pattern,
    )


def _chain(
    tmp_path: Path,
    filename: str,
    sheets: dict[str, list[list[object]]],
    *,
    standard_fields: tuple[StandardFieldDefinition, ...] = _BASIC_FIELDS,
    cleaning_rules: tuple[FieldCleaningRule, ...] = (),
    equipment_config: EquipmentTypeCodeConfig | None = None,
):
    path = build_workbook(tmp_path / filename, sheets)
    profile = profile_workbook(path)
    mapping = map_workbook_fields(
        profile, FieldMappingConfig(standard_fields)
    )
    cleaning = clean_workbook(
        profile,
        mapping,
        CleaningWorkbookConfig(cleaning_rules),
    )
    config = (
        equipment_config
        if equipment_config is not None
        else _equipment_config()
    )
    equipment = map_equipment_type_codes(cleaning, config)
    return path, profile, mapping, cleaning, equipment


def _run_template(
    tmp_path: Path,
    filename: str,
    sheets: dict[str, list[list[object]]],
    columns: tuple[TemplateColumnConfig, ...],
    *,
    template_sheets: tuple[str, ...] | None = None,
    standard_fields: tuple[StandardFieldDefinition, ...] = _BASIC_FIELDS,
    cleaning_rules: tuple[FieldCleaningRule, ...] = (),
    equipment_config: EquipmentTypeCodeConfig | None = None,
):
    path, profile, mapping, cleaning, equipment = _chain(
        tmp_path,
        filename,
        sheets,
        standard_fields=standard_fields,
        cleaning_rules=cleaning_rules,
        equipment_config=equipment_config,
    )
    template = build_standard_template(
        profile,
        mapping,
        cleaning,
        equipment,
        TemplateConfig(
            columns=columns,
            sheets=template_sheets,
        ),
    )
    return path, profile, mapping, cleaning, equipment, template


def _name_column() -> TemplateColumnConfig:
    return TemplateColumnConfig(standard_field="设备名称", header="设备名称")


def _date_column() -> TemplateColumnConfig:
    return TemplateColumnConfig(standard_field="投用日期", header="投用日期")


def _code_column() -> TemplateColumnConfig:
    return TemplateColumnConfig(
        standard_field="设备类型编码",
        header="设备类型编码",
        code_source_standard_field="设备类型名称",
    )


def _rule(
    rule_code: str,
    kind: QualityRuleKind,
    severity: QualitySeverity,
    parameters: dict[str, object],
    *,
    enabled: bool = True,
    blocking: bool | None = None,
) -> QualityRule:
    if blocking is None:
        blocking = severity in {
            QualitySeverity.ERROR,
            QualitySeverity.BLOCKER,
        }
    return QualityRule(
        rule_code=rule_code,
        severity=severity,
        applies_to=kind,
        blocking=blocking,
        enabled=enabled,
        parameters=parameters,
    )


def _template_rules(
    severity: QualitySeverity = QualitySeverity.ERROR,
) -> tuple[QualityRule, ...]:
    return tuple(
        _rule(
            code,
            QualityRuleKind.TEMPLATE_ISSUE,
            severity,
            {"issue_codes": [code]},
        )
        for code in _ALL_TEMPLATE_ISSUE_CODES
    )


def _required_rule(
    fields: tuple[str, ...],
    severity: QualitySeverity = QualitySeverity.BLOCKER,
    *,
    rule_code: str = "REQUIRED_FIELD_EMPTY",
    enabled: bool = True,
) -> QualityRule:
    return _rule(
        rule_code,
        QualityRuleKind.REQUIRED_FIELD,
        severity,
        {"standard_fields": list(fields)},
        enabled=enabled,
    )


def _duplicate_rule(
    fields: tuple[str, ...],
    severity: QualitySeverity = QualitySeverity.BLOCKER,
    *,
    skip_empty_keys: bool | None = True,
    rule_code: str = "DUPLICATE_KEY",
) -> QualityRule:
    parameters: dict[str, object] = {
        "key_standard_fields": list(fields)
    }
    if skip_empty_keys is not None:
        parameters["skip_empty_keys"] = skip_empty_keys
    return _rule(
        rule_code,
        QualityRuleKind.DUPLICATE_KEY,
        severity,
        parameters,
    )


def _config(*rules: QualityRule) -> QualityGateConfig:
    return QualityGateConfig(tuple(rules))


def _default_config(*extra: QualityRule) -> QualityGateConfig:
    return _config(*_template_rules(), *extra)


def _json_config(tmp_path: Path, payload: object) -> Path:
    config_path = tmp_path / "quality_gate.json"
    config_path.write_text(
        json.dumps(payload, ensure_ascii=False),
        encoding="utf-8",
    )
    return config_path


def _rule_payload(
    *,
    rule_code: str = "REQUIRED_FIELD_EMPTY",
    severity: str = "BLOCKER",
    applies_to: str = "required_field",
    enabled: object = True,
    blocking: object = True,
    parameters: object | None = None,
) -> dict[str, object]:
    return {
        "rule_code": rule_code,
        "severity": severity,
        "applies_to": applies_to,
        "enabled": enabled,
        "blocking": blocking,
        "parameters": (
            parameters
            if parameters is not None
            else {"standard_fields": ["设备名称"]}
        ),
    }


def test_load_config_valid_json(tmp_path: Path) -> None:
    config_path = _json_config(
        tmp_path,
        {
            "rules": [
                _rule_payload(),
                _rule_payload(
                    rule_code="EQUIPMENT_UNMATCHED",
                    severity="ERROR",
                    applies_to="template_issue",
                    parameters={"issue_codes": ["EQUIPMENT_UNMATCHED"]},
                ),
            ]
        },
    )

    config = load_quality_gate_config(config_path)

    assert len(config.rules) == 2
    assert config.rules[0].rule_code == "REQUIRED_FIELD_EMPTY"
    assert config.rules[0].severity is QualitySeverity.BLOCKER
    assert config.rules[1].applies_to is QualityRuleKind.TEMPLATE_ISSUE


def test_load_config_missing_file_raises(tmp_path: Path) -> None:
    with pytest.raises(QualityGateConfigError) as exc_info:
        load_quality_gate_config(tmp_path / "missing.json")

    assert "不存在" in str(exc_info.value)


@pytest.mark.parametrize(
    "payload",
    ["{bad json", "[]", '{"rules": 1}'],
)
def test_load_config_invalid_payload_raises(
    tmp_path: Path,
    payload: str,
) -> None:
    config_path = tmp_path / "bad.json"
    config_path.write_text(payload, encoding="utf-8")

    with pytest.raises(QualityGateConfigError):
        load_quality_gate_config(config_path)


def test_load_config_missing_rules_raises(tmp_path: Path) -> None:
    config_path = _json_config(tmp_path, {})

    with pytest.raises(QualityGateConfigError) as exc_info:
        load_quality_gate_config(config_path)

    assert "rules" in str(exc_info.value)


def test_load_config_empty_rules_raises(tmp_path: Path) -> None:
    config_path = _json_config(tmp_path, {"rules": []})

    with pytest.raises(QualityGateConfigError) as exc_info:
        load_quality_gate_config(config_path)

    assert "rules" in str(exc_info.value)


@pytest.mark.parametrize(
    "rule_code",
    ["", 5],
)
def test_load_config_invalid_rule_code_raises(
    tmp_path: Path,
    rule_code: object,
) -> None:
    payload = _rule_payload(rule_code=rule_code)
    config_path = _json_config(tmp_path, {"rules": [payload]})

    with pytest.raises(QualityGateConfigError) as exc_info:
        load_quality_gate_config(config_path)

    assert "rule_code" in str(exc_info.value)


def test_load_config_duplicate_rule_code_raises(tmp_path: Path) -> None:
    payload = _rule_payload()
    config_path = _json_config(tmp_path, {"rules": [payload, payload]})

    with pytest.raises(QualityGateConfigError) as exc_info:
        load_quality_gate_config(config_path)

    assert "重复" in str(exc_info.value)


@pytest.mark.parametrize(
    "severity",
    ["", "FATAL", 3],
)
def test_load_config_invalid_severity_raises(
    tmp_path: Path,
    severity: object,
) -> None:
    config_path = _json_config(
        tmp_path, {"rules": [_rule_payload(severity=severity)]}
    )

    with pytest.raises(QualityGateConfigError) as exc_info:
        load_quality_gate_config(config_path)

    assert "severity" in str(exc_info.value)


@pytest.mark.parametrize(
    "enabled",
    ["yes", 1],
)
def test_load_config_invalid_enabled_raises(
    tmp_path: Path,
    enabled: object,
) -> None:
    config_path = _json_config(
        tmp_path, {"rules": [_rule_payload(enabled=enabled)]}
    )

    with pytest.raises(QualityGateConfigError) as exc_info:
        load_quality_gate_config(config_path)

    assert "enabled" in str(exc_info.value)


@pytest.mark.parametrize(
    "blocking_payload",
    [
        {"blocking": "yes"},
        {},
    ],
)
def test_load_config_invalid_blocking_raises(
    tmp_path: Path,
    blocking_payload: dict[str, object],
) -> None:
    payload = {
        "rule_code": "R1",
        "severity": "BLOCKER",
        "applies_to": "required_field",
        "enabled": True,
        **blocking_payload,
        "parameters": {"standard_fields": ["设备名称"]},
    }
    config_path = _json_config(tmp_path, {"rules": [payload]})

    with pytest.raises(QualityGateConfigError) as exc_info:
        load_quality_gate_config(config_path)

    assert "blocking" in str(exc_info.value)


@pytest.mark.parametrize(
    "applies_to",
    ["", "unknown", 2],
)
def test_load_config_invalid_applies_to_raises(
    tmp_path: Path,
    applies_to: object,
) -> None:
    payload = _rule_payload(applies_to=applies_to)
    config_path = _json_config(tmp_path, {"rules": [payload]})

    with pytest.raises(QualityGateConfigError) as exc_info:
        load_quality_gate_config(config_path)

    assert "applies_to" in str(exc_info.value)


@pytest.mark.parametrize(
    "parameters",
    [
        {},
        {"standard_fields": []},
        {"standard_fields": [""]},
        {"standard_fields": ["设备名称", "设备名称"]},
        {"standard_fields": ["设备名称"], "extra": "x"},
    ],
)
def test_load_config_invalid_required_parameters_raises(
    tmp_path: Path,
    parameters: object,
) -> None:
    payload = _rule_payload(parameters=parameters)
    config_path = _json_config(tmp_path, {"rules": [payload]})

    with pytest.raises(QualityGateConfigError):
        load_quality_gate_config(config_path)


@pytest.mark.parametrize(
    "parameters",
    [
        {},
        {"key_standard_fields": []},
        {"key_standard_fields": ["设备编码"], "extra": 1},
    ],
)
def test_load_config_invalid_duplicate_parameters_raises(
    tmp_path: Path,
    parameters: object,
) -> None:
    payload = {
        "rule_code": "DUPLICATE_KEY",
        "severity": "BLOCKER",
        "applies_to": "duplicate_key",
        "enabled": True,
        "blocking": True,
        "parameters": parameters,
    }
    config_path = _json_config(tmp_path, {"rules": [payload]})

    with pytest.raises(QualityGateConfigError):
        load_quality_gate_config(config_path)


@pytest.mark.parametrize(
    "parameters",
    [
        {},
        {"issue_codes": []},
        {"issue_codes": [""]},
        {"source_stages": ["BAD_STAGE"]},
        {"issue_codes": ["X"], "extra": "y"},
    ],
)
def test_load_config_invalid_template_issue_parameters_raises(
    tmp_path: Path,
    parameters: object,
) -> None:
    payload = {
        "rule_code": "TEMPLATE",
        "severity": "ERROR",
        "applies_to": "template_issue",
        "enabled": True,
        "blocking": True,
        "parameters": parameters,
    }
    config_path = _json_config(tmp_path, {"rules": [payload]})

    with pytest.raises(QualityGateConfigError):
        load_quality_gate_config(config_path)


@pytest.mark.parametrize(
    ("severity", "blocking"),
    [
        ("INFO", True),
        ("WARNING", True),
        ("ERROR", False),
        ("BLOCKER", False),
    ],
)
def test_load_config_severity_blocking_contradiction_raises(
    tmp_path: Path,
    severity: str,
    blocking: bool,
) -> None:
    payload = _rule_payload(
        severity=severity,
        blocking=blocking,
    )
    config_path = _json_config(tmp_path, {"rules": [payload]})

    with pytest.raises(QualityGateConfigError) as exc_info:
        load_quality_gate_config(config_path)

    assert "blocking" in str(exc_info.value)


def _standard_sheets(
    rows: list[list[object]],
) -> dict[str, list[list[object]]]:
    return {
        "设备档案": [
            ["设备名称", "设备类型", "投用日期", "备注"],
            *rows,
        ]
    }


def test_input_sheet_order_mismatch_raises(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "order.xlsx",
        {
            "甲表": [["设备名称", "设备类型"], ["泵A", "除尘器"]],
            "乙表": [["设备名称", "设备类型"], ["风机", "空压机"]],
        },
        (_name_column(), _code_column()),
    )
    reversed_mapping = WorkbookFieldMappingResult(
        file_path=mapping.file_path,
        sheet_results=tuple(reversed(mapping.sheet_results)),
        summary=mapping.summary,
    )

    with pytest.raises(QualityGateDataError) as exc_info:
        run_quality_gate(
            profile,
            reversed_mapping,
            cleaning,
            equipment,
            template,
            _default_config(),
        )

    assert "Sheet" in str(exc_info.value)


def test_input_sheet_name_mismatch_raises(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "names.xlsx",
        _standard_sheets([["泵A", "除尘器", "2024-01-01", "备注"]]),
        (_name_column(), _code_column()),
    )
    cleaning_names = list(cleaning.sheet_results)
    partial_cleaning = CleaningWorkbookResult(
        source_file=cleaning.source_file,
        sheet_results=(),
        summary=cleaning.summary,
    )

    with pytest.raises(QualityGateDataError):
        run_quality_gate(
            profile,
            mapping,
            partial_cleaning,
            equipment,
            template,
            _default_config(),
        )


def test_input_source_file_mismatch_raises(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "source.xlsx",
        _standard_sheets([["泵A", "除尘器", "2024-01-01", "备注"]]),
        (_name_column(), _code_column()),
    )
    moved_cleaning = CleaningWorkbookResult(
        source_file=tmp_path / "other.xlsx",
        sheet_results=cleaning.sheet_results,
        summary=cleaning.summary,
    )

    with pytest.raises(QualityGateDataError) as exc_info:
        run_quality_gate(
            profile,
            mapping,
            moved_cleaning,
            equipment,
            template,
            _default_config(),
        )

    assert "源文件" in str(exc_info.value)


def test_input_row_range_mismatch_raises(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "rows.xlsx",
        _standard_sheets(
            [
                ["泵A", "除尘器", "2024-01-01", "备注"],
                ["风机", "空压机", "2024-02-02", "备注"],
            ]
        ),
        (_name_column(), _code_column()),
    )
    first_sheet = cleaning.sheet_results[0]
    changed_cleaning = CleaningWorkbookResult(
        source_file=cleaning.source_file,
        sheet_results=(
            replace(first_sheet, rows=first_sheet.rows[:-1]),
            *cleaning.sheet_results[1:],
        ),
        summary=cleaning.summary,
    )

    with pytest.raises(QualityGateDataError) as exc_info:
        run_quality_gate(
            profile,
            mapping,
            changed_cleaning,
            equipment,
            template,
            _default_config(),
        )

    assert "行覆盖范围不一致" in str(exc_info.value)


def test_template_row_range_mismatch_raises(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "template-rows.xlsx",
        _standard_sheets(
            [
                ["泵A", "除尘器", "2024-01-01", "备注"],
                ["风机", "空压机", "2024-02-02", "备注"],
            ]
        ),
        (_name_column(), _code_column()),
    )
    first_template_sheet = template.sheet_results[0]
    changed_template = replace(
        template,
        sheet_results=(
            replace(
                first_template_sheet,
                rows=first_template_sheet.rows[:-1],
            ),
            *template.sheet_results[1:],
        ),
    )

    with pytest.raises(QualityGateDataError) as exc_info:
        run_quality_gate(
            profile,
            mapping,
            cleaning,
            equipment,
            changed_template,
            _default_config(),
        )

    assert "模板结果行覆盖范围不一致" in str(exc_info.value)


def _unmatched_sheets() -> dict[str, list[list[object]]]:
    return _standard_sheets(
        [["泵A", "未知类型", "2024-01-01", "备注"]]
    )


def test_template_issue_unmatched_converted(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "unmatched.xlsx",
        _unmatched_sheets(),
        (_name_column(), _code_column()),
    )

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        _default_config(),
    )

    assert result.status is QualityStatus.FAIL
    assert result.can_proceed is False
    issue = result.issues[0]
    assert issue.rule_code == "EQUIPMENT_UNMATCHED"
    assert issue.severity is QualitySeverity.ERROR
    assert issue.source_stage is QualityIssueSourceStage.EQUIPMENT_TYPE_MAPPING
    assert issue.sheet == "设备档案"
    assert issue.row_number == 2
    assert issue.standard_field == "设备类型名称"
    assert issue.template_column == "设备类型编码"
    assert issue.column_number == 2
    assert issue.column_letter == "B"
    assert "未命中" in issue.reason


def test_field_mapping_missing_template_issue_converted(tmp_path: Path) -> None:
    missing_column = TemplateColumnConfig(
        standard_field="设备编码", header="设备编码"
    )
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "missing.xlsx",
        _standard_sheets([["泵A", "除尘器", "2024-01-01", "备注"]]),
        (_name_column(), missing_column),
    )

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        _default_config(),
    )

    issue = result.issues[0]
    assert issue.rule_code == "FIELD_MAPPING_MISSING"
    assert issue.source_stage is QualityIssueSourceStage.FIELD_MAPPING
    assert issue.row_number is None
    assert issue.standard_field == "设备编码"


def test_cleaning_invalid_template_issue_converted(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "invalid-date.xlsx",
        _standard_sheets(
            [["泵A", "除尘器", "不是日期", "备注"]]
        ),
        (_name_column(), _date_column(), _code_column()),
        cleaning_rules=(
            FieldCleaningRule(
                standard_field="投用日期",
                kind=CleaningKind.DATE,
                date_output_format="%Y-%m-%d",
            ),
        ),
    )

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        _default_config(),
    )

    issues = [
        issue
        for issue in result.issues
        if issue.rule_code == "CLEANING_INVALID"
    ]
    assert len(issues) == 1
    assert issues[0].source_stage is QualityIssueSourceStage.CLEANING
    assert issues[0].template_column == "投用日期"
    assert issues[0].row_number == 2
    assert issues[0].column_letter == "C"


def test_equipment_conflict_and_invalid_and_skipped_convert(
    tmp_path: Path,
) -> None:
    conflict_config = _equipment_config(
        rules=(
            EquipmentTypeRule(
                aliases=("COMMON",),
                standard_type_name="类型A",
                system_code="A-001",
            ),
            EquipmentTypeRule(
                aliases=("COMMON",),
                standard_type_name="类型B",
                system_code="B-001",
            ),
        )
    )
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "conflict.xlsx",
        _standard_sheets(
            [["泵A", "COMMON", "2024-01-01", "备注"]]
        ),
        (_name_column(), _code_column()),
        equipment_config=conflict_config,
    )

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        _default_config(),
    )

    issue = result.issues[0]
    assert issue.rule_code == "EQUIPMENT_CONFLICT"
    assert "命中多个规则" in issue.reason

    invalid_config = _equipment_config(
        code_pattern=r"^D-[A-Z0-9]{3}$",
        rules=(
            EquipmentTypeRule(
                aliases=("除尘器",),
                standard_type_name="除尘器",
                system_code="BAD-CODE",
            ),
        ),
    )
    _, profile2, mapping2, cleaning2, equipment2, template2 = _run_template(
        tmp_path,
        "invalid-code.xlsx",
        _standard_sheets(
            [["泵A", "除尘器", "2024-01-01", "备注"]]
        ),
        (_name_column(), _code_column()),
        equipment_config=invalid_config,
    )
    result2 = run_quality_gate(
        profile2,
        mapping2,
        cleaning2,
        equipment2,
        template2,
        _default_config(),
    )
    assert result2.issues[0].rule_code == "EQUIPMENT_INVALID"

    fields_only = (
        StandardFieldDefinition(name="设备名称"),
        StandardFieldDefinition(name="投用日期"),
        StandardFieldDefinition(name="备注"),
    )
    _, profile3, mapping3, cleaning3, equipment3, template3 = _run_template(
        tmp_path,
        "skipped.xlsx",
        {
            "设备档案": [
                ["设备名称", "投用日期", "备注"],
                ["泵A", "2024-01-01", "备注"],
            ]
        },
        (_name_column(), _code_column()),
        standard_fields=fields_only,
        equipment_config=_equipment_config(target="设备类型名称", rules=()),
    )
    result3 = run_quality_gate(
        profile3,
        mapping3,
        cleaning3,
        equipment3,
        template3,
        _default_config(),
    )
    assert result3.issues[0].rule_code == "EQUIPMENT_SKIPPED"


def test_unmapped_template_issue_raises_config_error(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "unmapped.xlsx",
        _unmatched_sheets(),
        (_name_column(), _code_column()),
    )
    config = _config(
        _rule(
            "OTHER_RULE",
            QualityRuleKind.TEMPLATE_ISSUE,
            QualitySeverity.ERROR,
            {"issue_codes": ["SOME_OTHER_CODE"]},
        )
    )

    with pytest.raises(QualityGateConfigError) as exc_info:
        run_quality_gate(
            profile,
            mapping,
            cleaning,
            equipment,
            template,
            config,
        )

    assert "未配置映射规则" in str(exc_info.value)


def test_first_matching_template_rule_wins(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "first-match.xlsx",
        _unmatched_sheets(),
        (_name_column(), _code_column()),
    )
    config = _config(
        _rule(
            "STAGE_RULE",
            QualityRuleKind.TEMPLATE_ISSUE,
            QualitySeverity.ERROR,
            {"source_stages": ["EQUIPMENT_TYPE_MAPPING"]},
        ),
        _rule(
            "CODE_RULE",
            QualityRuleKind.TEMPLATE_ISSUE,
            QualitySeverity.BLOCKER,
            {"issue_codes": ["EQUIPMENT_UNMATCHED"]},
        ),
    )

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        config,
    )

    assert len(result.issues) == 1
    assert result.issues[0].rule_code == "STAGE_RULE"


def test_clean_template_passes_without_issues(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "clean.xlsx",
        _standard_sheets(
            [["泵A", "除尘器", "2024-01-01", "备注"]]
        ),
        (_name_column(), _code_column()),
    )

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        _default_config(),
    )

    assert result.status is QualityStatus.PASS
    assert result.can_proceed is True
    assert result.summary.issue_count == 0


def test_required_field_empty_blocks(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "required.xlsx",
        _standard_sheets(
            [[None, "除尘器", "2024-01-01", "备注"]]
        ),
        (_name_column(), _code_column()),
    )
    config = _default_config(_required_rule(("设备名称",)))

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        config,
    )

    issue = result.issues[0]
    assert issue.rule_code == "REQUIRED_FIELD_EMPTY"
    assert issue.severity is QualitySeverity.BLOCKER
    assert issue.blocking is True
    assert issue.source_stage is QualityIssueSourceStage.QUALITY_GATE
    assert issue.row_number == 2
    assert issue.column_letter == "A"
    assert result.status is QualityStatus.FAIL
    assert result.can_proceed is False


def test_required_field_present_passes(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "required-ok.xlsx",
        _standard_sheets(
            [["泵A", "除尘器", "2024-01-01", "备注"]]
        ),
        (_name_column(), _code_column()),
    )
    config = _default_config(_required_rule(("设备名称",)))

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        config,
    )

    assert result.status is QualityStatus.PASS
    assert result.summary.issue_count == 0


def test_required_field_missing_from_template_raises(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "required-col.xlsx",
        _standard_sheets(
            [["泵A", "除尘器", "2024-01-01", "备注"]]
        ),
        (_name_column(), _code_column()),
    )
    config = _default_config(_required_rule(("设备编码",)))

    with pytest.raises(QualityGateConfigError) as exc_info:
        run_quality_gate(
            profile,
            mapping,
            cleaning,
            equipment,
            template,
            config,
        )

    assert "不在模板" in str(exc_info.value)


def test_inner_empty_row_hits_required_rule(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "inner-blank.xlsx",
        _standard_sheets(
            [
                ["泵A", "除尘器", "2024-01-01", "备注"],
                [None, None, None, None],
                ["风机", "空压机", "2024-02-02", "备注"],
            ]
        ),
        (_name_column(), _code_column()),
    )
    config = _default_config(_required_rule(("设备名称",)))

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        config,
    )

    rows = [
        issue.row_number
        for issue in result.issues
        if issue.rule_code == "REQUIRED_FIELD_EMPTY"
    ]
    assert rows == [3]
    assert result.status is QualityStatus.FAIL


def test_duplicate_key_single_field(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "dup.xlsx",
        _standard_sheets(
            [
                ["泵A", "除尘器", "2024-01-01", "备注"],
                ["泵A", "空压机", "2024-02-02", "备注"],
            ]
        ),
        (_name_column(), _code_column()),
    )
    config = _default_config(_duplicate_rule(("设备名称",)))

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        config,
    )

    issue = result.issues[0]
    assert issue.rule_code == "DUPLICATE_KEY"
    assert issue.row_number == 3
    assert "第 2 行重复" in issue.reason
    assert issue.blocking is True
    assert result.status is QualityStatus.FAIL


def test_duplicate_key_composite(tmp_path: Path) -> None:
    fields = (
        StandardFieldDefinition(name="装置编码"),
        StandardFieldDefinition(name="设备位号"),
        StandardFieldDefinition(name="设备类型名称", aliases=("设备类型",)),
    )
    columns = (
        TemplateColumnConfig(standard_field="装置编码", header="装置编码"),
        TemplateColumnConfig(standard_field="设备位号", header="设备位号"),
    )
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "composite.xlsx",
        {
            "设备档案": [
                ["装置编码", "设备位号", "设备类型"],
                ["U-01", "P-101", "除尘器"],
                ["U-01", "P-101", "空压机"],
                ["U-01", "P-102", "除尘器"],
            ]
        },
        columns,
        standard_fields=fields,
        equipment_config=_equipment_config(target="设备类型名称"),
    )
    config = _default_config(
        _duplicate_rule(("装置编码", "设备位号"))
    )

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        config,
    )

    duplicate_issues = [
        issue
        for issue in result.issues
        if issue.rule_code == "DUPLICATE_KEY"
    ]
    assert len(duplicate_issues) == 1
    assert duplicate_issues[0].row_number == 3
    assert "第 2 行重复" in duplicate_issues[0].reason


def test_empty_keys_skipped_by_default(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "empty-keys.xlsx",
        _standard_sheets(
            [
                [None, "除尘器", "2024-01-01", "备注"],
                [None, "空压机", "2024-02-02", "备注"],
            ]
        ),
        (_name_column(), _code_column()),
    )
    config = _default_config(_duplicate_rule(("设备名称",)))

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        config,
    )

    assert result.summary.issue_count == 0
    assert result.status is QualityStatus.PASS


def test_empty_keys_counted_when_not_skipped(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "empty-keys-counted.xlsx",
        _standard_sheets(
            [
                [None, "除尘器", "2024-01-01", "备注"],
                [None, "空压机", "2024-02-02", "备注"],
            ]
        ),
        (_name_column(), _code_column()),
    )
    config = _default_config(
        _duplicate_rule(("设备名称",), skip_empty_keys=False)
    )

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        config,
    )

    assert len(result.issues) == 1
    assert result.issues[0].rule_code == "DUPLICATE_KEY"


def test_duplicate_missing_key_column_raises(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "dup-missing.xlsx",
        _standard_sheets(
            [["泵A", "除尘器", "2024-01-01", "备注"]]
        ),
        (_name_column(), _code_column()),
    )
    config = _default_config(_duplicate_rule(("设备编码",)))

    with pytest.raises(QualityGateConfigError) as exc_info:
        run_quality_gate(
            profile,
            mapping,
            cleaning,
            equipment,
            template,
            config,
        )

    assert "不在模板" in str(exc_info.value)


@pytest.mark.parametrize(
    ("severity", "expected_status"),
    [
        (QualitySeverity.INFO, QualityStatus.PASS),
        (QualitySeverity.WARNING, QualityStatus.PASS),
    ],
)
def test_info_warning_do_not_fail(
    tmp_path: Path,
    severity: QualitySeverity,
    expected_status: QualityStatus,
) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        f"severity-{severity.value}.xlsx",
        _unmatched_sheets(),
        (_name_column(), _code_column()),
    )
    config = _default_config()
    template_only = _config(
        *_template_rules(severity=severity)
    )

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        template_only,
    )

    assert result.status is expected_status
    assert result.can_proceed is (expected_status is QualityStatus.PASS)


def test_error_fails_and_blocker_fails_with_blocking(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "errors.xlsx",
        _standard_sheets(
            [
                [None, "除尘器", "2024-01-01", "备注"],
                ["泵A", "未知类型", "2024-01-01", "备注"],
            ]
        ),
        (_name_column(), _code_column()),
    )
    config = _config(
        _rule(
            "REQUIRED_ERROR",
            QualityRuleKind.REQUIRED_FIELD,
            QualitySeverity.ERROR,
            {"standard_fields": ["设备名称"]},
        ),
        _rule(
            "EQUIPMENT_BLOCKER",
            QualityRuleKind.TEMPLATE_ISSUE,
            QualitySeverity.BLOCKER,
            {"issue_codes": ["EQUIPMENT_UNMATCHED"]},
        ),
    )

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        config,
    )

    assert result.status is QualityStatus.FAIL
    assert result.can_proceed is False
    error_issue = next(
        issue for issue in result.issues
        if issue.severity is QualitySeverity.ERROR
    )
    blocker_issue = next(
        issue for issue in result.issues
        if issue.severity is QualitySeverity.BLOCKER
    )
    assert error_issue.blocking is True
    assert blocker_issue.blocking is True


def test_mixed_severity_statistics(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "mixed.xlsx",
        _standard_sheets(
            [[None, "未知类型", "2024-01-01", "备注"]]
        ),
        (_name_column(), _code_column()),
    )
    config = _config(
        _rule(
            "UNMATCHED_WARN",
            QualityRuleKind.TEMPLATE_ISSUE,
            QualitySeverity.WARNING,
            {"issue_codes": ["EQUIPMENT_UNMATCHED"]},
        ),
        _required_rule(("设备名称",), severity=QualitySeverity.BLOCKER),
    )

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        config,
    )

    assert result.summary.issue_count == 2
    assert result.summary.warning_count == 1
    assert result.summary.blocker_count == 1
    assert result.summary.error_count == 0
    assert result.status is QualityStatus.FAIL


def test_disabled_rule_produces_no_issue(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "disabled.xlsx",
        _standard_sheets(
            [[None, "除尘器", "2024-01-01", "备注"]]
        ),
        (_name_column(), _code_column()),
    )
    config = _default_config(
        _required_rule(
            ("设备名称",),
            enabled=False,
            rule_code="DISABLED_REQUIRED",
        )
    )

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        config,
    )

    assert result.summary.issue_count == 0
    assert result.status is QualityStatus.PASS


def test_multi_sheet_aggregation_preserves_order(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "multi.xlsx",
        {
            "甲表": [
                ["设备名称", "设备类型"],
                ["泵A", "除尘器"],
            ],
            "乙表": [
                ["设备名称", "设备类型"],
                ["风机", "未知类型"],
            ],
        },
        (_name_column(), _code_column()),
    )

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        _default_config(),
    )

    assert [sheet.sheet_name for sheet in result.sheet_results] == [
        "甲表",
        "乙表",
    ]
    assert result.sheet_results[0].status is QualityStatus.PASS
    assert result.sheet_results[1].status is QualityStatus.FAIL
    assert result.status is QualityStatus.FAIL
    assert len(result.issues) == 1


def test_empty_sheet_produces_no_row_issue(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "empty-sheet.xlsx",
        {
            "空表": [
                ["设备名称", "设备类型"],
            ]
        },
        (_name_column(), _code_column()),
    )
    config = _default_config(_required_rule(("设备名称",)))

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        config,
    )

    sheet_result = result.sheet_results[0]
    assert sheet_result.status is QualityStatus.PASS
    assert result.summary.issue_count == 0


def test_template_issue_not_counted_from_upstream_duplicates(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "dedup.xlsx",
        _unmatched_sheets(),
        (_name_column(), _code_column()),
    )
    config = _default_config()

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        config,
    )

    assert len(result.issues) == 1
    assert len(equipment.sheet_results[0].audits) == 1
    assert result.summary.issue_count == 1


def test_severity_statistics_summary(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "stats.xlsx",
        _standard_sheets(
            [
                [None, "除尘器", "2024-01-01", "备注"],
                ["泵A", "未知类型", "2024-01-01", "备注"],
                ["泵B", "空压机", "2024-01-01", "备注"],
            ]
        ),
        (_name_column(), _code_column()),
    )
    config = _default_config(_required_rule(("设备名称",)))

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        config,
    )

    assert isinstance(result.summary, QualitySummary)
    assert result.summary.issue_count == 2
    assert result.summary.error_count == 1
    assert result.summary.blocker_count == 1
    assert result.summary.info_count == 0
    assert result.summary.warning_count == 0


def test_rule_code_statistics_summary(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "rule-stats.xlsx",
        _standard_sheets(
            [[None, "未知类型", "2024-01-01", "备注"]]
        ),
        (_name_column(), _code_column()),
    )
    config = _default_config(_required_rule(("设备名称",)))

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        config,
    )

    rule_counts = {
        item.rule_code: item.count for item in result.summary.rule_counts
    }
    assert rule_counts["EQUIPMENT_UNMATCHED"] == 1
    assert rule_counts["REQUIRED_FIELD_EMPTY"] == 1


def test_quality_issue_keeps_trace_fields(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "trace.xlsx",
        _standard_sheets(
            [["泵A", "未知类型", "2024-01-01", "备注"]]
        ),
        (_name_column(), _code_column()),
    )

    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        _default_config(),
    )

    issue = result.issues[0]
    assert isinstance(issue, QualityIssue)
    assert issue.rule_code == "EQUIPMENT_UNMATCHED"
    assert issue.sheet == "设备档案"
    assert issue.row_number == 2
    assert issue.column_number == 2
    assert issue.column_letter == "B"
    assert issue.standard_field == "设备类型名称"
    assert issue.template_column == "设备类型编码"
    assert issue.reason
    assert result.source_file == path


def test_run_does_not_modify_input_objects(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "pure.xlsx",
        _unmatched_sheets(),
        (_name_column(), _code_column()),
    )
    profile_before = profile
    mapping_before = mapping
    cleaning_before = cleaning
    equipment_before = equipment
    template_before = template

    run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        _default_config(),
    )

    assert profile == profile_before
    assert mapping == mapping_before
    assert cleaning == cleaning_before
    assert equipment == equipment_before
    assert template == template_before


def test_run_creates_no_files(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "no-write.xlsx",
        _unmatched_sheets(),
        (_name_column(), _code_column()),
    )
    before_files = {item.name for item in tmp_path.iterdir()}
    source_bytes = path.read_bytes()

    run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        _default_config(),
    )

    assert {item.name for item in tmp_path.iterdir()} == before_files
    assert path.read_bytes() == source_bytes


def test_result_object_shapes(tmp_path: Path) -> None:
    path, profile, mapping, cleaning, equipment, template = _run_template(
        tmp_path,
        "shapes.xlsx",
        _standard_sheets(
            [[None, "除尘器", "2024-01-01", "备注"]]
        ),
        (_name_column(), _code_column()),
    )
    result = run_quality_gate(
        profile,
        mapping,
        cleaning,
        equipment,
        template,
        _default_config(_required_rule(("设备名称",))),
    )

    assert isinstance(result, QualityWorkbookResult)
    assert isinstance(result.sheet_results[0], QualitySheetResult)
    assert result.sheet_results[0].issues
    assert result.sheet_results[0].summary.issue_count == 1
