"""P1-05 设备类型编码映射测试。"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from excel_qc.cleaning import (
    CleaningSummary,
    CleaningWorkbookConfig,
    CleaningWorkbookResult,
    FieldCleaningRule,
    clean_workbook,
)
from excel_qc.equipment_type_mapping import (
    EquipmentTypeCodeConfig,
    EquipmentTypeConfigError,
    EquipmentTypeRule,
    EquipmentTypeStatus,
    EquipmentTypeWorkbookResult,
    load_equipment_type_config,
    map_equipment_type_codes,
)
from excel_qc.field_mapping import (
    FieldMappingConfig,
    StandardFieldDefinition,
    map_workbook_fields,
)
from excel_qc.profiler import profile_workbook
from tests.qc.helpers import build_workbook


def _run_one(
    tmp_path: Path,
    filename: str,
    *,
    headers: list[str],
    rows: list[list[object]],
    standard_fields: tuple[StandardFieldDefinition, ...],
    equipment_config: EquipmentTypeCodeConfig,
    cleaning_rules: tuple[FieldCleaningRule, ...] = (),
    sheet_name: str = "Sheet1",
) -> EquipmentTypeWorkbookResult:
    path = build_workbook(
        tmp_path / filename,
        {sheet_name: [headers, *rows]},
    )
    profile = profile_workbook(path)
    mapping = map_workbook_fields(
        profile,
        FieldMappingConfig(standard_fields),
    )
    cleaning_result = clean_workbook(
        profile,
        mapping,
        CleaningWorkbookConfig(cleaning_rules),
    )
    return map_equipment_type_codes(cleaning_result, equipment_config)


def _simple_config(
    *,
    target: str = "设备类型",
    rules: tuple[EquipmentTypeRule, ...] = (
        EquipmentTypeRule(
            aliases=("除尘器", "除尘设备"),
            standard_type_name="除尘器",
            system_code="D-CCQ",
        ),
    ),
    code_pattern: str | None = None,
) -> EquipmentTypeCodeConfig:
    return EquipmentTypeCodeConfig(
        target_standard_field=target,
        rules=rules,
        code_pattern=code_pattern,
    )


def _empty_cleaning_result() -> CleaningWorkbookResult:
    return CleaningWorkbookResult(
        source_file=Path("sample.xlsx"),
        sheet_results=(),
        summary=CleaningSummary(),
    )


def test_exact_match(tmp_path: Path) -> None:
    result = _run_one(
        tmp_path,
        "exact.xlsx",
        headers=["设备类型"],
        rows=[["除尘器"]],
        standard_fields=(StandardFieldDefinition(name="设备类型"),),
        equipment_config=_simple_config(),
    )

    audit = result.sheet_results[0].audits[0]
    assert audit.status is EquipmentTypeStatus.MATCH
    assert audit.matched_standard_type == "除尘器"
    assert audit.system_code == "D-CCQ"
    assert result.summary.match_count == 1


def test_alias_match(tmp_path: Path) -> None:
    result = _run_one(
        tmp_path,
        "alias.xlsx",
        headers=["设备类型"],
        rows=[["除尘设备"]],
        standard_fields=(StandardFieldDefinition(name="设备类型"),),
        equipment_config=_simple_config(),
    )

    audit = result.sheet_results[0].audits[0]
    assert audit.status is EquipmentTypeStatus.MATCH
    assert audit.system_code == "D-CCQ"


def test_leading_trailing_whitespace_is_normalized(tmp_path: Path) -> None:
    result = _run_one(
        tmp_path,
        "space.xlsx",
        headers=["设备类型"],
        rows=[["  除尘器  "]],
        standard_fields=(StandardFieldDefinition(name="设备类型"),),
        equipment_config=_simple_config(),
        cleaning_rules=(
            FieldCleaningRule(standard_field="设备类型"),
        ),
    )

    audit = result.sheet_results[0].audits[0]
    assert audit.status is EquipmentTypeStatus.MATCH
    assert audit.cleaned_value == "除尘器"


def test_full_width_value_is_normalized(tmp_path: Path) -> None:
    result = _run_one(
        tmp_path,
        "fullwidth.xlsx",
        headers=["设备类型"],
        rows=[["ＡＩＲ＿ＣＯＭＰＲＥＳＳＯＲ"]],
        standard_fields=(StandardFieldDefinition(name="设备类型"),),
        equipment_config=_simple_config(
            rules=(
                EquipmentTypeRule(
                    aliases=("AIR_COMPRESSOR",),
                    standard_type_name="空压机",
                    system_code="D-KYJ",
                ),
            )
        ),
    )

    audit = result.sheet_results[0].audits[0]
    assert audit.status is EquipmentTypeStatus.MATCH
    assert audit.system_code == "D-KYJ"


def test_case_insensitive_match(tmp_path: Path) -> None:
    result = _run_one(
        tmp_path,
        "case.xlsx",
        headers=["设备类型"],
        rows=[["air_compressor"]],
        standard_fields=(StandardFieldDefinition(name="设备类型"),),
        equipment_config=_simple_config(
            rules=(
                EquipmentTypeRule(
                    aliases=("AIR_COMPRESSOR",),
                    standard_type_name="空压机",
                    system_code="D-KYJ",
                ),
            )
        ),
    )

    audit = result.sheet_results[0].audits[0]
    assert audit.status is EquipmentTypeStatus.MATCH
    assert audit.system_code == "D-KYJ"


def test_unmatched_value_is_reported(tmp_path: Path) -> None:
    result = _run_one(
        tmp_path,
        "unmatched.xlsx",
        headers=["设备类型"],
        rows=[["未知类型"]],
        standard_fields=(StandardFieldDefinition(name="设备类型"),),
        equipment_config=_simple_config(),
    )

    audit = result.sheet_results[0].audits[0]
    assert audit.status is EquipmentTypeStatus.UNMATCHED
    assert audit.system_code is None
    assert "未命中" in audit.reason


def test_conflicting_aliases_are_not_auto_selected(tmp_path: Path) -> None:
    result = _run_one(
        tmp_path,
        "conflict.xlsx",
        headers=["设备类型"],
        rows=[["COMMON"]],
        standard_fields=(StandardFieldDefinition(name="设备类型"),),
        equipment_config=_simple_config(
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
        ),
    )

    audit = result.sheet_results[0].audits[0]
    assert audit.status is EquipmentTypeStatus.CONFLICT
    assert audit.system_code is None
    assert "命中多个规则" in audit.reason
    assert result.summary.conflict_count == 1


def test_invalid_code_pattern_is_marked(tmp_path: Path) -> None:
    config = EquipmentTypeCodeConfig(
        target_standard_field="设备类型",
        code_pattern=r"^D-[A-Z0-9]{3}$",
        rules=(
            EquipmentTypeRule(
                aliases=("除尘器",),
                standard_type_name="除尘器",
                system_code="BAD-CODE",
            ),
        ),
    )
    result = _run_one(
        tmp_path,
        "invalid.xlsx",
        headers=["设备类型"],
        rows=[["除尘器"]],
        standard_fields=(StandardFieldDefinition(name="设备类型"),),
        equipment_config=config,
    )

    audit = result.sheet_results[0].audits[0]
    assert audit.status is EquipmentTypeStatus.INVALID
    assert audit.system_code == "BAD-CODE"
    assert "code_pattern" in audit.reason


def test_valid_code_passes_code_pattern(tmp_path: Path) -> None:
    result = _run_one(
        tmp_path,
        "valid-code.xlsx",
        headers=["设备类型"],
        rows=[["除尘器"]],
        standard_fields=(StandardFieldDefinition(name="设备类型"),),
        equipment_config=_simple_config(code_pattern=r"^D-[A-Z]{3}$"),
    )

    audit = result.sheet_results[0].audits[0]
    assert audit.status is EquipmentTypeStatus.MATCH


def test_empty_value_is_not_mapped(tmp_path: Path) -> None:
    result = _run_one(
        tmp_path,
        "empty.xlsx",
        headers=["设备类型", "其他字段"],
        rows=[[None, "保留值"]],
        standard_fields=(
            StandardFieldDefinition(name="设备类型"),
            StandardFieldDefinition(name="其他字段"),
        ),
        equipment_config=_simple_config(),
    )

    audit = result.sheet_results[0].audits[0]
    assert audit.status is EquipmentTypeStatus.UNMATCHED
    assert audit.system_code is None
    assert "空值" in audit.reason


def test_target_field_missing_is_skipped(tmp_path: Path) -> None:
    result = _run_one(
        tmp_path,
        "missing-target.xlsx",
        headers=["其他字段"],
        rows=[["值"]],
        standard_fields=(StandardFieldDefinition(name="其他字段"),),
        equipment_config=_simple_config(target="设备类型"),
    )

    audit = result.sheet_results[0].audits[0]
    assert audit.status is EquipmentTypeStatus.SKIPPED
    assert audit.system_code is None
    assert "未找到目标标准字段" in audit.reason


def test_ambiguous_target_field_is_skipped(tmp_path: Path) -> None:
    result = _run_one(
        tmp_path,
        "ambiguous-field.xlsx",
        headers=["设备类型"],
        rows=[["除尘器"]],
        standard_fields=(
            StandardFieldDefinition(name="设备A", aliases=("设备类型",)),
            StandardFieldDefinition(name="设备B", aliases=("设备类型",)),
        ),
        equipment_config=_simple_config(target="设备A"),
    )

    audit = result.sheet_results[0].audits[0]
    assert audit.status is EquipmentTypeStatus.SKIPPED
    assert audit.system_code is None


def test_conflicting_target_field_is_skipped_with_reason(tmp_path: Path) -> None:
    result = _run_one(
        tmp_path,
        "conflict-field.xlsx",
        headers=["设备名称", "别名"],
        rows=[["设备A", "别名A"]],
        standard_fields=(
            StandardFieldDefinition(
                name="设备类型名称",
                aliases=("设备名称", "别名"),
            ),
        ),
        equipment_config=_simple_config(target="设备类型名称"),
    )

    audits = result.sheet_results[0].audits
    assert len(audits) == 1
    assert audits[0].status is EquipmentTypeStatus.SKIPPED
    assert "CONFLICT" in audits[0].reason


def test_multi_sheet_mapping(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "multi.xlsx",
        {
            "设备档案": [["设备类型"], ["除尘器"]],
            "补充资料": [["设备类型"], ["空压机"]],
        },
    )
    profile = profile_workbook(path)
    mapping = map_workbook_fields(
        profile,
        FieldMappingConfig(
            (StandardFieldDefinition(name="设备类型"),)
        ),
    )
    cleaning_result = clean_workbook(
        profile,
        mapping,
        CleaningWorkbookConfig(
            (FieldCleaningRule(standard_field="设备类型"),)
        ),
    )
    equipment_config = EquipmentTypeCodeConfig(
        target_standard_field="设备类型",
        rules=(
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
        ),
    )

    result = map_equipment_type_codes(cleaning_result, equipment_config)

    assert len(result.sheet_results) == 2
    assert result.sheet_results[0].audits[0].system_code == "D-CCQ"
    assert result.sheet_results[1].audits[0].system_code == "D-KYJ"
    assert result.summary.match_count == 2


def test_audit_preserves_full_source_trace(tmp_path: Path) -> None:
    result = _run_one(
        tmp_path,
        "trace.xlsx",
        headers=["设备类型"],
        rows=[["  除尘器  "]],
        standard_fields=(StandardFieldDefinition(name="设备类型"),),
        equipment_config=_simple_config(),
        cleaning_rules=(
            FieldCleaningRule(standard_field="设备类型"),
        ),
    )

    audit = result.sheet_results[0].audits[0]
    assert audit.sheet_name == "Sheet1"
    assert audit.row_number == 2
    assert audit.source_column_number == 1
    assert audit.source_column_letter == "A"
    assert audit.source_field_name == "设备类型"
    assert audit.cleaned_value == "除尘器"
    assert audit.matched_standard_type == "除尘器"
    assert audit.system_code == "D-CCQ"
    assert audit.reason


def test_json_configuration_loads(tmp_path: Path) -> None:
    config_path = tmp_path / "equipment.json"
    config_path.write_text(
        json.dumps(
            {
                "target_standard_field": "设备类型",
                "code_pattern": "^D-[A-Z]{3}$",
                "rules": [
                    {
                        "aliases": ["除尘器", "除尘设备"],
                        "standard_type_name": "除尘器",
                        "system_code": "D-CCQ",
                    }
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    config = load_equipment_type_config(config_path)

    assert config.target_standard_field == "设备类型"
    assert len(config.rules) == 1
    assert config.rules[0].system_code == "D-CCQ"
    assert config.code_pattern == "^D-[A-Z]{3}$"


def test_missing_config_file_raises(tmp_path: Path) -> None:
    with pytest.raises(EquipmentTypeConfigError) as exc_info:
        load_equipment_type_config(tmp_path / "missing.json")

    assert "不存在" in str(exc_info.value)


@pytest.mark.parametrize(
    ("payload", "expected_text"),
    [
        ("{bad json", "无法解析"),
        ("[]", "顶层必须是 JSON 对象"),
        ('{"rules": []}', "target_standard_field"),
        ('{"target_standard_field": "设备类型"}', "rules"),
        (
            '{"target_standard_field": "设备类型", "rules": [{"standard_type_name": "除尘器"}]}',
            "aliases",
        ),
        (
            '{"target_standard_field": "设备类型", "rules": [{"aliases": ["除尘器"], "system_code": "D-CCQ"}]}',
            "standard_type_name",
        ),
        (
            '{"target_standard_field": "设备类型", "rules": [{"aliases": ["除尘器"], "standard_type_name": "除尘器"}]}',
            "system_code",
        ),
    ],
)
def test_invalid_json_config_raises(
    tmp_path: Path,
    payload: str,
    expected_text: str,
) -> None:
    config_path = tmp_path / "bad-equipment.json"
    config_path.write_text(payload, encoding="utf-8")

    with pytest.raises(EquipmentTypeConfigError) as exc_info:
        load_equipment_type_config(config_path)

    assert expected_text in str(exc_info.value)


def test_duplicate_rule_raises() -> None:
    config = EquipmentTypeCodeConfig(
        target_standard_field="设备类型",
        rules=(
            EquipmentTypeRule(
                aliases=("除尘器",),
                standard_type_name="除尘器",
                system_code="D-CCQ",
            ),
            EquipmentTypeRule(
                aliases=("除尘器",),
                standard_type_name="除尘器",
                system_code="D-CCQ",
            ),
        ),
    )

    with pytest.raises(EquipmentTypeConfigError) as exc_info:
        map_equipment_type_codes(_empty_cleaning_result(), config)

    assert "重复" in str(exc_info.value)


def test_empty_alias_raises_in_json(tmp_path: Path) -> None:
    config_path = tmp_path / "empty-alias.json"
    config_path.write_text(
        json.dumps(
            {
                "target_standard_field": "设备类型",
                "rules": [
                    {
                        "aliases": [""],
                        "standard_type_name": "除尘器",
                        "system_code": "D-CCQ",
                    }
                ],
            }
        ),
        encoding="utf-8",
    )

    with pytest.raises(EquipmentTypeConfigError) as exc_info:
        load_equipment_type_config(config_path)

    assert "aliases" in str(exc_info.value)


def test_invalid_code_pattern_raises(tmp_path: Path) -> None:
    config = EquipmentTypeCodeConfig(
        target_standard_field="设备类型",
        rules=(),
        code_pattern="[",
    )

    with pytest.raises(EquipmentTypeConfigError) as exc_info:
        map_equipment_type_codes(_empty_cleaning_result(), config)

    assert "code_pattern" in str(exc_info.value)


def test_input_cleaning_result_is_not_modified(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "readonly.xlsx",
        {"Sheet1": [["设备类型"], ["除尘器"]]},
    )
    profile = profile_workbook(path)
    mapping = map_workbook_fields(
        profile,
        FieldMappingConfig(
            (StandardFieldDefinition(name="设备类型"),)
        ),
    )
    cleaning_result = clean_workbook(
        profile,
        mapping,
        CleaningWorkbookConfig(
            (FieldCleaningRule(standard_field="设备类型"),)
        ),
    )
    before = cleaning_result.sheet_results[0].rows

    map_equipment_type_codes(cleaning_result, _simple_config())

    assert cleaning_result.sheet_results[0].rows == before


def test_mapping_does_not_write_excel(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "no-write.xlsx",
        {"Sheet1": [["设备类型"], ["除尘器"]]},
    )
    before = path.read_bytes()
    profile = profile_workbook(path)
    mapping = map_workbook_fields(
        profile,
        FieldMappingConfig(
            (StandardFieldDefinition(name="设备类型"),)
        ),
    )
    cleaning_result = clean_workbook(
        profile,
        mapping,
        CleaningWorkbookConfig(),
    )

    map_equipment_type_codes(cleaning_result, _simple_config())

    assert path.read_bytes() == before


def test_summary_counts(tmp_path: Path) -> None:
    result = _run_one(
        tmp_path,
        "summary.xlsx",
        headers=["设备类型"],
        rows=[["除尘器"], ["未知"], ["COMMON"]],
        standard_fields=(StandardFieldDefinition(name="设备类型"),),
        equipment_config=EquipmentTypeCodeConfig(
            target_standard_field="设备类型",
            rules=(
                EquipmentTypeRule(
                    aliases=("除尘器",),
                    standard_type_name="除尘器",
                    system_code="D-CCQ",
                ),
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
            ),
        ),
    )

    summary = result.summary
    assert summary.total == 3
    assert summary.match_count == 1
    assert summary.unmatched_count == 1
    assert summary.conflict_count == 1
