"""P1-04 数据清洗引擎测试。"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

import pytest

from excel_qc.cleaning import (
    CleanCellAudit,
    CleanedSheetResult,
    CleaningAction,
    CleaningConfigError,
    CleaningKind,
    CleaningWorkbookConfig,
    CleaningWorkbookResult,
    FieldCleaningRule,
    clean_workbook,
    load_cleaning_config,
)
from excel_qc.field_mapping import (
    FieldMappingConfig,
    StandardFieldDefinition,
    map_workbook_fields,
)
from excel_qc.loader import load_workbook
from excel_qc.profiler import WorkbookProfile, profile_workbook
from tests.qc.helpers import build_workbook


def _run(
    tmp_path: Path,
    filename: str,
    *,
    headers: list[str],
    rows: list[list[object]],
    standard_fields: tuple[StandardFieldDefinition, ...],
    cleaning_rules: tuple[FieldCleaningRule, ...],
    sheet_name: str = "Sheet1",
) -> CleaningWorkbookResult:
    path = build_workbook(
        tmp_path / filename,
        {sheet_name: [headers, *rows]},
    )
    profile = profile_workbook(path)
    mapping = map_workbook_fields(
        profile,
        FieldMappingConfig(standard_fields),
    )
    return clean_workbook(
        profile,
        mapping,
        CleaningWorkbookConfig(cleaning_rules),
    )


def _first_row(result: CleaningWorkbookResult):
    return result.sheet_results[0].rows[0]


def _first_cell(result: CleaningWorkbookResult) -> CleanCellAudit:
    return _first_row(result).cells[0]


def _empty_profile() -> WorkbookProfile:
    return WorkbookProfile(path=Path("sample.xlsx"), sheets=())


def test_text_whitespace_is_trimmed(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "trim.xlsx",
        headers=["设备名称"],
        rows=[["  空压机  "], [" 冷水机 "]],
        standard_fields=(StandardFieldDefinition(name="设备名称"),),
        cleaning_rules=(
            FieldCleaningRule(standard_field="设备名称", kind=CleaningKind.TEXT),
        ),
    )

    assert result.sheet_results[0].rows[0].values == ("空压机",)
    assert _first_cell(result).action is CleaningAction.CLEANED
    assert _first_cell(result).reason == "文本标准化"


def test_full_width_text_is_normalized(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "fullwidth.xlsx",
        headers=["代码"],
        rows=[["ＡＢＣ１２３"]],
        standard_fields=(StandardFieldDefinition(name="代码"),),
        cleaning_rules=(
            FieldCleaningRule(standard_field="代码", kind=CleaningKind.TEXT),
        ),
    )

    row = _first_row(result)
    assert row.values == ("ABC123",)
    assert _first_cell(result).action is CleaningAction.CLEANED


def test_invisible_characters_are_removed(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "invisible.xlsx",
        headers=["设备名称"],
        rows=[["设备\u200b名称"]],
        standard_fields=(StandardFieldDefinition(name="设备名称"),),
        cleaning_rules=(
            FieldCleaningRule(standard_field="设备名称", kind=CleaningKind.TEXT),
        ),
    )

    assert _first_row(result).values == ("设备名称",)


def test_newline_is_normalized(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "newline.xlsx",
        headers=["备注"],
        rows=[["第一行\r\n第二行"]],
        standard_fields=(StandardFieldDefinition(name="备注"),),
        cleaning_rules=(
            FieldCleaningRule(standard_field="备注", kind=CleaningKind.TEXT),
        ),
    )

    assert _first_row(result).values == ("第一行 第二行",)


def test_configured_special_characters_are_removed(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "special.xlsx",
        headers=["规格"],
        rows=[["10×20"]],
        standard_fields=(StandardFieldDefinition(name="规格"),),
        cleaning_rules=(
            FieldCleaningRule(
                standard_field="规格",
                kind=CleaningKind.TEXT,
                remove_characters=("×",),
            ),
        ),
    )

    assert _first_row(result).values == ("1020",)


def test_empty_marker_is_normalized_per_field(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "empty.xlsx",
        headers=["设备名称"],
        rows=[["无"]],
        standard_fields=(StandardFieldDefinition(name="设备名称"),),
        cleaning_rules=(
            FieldCleaningRule(
                standard_field="设备名称",
                kind=CleaningKind.TEXT,
                empty_markers=("无", "/", "NULL"),
            ),
        ),
    )

    cell = _first_cell(result)
    assert _first_row(result).values == ("",)
    assert cell.action is CleaningAction.EMPTY_NORMALIZED
    assert cell.cleaned_value == ""


def test_empty_marker_matching_is_case_insensitive(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "empty-case.xlsx",
        headers=["设备名称"],
        rows=[["NULL"]],
        standard_fields=(StandardFieldDefinition(name="设备名称"),),
        cleaning_rules=(
            FieldCleaningRule(
                standard_field="设备名称",
                kind=CleaningKind.TEXT,
                empty_markers=("null",),
            ),
        ),
    )

    assert _first_cell(result).action is CleaningAction.EMPTY_NORMALIZED


def test_valid_date_text_is_converted(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "date.xlsx",
        headers=["投用日期"],
        rows=[["2026/08/01"]],
        standard_fields=(StandardFieldDefinition(name="投用日期"),),
        cleaning_rules=(
            FieldCleaningRule(
                standard_field="投用日期",
                kind=CleaningKind.DATE,
                date_output_format="%Y-%m-%d",
            ),
        ),
    )

    assert _first_row(result).values == ("2026-08-01",)
    assert _first_cell(result).action is CleaningAction.CLEANED


def test_date_object_is_converted(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "date-object.xlsx",
        headers=["投用日期"],
        rows=[[datetime(2026, 8, 1, 10, 30)]],
        standard_fields=(StandardFieldDefinition(name="投用日期"),),
        cleaning_rules=(
            FieldCleaningRule(
                standard_field="投用日期",
                kind=CleaningKind.DATE,
                date_output_format="%Y-%m-%d",
            ),
        ),
    )

    assert _first_row(result).values == ("2026-08-01",)


def test_invalid_date_is_preserved_and_marked(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "bad-date.xlsx",
        headers=["投用日期"],
        rows=[["2026-13-45"]],
        standard_fields=(StandardFieldDefinition(name="投用日期"),),
        cleaning_rules=(
            FieldCleaningRule(
                standard_field="投用日期",
                kind=CleaningKind.DATE,
                date_output_format="%Y-%m-%d",
            ),
        ),
    )

    cell = _first_cell(result)
    assert cell.action is CleaningAction.INVALID
    assert cell.cleaned_value == "2026-13-45"
    assert "无法识别" in cell.reason


def test_numeric_text_is_normalized(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "number.xlsx",
        headers=["数量"],
        rows=[["100"]],
        standard_fields=(StandardFieldDefinition(name="数量"),),
        cleaning_rules=(
            FieldCleaningRule(standard_field="数量", kind=CleaningKind.NUMBER),
        ),
    )

    assert _first_row(result).values == ("100",)


def test_thousand_separators_are_removed(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "thousand.xlsx",
        headers=["金额"],
        rows=[["1,234.56"]],
        standard_fields=(StandardFieldDefinition(name="金额"),),
        cleaning_rules=(
            FieldCleaningRule(standard_field="金额", kind=CleaningKind.NUMBER),
        ),
    )

    assert _first_row(result).values == ("1234.56",)


def test_decimal_trailing_zeros_are_normalized(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "decimal.xlsx",
        headers=["单价"],
        rows=[["12.500"]],
        standard_fields=(StandardFieldDefinition(name="单价"),),
        cleaning_rules=(
            FieldCleaningRule(standard_field="单价", kind=CleaningKind.NUMBER),
        ),
    )

    assert _first_row(result).values == ("12.5",)


def test_configured_unit_is_stripped(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "unit.xlsx",
        headers=["数量"],
        rows=[["100台"]],
        standard_fields=(StandardFieldDefinition(name="数量"),),
        cleaning_rules=(
            FieldCleaningRule(
                standard_field="数量",
                kind=CleaningKind.NUMBER,
                number_units=("台",),
            ),
        ),
    )

    assert _first_row(result).values == ("100",)


def test_unconfigured_unit_is_preserved_and_marked_invalid(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "bad-unit.xlsx",
        headers=["数量"],
        rows=[["100台"]],
        standard_fields=(StandardFieldDefinition(name="数量"),),
        cleaning_rules=(
            FieldCleaningRule(standard_field="数量", kind=CleaningKind.NUMBER),
        ),
    )

    cell = _first_cell(result)
    assert cell.action is CleaningAction.INVALID
    assert cell.cleaned_value == "100台"


def test_matched_field_is_cleaned_and_unmatched_is_preserved(
    tmp_path: Path,
) -> None:
    result = _run(
        tmp_path,
        "mixed-status.xlsx",
        headers=["设备名称", "未知说明"],
        rows=[["  空压机  ", "原值"]],
        standard_fields=(StandardFieldDefinition(name="设备名称"),),
        cleaning_rules=(
            FieldCleaningRule(standard_field="设备名称", kind=CleaningKind.TEXT),
        ),
    )

    row = _first_row(result)
    assert row.values == ("空压机", "原值")
    assert row.cells[0].action is CleaningAction.CLEANED
    assert row.cells[1].action is CleaningAction.UNMAPPED
    assert row.cells[1].cleaned_value == "原值"


def test_unmatched_reason_is_explicit(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "unmatched.xlsx",
        headers=["未知字段"],
        rows=[["值"]],
        standard_fields=(StandardFieldDefinition(name="已知字段"),),
        cleaning_rules=(),
    )

    cell = _first_cell(result)
    assert cell.action is CleaningAction.UNMAPPED
    assert "未映射" in cell.reason


def test_ambiguous_mapping_preserves_value(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "ambiguous.xlsx",
        headers=["COMMON"],
        rows=[["原始值"]],
        standard_fields=(
            StandardFieldDefinition(name="A", aliases=("COMMON",)),
            StandardFieldDefinition(name="B", aliases=("COMMON",)),
        ),
        cleaning_rules=(),
    )

    cell = _first_cell(result)
    assert cell.action is CleaningAction.AMBIGUOUS
    assert cell.cleaned_value == "原始值"
    assert "歧义" in cell.reason


def test_conflict_mapping_preserves_all_values(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "conflict.xlsx",
        headers=["设备名称", "名称"],
        rows=[["空压机", "备选名称"]],
        standard_fields=(
            StandardFieldDefinition(
                name="设备档案.设备名称",
                aliases=("设备名称", "名称"),
            ),
        ),
        cleaning_rules=(
            FieldCleaningRule(
                standard_field="设备档案.设备名称",
                kind=CleaningKind.TEXT,
            ),
        ),
    )

    row = _first_row(result)
    assert row.values == ("空压机", "备选名称")
    assert [cell.action for cell in row.cells] == [
        CleaningAction.CONFLICT,
        CleaningAction.CONFLICT,
    ]


def test_matched_field_without_rule_is_not_changed(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "no-rule.xlsx",
        headers=["设备名称"],
        rows=[["  原值  "]],
        standard_fields=(StandardFieldDefinition(name="设备名称"),),
        cleaning_rules=(),
    )

    cell = _first_cell(result)
    assert cell.action is CleaningAction.NO_RULE
    assert cell.cleaned_value == "  原值  "
    assert "没有配置清洗规则" in cell.reason


def test_multi_sheet_cleaning(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "multi.xlsx",
        {
            "设备档案": [["设备名称"], ["  空压机  "]],
            "技术参数": [["参数名称"], [" 额定功率 "]],
        },
    )
    profile = profile_workbook(path)
    mapping_config = FieldMappingConfig(
        (
            StandardFieldDefinition(name="设备名称"),
            StandardFieldDefinition(name="参数名称"),
        )
    )
    mapping = map_workbook_fields(profile, mapping_config)
    cleaning_config = CleaningWorkbookConfig(
        (
            FieldCleaningRule(standard_field="设备名称", kind=CleaningKind.TEXT),
            FieldCleaningRule(standard_field="参数名称", kind=CleaningKind.TEXT),
        )
    )

    result = clean_workbook(profile, mapping, cleaning_config)

    assert len(result.sheet_results) == 2
    assert result.sheet_results[0].rows[0].values == ("空压机",)
    assert result.sheet_results[1].rows[0].values == ("额定功率",)
    assert result.summary.cleaned_count == 2


def test_audit_preserves_full_source_trace(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "audit.xlsx",
        headers=["设备名称"],
        rows=[["  空压机  "]],
        standard_fields=(StandardFieldDefinition(name="设备名称"),),
        cleaning_rules=(
            FieldCleaningRule(
                standard_field="设备名称",
                kind=CleaningKind.TEXT,
                rule_code="RULE_001",
            ),
        ),
    )

    cell = _first_cell(result)
    assert cell.sheet_name == "Sheet1"
    assert cell.row_number == 2
    assert cell.source_column_number == 1
    assert cell.source_column_letter == "A"
    assert cell.source_field_name == "设备名称"
    assert cell.standard_field_name == "设备名称"
    assert cell.original_value == "  空压机  "
    assert cell.cleaned_value == "空压机"
    assert cell.action is CleaningAction.CLEANED
    assert cell.rule_code == "RULE_001"


def test_empty_sheet_produces_empty_cleaned_result(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "empty.xlsx",
        {
            "空表": [],
            "Sheet1": [["设备名称"], ["空压机"]],
        },
    )
    profile = profile_workbook(path)
    mapping_config = FieldMappingConfig(
        (StandardFieldDefinition(name="设备名称"),)
    )
    mapping = map_workbook_fields(profile, mapping_config)
    cleaning_config = CleaningWorkbookConfig(
        (FieldCleaningRule(standard_field="设备名称", kind=CleaningKind.TEXT),)
    )

    result = clean_workbook(profile, mapping, cleaning_config)

    assert result.sheet_results[0].rows == ()
    assert result.sheet_results[1].rows[0].values == ("空压机",)


def test_json_cleaning_config_loads(tmp_path: Path) -> None:
    config_path = tmp_path / "cleaning.json"
    config_path.write_text(
        json.dumps(
            {
                "rules": [
                    {
                        "standard_field": "设备名称",
                        "kind": "text",
                        "rule_code": "RULE_TEXT",
                    },
                    {
                        "standard_field": "投用日期",
                        "kind": "date",
                        "date_output_format": "%Y-%m-%d",
                    },
                ]
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    config = load_cleaning_config(config_path)

    assert len(config.rules) == 2
    assert config.rules[0].standard_field == "设备名称"
    assert config.rules[1].kind is CleaningKind.DATE


@pytest.mark.parametrize(
    ("payload", "expected_text"),
    [
        ("{bad json", "无法解析"),
        ("[]", "顶层必须是 JSON 对象"),
        ('{"rules": [{"standard_field": "", "kind": "text"}]}', "standard_field"),
        ('{"rules": [{"standard_field": "设备名称", "kind": "bad"}]}', "text/date/number"),
        (
            '{"rules": [{"standard_field": "投用日期", "kind": "date"}]}',
            "date_output_format",
        ),
        (
            '{"rules": [{"standard_field": "设备名称", "kind": "text", '
            '"remove_characters": "bad"}]}',
            "字符串列表",
        ),
    ],
)
def test_invalid_json_cleaning_config_raises(
    tmp_path: Path,
    payload: str,
    expected_text: str,
) -> None:
    config_path = tmp_path / "bad-cleaning.json"
    config_path.write_text(payload, encoding="utf-8")
    if not expected_text:
        with pytest.raises(CleaningConfigError):
            load_cleaning_config(config_path)
        return
    with pytest.raises(CleaningConfigError) as exc_info:
        load_cleaning_config(config_path)
    assert expected_text in str(exc_info.value)


def test_duplicate_cleaning_rule_raises() -> None:
    config = CleaningWorkbookConfig(
        (
            FieldCleaningRule(standard_field="设备名称", kind=CleaningKind.TEXT),
            FieldCleaningRule(standard_field="设备名称", kind=CleaningKind.TEXT),
        )
    )

    with pytest.raises(CleaningConfigError) as exc_info:
        clean_workbook(
            _empty_profile(),
            mapping=None,  # type: ignore[arg-type]
            config=config,
        )

    assert "重复" in str(exc_info.value)


def test_date_rule_without_output_format_raises() -> None:
    config = CleaningWorkbookConfig(
        (
            FieldCleaningRule(
                standard_field="投用日期",
                kind=CleaningKind.DATE,
            ),
        )
    )

    with pytest.raises(CleaningConfigError) as exc_info:
        clean_workbook(
            _empty_profile(),
            mapping=None,  # type: ignore[arg-type]
            config=config,
        )

    assert "date_output_format" in str(exc_info.value)


def test_cleaning_does_not_modify_input_data(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "readonly.xlsx",
        {"Sheet1": [["设备名称"], ["  空压机  "]]},
    )
    profile = profile_workbook(path)
    mapping = map_workbook_fields(
        profile,
        FieldMappingConfig((StandardFieldDefinition(name="设备名称"),)),
    )
    cleaning_config = CleaningWorkbookConfig(
        (FieldCleaningRule(standard_field="设备名称", kind=CleaningKind.TEXT),)
    )
    raw = load_workbook(path)
    before_rows = raw.sheets[0].rows
    before_bytes = path.read_bytes()

    clean_workbook(profile, mapping, cleaning_config, raw=raw)

    assert raw.sheets[0].rows == before_rows
    assert path.read_bytes() == before_bytes


def test_cleaned_result_summary_counts(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        "summary.xlsx",
        headers=["设备名称"],
        rows=[["  空压机  "], ["冷水机"]],
        standard_fields=(StandardFieldDefinition(name="设备名称"),),
        cleaning_rules=(
            FieldCleaningRule(
                standard_field="设备名称",
                kind=CleaningKind.TEXT,
                rule_code="TEXT_001",
            ),
        ),
    )

    summary = result.summary
    assert summary.row_count == 2
    assert summary.cell_count == 2
    assert summary.cleaned_count == 1
    assert summary.unchanged_count == 1
    assert summary.invalid_count == 0
