"""格式错误检查（P0-07）测试。"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pytest

from excel_qc.errors import FormatConfigError
from excel_qc.format_check import check_formats
from excel_qc.inspector import inspect_workbook
from excel_qc.models import (
    FieldFormatRule,
    FieldSelectionRequest,
    FormatValueType,
)
from excel_qc.selection import select_fields
from tests.qc.helpers import build_workbook


def _build(
    path: Path,
    headers: list[str],
    rows: list[list[object]],
) -> Path:
    return build_workbook(path, {"Sheet1": [headers, *rows]})


def _select(path: Path, fields: tuple[str, ...]) -> object:
    structure = inspect_workbook(path)
    return select_fields(
        structure,
        [FieldSelectionRequest("Sheet1", fields)],
    )


def _check(
    path: Path,
    fields: tuple[str, ...],
    rules: list[FieldFormatRule],
):
    return check_formats(_select(path, fields), rules)


def test_valid_date_passes(tmp_path: Path) -> None:
    path = _build(
        tmp_path / "date-ok.xlsx",
        ["投用日期"],
        [[datetime(2026, 8, 1)]],
    )

    result = _check(
        path,
        ("投用日期",),
        [FieldFormatRule("Sheet1", "投用日期", FormatValueType.DATE)],
    )

    assert result.summary.total_checks == 1
    assert result.summary.pass_count == 1
    assert result.summary.error_count == 0
    assert result.issues == ()


def test_invalid_text_date_returns_format_error(tmp_path: Path) -> None:
    path = _build(
        tmp_path / "date-bad.xlsx",
        ["投用日期"],
        [["2026-13-45"]],
    )

    result = _check(
        path,
        ("投用日期",),
        [FieldFormatRule("Sheet1", "投用日期", FormatValueType.DATE)],
    )

    issue = result.issues[0]
    assert issue.error_code == "FORMAT_001"
    assert issue.error_type == "FORMAT_ERROR"
    assert issue.sheet_name == "Sheet1"
    assert issue.row_number == 2
    assert issue.field_name == "投用日期"
    assert issue.excel_column_letter == "A"
    assert issue.excel_column_number == 1
    assert issue.cell_address == "A2"
    assert issue.actual_value == "2026-13-45"
    assert issue.actual_data_type == "文本"
    assert issue.expected_format == "日期"
    assert "日期格式错误" in issue.reason


def test_valid_integer_passes(tmp_path: Path) -> None:
    path = _build(tmp_path / "int-ok.xlsx", ["数量"], [[100]])

    result = _check(
        path,
        ("数量",),
        [FieldFormatRule("Sheet1", "数量", FormatValueType.INTEGER)],
    )

    assert result.summary.pass_count == 1
    assert result.summary.error_count == 0


def test_integer_text_100_is_accepted_per_documented_rule(tmp_path: Path) -> None:
    path = _build(tmp_path / "int-text.xlsx", ["数量"], [["100"]])

    result = _check(
        path,
        ("数量",),
        [FieldFormatRule("Sheet1", "数量", FormatValueType.INTEGER)],
    )

    assert result.summary.pass_count == 1
    assert result.summary.error_count == 0


def test_float_in_integer_field_returns_format_error(tmp_path: Path) -> None:
    path = _build(tmp_path / "int-float.xlsx", ["数量"], [[100.5]])

    result = _check(
        path,
        ("数量",),
        [FieldFormatRule("Sheet1", "数量", FormatValueType.INTEGER)],
    )

    issue = result.issues[0]
    assert issue.error_code == "FORMAT_003"
    assert issue.actual_value == "100.5"
    assert "整数格式错误" in issue.reason


def test_valid_number_passes(tmp_path: Path) -> None:
    path = _build(tmp_path / "num-ok.xlsx", ["单价"], [[12.5]])

    result = _check(
        path,
        ("单价",),
        [FieldFormatRule("Sheet1", "单价", FormatValueType.NUMBER)],
    )

    assert result.summary.pass_count == 1
    assert result.summary.error_count == 0


def test_invalid_number_returns_format_error(tmp_path: Path) -> None:
    path = _build(tmp_path / "num-bad.xlsx", ["单价"], [["abc"]])

    result = _check(
        path,
        ("单价",),
        [FieldFormatRule("Sheet1", "单价", FormatValueType.NUMBER)],
    )

    assert result.issues[0].error_code == "FORMAT_002"


def test_valid_text_passes(tmp_path: Path) -> None:
    path = _build(tmp_path / "text-ok.xlsx", ["名称"], [["空压机"]])

    result = _check(
        path,
        ("名称",),
        [FieldFormatRule("Sheet1", "名称", FormatValueType.TEXT)],
    )

    assert result.summary.pass_count == 1
    assert result.summary.error_count == 0


def test_text_exceeding_max_length_returns_format_error(tmp_path: Path) -> None:
    long_text = "字" * 25
    path = _build(tmp_path / "text-long.xlsx", ["名称"], [[long_text]])

    result = _check(
        path,
        ("名称",),
        [FieldFormatRule("Sheet1", "名称", FormatValueType.TEXT, max_length=20)],
    )

    issue = result.issues[0]
    assert issue.error_code == "FORMAT_004"
    assert "超过最大允许长度 20" in issue.reason
    assert "实际长度 25" in issue.reason


def test_text_below_min_length_returns_format_error(tmp_path: Path) -> None:
    path = _build(tmp_path / "text-short.xlsx", ["代码"], [["AB"]])

    result = _check(
        path,
        ("代码",),
        [FieldFormatRule("Sheet1", "代码", FormatValueType.TEXT, min_length=5)],
    )

    assert result.issues[0].error_code == "FORMAT_004"


def test_required_field_with_value_passes(tmp_path: Path) -> None:
    path = _build(tmp_path / "required-ok.xlsx", ["编码"], [["D-CCQ"]])

    result = _check(
        path,
        ("编码",),
        [FieldFormatRule("Sheet1", "编码", FormatValueType.TEXT, required=True)],
    )

    assert result.summary.pass_count == 1
    assert result.summary.error_count == 0


def test_required_field_empty_returns_format_error(tmp_path: Path) -> None:
    path = _build(
        tmp_path / "required-bad.xlsx",
        ["编码", "名称"],
        [["", "空压机"]],
    )

    result = _check(
        path,
        ("编码", "名称"),
        [
            FieldFormatRule("Sheet1", "编码", FormatValueType.TEXT, required=True),
            FieldFormatRule("Sheet1", "名称", FormatValueType.TEXT),
        ],
    )

    issue = next(item for item in result.issues if item.error_code == "FORMAT_005")
    assert issue.field_name == "编码"
    assert "必填字段为空" in issue.reason


def test_optional_empty_field_does_not_raise_error(tmp_path: Path) -> None:
    path = _build(
        tmp_path / "optional-empty.xlsx",
        ["名称", "备注"],
        [["空压机", ""]],
    )

    result = _check(
        path,
        ("名称", "备注"),
        [
            FieldFormatRule("Sheet1", "名称", FormatValueType.TEXT),
            FieldFormatRule("Sheet1", "备注", FormatValueType.TEXT),
        ],
    )

    assert result.summary.pass_count == 2
    assert result.summary.error_count == 0


def test_valid_boolean_passes(tmp_path: Path) -> None:
    path = _build(tmp_path / "bool-ok.xlsx", ["在用"], [[True]])

    result = _check(
        path,
        ("在用",),
        [FieldFormatRule("Sheet1", "在用", FormatValueType.BOOLEAN)],
    )

    assert result.summary.pass_count == 1
    assert result.summary.error_count == 0


def test_invalid_boolean_returns_format_error(tmp_path: Path) -> None:
    path = _build(tmp_path / "bool-bad.xlsx", ["在用"], [["YES"]])

    result = _check(
        path,
        ("在用",),
        [FieldFormatRule("Sheet1", "在用", FormatValueType.BOOLEAN)],
    )

    issue = result.issues[0]
    assert issue.error_code == "FORMAT_006"
    assert "布尔值格式错误" in issue.reason


def test_fully_empty_row_is_skipped(tmp_path: Path) -> None:
    path = _build(
        tmp_path / "blank-row.xlsx",
        ["编码", "名称"],
        [
            ["D-001", "甲"],
            ["", ""],
            ["D-002", "乙"],
        ],
    )

    result = _check(
        path,
        ("编码", "名称"),
        [
            FieldFormatRule("Sheet1", "编码", FormatValueType.TEXT, required=True),
            FieldFormatRule("Sheet1", "名称", FormatValueType.TEXT, required=True),
        ],
    )

    assert result.summary.skipped_empty_row_count == 1
    assert result.summary.error_count == 0
    assert result.summary.total_checks == 4
    assert result.summary.pass_count == 4


def test_unselected_field_is_not_checked(tmp_path: Path) -> None:
    path = _build(
        tmp_path / "unselected.xlsx",
        ["编码", "名称", "投用日期"],
        [["D-001", "甲", "2026-13-45"]],
    )

    result = _check(
        path,
        ("编码", "名称"),
        [
            FieldFormatRule("Sheet1", "编码", FormatValueType.TEXT),
            FieldFormatRule("Sheet1", "名称", FormatValueType.TEXT),
        ],
    )

    assert result.summary.total_checks == 2
    assert result.summary.error_count == 0


def test_error_locations_are_recorded_completely(tmp_path: Path) -> None:
    path = _build(
        tmp_path / "locations.xlsx",
        ["字段甲", "字段乙", "字段丙", "字段丁", "字段戊", "字段己", "投用日期"],
        [["a", "b", "c", "d", "e", "f", "2026-13-45"]],
    )

    result = _check(
        path,
        ("投用日期",),
        [FieldFormatRule("Sheet1", "投用日期", FormatValueType.DATE)],
    )

    issue = result.issues[0]
    assert issue.sheet_name == "Sheet1"
    assert issue.row_number == 2
    assert issue.field_name == "投用日期"
    assert issue.excel_column_letter == "G"
    assert issue.excel_column_number == 7
    assert issue.cell_address == "G2"
    assert issue.actual_value == "2026-13-45"
    assert issue.actual_data_type == "文本"
    assert issue.expected_format == "日期"
    assert issue.error_code == "FORMAT_001"
    assert issue.error_type == "FORMAT_ERROR"
    assert issue.reason


def test_rule_for_unselected_field_raises_config_error(tmp_path: Path) -> None:
    path = _build(
        tmp_path / "not-selected.xlsx",
        ["编码", "名称"],
        [["D-001", "甲"]],
    )

    with pytest.raises(FormatConfigError):
        _check(
            path,
            ("编码",),
            [FieldFormatRule("Sheet1", "名称", FormatValueType.TEXT)],
        )


def test_format_check_does_not_modify_source_file(tmp_path: Path) -> None:
    path = _build(
        tmp_path / "readonly.xlsx",
        ["编码"],
        [["D-001"]],
    )
    before = path.read_bytes()

    _check(
        path,
        ("编码",),
        [FieldFormatRule("Sheet1", "编码", FormatValueType.TEXT)],
    )

    assert path.read_bytes() == before
