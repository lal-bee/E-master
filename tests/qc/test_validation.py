"""统一错误定位（P0-08）测试。"""

from __future__ import annotations

from pathlib import Path

from excel_qc.comparison import compare_standard_mapping
from excel_qc.format_check import check_formats
from excel_qc.inspector import inspect_workbook
from excel_qc.models import (
    CellLocation,
    ComparisonResult,
    FieldFormatRule,
    FieldSelectionRequest,
    FormatValueType,
    MappingCheckConfig,
    StandardDatasetConfig,
    ValidationErrorType,
    ValidationIssue,
    ValidationResult,
    ValidationSource,
)
from excel_qc.selection import select_fields
from excel_qc.standard import import_standard_dataset
from excel_qc.validation import (
    comparison_result_to_validation_result,
    format_check_result_to_validation_result,
    sort_validation_issues,
)
from tests.qc.helpers import build_workbook


def _build_business(path: Path, rows: list[list[str]]) -> Path:
    return build_workbook(
        path,
        {"设备档案": [["名称", "编码"], *rows]},
    )


def _build_standard(path: Path, rows: list[list[str]]) -> Path:
    return build_workbook(
        path,
        {"标准映射": [["原值", "标准值"], *rows]},
    )


def _compare_business(path: Path, standard_path: Path) -> ComparisonResult:
    structure = inspect_workbook(path)
    selection = select_fields(
        structure,
        [FieldSelectionRequest("设备档案", ("名称", "编码"))],
    )
    dataset = import_standard_dataset(
        standard_path,
        StandardDatasetConfig("标准集", "标准映射", "原值", "标准值"),
    )
    return compare_standard_mapping(
        selection,
        MappingCheckConfig("设备档案", "名称", "编码"),
        dataset,
    )


def _make_issue(
    *,
    sheet_name: str,
    row_number: int,
    column_number: int,
    column_letter: str,
) -> ValidationIssue:
    location = CellLocation(
        source_file=Path("sample.xlsx"),
        sheet_name=sheet_name,
        row_number=row_number,
        column_number=column_number,
        column_letter=column_letter,
        column_name=f"字段{column_letter}",
    )
    return ValidationIssue(
        location=location,
        error_type=ValidationErrorType.DATA_ERROR,
        error_code="DATA_002",
        actual_value="BAD",
        expected_value="OK",
        reason="测试原因",
        source=ValidationSource.COMPARISON,
    )


def test_cell_location_fields_and_cell_reference(tmp_path: Path) -> None:
    location = CellLocation(
        source_file=tmp_path / "设备系统导出.xlsx",
        sheet_name="设备档案",
        row_number=125,
        column_number=7,
        column_letter="G",
        column_name="设备类型编码",
    )

    assert location.source_file == tmp_path / "设备系统导出.xlsx"
    assert location.sheet_name == "设备档案"
    assert location.row_number == 125
    assert location.column_number == 7
    assert location.column_letter == "G"
    assert location.column_name == "设备类型编码"
    assert location.cell_reference == "G125"


def test_validation_issue_keeps_complete_information() -> None:
    issue = _make_issue(
        sheet_name="设备档案",
        row_number=125,
        column_number=7,
        column_letter="G",
    )

    assert issue.location.sheet_name == "设备档案"
    assert issue.location.row_number == 125
    assert issue.location.column_number == 7
    assert issue.location.column_letter == "G"
    assert issue.location.cell_reference == "G125"
    assert issue.error_type is ValidationErrorType.DATA_ERROR
    assert issue.error_code == "DATA_002"
    assert issue.actual_value == "BAD"
    assert issue.expected_value == "OK"
    assert issue.reason == "测试原因"
    assert issue.source is ValidationSource.COMPARISON


def test_comparison_result_converts_to_validation_result(tmp_path: Path) -> None:
    business = _build_business(
        tmp_path / "business.xlsx",
        [
            ["除尘器", "D-CCQ"],
            ["未知设备", "D-XXX"],
            ["离心泵", "D-WRONG"],
        ],
    )
    standard = _build_standard(
        tmp_path / "standard.xlsx",
        [["除尘器", "D-CCQ"], ["离心泵", "D-LXB"]],
    )
    comparison = _compare_business(business, standard)

    validation = comparison_result_to_validation_result(comparison)

    assert isinstance(validation, ValidationResult)
    assert validation.file == business
    assert validation.summary.total_checks == 3
    assert validation.summary.pass_count == 1
    assert validation.summary.error_count == 2
    assert validation.summary.data_not_found_count == 1
    assert validation.summary.data_error_count == 1
    assert [issue.error_type for issue in validation.issues] == [
        ValidationErrorType.DATA_NOT_FOUND,
        ValidationErrorType.DATA_ERROR,
    ]
    assert [issue.error_code for issue in validation.issues] == ["DATA_001", "DATA_002"]
    assert all(issue.source is ValidationSource.COMPARISON for issue in validation.issues)
    first = validation.issues[0]
    assert first.location.sheet_name == "设备档案"
    assert first.location.row_number == 3
    assert first.location.column_letter == "B"
    assert first.location.column_number == 2
    assert first.location.column_name == "编码"
    assert first.location.cell_reference == "B3"
    assert first.actual_value == "D-XXX"
    assert first.expected_value == ""
    assert first.reason


def test_comparison_uses_real_physical_row_when_header_not_first_row(
    tmp_path: Path,
) -> None:
    path = build_workbook(
        tmp_path / "deep-header.xlsx",
        {
            "设备档案": [
                ["设备档案台账"],
                ["名称", "编码"],
                ["未知设备", "D-XXX"],
            ]
        },
    )
    standard = _build_standard(
        tmp_path / "standard.xlsx",
        [["除尘器", "D-CCQ"]],
    )
    comparison = _compare_business(path, standard)
    validation = comparison_result_to_validation_result(comparison)

    issue = validation.issues[0]
    assert issue.location.row_number == 3
    assert issue.location.cell_reference == "B3"


def test_format_check_result_converts_to_validation_result(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "format.xlsx",
        {
            "Sheet1": [
                ["编码", "日期"],
                ["D-001", "2026-13-45"],
            ]
        },
    )
    structure = inspect_workbook(path)
    selection = select_fields(
        structure,
        [FieldSelectionRequest("Sheet1", ("编码", "日期"))],
    )
    format_result = check_formats(
        selection,
        [
            FieldFormatRule("Sheet1", "编码", FormatValueType.TEXT),
            FieldFormatRule("Sheet1", "日期", FormatValueType.DATE),
        ],
    )

    validation = format_check_result_to_validation_result(format_result)

    assert validation.file == path
    assert validation.summary.total_checks == 2
    assert validation.summary.pass_count == 1
    assert validation.summary.error_count == 1
    assert validation.summary.format_error_count == 1
    issue = validation.issues[0]
    assert issue.error_type is ValidationErrorType.FORMAT_ERROR
    assert issue.error_code == "FORMAT_001"
    assert issue.source is ValidationSource.FORMAT_CHECK
    assert issue.location.sheet_name == "Sheet1"
    assert issue.location.row_number == 2
    assert issue.location.column_letter == "B"
    assert issue.location.column_number == 2
    assert issue.location.column_name == "日期"
    assert issue.location.cell_reference == "B2"
    assert issue.actual_value == "2026-13-45"
    assert issue.reason


def test_validation_result_contains_issues_from_multiple_sheets(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "multi-sheet.xlsx",
        {
            "设备档案": [["日期"], ["2026-13-45"]],
            "技术参数": [["数量"], ["abc"]],
            "特种设备": [["在用"], ["YES"]],
        },
    )
    structure = inspect_workbook(path)
    selection = select_fields(
        structure,
        [
            FieldSelectionRequest("设备档案", ("日期",)),
            FieldSelectionRequest("技术参数", ("数量",)),
            FieldSelectionRequest("特种设备", ("在用",)),
        ],
    )
    format_result = check_formats(
        selection,
        [
            FieldFormatRule("设备档案", "日期", FormatValueType.DATE),
            FieldFormatRule("技术参数", "数量", FormatValueType.NUMBER),
            FieldFormatRule("特种设备", "在用", FormatValueType.BOOLEAN),
        ],
    )

    validation = format_check_result_to_validation_result(format_result)

    assert [issue.location.sheet_name for issue in validation.issues] == [
        "设备档案",
        "技术参数",
        "特种设备",
    ]
    assert [issue.error_code for issue in validation.issues] == [
        "FORMAT_001",
        "FORMAT_002",
        "FORMAT_006",
    ]
    assert [issue.location.cell_reference for issue in validation.issues] == [
        "A2",
        "A2",
        "A2",
    ]


def test_issues_are_sorted_by_sheet_row_column(tmp_path: Path) -> None:
    issues = [
        _make_issue(sheet_name="特种设备", row_number=5, column_number=9, column_letter="I"),
        _make_issue(sheet_name="设备档案", row_number=12, column_number=6, column_letter="F"),
        _make_issue(sheet_name="设备档案", row_number=12, column_number=3, column_letter="C"),
        _make_issue(sheet_name="技术参数", row_number=4, column_number=2, column_letter="B"),
    ]

    ordered = sort_validation_issues(
        issues,
        sheet_order=["设备档案", "技术参数", "特种设备"],
    )

    assert [(issue.location.sheet_name, issue.location.row_number, issue.location.column_number)
            for issue in ordered] == [
        ("设备档案", 12, 3),
        ("设备档案", 12, 6),
        ("技术参数", 4, 2),
        ("特种设备", 5, 9),
    ]


def test_conversion_does_not_modify_business_or_standard_files(
    tmp_path: Path,
) -> None:
    business = _build_business(
        tmp_path / "business.xlsx",
        [["未知设备", "D-XXX"]],
    )
    standard = _build_standard(
        tmp_path / "standard.xlsx",
        [["除尘器", "D-CCQ"]],
    )
    business_before = business.read_bytes()
    standard_before = standard.read_bytes()

    comparison = _compare_business(business, standard)
    comparison_result_to_validation_result(comparison)

    assert business.read_bytes() == business_before
    assert standard.read_bytes() == standard_before
