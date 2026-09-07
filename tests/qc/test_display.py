"""核验结果展示（P0-09）测试。"""

from __future__ import annotations

from pathlib import Path

from excel_qc.models import (
    CellLocation,
    ValidationErrorType,
    ValidationIssue,
    ValidationResult,
    ValidationSource,
    ValidationSummary,
)
from ui.display import build_display, error_detail, filter_errors
from ui.html import render_html


def _issue(
    *,
    sheet_name: str,
    row_number: int,
    column_letter: str,
    column_number: int,
    field_name: str,
    error_type: ValidationErrorType,
    error_code: str,
    actual_value: str,
    expected_value: str = "",
    source: ValidationSource = ValidationSource.COMPARISON,
) -> ValidationIssue:
    return ValidationIssue(
        location=CellLocation(
            source_file=Path("设备系统导出.xlsx"),
            sheet_name=sheet_name,
            row_number=row_number,
            column_number=column_number,
            column_letter=column_letter,
            column_name=field_name,
        ),
        error_type=error_type,
        error_code=error_code,
        actual_value=actual_value,
        expected_value=expected_value,
        reason=f"{field_name}校验失败",
        source=source,
    )


def _demo_result() -> ValidationResult:
    return ValidationResult(
        file=Path("设备系统导出.xlsx"),
        summary=ValidationSummary(
            total_checks=10,
            pass_count=6,
            error_count=4,
            format_error_count=2,
            data_not_found_count=1,
            data_error_count=1,
            skipped_empty_row_count=3,
        ),
        issues=(
            _issue(
                sheet_name="设备档案",
                row_number=125,
                column_letter="G",
                column_number=7,
                field_name="设备类型编码",
                error_type=ValidationErrorType.DATA_ERROR,
                error_code="DATA_002",
                actual_value="D-CQJ",
                expected_value="D-CCQ",
            ),
            _issue(
                sheet_name="设备档案",
                row_number=128,
                column_letter="C",
                column_number=3,
                field_name="投用日期",
                error_type=ValidationErrorType.FORMAT_ERROR,
                error_code="FORMAT_001",
                actual_value="2026-13-45",
                source=ValidationSource.FORMAT_CHECK,
            ),
            _issue(
                sheet_name="技术参数",
                row_number=86,
                column_letter="F",
                column_number=6,
                field_name="数量",
                error_type=ValidationErrorType.DATA_NOT_FOUND,
                error_code="DATA_001",
                actual_value="D-XXX",
            ),
            _issue(
                sheet_name="技术参数",
                row_number=88,
                column_letter="F",
                column_number=6,
                field_name="数量",
                error_type=ValidationErrorType.FORMAT_ERROR,
                error_code="FORMAT_003",
                actual_value="1.5",
                source=ValidationSource.FORMAT_CHECK,
            ),
        ),
    )


def test_display_builds_file_and_summary() -> None:
    display = build_display(_demo_result())

    assert display.file_path == Path("设备系统导出.xlsx")
    assert display.file_name == "设备系统导出.xlsx"
    assert display.summary.total_checks == 10
    assert display.summary.pass_count == 6
    assert display.summary.error_count == 4
    assert display.summary.format_error_count == 2
    assert display.summary.data_not_found_count == 1
    assert display.summary.data_error_count == 1
    assert display.summary.skipped_empty_row_count == 3


def test_error_detail_rows_contain_all_table_fields() -> None:
    display = build_display(_demo_result())

    error = display.errors[0]
    assert error.sequence_number == 1
    assert error.sheet_name == "设备档案"
    assert error.row_number == 125
    assert error.column_letter == "G"
    assert error.column_number == 7
    assert error.field_name == "设备类型编码"
    assert error.cell_reference == "G125"
    assert error.error_type == "DATA_ERROR"
    assert error.error_code == "DATA_002"
    assert error.actual_value == "D-CQJ"
    assert error.expected_value == "D-CCQ"
    assert error.reason == "设备类型编码校验失败"
    assert error.source == "COMPARISON"


def test_filter_by_error_type() -> None:
    display = build_display(_demo_result())

    filtered = filter_errors(display, error_type="FORMAT_ERROR")

    assert [error.error_code for error in filtered] == ["FORMAT_001", "FORMAT_003"]


def test_filter_by_sheet() -> None:
    display = build_display(_demo_result())

    filtered = filter_errors(display, sheet="技术参数")

    assert len(filtered) == 2
    assert all(error.sheet_name == "技术参数" for error in filtered)


def test_filter_by_field() -> None:
    display = build_display(_demo_result())

    filtered = filter_errors(display, field="数量")

    assert len(filtered) == 2
    assert all(error.field_name == "数量" for error in filtered)


def test_filter_by_error_code() -> None:
    display = build_display(_demo_result())

    filtered = filter_errors(display, error_code="DATA_001")

    assert len(filtered) == 1
    assert filtered[0].row_number == 86


def test_filter_does_not_modify_underlying_result() -> None:
    result = _demo_result()
    display = build_display(result)
    original_issues = result.issues

    filtered = filter_errors(display, error_type="DATA_ERROR")

    assert result.issues == original_issues
    assert display.errors is not filtered
    assert len(display.errors) == 4
    assert len(filtered) == 1


def test_display_preserves_validation_result_order() -> None:
    result = _demo_result()
    display = build_display(result)

    assert [(error.sheet_name, error.row_number, error.column_number)
            for error in display.errors] == [
        ("设备档案", 125, 7),
        ("设备档案", 128, 3),
        ("技术参数", 86, 6),
        ("技术参数", 88, 6),
    ]


def test_error_detail_lookup_by_sequence() -> None:
    display = build_display(_demo_result())

    detail = error_detail(display, 3)
    assert detail is not None
    assert detail.sheet_name == "技术参数"
    assert detail.cell_reference == "F86"
    assert error_detail(display, 999) is None


def test_all_passed_when_no_errors() -> None:
    result = ValidationResult(
        file=Path("全部通过.xlsx"),
        summary=ValidationSummary(
            total_checks=10,
            pass_count=10,
            error_count=0,
        ),
        issues=(),
    )

    display = build_display(result)

    assert display.all_passed is True
    assert display.errors == ()
    assert display.sheet_options == ()


def test_multi_sheet_display_and_options() -> None:
    display = build_display(_demo_result())

    assert display.sheet_options == ("设备档案", "技术参数")
    assert display.field_options == ("设备类型编码", "投用日期", "数量")
    assert display.error_type_options == (
        "DATA_ERROR",
        "FORMAT_ERROR",
        "DATA_NOT_FOUND",
    )
    assert display.error_code_options == (
        "DATA_002",
        "FORMAT_001",
        "DATA_001",
        "FORMAT_003",
    )


def test_html_render_contains_summary_and_table() -> None:
    display = build_display(_demo_result())

    html_page = render_html(display)

    assert "核验结果" in html_page
    assert "设备系统导出.xlsx" in html_page
    assert "总数" in html_page
    assert "错误明细" in html_page
    assert "DATA_002" in html_page
    assert "错误原因" in html_page
    assert "来源" in html_page


def test_html_render_shows_all_passed_when_no_errors() -> None:
    result = ValidationResult(
        file=Path("全部通过.xlsx"),
        summary=ValidationSummary(
            total_checks=5,
            pass_count=5,
            error_count=0,
        ),
        issues=(),
    )

    html_page = render_html(build_display(result))

    assert "全部通过，无错误" in html_page
