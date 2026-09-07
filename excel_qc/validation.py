"""P0-08 统一错误定位：公共模型与 P0-06/P0-07 结果的适配出口。"""

from __future__ import annotations

from collections.abc import Iterable, Sequence
from pathlib import Path

from excel_qc.models import (
    CellLocation,
    ComparisonResult,
    ComparisonRowResult,
    ComparisonStatus,
    FormatCheckResult,
    FormatIssue,
    ValidationErrorType,
    ValidationIssue,
    ValidationResult,
    ValidationSource,
    ValidationSummary,
)

_COMPARISON_ERROR_CODES = {
    ComparisonStatus.DATA_NOT_FOUND: "DATA_001",
    ComparisonStatus.DATA_ERROR: "DATA_002",
}


def comparison_result_to_validation_result(
    result: ComparisonResult,
    sheet_order: Sequence[str] | None = None,
) -> ValidationResult:
    """把 P0-06 ComparisonResult 转为统一 ValidationResult（PASS 不入 issues）。"""
    issues: list[ValidationIssue] = []
    for row in result.rows:
        if row.status is ComparisonStatus.PASS:
            continue
        issues.append(_comparison_issue(result, row))

    issues = list(sort_validation_issues(issues, sheet_order))
    return ValidationResult(
        file=result.business_file,
        summary=ValidationSummary(
            total_checks=result.summary.total_rows,
            pass_count=result.summary.pass_count,
            error_count=len(issues),
            data_not_found_count=result.summary.data_not_found_count,
            data_error_count=result.summary.data_error_count,
            skipped_empty_row_count=result.skipped_empty_row_count,
        ),
        issues=tuple(issues),
    )


def format_check_result_to_validation_result(
    result: FormatCheckResult,
    sheet_order: Sequence[str] | None = None,
) -> ValidationResult:
    """把 P0-07 FormatCheckResult 转为统一 ValidationResult。"""
    issues = [_format_issue(result.business_file, issue) for issue in result.issues]
    issues = list(sort_validation_issues(issues, sheet_order))
    return ValidationResult(
        file=result.business_file,
        summary=ValidationSummary(
            total_checks=result.summary.total_checks,
            pass_count=result.summary.pass_count,
            error_count=len(issues),
            format_error_count=len(issues),
            skipped_empty_row_count=result.summary.skipped_empty_row_count,
        ),
        issues=tuple(issues),
    )


def sort_validation_issues(
    issues: Iterable[ValidationIssue],
    sheet_order: Sequence[str] | None = None,
) -> tuple[ValidationIssue, ...]:
    """按 Sheet 原始顺序、物理行号、列号稳定排序。"""
    issue_list = list(issues)
    if sheet_order is not None:
        order = {sheet_name: index for index, sheet_name in enumerate(sheet_order)}
    else:
        order = {
            sheet_name: index
            for index, sheet_name in enumerate(
                dict.fromkeys(issue.location.sheet_name for issue in issue_list)
            )
        }

    def sort_key(issue: ValidationIssue) -> tuple[int, int, int]:
        location = issue.location
        return (
            order.get(location.sheet_name, len(order)),
            location.row_number,
            location.column_number,
        )

    return tuple(sorted(issue_list, key=sort_key))


def _comparison_issue(
    result: ComparisonResult,
    row: ComparisonRowResult,
) -> ValidationIssue:
    error_type = ValidationErrorType(row.status.value)
    return ValidationIssue(
        location=CellLocation(
            source_file=result.business_file,
            sheet_name=row.sheet_name,
            row_number=row.row_number,
            column_number=row.standard_column_number,
            column_letter=row.standard_column_letter,
            column_name=row.standard_field_name or f"列{row.standard_column_letter}",
        ),
        error_type=error_type,
        error_code=_COMPARISON_ERROR_CODES[row.status],
        actual_value=row.actual_value,
        expected_value=row.expected_value,
        reason=row.reason,
        source=ValidationSource.COMPARISON,
    )


def _format_issue(
    source_file: Path,
    issue: FormatIssue,
) -> ValidationIssue:
    column_name = issue.field_name or f"列{issue.excel_column_letter}"
    return ValidationIssue(
        location=CellLocation(
            source_file=source_file,
            sheet_name=issue.sheet_name,
            row_number=issue.row_number,
            column_number=issue.excel_column_number,
            column_letter=issue.excel_column_letter,
            column_name=column_name,
        ),
        error_type=ValidationErrorType(issue.error_type),
        error_code=issue.error_code,
        actual_value=issue.actual_value,
        expected_value="",
        reason=issue.reason,
        source=ValidationSource.FORMAT_CHECK,
    )
