"""P0-09 展示层：只消费 ValidationResult，不执行任何核验逻辑。"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from excel_qc.models import ValidationResult, ValidationSummary


@dataclass(frozen=True)
class ErrorDisplayRow:
    """错误明细表中一行的展示数据。"""

    sequence_number: int  # 明细表序号，从 1 开始
    sheet_name: str
    row_number: int
    column_letter: str
    column_number: int
    field_name: str
    cell_reference: str
    error_type: str
    error_code: str
    actual_value: str
    expected_value: str
    reason: str
    source: str


@dataclass(frozen=True)
class ValidationDisplay:
    """ValidationResult 的只读展示快照。"""

    file_path: Path
    summary: ValidationSummary
    errors: tuple[ErrorDisplayRow, ...]
    sheet_options: tuple[str, ...]
    field_options: tuple[str, ...]
    error_type_options: tuple[str, ...]
    error_code_options: tuple[str, ...]

    @property
    def file_name(self) -> str:
        return self.file_path.name

    @property
    def all_passed(self) -> bool:
        return self.summary.error_count == 0


def build_display(result: ValidationResult) -> ValidationDisplay:
    """把统一 ValidationResult 映射为展示层数据模型（保持 issue 顺序）。"""
    errors = tuple(
        ErrorDisplayRow(
            sequence_number=index,
            sheet_name=issue.location.sheet_name,
            row_number=issue.location.row_number,
            column_letter=issue.location.column_letter,
            column_number=issue.location.column_number,
            field_name=issue.location.column_name,
            cell_reference=issue.location.cell_reference,
            error_type=issue.error_type.value,
            error_code=issue.error_code,
            actual_value=issue.actual_value,
            expected_value=issue.expected_value,
            reason=issue.reason,
            source=issue.source.value,
        )
        for index, issue in enumerate(result.issues, start=1)
    )
    return ValidationDisplay(
        file_path=result.file,
        summary=result.summary,
        errors=errors,
        sheet_options=_unique_options(error.sheet_name for error in errors),
        field_options=_unique_options(error.field_name for error in errors),
        error_type_options=_unique_options(error.error_type for error in errors),
        error_code_options=_unique_options(error.error_code for error in errors),
    )


def filter_errors(
    display: ValidationDisplay,
    *,
    error_type: str | None = None,
    sheet: str | None = None,
    field: str | None = None,
    error_code: str | None = None,
) -> tuple[ErrorDisplayRow, ...]:
    """按错误类型/Sheet/字段/错误编码筛选，保持 ValidationResult 既定顺序。"""
    return tuple(
        error
        for error in display.errors
        if (error_type is None or error.error_type == error_type)
        and (sheet is None or error.sheet_name == sheet)
        and (field is None or error.field_name == field)
        and (error_code is None or error.error_code == error_code)
    )


def error_detail(
    display: ValidationDisplay,
    sequence_number: int,
) -> ErrorDisplayRow | None:
    """按明细序号返回错误详情；不存在时返回 None。"""
    for error in display.errors:
        if error.sequence_number == sequence_number:
            return error
    return None


def _unique_options(values) -> tuple[str, ...]:
    return tuple(dict.fromkeys(value for value in values))
