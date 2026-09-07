"""P0-07 格式错误检查：对 P0-04 选中的字段执行通用格式规则。"""

from __future__ import annotations

import math
import re
from collections.abc import Iterable
from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal
from typing import Any

from excel_qc.coordinates import cell_address, column_letter
from excel_qc.errors import FormatConfigError
from excel_qc.loader import LoadedSheet, load_workbook
from excel_qc.models import (
    FieldFormatRule,
    FieldSelection,
    FormatCheckResult,
    FormatIssue,
    FormatSummary,
    FormatValueType,
    SheetSelection,
    WorkbookSelection,
)
from excel_qc.text import cell_text, is_blank_value

_DATE_FORMATS = ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d")
_DATETIME_FORMATS = _DATE_FORMATS + (
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y/%m/%d %H:%M:%S",
    "%Y/%m/%d %H:%M",
)
_BOOLEAN_TEXT_VALUES = {"true", "false"}


@dataclass(frozen=True)
class _ResolvedRule:
    sheet: SheetSelection
    field: FieldSelection
    value_type: FormatValueType
    required: bool
    min_length: int | None
    max_length: int | None
    date_format: str | None


def check_formats(
    selection: WorkbookSelection,
    rules: Iterable[FieldFormatRule],
) -> FormatCheckResult:
    """只针对 P0-04 已选 Sheet/字段执行格式规则。"""
    rule_list = list(rules)
    if not rule_list:
        raise FormatConfigError("未提供任何格式规则")

    resolved_rules = _resolve_rules(selection, rule_list)
    raw = load_workbook(selection.path)
    issues: list[FormatIssue] = []
    total_checks = 0
    pass_count = 0
    skipped_empty_count = 0

    for sheet_name in _ordered_sheet_names(raw.sheets, resolved_rules):
        sheet_rules = [rule for rule in resolved_rules if rule.sheet.sheet_name == sheet_name]
        loaded_sheet = _find_loaded_sheet(raw.sheets, sheet_name)
        rows = [list(row) for row in loaded_sheet.rows]
        structure = sheet_rules[0].sheet.structure
        if structure.header_row is None or structure.used_last_row is None:
            continue

        for row_number in range(structure.header_row + 1, structure.used_last_row + 1):
            row_index = row_number - 1
            row = rows[row_index] if row_index < len(rows) else []
            if _row_is_empty(row, structure):
                skipped_empty_count += 1
                continue
            for rule in sheet_rules:
                total_checks += 1
                issue = _check_cell(row, row_number, rule)
                if issue is not None:
                    issues.append(issue)
                else:
                    pass_count += 1

    return FormatCheckResult(
        business_file=selection.path,
        issues=tuple(issues),
        summary=FormatSummary(
            total_checks=total_checks,
            pass_count=pass_count,
            error_count=len(issues),
            skipped_empty_row_count=skipped_empty_count,
        ),
    )


def _resolve_rules(
    selection: WorkbookSelection,
    rules: list[FieldFormatRule],
) -> list[_ResolvedRule]:
    resolved: list[_ResolvedRule] = []
    seen_keys: set[tuple[str, int]] = set()
    for rule in rules:
        _validate_rule(rule)
        sheet = _find_selected_sheet(selection, rule.sheet_name)
        field = _resolve_selected_field(sheet, rule.field)
        key = (sheet.sheet_name, field.column_index)
        if key in seen_keys:
            raise FormatConfigError(
                f"Sheet“{sheet.sheet_name}”字段“{field.display_name}”"
                "存在重复格式规则"
            )
        seen_keys.add(key)
        resolved.append(
            _ResolvedRule(
                sheet=sheet,
                field=field,
                value_type=rule.value_type,
                required=rule.required,
                min_length=rule.min_length,
                max_length=rule.max_length,
                date_format=rule.date_format,
            )
        )
    return resolved


def _validate_rule(rule: FieldFormatRule) -> None:
    if not rule.sheet_name.strip():
        raise FormatConfigError("格式规则缺少 Sheet 名称")
    if not rule.field.strip():
        raise FormatConfigError("格式规则缺少字段")
    if rule.min_length is not None and rule.min_length < 0:
        raise FormatConfigError("min_length 不能为负数")
    if rule.max_length is not None and rule.max_length < 0:
        raise FormatConfigError("max_length 不能为负数")
    if (
        rule.min_length is not None
        and rule.max_length is not None
        and rule.min_length > rule.max_length
    ):
        raise FormatConfigError("min_length 不能大于 max_length")


def _find_selected_sheet(
    selection: WorkbookSelection,
    sheet_name: str,
) -> SheetSelection:
    for sheet in selection.selections:
        if sheet.sheet_name == sheet_name:
            return sheet
    available = "、".join(selection.sheet_names) or "（无）"
    raise FormatConfigError(
        f"Sheet“{sheet_name}”不在当前 WorkbookSelection 中，已选 Sheet：{available}"
    )


def _resolve_selected_field(sheet: SheetSelection, identifier: str) -> FieldSelection:
    key = identifier.strip()
    by_name = [
        field
        for field in sheet.selected_fields
        if field.name == key or field.display_name == key
    ]
    if len(by_name) > 1:
        letters = "、".join(field.excel_column for field in by_name)
        raise FormatConfigError(
            f"Sheet“{sheet.sheet_name}”字段“{key}”对应多列（{letters}），"
            "请改用 Excel 列字母指定"
        )
    field = by_name[0] if by_name else _match_field_by_letter(sheet, key)
    if field is None:
        options = _format_selected_fields(sheet)
        raise FormatConfigError(
            f"Sheet“{sheet.sheet_name}”中不存在已选择的字段“{key}”。"
            f"该 Sheet 已选字段（Excel 列：字段名）：{options}"
        )
    return field


def _match_field_by_letter(sheet: SheetSelection, key: str) -> FieldSelection | None:
    column_index = _parse_column_letter(key)
    if column_index is None:
        return None
    for field in sheet.selected_fields:
        if field.column_index == column_index:
            return field
    return None


def _parse_column_letter(identifier: str) -> int | None:
    letters = identifier.upper()
    if not letters or not letters.isascii() or not letters.isalpha():
        return None
    index = 0
    for char in letters:
        index = index * 26 + (ord(char) - ord("A") + 1)
    column_index = index - 1
    return column_index if column_letter(column_index) == letters else None


def _ordered_sheet_names(
    sheets: tuple[LoadedSheet, ...],
    rules: list[_ResolvedRule],
) -> list[str]:
    """按业务 Excel 原始 Sheet 顺序返回有规则参与的 Sheet。"""
    rule_sheet_names = {rule.sheet.sheet_name for rule in rules}
    return [
        sheet.sheet_name
        for sheet in sheets
        if sheet.sheet_name in rule_sheet_names
    ]


def _find_loaded_sheet(sheets: tuple[LoadedSheet, ...], sheet_name: str) -> LoadedSheet:
    for sheet in sheets:
        if sheet.sheet_name == sheet_name:
            return sheet
    raise FormatConfigError(f"无法在业务文件中找到 Sheet“{sheet_name}”")


def _check_cell(row: list[Any], row_number: int, rule: _ResolvedRule) -> FormatIssue | None:
    raw_value = _cell(row, rule.field.column_index)
    actual_text = cell_text(raw_value)
    actual_type = _actual_data_type(raw_value)
    expected_format = rule.value_type.label

    if not actual_text:
        if rule.required:
            return _issue(
                row_number=row_number,
                rule=rule,
                actual_value=actual_text,
                actual_data_type=actual_type,
                expected_format=expected_format,
                error_code="FORMAT_005",
                reason=f"必填字段为空（{rule.sheet.sheet_name} 第 {row_number} 行）",
            )
        return None

    if rule.value_type in (FormatValueType.DATE, FormatValueType.DATETIME):
        if not _is_date_like(raw_value, rule.value_type, rule.date_format):
            return _issue(
                row_number, rule, actual_text, actual_type, expected_format,
                "FORMAT_001",
                f"日期格式错误：实际值“{actual_text}”不是合法的"
                f"{expected_format}格式",
            )
        return None

    if rule.value_type is FormatValueType.INTEGER:
        if not _is_integer_like(raw_value):
            return _issue(
                row_number, rule, actual_text, actual_type, expected_format,
                "FORMAT_003",
                f"整数格式错误：期望整数，实际值“{actual_text}”",
            )
        return None

    if rule.value_type is FormatValueType.NUMBER:
        if not _is_number_like(raw_value):
            return _issue(
                row_number, rule, actual_text, actual_type, expected_format,
                "FORMAT_002",
                f"数字格式错误：期望数值，实际值“{actual_text}”",
            )
        return None

    if rule.value_type is FormatValueType.BOOLEAN:
        if not _is_boolean_like(raw_value):
            return _issue(
                row_number, rule, actual_text, actual_type, expected_format,
                "FORMAT_006",
                f"布尔值格式错误：期望 TRUE/FALSE，实际值“{actual_text}”",
            )
        return None

    # TEXT：长度检查
    length = len(actual_text)
    if rule.max_length is not None and length > rule.max_length:
        return _issue(
            row_number, rule, actual_text, actual_type, expected_format,
            "FORMAT_004",
            f"文本长度超过最大允许长度 {rule.max_length}，实际长度 {length}",
        )
    if rule.min_length is not None and length < rule.min_length:
        return _issue(
            row_number, rule, actual_text, actual_type, expected_format,
            "FORMAT_004",
            f"文本长度小于最小允许长度 {rule.min_length}，实际长度 {length}",
        )
    return None


def _issue(
    row_number: int,
    rule: _ResolvedRule,
    actual_value: str,
    actual_data_type: str,
    expected_format: str,
    error_code: str,
    reason: str,
) -> FormatIssue:
    return FormatIssue(
        sheet_name=rule.sheet.sheet_name,
        row_number=row_number,
        field_name=rule.field.name,
        excel_column_letter=rule.field.excel_column,
        excel_column_number=rule.field.excel_column_number,
        cell_address=cell_address(row_number, rule.field.column_index),
        actual_value=actual_value,
        actual_data_type=actual_data_type,
        expected_format=expected_format,
        error_code=error_code,
        reason=reason,
    )


def _is_date_like(
    value: Any,
    value_type: FormatValueType,
    date_format: str | None,
) -> bool:
    if isinstance(value, datetime):
        if value_type is FormatValueType.DATE:
            return _is_midnight(value)
        return True
    if isinstance(value, date):
        return True
    if not isinstance(value, str):
        return False
    text = value.strip()
    if not text:
        return False
    if date_format:
        return _parse_with_formats(text, (date_format,)) is not None
    formats = _DATETIME_FORMATS if value_type is FormatValueType.DATETIME else _DATE_FORMATS
    return _parse_with_formats(text, formats) is not None


def _parse_with_formats(text: str, formats: tuple[str, ...]) -> datetime | None:
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _is_midnight(value: datetime) -> bool:
    return (
        value.hour == 0
        and value.minute == 0
        and value.second == 0
        and value.microsecond == 0
    )


def _is_integer_like(value: Any) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, int):
        return True
    if isinstance(value, float):
        return math.isfinite(value) and value.is_integer()
    if isinstance(value, Decimal):
        return value == value.to_integral_value()
    if isinstance(value, str):
        return re.fullmatch(r"[+-]?\d+", value.strip()) is not None
    return False


def _is_number_like(value: Any) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, Decimal)):
        return True
    if isinstance(value, float):
        return math.isfinite(value)
    if isinstance(value, str):
        try:
            number = float(value.strip())
        except ValueError:
            return False
        return math.isfinite(number)
    return False


def _is_boolean_like(value: Any) -> bool:
    if isinstance(value, bool):
        return True
    return isinstance(value, str) and value.strip().lower() in _BOOLEAN_TEXT_VALUES


def _actual_data_type(value: Any) -> str:
    if value is None:
        return "空"
    if isinstance(value, bool):
        return "布尔"
    if isinstance(value, datetime):
        return "日期时间"
    if isinstance(value, date):
        return "日期"
    if isinstance(value, time):
        return "时间"
    if isinstance(value, int):
        return "整数"
    if isinstance(value, (float, Decimal)):
        return "数值"
    if isinstance(value, str):
        return "文本"
    return "其他"


def _row_is_empty(row: list[Any], structure) -> bool:
    if structure.used_first_col is None or structure.used_last_col is None:
        return True
    for col_index in range(structure.used_first_col - 1, structure.used_last_col):
        if not is_blank_value(_cell(row, col_index)):
            return False
    return True


def _cell(row: list[Any], column_index: int) -> Any:
    return row[column_index] if column_index < len(row) else None


def _format_selected_fields(sheet: SheetSelection) -> str:
    options = [
        f"{field.excel_column}：{field.display_name}" for field in sheet.selected_fields
    ]
    return "、".join(options) or "（无）"
