"""P0-06 标准映射比对：业务 Excel 选中字段 → 标准数据集 → 逐行比对。"""

from __future__ import annotations

from typing import Any

from excel_qc.coordinates import column_letter
from excel_qc.errors import (
    ComparisonBlankValueError,
    ComparisonConfigError,
    StandardMappingConflictError,
    StandardNoDataError,
)
from excel_qc.loader import LoadedSheet, load_workbook
from excel_qc.models import (
    ComparisonResult,
    ComparisonRowResult,
    ComparisonStatus,
    ComparisonSummary,
    FieldSelection,
    MappingCheckConfig,
    SheetSelection,
    StandardDataset,
    WorkbookSelection,
)
from excel_qc.text import cell_text, is_blank_value


def compare_standard_mapping(
    selection: WorkbookSelection,
    config: MappingCheckConfig,
    dataset: StandardDataset,
) -> ComparisonResult:
    """对业务 Excel 中一个 Sheet 的两列执行标准映射比对。

    - 业务字段必须已包含在 P0-04 的 SheetSelection 中；
    - 标准数据必须来自 P0-05 的 StandardDataset；
    - 本函数不修改业务 Excel，也不修改标准 Excel。
    """
    _validate_config(config)
    if not selection.selections:
        raise ComparisonConfigError("WorkbookSelection 中未选择任何 Sheet")

    business_sheet = _find_selected_sheet(selection, config.sheet_name)
    match_field = _resolve_selected_field(
        business_sheet, config.match_field, "匹配字段"
    )
    standard_field = _resolve_selected_field(
        business_sheet, config.standard_field, "标准值字段"
    )
    if match_field.column_index == standard_field.column_index:
        raise ComparisonConfigError(
            f"Sheet“{business_sheet.sheet_name}”中匹配字段与标准值字段不能是同一列"
        )

    lookup = _build_lookup(dataset)
    loaded_sheet = _load_business_sheet(selection, business_sheet.sheet_name)
    row_results, skipped_empty_count = _compare_rows(
        selection=selection,
        sheet=business_sheet,
        loaded_sheet=loaded_sheet,
        match_field=match_field,
        standard_field=standard_field,
        lookup=lookup,
        dataset_name=dataset.dataset_name,
    )

    summary = ComparisonSummary(
        total_rows=len(row_results),
        pass_count=sum(row.status is ComparisonStatus.PASS for row in row_results),
        data_not_found_count=sum(
            row.status is ComparisonStatus.DATA_NOT_FOUND for row in row_results
        ),
        data_error_count=sum(
            row.status is ComparisonStatus.DATA_ERROR for row in row_results
        ),
    )
    return ComparisonResult(
        business_file=selection.path,
        sheet_name=business_sheet.sheet_name,
        dataset_name=dataset.dataset_name,
        match_column=match_field,
        standard_column=standard_field,
        rows=tuple(row_results),
        summary=summary,
        skipped_empty_row_count=skipped_empty_count,
    )


def _validate_config(config: MappingCheckConfig) -> None:
    if not config.sheet_name.strip():
        raise ComparisonConfigError("业务 Sheet 名称不能为空")
    if not config.match_field.strip():
        raise ComparisonConfigError("业务匹配字段不能为空")
    if not config.standard_field.strip():
        raise ComparisonConfigError("业务标准值字段不能为空")


def _find_selected_sheet(
    selection: WorkbookSelection,
    sheet_name: str,
) -> SheetSelection:
    for sheet in selection.selections:
        if sheet.sheet_name == sheet_name:
            return sheet
    available = "、".join(selection.sheet_names) or "（无）"
    raise ComparisonConfigError(
        f"Sheet“{sheet_name}”不在当前 WorkbookSelection 中，"
        f"已选 Sheet：{available}"
    )


def _resolve_selected_field(
    sheet: SheetSelection,
    identifier: str,
    role: str,
) -> FieldSelection:
    key = identifier.strip()
    by_name = [
        field
        for field in sheet.selected_fields
        if field.name == key or field.display_name == key
    ]
    if len(by_name) > 1:
        letters = "、".join(field.excel_column for field in by_name)
        raise ComparisonConfigError(
            f"Sheet“{sheet.sheet_name}”中{role}“{key}”对应多列（{letters}），"
            "请改用 Excel 列字母指定"
        )
    field = by_name[0] if by_name else _match_field_by_letter(sheet, key)
    if field is None:
        options = _format_selected_fields(sheet)
        raise ComparisonConfigError(
            f"Sheet“{sheet.sheet_name}”中不存在已选择的{role}“{key}”。"
            f"该 Sheet 已选字段（Excel 列：字段名）：{options}"
        )
    return field


def _match_field_by_letter(
    sheet: SheetSelection,
    key: str,
) -> FieldSelection | None:
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


def _build_lookup(dataset: StandardDataset) -> dict[str, str]:
    """构建唯一“匹配值 → 标准值”映射；理论上遇到冲突必须显式失败。"""
    if not dataset.records:
        raise StandardNoDataError(
            f"标准数据集“{dataset.dataset_name}”中没有有效数据，无法执行比对"
        )
    lookup: dict[str, str] = {}
    for record in dataset.records:
        previous = lookup.get(record.match_value)
        if previous is not None and previous != record.standard_value:
            raise StandardMappingConflictError(
                f"标准数据集“{dataset.dataset_name}”中匹配值"
                f"“{record.match_value}”对应多个不同标准值"
                f"（“{previous}”与“{record.standard_value}”），无法比对"
            )
        lookup[record.match_value] = record.standard_value
    return lookup


def _load_business_sheet(
    selection: WorkbookSelection,
    sheet_name: str,
) -> LoadedSheet:
    raw = load_workbook(selection.path)
    for sheet in raw.sheets:
        if sheet.sheet_name == sheet_name:
            return sheet
    raise ComparisonConfigError(
        f"无法在业务文件 {selection.path} 中找到 Sheet“{sheet_name}”"
    )


def _compare_rows(
    selection: WorkbookSelection,
    sheet: SheetSelection,
    loaded_sheet: LoadedSheet,
    match_field: FieldSelection,
    standard_field: FieldSelection,
    lookup: dict[str, str],
    dataset_name: str,
) -> tuple[list[ComparisonRowResult], int]:
    structure = sheet.structure
    rows = [list(row) for row in loaded_sheet.rows]
    if structure.header_row is None or structure.used_last_row is None:
        raise ComparisonConfigError(
            f"Sheet“{sheet.sheet_name}”没有可用于比对的数据区域"
        )

    results: list[ComparisonRowResult] = []
    skipped_empty_count = 0
    for row_number in range(structure.header_row + 1, structure.used_last_row + 1):
        row_index = row_number - 1
        row = rows[row_index] if row_index < len(rows) else []
        if _row_is_empty(row, structure):
            skipped_empty_count += 1
            continue

        match_text = cell_text(_cell(row, match_field.column_index))
        actual_text = cell_text(_cell(row, standard_field.column_index))
        if not match_text:
            raise ComparisonBlankValueError(
                f"业务 Sheet“{sheet.sheet_name}”第 {row_number} 行"
                f"（Excel {match_field.excel_column} 列：{match_field.display_name}）"
                "匹配字段为空，无法执行标准映射比对"
            )
        if not actual_text:
            raise ComparisonBlankValueError(
                f"业务 Sheet“{sheet.sheet_name}”第 {row_number} 行"
                f"（Excel {standard_field.excel_column} 列："
                f"{standard_field.display_name}）标准值字段为空，无法执行比对"
            )

        expected_value = lookup.get(match_text)
        if expected_value is None:
            status = ComparisonStatus.DATA_NOT_FOUND
            expected_text = ""
            reason = (
                f"匹配值“{match_text}”不存在于当前标准数据集"
                f"“{dataset_name}”"
            )
        elif actual_text == expected_value:
            status = ComparisonStatus.PASS
            expected_text = expected_value
            reason = (
                f"匹配值“{match_text}”与标准值“{actual_text}”一致"
            )
        else:
            status = ComparisonStatus.DATA_ERROR
            expected_text = expected_value
            reason = (
                f"匹配值“{match_text}”对应的标准值为“{expected_value}”，"
                f"实际值为“{actual_text}”"
            )

        results.append(
            ComparisonRowResult(
                sheet_name=sheet.sheet_name,
                row_number=row_number,
                match_field_name=match_field.name,
                match_value=match_text,
                standard_field_name=standard_field.name,
                actual_value=actual_text,
                expected_value=expected_text,
                status=status,
                reason=reason,
                match_column_letter=match_field.excel_column,
                match_column_number=match_field.excel_column_number,
                standard_column_letter=standard_field.excel_column,
                standard_column_number=standard_field.excel_column_number,
            )
        )
    return results, skipped_empty_count


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
        f"{field.excel_column}：{field.display_name}"
        for field in sheet.selected_fields
    ]
    return "、".join(options) or "（无）"
