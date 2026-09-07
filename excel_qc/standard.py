"""P0-05 标准 Excel 导入：把用户指定的标准数据集转为可被 P0-06 读取的数据集。"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from excel_qc.coordinates import column_letter
from excel_qc.errors import (
    StandardBlankValueError,
    StandardConfigError,
    StandardEmptySheetError,
    StandardFieldAmbiguousError,
    StandardFieldNotFoundError,
    StandardMappingConflictError,
    StandardNoDataError,
    StandardSheetNotFoundError,
)
from excel_qc.inspector import WorkbookInspector
from excel_qc.loader import LoadedSheet, load_workbook
from excel_qc.models import (
    ColumnField,
    SheetStructure,
    StandardColumn,
    StandardDataset,
    StandardDatasetConfig,
    StandardRecord,
)
from excel_qc.text import cell_text, is_blank_value


def import_standard_dataset(
    path: str | Path,
    config: StandardDatasetConfig,
) -> StandardDataset:
    """只读导入标准 Excel 并生成标准数据集。

    不会修改原始标准文件；Sheet、匹配字段、标准值字段均由用户配置指定，
    不硬编码任何业务字段。
    """
    _validate_config(config)
    source = Path(path)
    raw = load_workbook(source)
    structure = WorkbookInspector().inspect(raw)

    sheet = _find_sheet(structure, config.sheet_name)
    loaded_sheet = _find_loaded_sheet(raw.sheets, config.sheet_name)
    if sheet.empty:
        raise StandardEmptySheetError(
            f"标准数据集“{config.dataset_name}”：Sheet“{config.sheet_name}”为空"
        )
    if sheet.header_row is None or not sheet.fields:
        raise StandardEmptySheetError(
            f"标准数据集“{config.dataset_name}”：Sheet“{config.sheet_name}”"
            "未识别到有效表头/字段"
        )

    match_column = _resolve_column(sheet, config.match_field, "匹配字段")
    standard_column = _resolve_column(sheet, config.standard_field, "标准值字段")
    if match_column.column_index == standard_column.column_index:
        raise StandardConfigError(
            f"标准数据集“{config.dataset_name}”：匹配字段与标准值字段不能是同一列"
        )

    records, skipped_empty_count, duplicate_same_count = _read_records(
        dataset_name=config.dataset_name,
        sheet=sheet,
        loaded_sheet=loaded_sheet,
        match_column=match_column,
        standard_column=standard_column,
    )
    if not records:
        raise StandardNoDataError(
            f"标准数据集“{config.dataset_name}”：Sheet“{config.sheet_name}”"
            "中没有有效数据记录"
        )

    return StandardDataset(
        dataset_name=config.dataset_name,
        source_file=source,
        sheet_name=sheet.sheet_name,
        header_row=sheet.header_row,
        match_column=match_column,
        standard_column=standard_column,
        records=tuple(records),
        skipped_empty_row_count=skipped_empty_count,
        duplicate_same_value_count=duplicate_same_count,
    )


def _validate_config(config: StandardDatasetConfig) -> None:
    if not config.dataset_name.strip():
        raise StandardConfigError("标准数据集名称不能为空")
    if not config.sheet_name.strip():
        raise StandardConfigError("标准数据 Sheet 名称不能为空")
    if not config.match_field.strip():
        raise StandardConfigError("匹配字段不能为空")
    if not config.standard_field.strip():
        raise StandardConfigError("标准值字段不能为空")


def _find_sheet(structure, sheet_name: str) -> SheetStructure:
    for sheet in structure.sheets:
        if sheet.sheet_name == sheet_name:
            return sheet
    available = "、".join(sheet.sheet_name for sheet in structure.sheets) or "（无）"
    raise StandardSheetNotFoundError(
        f"标准数据 Sheet“{sheet_name}”不存在，可选 Sheet：{available}"
    )


def _find_loaded_sheet(sheets: tuple[LoadedSheet, ...], sheet_name: str) -> LoadedSheet:
    for sheet in sheets:
        if sheet.sheet_name == sheet_name:
            return sheet
    raise StandardSheetNotFoundError(f"Sheet“{sheet_name}”不存在")


def _resolve_column(
    sheet: SheetStructure,
    identifier: str,
    role: str,
) -> StandardColumn:
    """把用户输入的字段名/列字母解析为唯一标准数据列。"""
    key = identifier.strip()
    by_name = [
        field
        for field in sheet.fields
        if field.name == key or field.display_name == key
    ]
    if len(by_name) > 1:
        letters = "、".join(field.excel_column for field in by_name)
        raise StandardFieldAmbiguousError(
            f"Sheet“{sheet.sheet_name}”中{role}“{key}”对应多列（{letters}），"
            "请改用 Excel 列字母指定"
        )
    field = by_name[0] if by_name else _match_column_letter(sheet, key)
    if field is None:
        options = _format_available_fields(sheet)
        raise StandardFieldNotFoundError(
            f"Sheet“{sheet.sheet_name}”中不存在{role}“{key}”。"
            f"可选字段（Excel 列：字段名）：{options}"
        )
    return StandardColumn(
        name=field.name,
        excel_column_letter=field.excel_column,
        excel_column_number=field.column_index + 1,
        column_index=field.column_index,
    )


def _match_column_letter(sheet: SheetStructure, key: str) -> ColumnField | None:
    column_index = _parse_column_letter(key)
    if column_index is None:
        return None
    for field in sheet.fields:
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


def _read_records(
    dataset_name: str,
    sheet: SheetStructure,
    loaded_sheet: LoadedSheet,
    match_column: StandardColumn,
    standard_column: StandardColumn,
) -> tuple[list[StandardRecord], int, int]:
    """按物理行顺序读取表头下方的数据记录。"""
    rows = [list(row) for row in loaded_sheet.rows]
    records: list[StandardRecord] = []
    by_match: dict[str, StandardRecord] = {}
    skipped_empty_count = 0
    duplicate_same_count = 0

    if sheet.used_last_row is None or sheet.header_row is None:
        raise StandardNoDataError(
            f"标准数据集“{dataset_name}”：Sheet“{sheet.sheet_name}”中没有数据区域"
        )

    for row_number in range(sheet.header_row + 1, sheet.used_last_row + 1):
        row_index = row_number - 1
        row = rows[row_index] if row_index < len(rows) else []
        match_text = cell_text(_cell(row, match_column.column_index))
        standard_text = cell_text(_cell(row, standard_column.column_index))

        if not match_text and not standard_text:
            if _row_is_empty(row, sheet):
                skipped_empty_count += 1
                continue
            _raise_blank_value(
                dataset_name, sheet.sheet_name, row_number,
                match_column, "匹配字段",
            )
        if not match_text:
            _raise_blank_value(
                dataset_name, sheet.sheet_name, row_number,
                match_column, "匹配字段",
            )
        if not standard_text:
            _raise_blank_value(
                dataset_name, sheet.sheet_name, row_number,
                standard_column, "标准值字段",
            )

        previous = by_match.get(match_text)
        if previous is not None:
            if previous.standard_value != standard_text:
                raise StandardMappingConflictError(
                    f"标准数据集“{dataset_name}”Sheet“{sheet.sheet_name}”："
                    f"匹配值“{match_text}”在第 {previous.row_number} 行和第 "
                    f"{row_number} 行对应不同标准值"
                    f"（“{previous.standard_value}”与“{standard_text}”），"
                    f"匹配字段列为 {match_column.excel_column_letter}，"
                    "无法确定唯一标准值"
                )
            duplicate_same_count += 1
            continue

        record = StandardRecord(
            row_number=row_number,
            match_value=match_text,
            standard_value=standard_text,
        )
        records.append(record)
        by_match[match_text] = record

    return records, skipped_empty_count, duplicate_same_count


def _raise_blank_value(
    dataset_name: str,
    sheet_name: str,
    row_number: int,
    column: StandardColumn,
    role: str,
) -> None:
    raise StandardBlankValueError(
        f"标准数据集“{dataset_name}”Sheet“{sheet_name}”：{role}为空，"
        f"第 {row_number} 行、Excel 第 {column.excel_column_number} 列"
        f"（{column.excel_column_letter}：{column.display_name}）"
    )


def _row_is_empty(row: list[Any], sheet: SheetStructure) -> bool:
    if sheet.used_first_col is None or sheet.used_last_col is None:
        return True
    for col_index in range(sheet.used_first_col - 1, sheet.used_last_col):
        if not is_blank_value(_cell(row, col_index)):
            return False
    return True


def _cell(row: list[Any], column_index: int) -> Any:
    return row[column_index] if column_index < len(row) else None


def _format_available_fields(sheet: SheetStructure) -> str:
    options = [
        f"{field.excel_column}：{field.display_name}" for field in sheet.fields
    ]
    if len(options) > 8:
        options = options[:8] + [f"…共 {len(sheet.fields)} 列"]
    return "、".join(options) or "（无）"
