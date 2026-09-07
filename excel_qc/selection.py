"""P0-04 字段选择：把用户选择的 Sheet/字段解析为后续核验可直接使用的范围。"""

from __future__ import annotations

from collections.abc import Iterable

from excel_qc.coordinates import column_letter
from excel_qc.errors import (
    AmbiguousFieldError,
    DuplicateSheetSelectionError,
    EmptySheetSelectionError,
    NoFieldSelectedError,
    UnknownFieldError,
    UnknownSheetError,
)
from excel_qc.models import (
    ColumnField,
    FieldSelection,
    FieldSelectionRequest,
    Issue,
    SheetSelection,
    SheetStructure,
    WorkbookSelection,
    WorkbookStructure,
)


def select_fields(
    structure: WorkbookStructure,
    requests: Iterable[FieldSelectionRequest],
) -> WorkbookSelection:
    """根据用户请求生成核验范围。

    - 未出现在请求中的 Sheet 不会参与后续核验；
    - 每个 Sheet 内未出现的字段不会参与后续核验；
    - “Sheet 已选中但未选择任何字段”是可表达状态，选择结果中保留该
      Sheet，并通过 NO_FIELD_SELECTED 提示；
    - 选择结果保存 Sheet 名、表头物理行号、唯一 Excel 列坐标与原始
      Sheet 结构，后续比对引擎可据此直接回溯原始文件坐标。
    """
    request_list = list(requests)
    sheets_by_name = {sheet.sheet_name: sheet for sheet in structure.sheets}
    _validate_unique_sheets(request_list, sheets_by_name)

    selections: list[SheetSelection] = []
    issues: list[Issue] = []
    for request in request_list:
        sheet = sheets_by_name[request.sheet_name]
        if sheet.empty:
            raise EmptySheetSelectionError(
                f"Sheet“{request.sheet_name}”为空，无法选择核验字段"
            )
        selected_fields = _resolve_fields(sheet, request.fields)
        if not selected_fields:
            issues.append(
                Issue(
                    code="NO_FIELD_SELECTED",
                    message=f"Sheet“{request.sheet_name}”已选中，但未选择任何字段，"
                    "该 Sheet 暂无可核验字段",
                    sheet=request.sheet_name,
                )
            )
        selections.append(
            SheetSelection(
                sheet_name=sheet.sheet_name,
                header_row=sheet.header_row if sheet.header_row is not None else 0,
                selected_fields=selected_fields,
                structure=sheet,
            )
        )
    return WorkbookSelection(
        path=structure.path,
        selections=tuple(selections),
        issues=tuple(issues),
    )


def validate_selection(selection: WorkbookSelection) -> None:
    """校验核验范围是否可用于后续比对。

    - 未选择任何 Sheet 时拒绝；
    - 某个 Sheet 被选中但没有字段时拒绝，并提示先选择字段。
    """
    if not selection.selections:
        raise NoFieldSelectedError("未选择任何参与核验的 Sheet")
    for sheet in selection.selections:
        if not sheet.selected_fields:
            raise NoFieldSelectedError(
                f"Sheet“{sheet.sheet_name}”已选中，但未选择任何字段，"
                "无法参与后续核验"
            )


def _validate_unique_sheets(
    requests: list[FieldSelectionRequest],
    sheets_by_name: dict[str, SheetStructure],
) -> None:
    seen: set[str] = set()
    for request in requests:
        if request.sheet_name not in sheets_by_name:
            available = "、".join(sorted(sheets_by_name)) or "（无）"
            raise UnknownSheetError(
                f"Sheet“{request.sheet_name}”不存在，可选 Sheet：{available}"
            )
        if request.sheet_name in seen:
            raise DuplicateSheetSelectionError(
                f"Sheet“{request.sheet_name}”被重复选择，请合并为一条请求"
            )
        seen.add(request.sheet_name)


def _resolve_fields(
    sheet: SheetStructure,
    identifiers: Iterable[str],
) -> tuple[FieldSelection, ...]:
    """把用户输入解析为唯一列。

    用户输入支持：
    1. 字段名（如“设备编码”）或占位名（如“列B”）；
    2. Excel 列字母（如“B”）。

    字段名重复时，必须改用 Excel 列字母以保证唯一定位。
    """
    resolved: list[FieldSelection] = []
    seen_columns: set[int] = set()
    for raw_identifier in identifiers:
        identifier = raw_identifier.strip()
        if not identifier:
            raise UnknownFieldError(
                f"Sheet“{sheet.sheet_name}”中存在空的字段标识"
            )

        column_field = _match_single_field(sheet, identifier)
        if column_field.column_index in seen_columns:
            continue  # 同一列被重复勾选时幂等处理
        seen_columns.add(column_field.column_index)
        resolved.append(_to_field_selection(column_field))
    return tuple(resolved)


def _match_single_field(sheet: SheetStructure, identifier: str) -> ColumnField:
    by_name = [
        field
        for field in sheet.fields
        if field.name == identifier or field.display_name == identifier
    ]
    if len(by_name) > 1:
        letters = "、".join(field.excel_column for field in by_name)
        raise AmbiguousFieldError(
            f"Sheet“{sheet.sheet_name}”中字段名“{identifier}”对应多列（{letters}），"
            "请改用 Excel 列字母选择"
        )
    if by_name:
        return by_name[0]

    column_index = _parse_column_letter(identifier)
    if column_index is not None:
        for field in sheet.fields:
            if field.column_index == column_index:
                return field

    available = _format_available_fields(sheet)
    raise UnknownFieldError(
        f"Sheet“{sheet.sheet_name}”中不存在字段“{identifier}”。"
        f"可选字段（Excel 列：字段名）：{available}"
    )


def _to_field_selection(field: ColumnField) -> FieldSelection:
    return FieldSelection(
        name=field.name,
        excel_column_letter=field.excel_column,
        excel_column_number=field.column_index + 1,
        column_index=field.column_index,
        data_type=field.data_type,
    )


def _parse_column_letter(identifier: str) -> int | None:
    letters = identifier.upper()
    if not letters or not letters.isascii() or not letters.isalpha():
        return None
    index = 0
    for char in letters:
        index = index * 26 + (ord(char) - ord("A") + 1)
    column_index = index - 1
    return column_index if column_letter(column_index) == letters else None


def _format_available_fields(sheet: SheetStructure) -> str:
    options = [
        f"{field.excel_column}：{field.display_name}" for field in sheet.fields
    ]
    if len(options) > 8:
        options = options[:8] + [f"…共 {len(sheet.fields)} 列"]
    return "、".join(options) or "（无）"
