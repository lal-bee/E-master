"""Excel 结构与列名识别（P0-02、P0-03）。"""

from __future__ import annotations

from collections import Counter
from datetime import date, datetime, time
from decimal import Decimal
from pathlib import Path
from typing import Any

from excel_qc.loader import LoadedSheet, LoadedWorkbook, load_workbook
from excel_qc.models import (
    BasicDataType,
    ColumnField,
    Issue,
    SheetStructure,
    WorkbookStructure,
)


def inspect_workbook(path: str | Path) -> WorkbookStructure:
    """导入 .xlsx 并完成 Sheet 与列名结构识别。"""
    raw = load_workbook(path)
    return WorkbookInspector().inspect(raw)


class WorkbookInspector:
    """把已导入的原始 Sheet 转为结构识别结果。"""

    def inspect(self, raw: LoadedWorkbook) -> WorkbookStructure:
        sheets = [self._inspect_sheet(sheet) for sheet in raw.sheets]
        issues = [issue for sheet in sheets for issue in sheet.issues]
        if not sheets:
            issues.append(Issue(code="NO_SHEET", message="文件中未发现任何 Sheet"))
        return WorkbookStructure(
            path=raw.path,
            sheet_count=len(sheets),
            sheets=sheets,
            issues=issues,
        )

    def _inspect_sheet(self, loaded: LoadedSheet) -> SheetStructure:
        rows = [list(row) for row in loaded.rows]
        result = SheetStructure(sheet_name=loaded.sheet_name)
        region = self._used_region(rows)
        if region is None:
            result.empty = True
            result.issues.append(
                Issue(
                    code="EMPTY_SHEET",
                    message="该 Sheet 没有任何内容",
                    sheet=result.sheet_name,
                )
            )
            return result

        (
            content_rows,
            content_cols,
            used_first_row,
            used_last_row,
            used_first_col,
            used_last_col,
        ) = region
        result.used_first_row = used_first_row
        result.used_last_row = used_last_row
        result.used_first_col = used_first_col
        result.used_last_col = used_last_col

        header_row_index = self._find_header_row(rows, content_rows, content_cols)
        if header_row_index is None:
            result.issues.append(
                Issue(
                    code="NO_HEADER_FOUND",
                    message="未识别到有效的表头行，字段名需人工确认",
                    sheet=result.sheet_name,
                )
            )
            return result

        result.header_row = header_row_index + 1
        data_rows = [
            row_index
            for row_index in content_rows
            if row_index > header_row_index
            and any(
                not _is_blank(self._cell(rows, row_index, col_index))
                for col_index in content_cols
            )
        ]
        result.data_row_count = len(data_rows)

        normalized_names: list[str] = []
        for col_index in content_cols:
            raw_name = self._cell(rows, header_row_index, col_index)
            name = "" if _is_blank(raw_name) else str(raw_name).strip()
            normalized_names.append(name.strip().lower())
            if not name:
                excel_column = self._excel_column(col_index)
                result.issues.append(
                    Issue(
                        code="BLANK_HEADER_CELL",
                        message=(
                            f"表头第 {excel_column} 列缺少字段名，"
                            f"将使用占位名“列{excel_column}”"
                        ),
                        sheet=result.sheet_name,
                    )
                )
            data_type, non_empty_count, samples = self._column_summary(
                rows, data_rows, col_index
            )
            result.fields.append(
                ColumnField(
                    column_index=col_index,
                    name=name,
                    data_type=data_type,
                    non_empty_count=non_empty_count,
                    sample_values=samples,
                )
            )

        duplicates = sorted(
            {
                name
                for name, count in Counter(
                    normalized for normalized in normalized_names if normalized
                ).items()
                if count > 1
            }
        )
        if duplicates:
            result.issues.append(
                Issue(
                    code="DUPLICATE_HEADER_NAME",
                    message="存在重复字段名: " + "、".join(duplicates),
                    sheet=result.sheet_name,
                )
            )
        return result

    @staticmethod
    def _used_region(
        rows: list[list[Any]],
    ) -> tuple[list[int], list[int], int, int, int, int] | None:
        """返回有内容行的 0 起始下标、有内容列的 0 起始下标与物理行列范围。"""
        height = len(rows)
        width = max((len(row) for row in rows), default=0)
        if height == 0 or width == 0:
            return None

        content_rows = [
            row_index
            for row_index in range(height)
            if any(not _is_blank(value) for value in rows[row_index])
        ]
        if not content_rows:
            return None

        content_cols = [
            col_index
            for col_index in range(width)
            if any(
                col_index < len(rows[row_index])
                and not _is_blank(rows[row_index][col_index])
                for row_index in content_rows
            )
        ]
        if not content_cols:
            return None

        return (
            content_rows,
            content_cols,
            content_rows[0] + 1,
            content_rows[-1] + 1,
            content_cols[0] + 1,
            content_cols[-1] + 1,
        )

    @staticmethod
    def _find_header_row(
        rows: list[list[Any]],
        content_rows: list[int],
        content_cols: list[int],
    ) -> int | None:
        """返回首个有效表头的 0 起始行下标。

        当前阶段的确定性约定：
        1. 多列表格：该行至少有 2 个非空单元格，且非空单元格均为文本型，
           避免把“标题/说明”行当作表头；
        2. 单列表格：无法区分标题与表头，取首个有内容行。
        """
        if len(content_cols) == 1:
            return content_rows[0]

        for row_index in content_rows:
            values = [
                WorkbookInspector._cell(rows, row_index, col_index)
                for col_index in content_cols
                if not _is_blank(
                    WorkbookInspector._cell(rows, row_index, col_index)
                )
            ]
            if len(values) < 2:
                continue
            if all(_is_text_like(value) for value in values):
                return row_index
        return None

    @staticmethod
    def _column_summary(
        rows: list[list[Any]],
        data_rows: list[int],
        col_index: int,
    ) -> tuple[BasicDataType, int, tuple[str, ...]]:
        types: set[BasicDataType] = set()
        samples: list[str] = []
        non_empty_count = 0
        for row_index in data_rows:
            value = WorkbookInspector._cell(rows, row_index, col_index)
            if _is_blank(value):
                continue
            non_empty_count += 1
            types.add(_basic_type(value))
            if len(samples) < 3:
                samples.append(str(value))

        if not types:
            return BasicDataType.BLANK, 0, ()

        # Excel 单元格中整数与小数混存很常见，统一归为“数值”更符合直觉。
        if types <= {BasicDataType.INTEGER, BasicDataType.NUMBER}:
            if BasicDataType.NUMBER in types:
                data_type = BasicDataType.NUMBER
            else:
                data_type = BasicDataType.INTEGER
        elif len(types) == 1:
            data_type = next(iter(types))
        else:
            data_type = BasicDataType.MIXED
        return data_type, non_empty_count, tuple(samples)

    @staticmethod
    def _cell(rows: list[list[Any]], row_index: int, col_index: int) -> Any:
        row = rows[row_index]
        return row[col_index] if col_index < len(row) else None

    @staticmethod
    def _excel_column(col_index: int) -> str:
        letters = ""
        index = col_index + 1
        while index > 0:
            index, remainder = divmod(index - 1, 26)
            letters = chr(65 + remainder) + letters
        return letters


def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and value != value:  # NaN
        return True
    return isinstance(value, str) and value.strip() == ""


def _is_text_like(value: Any) -> bool:
    """判断单元格值是否可作为字段名（数字、日期等不作为文本表头）。"""
    if not isinstance(value, str):
        return False
    stripped = value.strip()
    if not stripped:
        return False
    return not _looks_numeric(stripped)


def _looks_numeric(text: str) -> bool:
    try:
        float(text.replace(",", ""))
    except ValueError:
        return False
    return True


def _basic_type(value: Any) -> BasicDataType:
    if isinstance(value, bool):
        return BasicDataType.BOOLEAN
    if isinstance(value, datetime):
        return BasicDataType.DATETIME
    if isinstance(value, date):
        return BasicDataType.DATE
    if isinstance(value, time):
        return BasicDataType.TIME
    if isinstance(value, int):
        return BasicDataType.INTEGER
    if isinstance(value, (float, Decimal)):
        return BasicDataType.NUMBER
    if isinstance(value, str):
        return BasicDataType.TEXT
    return BasicDataType.OTHER
