"""P1-02 Excel 结构探查增强：发现并描述 Excel 文件结构。

本模块只负责“Excel 里有什么结构”，不负责字段语义映射（P1-03）或
后续清洗/转换。它复用 V1.0 的 loader、BasicDataType 与文本工具，
不建立第二套 Excel 读取体系。
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal
from enum import Enum
from pathlib import Path
from typing import Any

from excel_qc.coordinates import column_letter
from excel_qc.loader import LoadedSheet, load_workbook
from excel_qc.models import BasicDataType
from excel_qc.text import cell_text, is_blank_value

_HEADER_SCAN_CONTENT_ROW_LIMIT = 60
_HEADER_BLOCK_MIN_SCORE = 0.75
_TRAILING_NOTE_PREFIXES = (
    "说明",
    "备注",
    "注意",
    "提示",
    "附注",
    "数据来源",
    "统计口径",
    "填表",
)
_TRAILING_NOTE_PATTERN = re.compile(r"^注\s*[:：]")

_DATE_TEXT_FORMATS = (
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%Y.%m.%d",
    "%Y年%m月%d日",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
)
_NUMBER_TEXT_PATTERN = re.compile(
    r"^[+-]?(?:\d[\d,]*(?:\.\d*)?|\.\d+)$"
)
_BOOLEAN_TEXT_VALUES = {"true", "false"}


class ProfileIssueLevel(str, Enum):
    """结构探查提示级别。"""

    INFO = "INFO"
    WARNING = "WARNING"
    ERROR = "ERROR"


@dataclass(frozen=True)
class ProfileIssue:
    """一条结构异常/提示，结构化保留代码与位置。"""

    code: str
    message: str
    level: ProfileIssueLevel
    sheet_name: str | None = None
    row_number: int | None = None
    column_letter: str | None = None


@dataclass(frozen=True)
class HeaderCandidate:
    """一个连续表头候选块（支持多行表头）。"""

    start_row: int  # Excel 物理行号，1 起始
    end_row: int
    row_numbers: tuple[int, ...]
    score: float

    @property
    def row_count(self) -> int:
        return len(self.row_numbers)


@dataclass(frozen=True)
class FieldProfile:
    """结构探查阶段的一个字段特征。"""

    name: str  # 原始字段名；空表头为 ""
    column_index: int  # 0 起始列下标
    column_number: int  # Excel 1 起始列号
    column_letter: str
    data_type: BasicDataType
    non_empty_count: int
    empty_count: int
    sample_values: tuple[str, ...]
    suspected_date: bool = False
    suspected_number: bool = False
    suspected_text: bool = False
    suspected_boolean: bool = False

    @property
    def display_name(self) -> str:
        return self.name or f"列{self.column_letter}"


@dataclass(frozen=True)
class SheetProfile:
    """单个 Sheet 的结构探查结果。"""

    sheet_name: str
    sheet_index: int
    empty: bool
    total_row_count: int | None
    total_column_count: int | None
    used_first_row: int | None
    used_last_row: int | None
    used_first_col: int | None
    used_last_col: int | None
    merged_cell_ranges: tuple[str, ...] = ()
    hidden_row_numbers: tuple[int, ...] = ()
    hidden_column_letters: tuple[str, ...] = ()
    header_candidates: tuple[HeaderCandidate, ...] = ()
    header_start_row: int | None = None
    header_end_row: int | None = None
    header_row: int | None = None
    header_row_count: int = 0
    multi_row_header: bool = False
    data_first_row: int | None = None
    data_last_row: int | None = None
    data_first_col: int | None = None
    data_last_col: int | None = None
    data_row_count: int = 0
    fields: tuple[FieldProfile, ...] = ()
    issues: tuple[ProfileIssue, ...] = ()

    @property
    def sheet_number(self) -> int:
        return self.sheet_index + 1

    @property
    def field_count(self) -> int:
        return len(self.fields)


@dataclass(frozen=True)
class WorkbookProfile:
    """整个 Excel 文件的结构探查结果。"""

    path: Path
    sheets: tuple[SheetProfile, ...]
    issues: tuple[ProfileIssue, ...] = ()

    @property
    def sheet_count(self) -> int:
        return len(self.sheets)

    @property
    def sheet_names(self) -> tuple[str, ...]:
        return tuple(sheet.sheet_name for sheet in self.sheets)


def profile_workbook(path: str | Path) -> WorkbookProfile:
    """对 .xlsx 执行增强结构探查，返回 WorkbookProfile。

    - 原始文件只读；
    - 只做结构发现与描述，不做字段语义映射；
    - 读取元数据时需要常规模式打开文件，因此仅供探查调用。
    """
    raw = load_workbook(path, include_sheet_metadata=True)
    sheets = tuple(_profile_sheet(sheet) for sheet in raw.sheets)
    all_issues = tuple(issue for sheet in sheets for issue in sheet.issues)
    if not sheets:
        all_issues = (
            ProfileIssue(
                code="NO_SHEET",
                message="文件中未发现任何 Sheet",
                level=ProfileIssueLevel.ERROR,
            ),
        )
    return WorkbookProfile(path=raw.path, sheets=sheets, issues=all_issues)


def _profile_sheet(sheet: LoadedSheet) -> SheetProfile:
    rows = [list(row) for row in sheet.rows]
    issues: list[ProfileIssue] = []
    issues.extend(_sheet_metadata_issues(sheet))

    region = _used_region(rows)
    if region is None:
        issues.append(
            ProfileIssue(
                code="EMPTY_SHEET",
                message="该 Sheet 没有任何内容",
                level=ProfileIssueLevel.WARNING,
                sheet_name=sheet.sheet_name,
            )
        )
        return SheetProfile(
            sheet_name=sheet.sheet_name,
            sheet_index=sheet.sheet_index,
            empty=True,
            total_row_count=sheet.total_row_count,
            total_column_count=sheet.total_column_count,
            used_first_row=None,
            used_last_row=None,
            used_first_col=None,
            used_last_col=None,
            merged_cell_ranges=sheet.merged_cell_ranges,
            hidden_row_numbers=sheet.hidden_row_numbers,
            hidden_column_letters=sheet.hidden_column_letters,
            issues=tuple(issues),
        )

    content_rows, content_cols, used_first_row, used_last_row, used_first_col, used_last_col = region
    header_candidates: tuple[HeaderCandidate, ...] = ()
    header_start_row: int | None = None
    header_end_row: int | None = None
    header_row: int | None = None
    header_row_count = 0
    multi_row_header = False

    if len(content_cols) == 1:
        header_row = content_rows[0] + 1
        header_start_row = header_row
        header_end_row = header_row
        header_row_count = 1
        candidate = HeaderCandidate(
            start_row=header_row,
            end_row=header_row,
            row_numbers=(header_row,),
            score=1.0,
        )
        header_candidates = (candidate,)
    else:
        candidates, selected = _find_header_block(
            rows=rows,
            content_rows=content_rows,
            content_cols=content_cols,
        )
        header_candidates = candidates
        if selected is not None:
            header_start_row = selected.start_row
            header_end_row = selected.end_row
            header_row = selected.end_row
            header_row_count = selected.row_count
            multi_row_header = header_row_count > 1
            issues.extend(_header_block_issues(sheet.sheet_name, selected, used_first_row))
        else:
            issues.append(
                ProfileIssue(
                    code="NO_HEADER_FOUND",
                    message="未识别到有效表头，字段名需人工确认",
                    level=ProfileIssueLevel.WARNING,
                    sheet_name=sheet.sheet_name,
                )
            )

    if header_row is None:
        return SheetProfile(
            sheet_name=sheet.sheet_name,
            sheet_index=sheet.sheet_index,
            empty=False,
            total_row_count=sheet.total_row_count,
            total_column_count=sheet.total_column_count,
            used_first_row=used_first_row,
            used_last_row=used_last_row,
            used_first_col=used_first_col,
            used_last_col=used_last_col,
            merged_cell_ranges=sheet.merged_cell_ranges,
            hidden_row_numbers=sheet.hidden_row_numbers,
            hidden_column_letters=sheet.hidden_column_letters,
            header_candidates=header_candidates,
            issues=tuple(issues),
        )

    if multi_row_header:
        issues.append(
            ProfileIssue(
                code="MULTI_ROW_HEADER",
                message=(
                    f"识别到多行表头（第 {header_start_row} 至 "
                    f"{header_end_row} 行），字段名以第 {header_row} 行为准"
                ),
                level=ProfileIssueLevel.INFO,
                sheet_name=sheet.sheet_name,
                row_number=header_start_row,
            )
        )

    data_content_rows, trailing_rows = _data_rows_after_trailing_notes(
        rows=rows,
        header_row=header_row,
        used_last_row=used_last_row,
        content_cols=content_cols,
    )
    for trailing_row in trailing_rows:
        issues.append(
            ProfileIssue(
                code="TRAILING_CONTENT",
                message=(
                    f"数据区末尾第 {trailing_row} 行疑似说明/备注内容，"
                    "已从有效数据行中排除"
                ),
                level=ProfileIssueLevel.INFO,
                sheet_name=sheet.sheet_name,
                row_number=trailing_row,
            )
        )

    data_first_row = data_content_rows[0] + 1 if data_content_rows else None
    data_last_row = data_content_rows[-1] + 1 if data_content_rows else None
    data_row_count = len(data_content_rows)
    if not data_content_rows:
        issues.append(
            ProfileIssue(
                code="NO_DATA_REGION",
                message=f"表头位于第 {header_row} 行，但其下方没有有效数据行",
                level=ProfileIssueLevel.WARNING,
                sheet_name=sheet.sheet_name,
                row_number=header_row,
            )
        )

    fields, field_issues = _profile_fields(
        sheet_name=sheet.sheet_name,
        rows=rows,
        header_row=header_row,
        data_rows=data_content_rows,
        content_cols=content_cols,
    )
    issues.extend(field_issues)

    return SheetProfile(
        sheet_name=sheet.sheet_name,
        sheet_index=sheet.sheet_index,
        empty=False,
        total_row_count=sheet.total_row_count,
        total_column_count=sheet.total_column_count,
        used_first_row=used_first_row,
        used_last_row=used_last_row,
        used_first_col=used_first_col,
        used_last_col=used_last_col,
        merged_cell_ranges=sheet.merged_cell_ranges,
        hidden_row_numbers=sheet.hidden_row_numbers,
        hidden_column_letters=sheet.hidden_column_letters,
        header_candidates=header_candidates,
        header_start_row=header_start_row,
        header_end_row=header_end_row,
        header_row=header_row,
        header_row_count=header_row_count,
        multi_row_header=multi_row_header,
        data_first_row=data_first_row,
        data_last_row=data_last_row,
        data_first_col=used_first_col,
        data_last_col=used_last_col,
        data_row_count=data_row_count,
        fields=fields,
        issues=tuple(issues),
    )


def _sheet_metadata_issues(sheet: LoadedSheet) -> list[ProfileIssue]:
    issues: list[ProfileIssue] = []
    if sheet.merged_cell_ranges:
        issues.append(
            ProfileIssue(
                code="MERGED_CELL_RANGES",
                message=(
                    f"检测到 {len(sheet.merged_cell_ranges)} 个合并区域："
                    + "、".join(sheet.merged_cell_ranges)
                ),
                level=ProfileIssueLevel.INFO,
                sheet_name=sheet.sheet_name,
            )
        )
    if sheet.hidden_row_numbers:
        issues.append(
            ProfileIssue(
                code="HIDDEN_ROWS",
                message=(
                    f"检测到隐藏行："
                    + "、".join(str(row) for row in sheet.hidden_row_numbers)
                ),
                level=ProfileIssueLevel.INFO,
                sheet_name=sheet.sheet_name,
                row_number=sheet.hidden_row_numbers[0],
            )
        )
    if sheet.hidden_column_letters:
        issues.append(
            ProfileIssue(
                code="HIDDEN_COLUMNS",
                message=(
                    "检测到隐藏列："
                    + "、".join(sheet.hidden_column_letters)
                ),
                level=ProfileIssueLevel.INFO,
                sheet_name=sheet.sheet_name,
                column_letter=sheet.hidden_column_letters[0],
            )
        )
    return issues


def _used_region(
    rows: list[list[Any]],
) -> tuple[list[int], list[int], int, int, int, int] | None:
    content_rows = [
        row_index
        for row_index, row in enumerate(rows)
        if any(not is_blank_value(value) for value in row)
    ]
    if not content_rows:
        return None
    width = max((len(row) for row in rows), default=0)
    content_cols = [
        col_index
        for col_index in range(width)
        if any(
            col_index < len(rows[row_index])
            and not is_blank_value(rows[row_index][col_index])
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


def _find_header_block(
    rows: list[list[Any]],
    content_rows: list[int],
    content_cols: list[int],
) -> tuple[tuple[HeaderCandidate, ...], HeaderCandidate | None]:
    scan_rows = content_rows[:_HEADER_SCAN_CONTENT_ROW_LIMIT]
    scored_rows: dict[int, float] = {}
    for row_index in scan_rows:
        score = _header_row_score(rows[row_index], content_cols)
        if score is not None:
            scored_rows[row_index] = score

    if not scored_rows:
        return (), None

    selected_row = next(
        (
            row_index
            for row_index in scan_rows
            if scored_rows.get(row_index, 0.0) >= _HEADER_BLOCK_MIN_SCORE
        ),
        None,
    )
    if selected_row is None:
        return (), None

    block = [selected_row]
    # 向上并入紧邻的候选行，用于描述多行表头。
    for previous_row in reversed(scan_rows[: scan_rows.index(selected_row)]):
        if previous_row == block[-1] - 1 and previous_row in scored_rows:
            block.append(previous_row)
        elif previous_row != block[-1] - 1:
            break
    block.sort()
    average_score = sum(scored_rows[row] for row in block) / len(block)
    candidate = HeaderCandidate(
        start_row=block[0] + 1,
        end_row=block[-1] + 1,
        row_numbers=tuple(row + 1 for row in block),
        score=round(average_score, 4),
    )
    return (candidate,), candidate


def _header_row_score(
    row: list[Any],
    content_cols: list[int],
) -> float | None:
    non_blank = [
        value
        for col_index in content_cols
        if col_index < len(row)
        for value in [row[col_index]]
        if not is_blank_value(value)
    ]
    if len(non_blank) < 2:
        return None
    text_like_count = sum(1 for value in non_blank if _is_text_like(value))
    if text_like_count < 2:
        return None
    fill_ratio = len(non_blank) / len(content_cols)
    text_ratio = text_like_count / len(non_blank)
    # 不使用唯一值评分，避免“数据行名称更唯一”反而被误当成表头。
    score = 0.6 * fill_ratio + 0.4 * text_ratio
    return score


def _is_text_like(value: Any) -> bool:
    if not isinstance(value, str):
        return False
    text = value.strip()
    return bool(text) and not _looks_numeric(text)


def _looks_numeric(text: str) -> bool:
    try:
        float(text.replace(",", ""))
    except ValueError:
        return False
    return True


def _header_block_issues(
    sheet_name: str,
    candidate: HeaderCandidate,
    used_first_row: int,
) -> list[ProfileIssue]:
    issues: list[ProfileIssue] = []
    if candidate.start_row > used_first_row:
        issues.append(
            ProfileIssue(
                code="CONTENT_ABOVE_HEADER",
                message=(
                    f"表头前存在内容行（第 {used_first_row} 至 "
                    f"{candidate.start_row - 1} 行），疑似标题或说明"
                ),
                level=ProfileIssueLevel.INFO,
                sheet_name=sheet_name,
                row_number=used_first_row,
            )
        )
    return issues


def _data_rows_after_trailing_notes(
    rows: list[list[Any]],
    *,
    header_row: int,
    used_last_row: int,
    content_cols: list[int],
) -> tuple[list[int], list[int]]:
    content_rows = [
        row_index
        for row_index in range(header_row, used_last_row)
        if any(
            col_index < len(rows[row_index])
            and not is_blank_value(rows[row_index][col_index])
            for col_index in content_cols
        )
    ]
    trailing_rows: list[int] = []
    for row_index in reversed(content_rows):
        if not _is_trailing_note_row(rows[row_index], content_cols):
            break
        trailing_rows.append(row_index)
    if trailing_rows:
        trailing_rows.reverse()
        content_rows = content_rows[: -len(trailing_rows)]
    return content_rows, [row + 1 for row in trailing_rows]


def _is_trailing_note_row(row: list[Any], content_cols: list[int]) -> bool:
    non_blank = [
        value
        for col_index in content_cols
        if col_index < len(row)
        for value in [row[col_index]]
        if not is_blank_value(value)
    ]
    if len(non_blank) != 1:
        return False
    value = non_blank[0]
    if not isinstance(value, str):
        return False
    text = value.strip()
    if not text:
        return False
    return text.startswith(_TRAILING_NOTE_PREFIXES) or bool(
        _TRAILING_NOTE_PATTERN.match(text)
    )


def _profile_fields(
    *,
    sheet_name: str,
    rows: list[list[Any]],
    header_row: int,
    data_rows: list[int],
    content_cols: list[int],
) -> tuple[tuple[FieldProfile, ...], list[ProfileIssue]]:
    fields: list[FieldProfile] = []
    issues: list[ProfileIssue] = []
    normalized_names: list[str] = []
    header_row_index = header_row - 1

    for col_index in content_cols:
        header_value = _cell(rows, header_row_index, col_index)
        name = "" if is_blank_value(header_value) else str(header_value).strip()
        normalized_names.append(name.lower())
        if not name:
            issues.append(
                ProfileIssue(
                    code="EMPTY_HEADER_CELL",
                    message=(
                        f"第 {column_letter(col_index)} 列表头为空，"
                        f"将使用占位名“列{column_letter(col_index)}”"
                    ),
                    level=ProfileIssueLevel.WARNING,
                    sheet_name=sheet_name,
                    row_number=header_row,
                    column_letter=column_letter(col_index),
                )
            )

        values = [
            _cell(rows, row_index, col_index)
            for row_index in data_rows
            if not is_blank_value(_cell(rows, row_index, col_index))
        ]
        samples = tuple(cell_text(value) for value in values[:3])
        data_type = _field_data_type(values)
        non_empty_count = len(values)
        empty_count = max(len(data_rows) - non_empty_count, 0)
        suspected = _suspected_kinds(values)
        fields.append(
            FieldProfile(
                name=name,
                column_index=col_index,
                column_number=col_index + 1,
                column_letter=column_letter(col_index),
                data_type=data_type,
                non_empty_count=non_empty_count,
                empty_count=empty_count,
                sample_values=samples,
                suspected_date=suspected["date"],
                suspected_number=suspected["number"],
                suspected_text=suspected["text"],
                suspected_boolean=suspected["boolean"],
            )
        )

    duplicate_names = sorted(
        {
            name
            for name, count in Counter(
                normalized for normalized in normalized_names if normalized
            ).items()
            if count > 1
        }
    )
    if duplicate_names:
        issues.append(
            ProfileIssue(
                code="DUPLICATE_HEADER_NAME",
                message="存在重复字段名: " + "、".join(duplicate_names),
                level=ProfileIssueLevel.WARNING,
                sheet_name=sheet_name,
                row_number=header_row,
            )
        )
    return tuple(fields), issues


def _field_data_type(values: list[Any]) -> BasicDataType:
    if not values:
        return BasicDataType.BLANK
    types = {_basic_data_type(value) for value in values}
    if types <= {BasicDataType.INTEGER, BasicDataType.NUMBER}:
        return (
            BasicDataType.NUMBER
            if BasicDataType.NUMBER in types
            else BasicDataType.INTEGER
        )
    if len(types) == 1:
        return next(iter(types))
    return BasicDataType.MIXED


def _basic_data_type(value: Any) -> BasicDataType:
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


def _suspected_kinds(values: list[Any]) -> dict[str, bool]:
    if not values:
        return {"date": False, "number": False, "text": False, "boolean": False}
    suspected_date = all(_looks_like_date(value) for value in values)
    suspected_number = all(_looks_like_number(value) for value in values)
    suspected_boolean = all(_looks_like_boolean(value) for value in values)
    suspected_text = all(
        isinstance(value, str)
        for value in values
    ) and not (suspected_date or suspected_number or suspected_boolean)
    return {
        "date": suspected_date,
        "number": suspected_number,
        "text": suspected_text,
        "boolean": suspected_boolean,
    }


def _looks_like_date(value: Any) -> bool:
    if isinstance(value, (datetime, date, time)):
        return True
    if not isinstance(value, str):
        return False
    text = value.strip()
    if not text:
        return False
    for fmt in _DATE_TEXT_FORMATS:
        try:
            datetime.strptime(text, fmt)
        except ValueError:
            continue
        return True
    return False


def _looks_like_number(value: Any) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float, Decimal)):
        return True
    if not isinstance(value, str):
        return False
    text = value.strip()
    return bool(_NUMBER_TEXT_PATTERN.match(text))


def _looks_like_boolean(value: Any) -> bool:
    if isinstance(value, bool):
        return True
    return isinstance(value, str) and value.strip().lower() in _BOOLEAN_TEXT_VALUES


def _cell(rows: list[list[Any]], row_index: int, col_index: int) -> Any:
    row = rows[row_index]
    return row[col_index] if col_index < len(row) else None
