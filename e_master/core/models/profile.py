"""结构探查与问题报告的数据契约。"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path


class IssueLevel(str, Enum):
    """问题严重级别。"""

    INFO = "info"
    WARNING = "warning"
    ERROR = "error"


@dataclass(frozen=True)
class DataIssue:
    """一条可定位的结构/质量问题。"""

    code: str
    message: str
    level: IssueLevel = IssueLevel.WARNING
    sheet: str | None = None
    row: int | None = None
    column: str | None = None
    details: dict[str, object] = field(default_factory=dict)


@dataclass(frozen=True)
class HeaderCandidate:
    """候选表头行。"""

    row_number: int
    score: float
    reasons: tuple[str, ...] = ()
    sample_values: tuple[str, ...] = ()


@dataclass(frozen=True)
class ColumnProfile:
    """按推荐表头识别出的列概况。"""

    excel_column: str
    name: str
    non_empty_count: int
    sample_values: tuple[str, ...] = ()


@dataclass
class SheetProfile:
    """单个 Sheet 的结构画像。"""

    sheet_name: str
    used_first_row: int | None = None
    used_last_row: int | None = None
    used_first_col: int | None = None
    used_last_col: int | None = None
    header_candidates: list[HeaderCandidate] = field(default_factory=list)
    recommended_header: HeaderCandidate | None = None
    data_start_row: int | None = None
    data_row_count: int = 0
    column_count: int = 0
    empty_cell_count: int = 0
    duplicate_data_rows: int = 0
    columns: list[ColumnProfile] = field(default_factory=list)
    issues: list[DataIssue] = field(default_factory=list)


@dataclass
class WorkbookProfile:
    """一个源文件的结构探查结果。"""

    path: Path
    file_type: str
    total_sheets: int = 0
    sheets: list[SheetProfile] = field(default_factory=list)
    issues: list[DataIssue] = field(default_factory=list)
