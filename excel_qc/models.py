"""P0 阶段的数据契约：文件、Sheet、字段与基础类型。"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path

from excel_qc.coordinates import column_letter


class BasicDataType(str, Enum):
    """字段的基础数据类型（供结构识别阶段展示，不替代后续格式校验）。"""

    TEXT = "text"
    INTEGER = "integer"
    NUMBER = "number"
    DATE = "date"
    DATETIME = "datetime"
    TIME = "time"
    BOOLEAN = "boolean"
    BLANK = "blank"
    MIXED = "mixed"
    OTHER = "other"

    @property
    def label(self) -> str:
        labels = {
            BasicDataType.TEXT: "文本",
            BasicDataType.INTEGER: "整数",
            BasicDataType.NUMBER: "数值",
            BasicDataType.DATE: "日期",
            BasicDataType.DATETIME: "日期时间",
            BasicDataType.TIME: "时间",
            BasicDataType.BOOLEAN: "布尔",
            BasicDataType.BLANK: "空",
            BasicDataType.MIXED: "混合",
            BasicDataType.OTHER: "其他",
        }
        return labels[self]


@dataclass(frozen=True)
class Issue:
    """结构识别阶段的一条提示/告警。"""

    code: str
    message: str
    sheet: str | None = None


@dataclass(frozen=True)
class ColumnField:
    """一个参与结构识别的字段（对应一个 Excel 列）。"""

    column_index: int  # 0 起始列下标
    name: str  # 表头名称；为空时使用占位名
    data_type: BasicDataType
    non_empty_count: int
    sample_values: tuple[str, ...] = ()

    @property
    def excel_column(self) -> str:
        return column_letter(self.column_index)

    @property
    def display_name(self) -> str:
        return self.name or f"列{self.excel_column}"


@dataclass
class SheetStructure:
    """单个 Sheet 的结构识别结果。"""

    sheet_name: str
    empty: bool = False
    used_first_row: int | None = None  # Excel 物理行号（1 起始）
    used_last_row: int | None = None
    used_first_col: int | None = None  # Excel 物理列号（1 起始）
    used_last_col: int | None = None
    header_row: int | None = None  # 首个有效表头的 Excel 物理行号
    data_row_count: int = 0  # 表头下方含内容的非空数据行数
    fields: list[ColumnField] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)

    @property
    def field_count(self) -> int:
        return len(self.fields)


@dataclass
class WorkbookStructure:
    """一个 Excel 文件的结构识别结果。"""

    path: Path
    sheet_count: int
    sheets: list[SheetStructure]
    issues: list[Issue] = field(default_factory=list)

    @property
    def sheet_names(self) -> list[str]:
        return [sheet.sheet_name for sheet in self.sheets]


@dataclass(frozen=True)
class FieldSelectionRequest:
    """用户对一个 Sheet 的字段选择请求。

    fields 中的每个元素可以是字段名（如“设备编码”）或 Excel 列字母
    （如“C”）。字段名重复时必须使用列字母。
    """

    sheet_name: str
    fields: tuple[str, ...] = ()


@dataclass(frozen=True)
class FieldSelection:
    """一个被用户选中、已解析为唯一坐标的字段。

    P0-05/P0-06 直接读取本对象即可获得字段名与原始 Excel 列坐标，
    不需要再次扫描或猜测表头。
    """

    name: str  # 表头原始字段名；空表头为 ""
    excel_column_letter: str  # Excel 列字母，如 G
    excel_column_number: int  # Excel 1 起始列号，如 7
    column_index: int  # 0 起始列下标，供 pandas/内部切片使用
    data_type: BasicDataType

    @property
    def excel_column(self) -> str:
        return self.excel_column_letter

    @property
    def display_name(self) -> str:
        return self.name or f"列{self.excel_column_letter}"


@dataclass(frozen=True)
class SheetSelection:
    """一个参与核验的 Sheet 及其已选字段。

    structure 保留该 Sheet 的原始结构识别结果；selected_fields 中保留
    Excel 列坐标快照，后续模块无需依赖重新识别。
    """

    sheet_name: str
    header_row: int
    selected_fields: tuple[FieldSelection, ...]
    structure: SheetStructure

    @property
    def field_count(self) -> int:
        return len(self.selected_fields)


@dataclass(frozen=True)
class WorkbookSelection:
    """字段选择完成后的核验范围（整个 Excel 核验任务的选择结果）。"""

    path: Path
    selections: tuple[SheetSelection, ...] = ()
    issues: tuple[Issue, ...] = ()

    @property
    def sheet_count(self) -> int:
        return len(self.selections)

    @property
    def sheet_names(self) -> list[str]:
        return [selection.sheet_name for selection in self.selections]


@dataclass(frozen=True)
class StandardDatasetConfig:
    """标准数据集导入配置。

    match_field / standard_field 可以是字段名（如“设备类型名称”）或
    Excel 列字母（如“B”），导入时会解析为唯一列。
    """

    dataset_name: str
    sheet_name: str
    match_field: str
    standard_field: str


@dataclass(frozen=True)
class StandardColumn:
    """标准数据集中被指定为匹配字段/标准值字段的列。"""

    name: str  # 字段名；空表头为 ""
    excel_column_letter: str  # Excel 列字母，如 B
    excel_column_number: int  # Excel 1 起始列号，如 2
    column_index: int  # 0 起始列下标

    @property
    def excel_column(self) -> str:
        return self.excel_column_letter

    @property
    def display_name(self) -> str:
        return self.name or f"列{self.excel_column_letter}"


@dataclass(frozen=True)
class StandardRecord:
    """一条标准映射记录。

    row_number 为源标准 Excel 中的物理行号（1 起始），用于追溯。
    """

    row_number: int
    match_value: str
    standard_value: str


@dataclass(frozen=True)
class StandardDataset:
    """导入完成的标准数据集，是 P0-06 比对引擎的标准值数据来源。"""

    dataset_name: str
    source_file: Path
    sheet_name: str
    header_row: int  # 标准 Excel 中表头的物理行号
    match_column: StandardColumn
    standard_column: StandardColumn
    records: tuple[StandardRecord, ...]
    skipped_empty_row_count: int = 0
    duplicate_same_value_count: int = 0

    @property
    def record_count(self) -> int:
        return len(self.records)


class ComparisonStatus(str, Enum):
    """单条业务数据的标准映射比对结果。"""

    PASS = "PASS"
    DATA_NOT_FOUND = "DATA_NOT_FOUND"
    DATA_ERROR = "DATA_ERROR"


@dataclass(frozen=True)
class MappingCheckConfig:
    """一次标准映射比对中，业务 Excel 的两个参与字段。"""

    sheet_name: str  # 业务 Excel 中参与比对的 Sheet
    match_field: str  # 业务字段：提供匹配值（字段名或 Excel 列字母）
    standard_field: str  # 业务字段：提供实际标准值（字段名或 Excel 列字母）


@dataclass(frozen=True)
class ComparisonRowResult:
    """一条业务数据行的标准映射比对结果。"""

    sheet_name: str
    row_number: int  # 业务 Excel 物理行号（1 起始）
    match_field_name: str
    match_value: str
    standard_field_name: str
    actual_value: str
    expected_value: str  # DATA_NOT_FOUND 时为空字符串
    status: ComparisonStatus
    reason: str
    match_column_letter: str
    match_column_number: int
    standard_column_letter: str
    standard_column_number: int

    @property
    def error_type(self) -> str | None:
        """PASS 返回 None，否则返回 DATA_NOT_FOUND / DATA_ERROR。"""
        return self.status.value if self.status is not ComparisonStatus.PASS else None


@dataclass(frozen=True)
class ComparisonSummary:
    """一次比对任务的统计结果。"""

    total_rows: int
    pass_count: int
    data_not_found_count: int
    data_error_count: int


@dataclass(frozen=True)
class ComparisonResult:
    """一次“业务 Sheet → 标准数据集”比对任务的完整结果。"""

    business_file: Path
    sheet_name: str
    dataset_name: str
    match_column: FieldSelection
    standard_column: FieldSelection
    rows: tuple[ComparisonRowResult, ...]
    summary: ComparisonSummary
    skipped_empty_row_count: int = 0


class FormatValueType(str, Enum):
    """P0-07 支持的字段格式类型。"""

    DATE = "date"
    DATETIME = "datetime"
    NUMBER = "number"
    INTEGER = "integer"
    TEXT = "text"
    BOOLEAN = "boolean"

    @property
    def label(self) -> str:
        labels = {
            FormatValueType.DATE: "日期",
            FormatValueType.DATETIME: "日期时间",
            FormatValueType.NUMBER: "数值",
            FormatValueType.INTEGER: "整数",
            FormatValueType.TEXT: "文本",
            FormatValueType.BOOLEAN: "布尔",
        }
        return labels[self]


@dataclass(frozen=True)
class FieldFormatRule:
    """一条通用字段格式规则，只描述格式要求，不绑定具体业务名称。"""

    sheet_name: str  # 业务 Excel 中参与格式检查的 Sheet
    field: str  # 字段名或 Excel 列字母（必须已在 P0-04 中被选中）
    value_type: FormatValueType
    required: bool = False
    min_length: int | None = None
    max_length: int | None = None
    date_format: str | None = None  # 可选，仅 DATE/DATETIME 使用


@dataclass(frozen=True)
class FormatIssue:
    """一条格式错误，保留完整定位信息供 P0-08 使用。"""

    sheet_name: str
    row_number: int  # Excel 物理行号
    field_name: str
    excel_column_letter: str
    excel_column_number: int
    cell_address: str  # 如 G125
    actual_value: str
    actual_data_type: str
    expected_format: str
    error_code: str
    reason: str
    error_type: str = "FORMAT_ERROR"


@dataclass(frozen=True)
class FormatSummary:
    """一次格式检查的统计结果。"""

    total_checks: int
    pass_count: int
    error_count: int
    skipped_empty_row_count: int = 0


@dataclass(frozen=True)
class FormatCheckResult:
    """针对 P0-04 选中字段执行格式规则后的完整结果。"""

    business_file: Path
    issues: tuple[FormatIssue, ...]
    summary: FormatSummary


@dataclass(frozen=True)
class CellLocation:
    """统一错误定位：文件 → Sheet → 物理行 → 字段/列 → 单元格。"""

    source_file: Path
    sheet_name: str
    row_number: int  # Excel 物理行号（1 起始）
    column_number: int  # Excel 1 起始列号
    column_letter: str  # Excel 列字母
    column_name: str  # 字段名；空表头使用占位名

    @property
    def cell_reference(self) -> str:
        """统一由列字母 + 物理行号生成，如 G125。"""
        return f"{self.column_letter}{self.row_number}"


class ValidationErrorType(str, Enum):
    """统一核验错误类型。"""

    FORMAT_ERROR = "FORMAT_ERROR"
    DATA_NOT_FOUND = "DATA_NOT_FOUND"
    DATA_ERROR = "DATA_ERROR"


class ValidationSource(str, Enum):
    """错误来源模块。"""

    COMPARISON = "COMPARISON"
    FORMAT_CHECK = "FORMAT_CHECK"


@dataclass(frozen=True)
class ValidationIssue:
    """一条统一核验错误，完整保留定位与实际/期望信息。"""

    location: CellLocation
    error_type: ValidationErrorType
    error_code: str
    actual_value: str
    reason: str
    source: ValidationSource
    expected_value: str = ""


@dataclass(frozen=True)
class ValidationSummary:
    """统一核验统计。"""

    total_checks: int
    pass_count: int
    error_count: int
    format_error_count: int = 0
    data_not_found_count: int = 0
    data_error_count: int = 0
    skipped_empty_row_count: int = 0


@dataclass(frozen=True)
class ValidationResult:
    """统一核验结果，可承载多个 Sheet 的 ValidationIssue。"""

    file: Path
    summary: ValidationSummary
    issues: tuple[ValidationIssue, ...] = ()
