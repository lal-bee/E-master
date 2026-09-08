"""P1-04 数据清洗引擎：按标准字段规则清洗内存中的业务数据。

只处理已明确映射（MATCH）的字段；UNMATCHED/AMBIGUOUS/CONFLICT
保留原值并在审计中说明。清洗过程不修改任何输入对象，也不写 Excel。
"""

from __future__ import annotations

import json
import math
import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from enum import Enum
from pathlib import Path
from typing import Any

from excel_qc.field_mapping import (
    FieldMappingEntry,
    FieldMappingResult,
    FieldMappingStatus,
    FieldMappingSummary,
    WorkbookFieldMappingResult,
)
from excel_qc.loader import LoadedWorkbook, load_workbook
from excel_qc.profiler import SheetProfile, WorkbookProfile

_DEFAULT_DATE_INPUT_FORMATS = (
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%Y.%m.%d",
    "%Y年%m月%d日",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
)
_INVISIBLE_WHITESPACE_PATTERN = re.compile(r"[\u200b-\u200d\u2060\ufeff]+")
_CRLF_PATTERN = re.compile(r"\r\n|\r")


class CleaningError(Exception):
    """数据清洗领域异常基类。"""


class CleaningConfigError(CleaningError):
    """清洗规则配置不合法。"""


class CleaningDataError(CleaningError):
    """清洗输入数据与映射/结构不一致。"""


class CleaningKind(str, Enum):
    """字段清洗类型。"""

    TEXT = "text"
    DATE = "date"
    NUMBER = "number"


class CleaningAction(str, Enum):
    """单个单元格清洗动作。"""

    CLEANED = "CLEANED"
    UNCHANGED = "UNCHANGED"
    EMPTY_NORMALIZED = "EMPTY_NORMALIZED"
    INVALID = "INVALID"
    NO_RULE = "NO_RULE"
    UNMAPPED = "UNMAPPED"
    AMBIGUOUS = "AMBIGUOUS"
    CONFLICT = "CONFLICT"


@dataclass(frozen=True)
class FieldCleaningRule:
    """一个标准字段的确定性清洗规则。"""

    standard_field: str
    kind: CleaningKind = CleaningKind.TEXT
    full_width_to_half_width: bool = True
    collapse_whitespace: bool = True
    remove_characters: tuple[str, ...] = ()
    empty_markers: tuple[str, ...] = ()
    empty_replacement: str = ""
    date_input_formats: tuple[str, ...] = ()
    date_output_format: str | None = None
    number_units: tuple[str, ...] = ()
    strip_thousands_separator: bool = True
    rule_code: str | None = None


@dataclass(frozen=True)
class CleaningWorkbookConfig:
    """P1-04 清洗配置：按标准字段组织规则。"""

    rules: tuple[FieldCleaningRule, ...] = ()


@dataclass(frozen=True)
class CleanCellAudit:
    """单个单元格的清洗审计记录。"""

    sheet_name: str
    row_number: int
    source_column_number: int
    source_column_letter: str
    source_field_name: str
    source_display_name: str
    standard_field_name: str | None
    original_value: Any
    cleaned_value: Any
    action: CleaningAction
    rule_code: str | None
    reason: str


@dataclass(frozen=True)
class CleanedRow:
    """一行清洗后的数据及逐单元格审计。"""

    source_row_number: int
    values: tuple[Any, ...]
    cells: tuple[CleanCellAudit, ...]


@dataclass(frozen=True)
class CleaningSummary:
    """清洗结果统计。"""

    row_count: int = 0
    cell_count: int = 0
    cleaned_count: int = 0
    unchanged_count: int = 0
    empty_normalized_count: int = 0
    invalid_count: int = 0
    unresolved_count: int = 0
    no_rule_count: int = 0


@dataclass(frozen=True)
class CleanedSheetResult:
    """一个 Sheet 的清洗结果。"""

    sheet_name: str
    rows: tuple[CleanedRow, ...]
    summary: CleaningSummary

    @property
    def audits(self) -> tuple[CleanCellAudit, ...]:
        return tuple(cell for row in self.rows for cell in row.cells)


@dataclass(frozen=True)
class CleaningWorkbookResult:
    """整个工作簿的清洗结果。"""

    source_file: Path
    sheet_results: tuple[CleanedSheetResult, ...]
    summary: CleaningSummary


def load_cleaning_config(path: str | Path) -> CleaningWorkbookConfig:
    """从 JSON 文件加载清洗规则配置。

    JSON 结构：
    {
      "rules": [
        {
          "standard_field": "设备名称",
          "kind": "text",
          "empty_markers": ["无", "/"]
        }
      ]
    }
    """
    config_path = Path(path)
    if not config_path.exists() or not config_path.is_file():
        raise CleaningConfigError(f"清洗规则配置文件不存在: {config_path}")
    try:
        payload = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise CleaningConfigError(
            f"清洗规则配置文件无法解析: {config_path}：{exc}"
        ) from exc
    if not isinstance(payload, dict):
        raise CleaningConfigError("清洗规则配置顶层必须是 JSON 对象")
    raw_rules = payload.get("rules")
    if not isinstance(raw_rules, list):
        raise CleaningConfigError("清洗规则配置缺少 rules 列表")

    rules: list[FieldCleaningRule] = []
    for index, raw_rule in enumerate(raw_rules, start=1):
        if not isinstance(raw_rule, dict):
            raise CleaningConfigError(f"rules 第 {index} 项必须是对象")
        standard_field = raw_rule.get("standard_field")
        if not isinstance(standard_field, str) or not standard_field.strip():
            raise CleaningConfigError(
                f"rules 第 {index} 项 standard_field 必须是非空字符串"
            )
        kind_text = raw_rule.get("kind", "text")
        if kind_text not in {item.value for item in CleaningKind}:
            raise CleaningConfigError(
                f"rules 第 {index} 项 kind 必须是 text/date/number"
            )
        kind = CleaningKind(kind_text)
        date_output_format = raw_rule.get("date_output_format")
        if kind is CleaningKind.DATE and (
            not isinstance(date_output_format, str)
            or not date_output_format
        ):
            raise CleaningConfigError(
                f"rules 第 {index} 项 date 类型必须提供 date_output_format"
            )
        rule = FieldCleaningRule(
            standard_field=standard_field.strip(),
            kind=kind,
            full_width_to_half_width=bool(
                raw_rule.get("full_width_to_half_width", True)
            ),
            collapse_whitespace=bool(raw_rule.get("collapse_whitespace", True)),
            remove_characters=_string_tuple(
                raw_rule.get("remove_characters", []),
                f"rules 第 {index} 项 remove_characters",
            ),
            empty_markers=_string_tuple(
                raw_rule.get("empty_markers", []),
                f"rules 第 {index} 项 empty_markers",
            ),
            empty_replacement=_optional_string(
                raw_rule.get("empty_replacement", ""),
                f"rules 第 {index} 项 empty_replacement",
            ),
            date_input_formats=_string_tuple(
                raw_rule.get("date_input_formats", []),
                f"rules 第 {index} 项 date_input_formats",
            ),
            date_output_format=date_output_format,
            number_units=_string_tuple(
                raw_rule.get("number_units", []),
                f"rules 第 {index} 项 number_units",
            ),
            strip_thousands_separator=bool(
                raw_rule.get("strip_thousands_separator", True)
            ),
            rule_code=_optional_string(
                raw_rule.get("rule_code", ""),
                f"rules 第 {index} 项 rule_code",
            ),
        )
        rules.append(rule)
    config = CleaningWorkbookConfig(tuple(rules))
    _validate_config(config)
    return config


def clean_workbook(
    profile: WorkbookProfile,
    mapping: WorkbookFieldMappingResult,
    config: CleaningWorkbookConfig,
    *,
    raw: LoadedWorkbook | None = None,
) -> CleaningWorkbookResult:
    """按字段映射与清洗规则清洗整个工作簿数据。

    profile/mapping/raw 应来自同一源文件；raw 缺省时使用 loader
    重新只读加载 profile.path。不会修改任何输入对象。
    """
    _validate_config(config)
    _validate_mapping_shape(profile, mapping)
    source = raw if raw is not None else load_workbook(profile.path)
    mapping_by_sheet = {
        result.sheet_name: result for result in mapping.sheet_results
    }
    sheet_results: list[CleanedSheetResult] = []
    for sheet_profile in profile.sheets:
        sheet_mapping = mapping_by_sheet.get(sheet_profile.sheet_name)
        if sheet_mapping is None:
            sheet_mapping = FieldMappingResult(
                sheet_name=sheet_profile.sheet_name,
                entries=(),
                summary=FieldMappingSummary(),
            )
        raw_sheet = _find_raw_sheet(source, sheet_profile.sheet_name)
        sheet_results.append(
            _clean_sheet(
                sheet_profile=sheet_profile,
                raw_rows=raw_sheet.rows,
                mapping=sheet_mapping,
                config=config,
            )
        )
    merged = _merge_summaries(tuple(result.summary for result in sheet_results))
    return CleaningWorkbookResult(
        source_file=profile.path,
        sheet_results=tuple(sheet_results),
        summary=merged,
    )


def _clean_sheet(
    *,
    sheet_profile: SheetProfile,
    raw_rows: tuple[tuple[Any, ...], ...],
    mapping: FieldMappingResult,
    config: CleaningWorkbookConfig,
) -> CleanedSheetResult:
    if (
        sheet_profile.data_first_row is None
        or sheet_profile.data_last_row is None
        or sheet_profile.data_first_row > sheet_profile.data_last_row
    ):
        return CleanedSheetResult(
            sheet_name=sheet_profile.sheet_name,
            rows=(),
            summary=CleaningSummary(),
        )

    rules_by_field = {
        rule.standard_field: rule for rule in config.rules
    }
    rows: list[CleanedRow] = []
    for row_number in range(
        sheet_profile.data_first_row,
        sheet_profile.data_last_row + 1,
    ):
        row_index = row_number - 1
        raw_values = list(raw_rows[row_index]) if row_index < len(raw_rows) else []
        cleaned_values = list(raw_values)
        audits: list[CleanCellAudit] = []
        for entry in mapping.entries:
            column_index = entry.source_column_number - 1
            if column_index < 0:
                continue
            original = (
                raw_values[column_index]
                if column_index < len(raw_values)
                else None
            )
            audit = _clean_mapped_cell(
                sheet_name=sheet_profile.sheet_name,
                row_number=row_number,
                entry=entry,
                original=original,
                rules_by_field=rules_by_field,
            )
            if column_index < len(cleaned_values):
                cleaned_values[column_index] = audit.cleaned_value
            audits.append(audit)
        rows.append(
            CleanedRow(
                source_row_number=row_number,
                values=tuple(cleaned_values),
                cells=tuple(audits),
            )
        )

    return CleanedSheetResult(
        sheet_name=sheet_profile.sheet_name,
        rows=tuple(rows),
        summary=_summary_from_rows(tuple(rows)),
    )


def _clean_mapped_cell(
    *,
    sheet_name: str,
    row_number: int,
    entry: FieldMappingEntry,
    original: Any,
    rules_by_field: dict[str, FieldCleaningRule],
) -> CleanCellAudit:
    if entry.status is FieldMappingStatus.UNMATCHED:
        return _audit(
            sheet_name,
            row_number,
            entry,
            original,
            original,
            CleaningAction.UNMAPPED,
            None,
            "字段未映射到标准字段，保留原值且不执行清洗",
        )
    if entry.status is FieldMappingStatus.AMBIGUOUS:
        return _audit(
            sheet_name,
            row_number,
            entry,
            original,
            original,
            CleaningAction.AMBIGUOUS,
            None,
            "字段映射存在歧义，保留原值且不执行清洗",
        )
    if entry.status is FieldMappingStatus.CONFLICT:
        return _audit(
            sheet_name,
            row_number,
            entry,
            original,
            original,
            CleaningAction.CONFLICT,
            None,
            "字段映射存在冲突，保留原值且不执行清洗",
        )

    rule = rules_by_field.get(entry.standard_field_name or "")
    if rule is None:
        return _audit(
            sheet_name,
            row_number,
            entry,
            original,
            original,
            CleaningAction.NO_RULE,
            None,
            f"标准字段“{entry.standard_field_name}”没有配置清洗规则，保留原值",
        )
    return _clean_value_with_rule(
        sheet_name=sheet_name,
        row_number=row_number,
        entry=entry,
        original=original,
        rule=rule,
    )


def _clean_value_with_rule(
    *,
    sheet_name: str,
    row_number: int,
    entry: FieldMappingEntry,
    original: Any,
    rule: FieldCleaningRule,
) -> CleanCellAudit:
    if original is None:
        return _audit(
            sheet_name,
            row_number,
            entry,
            original,
            original,
            CleaningAction.UNCHANGED,
            rule.rule_code,
            "空值保持不变",
        )

    text_value = _as_normalized_text(original, rule)
    if rule.empty_markers:
        for marker in rule.empty_markers:
            if _normalize_marker(marker) == text_value:
                cleaned = rule.empty_replacement
                return _audit(
                    sheet_name,
                    row_number,
                    entry,
                    original,
                    cleaned,
                    CleaningAction.EMPTY_NORMALIZED,
                    rule.rule_code,
                    f"空值占位符“{original}”按字段规则归一化为空",
                )

    if rule.kind is CleaningKind.DATE:
        return _clean_date_cell(
            sheet_name, row_number, entry, original, rule
        )
    if rule.kind is CleaningKind.NUMBER:
        return _clean_number_cell(
            sheet_name, row_number, entry, original, rule
        )
    return _clean_text_cell(
        sheet_name, row_number, entry, original, rule
    )


def _clean_text_cell(
    sheet_name: str,
    row_number: int,
    entry: FieldMappingEntry,
    original: Any,
    rule: FieldCleaningRule,
) -> CleanCellAudit:
    if not isinstance(original, str):
        return _audit(
            sheet_name,
            row_number,
            entry,
            original,
            original,
            CleaningAction.UNCHANGED,
            rule.rule_code,
            "非文本值保持不变",
        )
    cleaned = _normalize_text(original, rule)
    action = (
        CleaningAction.UNCHANGED
        if cleaned == original
        else CleaningAction.CLEANED
    )
    reason = "文本标准化" if action is CleaningAction.CLEANED else "文本无需清洗"
    return _audit(
        sheet_name,
        row_number,
        entry,
        original,
        cleaned,
        action,
        rule.rule_code,
        reason,
    )


def _clean_date_cell(
    sheet_name: str,
    row_number: int,
    entry: FieldMappingEntry,
    original: Any,
    rule: FieldCleaningRule,
) -> CleanCellAudit:
    output_format = rule.date_output_format or "%Y-%m-%d"
    parsed: datetime | None = None
    if isinstance(original, datetime):
        parsed = original
    elif isinstance(original, date):
        parsed = datetime(original.year, original.month, original.day)
    elif isinstance(original, str):
        text = _normalize_text(original, rule)
        formats = rule.date_input_formats or _DEFAULT_DATE_INPUT_FORMATS
        for fmt in formats:
            try:
                parsed = datetime.strptime(text, fmt)
                break
            except ValueError:
                continue
    if parsed is None:
        return _audit(
            sheet_name,
            row_number,
            entry,
            original,
            original,
            CleaningAction.INVALID,
            rule.rule_code,
            "无法识别的日期值，保留原值",
        )
    cleaned = parsed.strftime(output_format)
    action = (
        CleaningAction.UNCHANGED
        if cleaned == str(original)
        else CleaningAction.CLEANED
    )
    return _audit(
        sheet_name,
        row_number,
        entry,
        original,
        cleaned,
        action,
        rule.rule_code,
        "日期转换完成" if action is CleaningAction.CLEANED else "日期无需转换",
    )


def _clean_number_cell(
    sheet_name: str,
    row_number: int,
    entry: FieldMappingEntry,
    original: Any,
    rule: FieldCleaningRule,
) -> CleanCellAudit:
    try:
        cleaned = _normalize_number(original, rule)
    except CleaningError:
        return _audit(
            sheet_name,
            row_number,
            entry,
            original,
            original,
            CleaningAction.INVALID,
            rule.rule_code,
            "无法确定的数值/单位，保留原值",
        )
    action = (
        CleaningAction.UNCHANGED
        if cleaned == str(original)
        else CleaningAction.CLEANED
    )
    reason = "数值标准化完成" if action is CleaningAction.CLEANED else "数值无需清洗"
    return _audit(
        sheet_name,
        row_number,
        entry,
        original,
        cleaned,
        action,
        rule.rule_code,
        reason,
    )


def _normalize_number(value: Any, rule: FieldCleaningRule) -> str:
    if isinstance(value, bool):
        raise CleaningError("布尔值不能按数值清洗")
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if not math.isfinite(value):
            raise CleaningError("非有限浮点数")
        return _format_decimal(Decimal(str(value)))
    if isinstance(value, Decimal):
        return _format_decimal(value)
    if not isinstance(value, str):
        raise CleaningError("不支持的值类型")

    text = _normalize_text(value, rule)
    if not text:
        raise CleaningError("空字符串不是数值")
    if rule.number_units:
        for unit in sorted(rule.number_units, key=len, reverse=True):
            if text.endswith(unit) and len(text) > len(unit):
                text = text[: -len(unit)].strip()
                break
        else:
            if any(char.isalpha() for char in text):
                raise CleaningError("文本包含未配置单位")
    elif any(char.isalpha() for char in text):
        raise CleaningError("文本包含未配置单位")

    if rule.strip_thousands_separator:
        if re.fullmatch(r"[+-]?\d{1,3}(,\d{3})+(\.\d+)?", text):
            text = text.replace(",", "")
    text = text.replace(" ", "")
    try:
        number = Decimal(text)
    except InvalidOperation as exc:
        raise CleaningError("无法解析数值") from exc
    return _format_decimal(number)


def _format_decimal(value: Decimal) -> str:
    normalized = value.normalize()
    if normalized == normalized.to_integral_value():
        return format(normalized.quantize(Decimal(1)), "f")
    return format(normalized, "f")


def _as_normalized_text(value: Any, rule: FieldCleaningRule) -> str:
    if isinstance(value, str):
        return _normalize_text(value, rule).casefold()
    if isinstance(value, bool):
        return "true" if value else "false"
    if value is None:
        return ""
    return _normalize_text(str(value), rule).casefold()


def _normalize_text(value: str, rule: FieldCleaningRule) -> str:
    text = value
    if rule.full_width_to_half_width:
        text = unicodedata.normalize("NFKC", text)
    text = _INVISIBLE_WHITESPACE_PATTERN.sub("", text)
    text = _CRLF_PATTERN.sub("\n", text)
    if rule.collapse_whitespace:
        text = " ".join(text.split())
    else:
        text = text.strip()
    for character in rule.remove_characters:
        text = text.replace(character, "")
    return text


def _normalize_marker(value: str) -> str:
    return _normalize_text(
        value, FieldCleaningRule(standard_field="__marker__")
    ).casefold()


def _audit(
    sheet_name: str,
    row_number: int,
    entry: FieldMappingEntry,
    original: Any,
    cleaned: Any,
    action: CleaningAction,
    rule_code: str | None,
    reason: str,
) -> CleanCellAudit:
    return CleanCellAudit(
        sheet_name=sheet_name,
        row_number=row_number,
        source_column_number=entry.source_column_number,
        source_column_letter=entry.source_column_letter,
        source_field_name=entry.source_field_name,
        source_display_name=entry.source_display_name,
        standard_field_name=entry.standard_field_name,
        original_value=original,
        cleaned_value=cleaned,
        action=action,
        rule_code=rule_code,
        reason=reason,
    )


def _validate_config(config: CleaningWorkbookConfig) -> None:
    if not isinstance(config, CleaningWorkbookConfig):
        raise CleaningConfigError("config 必须是 CleaningWorkbookConfig")
    seen: set[str] = set()
    for rule in config.rules:
        if not isinstance(rule, FieldCleaningRule):
            raise CleaningConfigError("rules 只允许 FieldCleaningRule")
        if not isinstance(rule.standard_field, str):
            raise CleaningConfigError("清洗规则 standard_field 必须是字符串")
        field = rule.standard_field.strip()
        if not field:
            raise CleaningConfigError("清洗规则 standard_field 不能为空")
        if field in seen:
            raise CleaningConfigError(f"清洗规则标准字段重复: {field}")
        seen.add(field)
        if rule.kind is CleaningKind.DATE and not rule.date_output_format:
            raise CleaningConfigError(
                f"标准字段“{field}”的 date 清洗规则缺少 date_output_format"
            )
        if not isinstance(rule.date_output_format, (str, type(None))):
            raise CleaningConfigError(
                f"标准字段“{field}”的 date_output_format 必须是字符串"
            )
        _validate_string_sequence(
            rule.remove_characters, f"标准字段“{field}”的 remove_characters"
        )
        _validate_string_sequence(
            rule.empty_markers, f"标准字段“{field}”的 empty_markers"
        )
        _validate_string_sequence(
            rule.date_input_formats, f"标准字段“{field}”的 date_input_formats"
        )
        _validate_string_sequence(
            rule.number_units, f"标准字段“{field}”的 number_units"
        )


def _validate_string_sequence(value: Any, label: str) -> None:
    if not isinstance(value, (tuple, list)) or not all(
        isinstance(item, str) for item in value
    ):
        raise CleaningConfigError(f"{label} 必须是字符串元组/列表")


def _validate_mapping_shape(
    profile: WorkbookProfile,
    mapping: WorkbookFieldMappingResult,
) -> None:
    profile_names = tuple(sheet.sheet_name for sheet in profile.sheets)
    mapping_names = tuple(result.sheet_name for result in mapping.sheet_results)
    if profile_names != mapping_names:
        raise CleaningDataError(
            "profile 与 mapping 的 Sheet 顺序/名称不一致"
        )


def _find_raw_sheet(
    raw: LoadedWorkbook,
    sheet_name: str,
):
    for sheet in raw.sheets:
        if sheet.sheet_name == sheet_name:
            return sheet
    raise CleaningDataError(f"raw 中找不到 Sheet“{sheet_name}”")


def _summary_from_rows(rows: tuple[CleanedRow, ...]) -> CleaningSummary:
    cells = tuple(cell for row in rows for cell in row.cells)
    counts = {action: 0 for action in CleaningAction}
    for cell in cells:
        counts[cell.action] += 1
    return CleaningSummary(
        row_count=len(rows),
        cell_count=len(cells),
        cleaned_count=counts[CleaningAction.CLEANED],
        unchanged_count=counts[CleaningAction.UNCHANGED],
        empty_normalized_count=counts[CleaningAction.EMPTY_NORMALIZED],
        invalid_count=counts[CleaningAction.INVALID],
        unresolved_count=(
            counts[CleaningAction.UNMAPPED]
            + counts[CleaningAction.AMBIGUOUS]
            + counts[CleaningAction.CONFLICT]
        ),
        no_rule_count=counts[CleaningAction.NO_RULE],
    )


def _merge_summaries(
    summaries: tuple[CleaningSummary, ...],
) -> CleaningSummary:
    return CleaningSummary(
        row_count=sum(item.row_count for item in summaries),
        cell_count=sum(item.cell_count for item in summaries),
        cleaned_count=sum(item.cleaned_count for item in summaries),
        unchanged_count=sum(item.unchanged_count for item in summaries),
        empty_normalized_count=sum(
            item.empty_normalized_count for item in summaries
        ),
        invalid_count=sum(item.invalid_count for item in summaries),
        unresolved_count=sum(item.unresolved_count for item in summaries),
        no_rule_count=sum(item.no_rule_count for item in summaries),
    )


def _string_tuple(value: Any, label: str) -> tuple[str, ...]:
    if not isinstance(value, list) or not all(
        isinstance(item, str) for item in value
    ):
        raise CleaningConfigError(f"{label} 必须是字符串列表")
    result = tuple(item.strip() for item in value if item.strip())
    return result


def _optional_string(value: Any, label: str) -> str | None:
    if value is None or value == "":
        return None if value is None else ""
    if not isinstance(value, str):
        raise CleaningConfigError(f"{label} 必须是字符串")
    return value
