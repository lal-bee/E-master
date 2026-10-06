"""P1-08 系统导出 Excel 接入：将第二数据源标准化并保留来源追溯。"""

from __future__ import annotations

import json
from collections import Counter
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any

from excel_qc.cleaning import (
    CleaningAction,
    CleaningConfigError,
    CleaningWorkbookConfig,
    CleaningWorkbookResult,
    FieldCleaningRule,
    clean_workbook,
)
from excel_qc.coordinates import column_letter
from excel_qc.field_mapping import (
    FieldMappingConfig,
    FieldMappingConfigError,
    FieldMappingEntry,
    FieldMappingStatus,
    WorkbookFieldMappingResult,
    map_workbook_fields,
)
from excel_qc.loader import LoadedSheet, load_workbook
from excel_qc.models import BasicDataType
from excel_qc.profiler import (
    FieldProfile,
    ProfileIssue,
    ProfileIssueLevel,
    SheetProfile,
    WorkbookProfile,
    profile_workbook,
)
from excel_qc.text import is_blank_value


class SystemExportError(Exception):
    """系统导出接入领域异常。"""


class SystemExportConfigError(SystemExportError):
    """系统导出接入配置不合法。"""


class SystemExportDataError(SystemExportError):
    """系统导出工作簿与指定配置不一致。"""


@dataclass(frozen=True)
class SystemExportConfig:
    """系统导出接入配置。Sheet 名和表头行为 Excel 物理坐标。"""

    sheet_name: str
    header_row_number: int
    field_mapping: FieldMappingConfig
    cleaning: CleaningWorkbookConfig = CleaningWorkbookConfig()
    primary_key_fields: tuple[str, ...] = ()


@dataclass(frozen=True)
class SystemExportIssue:
    """系统导入结果中的问题；行号/列号为空表示工作簿级问题。"""

    code: str
    message: str
    sheet_name: str | None = None
    row_number: int | None = None
    column_number: int | None = None
    column_letter: str | None = None
    standard_field_name: str | None = None


@dataclass(frozen=True)
class SystemExportFieldValue:
    """一项标准字段值及其系统导出源单元格追溯。"""

    source_field_name: str
    source_column_number: int
    source_column_letter: str
    standard_field_name: str | None
    original_value: Any
    normalized_value: Any
    applied_rules: tuple[str, ...] = ()
    mapping_status: FieldMappingStatus = FieldMappingStatus.UNMATCHED


@dataclass(frozen=True)
class SystemExportRow:
    """保留系统导出物理行坐标和标准字段数据的一行记录。"""

    source_file: Path
    sheet_name: str
    source_row_number: int
    fields: tuple[SystemExportFieldValue, ...]
    primary_key_fields: tuple[str, ...]
    primary_key_values: tuple[Any, ...]

    @property
    def values(self) -> dict[str, Any]:
        """返回已唯一映射标准字段的标准化值。"""
        return {
            field.standard_field_name: field.normalized_value
            for field in self.fields
            if field.standard_field_name is not None
            and field.mapping_status is FieldMappingStatus.MATCH
        }

    @property
    def primary_key_value(self) -> Any:
        if len(self.primary_key_values) == 1:
            return self.primary_key_values[0]
        return self.primary_key_values


@dataclass(frozen=True)
class SystemExportSummary:
    row_count: int = 0
    issue_count: int = 0
    error_count: int = 0
    warning_count: int = 0
    issue_counts: tuple[tuple[str, int], ...] = ()


@dataclass(frozen=True)
class SystemExportWorkbookResult:
    """供后续核验消费的系统导出数据、来源、问题与汇总。"""

    source_file: Path
    sheet_name: str
    header_row_number: int
    rows: tuple[SystemExportRow, ...]
    issues: tuple[SystemExportIssue, ...]
    summary: SystemExportSummary
    mapping_result: WorkbookFieldMappingResult
    cleaning_result: CleaningWorkbookResult


def load_system_export_config(path: str | Path) -> SystemExportConfig:
    """读取系统导出接入 JSON 配置。

    示例 JSON 结构见 examples/system_export_demo.json；示例仅用于演示。
    """
    config_path = Path(path)
    if not config_path.is_file():
        raise SystemExportConfigError(f"系统导出配置文件不存在: {config_path}")
    try:
        payload = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise SystemExportConfigError(
            f"系统导出配置文件无法解析: {config_path}：{exc}"
        ) from exc
    if not isinstance(payload, dict):
        raise SystemExportConfigError("系统导出配置顶层必须是 JSON 对象")
    sheet_name = payload.get("sheet_name")
    header_row = payload.get("header_row_number")
    if not isinstance(sheet_name, str) or not sheet_name.strip():
        raise SystemExportConfigError("sheet_name 必须是非空字符串")
    if isinstance(header_row, bool) or not isinstance(header_row, int) or header_row < 1:
        raise SystemExportConfigError("header_row_number 必须是大于 0 的整数")

    raw_fields = payload.get("standard_fields")
    if not isinstance(raw_fields, list):
        raise SystemExportConfigError("standard_fields 必须是列表")
    try:
        fields = tuple(
            _standard_field(item, index)
            for index, item in enumerate(raw_fields, start=1)
        )
        raw_cleaning = payload.get("cleaning_rules", [])
        if not isinstance(raw_cleaning, list):
            raise SystemExportConfigError("cleaning_rules 必须是列表")
        cleaning_rules = tuple(
            _cleaning_rule(item, index)
            for index, item in enumerate(raw_cleaning, start=1)
        )
        key_fields = payload.get("primary_key_fields", [])
        if not isinstance(key_fields, list) or not all(
            isinstance(item, str) and item.strip() for item in key_fields
        ):
            raise SystemExportConfigError(
                "primary_key_fields 必须是非空字符串组成的列表"
            )
        config = SystemExportConfig(
            sheet_name=sheet_name.strip(),
            header_row_number=header_row,
            field_mapping=FieldMappingConfig(fields),
            cleaning=CleaningWorkbookConfig(cleaning_rules),
            primary_key_fields=tuple(item.strip() for item in key_fields),
        )
        _validate_config(config)
        return config
    except (FieldMappingConfigError, CleaningConfigError) as exc:
        raise SystemExportConfigError(str(exc)) from exc


def ingest_system_export(
    path: str | Path,
    config: SystemExportConfig,
) -> SystemExportWorkbookResult:
    """读取、映射并标准化系统导出 .xlsx，保留行和单元格问题。

    读取只经过现有 loader/profiler/field_mapping/cleaning 模块；不会写入
    输入文件。重复主键和清洗失败均保留记录并产生问题。
    """
    _validate_config(config)
    source = load_workbook(path)
    raw_sheet = next(
        (sheet for sheet in source.sheets if sheet.sheet_name == config.sheet_name),
        None,
    )
    issues: list[SystemExportIssue] = []
    if raw_sheet is None:
        issues.append(
            SystemExportIssue(
                code="SHEET_NOT_FOUND",
                message=(
                    f"未找到配置的 Sheet“{config.sheet_name}”；"
                    f"可用 Sheet：{', '.join(source.sheet_names) or '无'}"
                ),
            )
        )
        empty_profile = WorkbookProfile(source.path, (), ())
        empty_mapping = map_workbook_fields(empty_profile, config.field_mapping)
        empty_cleaning = _clean_profile(
            empty_profile, empty_mapping, config.cleaning, source
        )
        return _result(source, config, (), tuple(issues), empty_mapping, empty_cleaning)

    profile = profile_workbook(path)
    configured_sheet = _configured_sheet_profile(
        raw_sheet, profile, config.header_row_number
    )
    configured_profile = WorkbookProfile(
        path=source.path,
        sheets=(configured_sheet,),
        issues=configured_sheet.issues,
    )
    mapping = map_workbook_fields(configured_profile, config.field_mapping)
    mapping_sheet = mapping.sheet_results[0]
    cleaning = _clean_profile(configured_profile, mapping, config.cleaning, source)
    cleaned_sheet = cleaning.sheet_results[0]

    issues.extend(_mapping_issues(mapping_sheet, config))
    rows: list[SystemExportRow] = []
    for cleaned_row in cleaned_sheet.rows:
        fields: list[SystemExportFieldValue] = []
        for audit in cleaned_row.cells:
            entry = _entry_for_audit(mapping_sheet.entries, audit.source_column_number)
            fields.append(
                SystemExportFieldValue(
                    source_field_name=audit.source_field_name,
                    source_column_number=audit.source_column_number,
                    source_column_letter=audit.source_column_letter,
                    standard_field_name=audit.standard_field_name,
                    original_value=audit.original_value,
                    normalized_value=audit.cleaned_value,
                    applied_rules=_applied_rules(audit),
                    mapping_status=entry.status if entry else FieldMappingStatus.UNMATCHED,
                )
            )
            if audit.action is CleaningAction.INVALID:
                issues.append(
                    SystemExportIssue(
                        code="CLEANING_FAILED",
                        message=audit.reason,
                        sheet_name=config.sheet_name,
                        row_number=audit.row_number,
                        column_number=audit.source_column_number,
                        column_letter=audit.source_column_letter,
                        standard_field_name=audit.standard_field_name,
                    )
                )
        values_by_name = {
            item.standard_field_name: item.normalized_value
            for item in fields
            if item.standard_field_name is not None
            and item.mapping_status is FieldMappingStatus.MATCH
        }
        key_values = tuple(values_by_name.get(name) for name in config.primary_key_fields)
        row = SystemExportRow(
            source_file=source.path,
            sheet_name=config.sheet_name,
            source_row_number=cleaned_row.source_row_number,
            fields=tuple(fields),
            primary_key_fields=config.primary_key_fields,
            primary_key_values=key_values,
        )
        rows.append(row)
        for field_name, value in zip(config.primary_key_fields, key_values):
            if is_blank_value(value):
                field_trace = next(
                    (item for item in fields if item.standard_field_name == field_name),
                    None,
                )
                issues.append(
                    SystemExportIssue(
                        code="EMPTY_PRIMARY_KEY",
                        message=f"主键字段“{field_name}”为空",
                        sheet_name=config.sheet_name,
                        row_number=row.source_row_number,
                        column_number=field_trace.source_column_number if field_trace else None,
                        column_letter=field_trace.source_column_letter if field_trace else None,
                        standard_field_name=field_name,
                    )
                )

    issues.extend(_duplicate_key_issues(tuple(rows), config.primary_key_fields, config.sheet_name))
    return _result(source, config, tuple(rows), tuple(issues), mapping, cleaning)


def _configured_sheet_profile(
    raw_sheet: LoadedSheet,
    detected_profile: WorkbookProfile,
    header_row_number: int,
) -> SheetProfile:
    if header_row_number > len(raw_sheet.rows):
        raise SystemExportDataError(
            f"配置的表头行 {header_row_number} 超出 Sheet“{raw_sheet.sheet_name}”范围"
        )
    base = next(
        (item for item in detected_profile.sheets if item.sheet_name == raw_sheet.sheet_name),
        None,
    )
    header = raw_sheet.rows[header_row_number - 1]
    last_row = max(
        (
            header_row_number + index + 1
            for index, row in enumerate(raw_sheet.rows[header_row_number:])
            if any(not is_blank_value(value) for value in row)
        ),
        default=header_row_number,
    )
    data_first = header_row_number + 1
    data_last = last_row if last_row >= data_first else None
    width = max(len(header), max((len(row) for row in raw_sheet.rows), default=0))
    old_fields = {field.column_index: field for field in (base.fields if base else ())}
    fields: list[FieldProfile] = []
    header_issues: list[ProfileIssue] = []
    names: Counter[str] = Counter()
    for col_index in range(width):
        value = header[col_index] if col_index < len(header) else None
        name = "" if is_blank_value(value) else str(value).strip()
        names[name.strip().casefold()] += bool(name)
        old = old_fields.get(col_index)
        fields.append(
            replace(old, name=name)
            if old is not None
            else FieldProfile(
                name=name,
                column_index=col_index,
                column_number=col_index + 1,
                column_letter=column_letter(col_index),
                data_type=BasicDataType.OTHER,
                non_empty_count=0,
                empty_count=0,
                sample_values=(),
            )
        )
        if not name:
            header_issues.append(
                ProfileIssue(
                    code="EMPTY_HEADER_CELL",
                    message=f"第 {column_letter(col_index)} 列配置表头为空",
                    level=ProfileIssueLevel.WARNING,
                    sheet_name=raw_sheet.sheet_name,
                    row_number=header_row_number,
                    column_letter=column_letter(col_index),
                )
            )
    duplicates = sorted(name for name, count in names.items() if name and count > 1)
    if duplicates:
        header_issues.append(
            ProfileIssue(
                code="DUPLICATE_HEADER_NAME",
                message="配置表头行存在同名列：" + "、".join(duplicates),
                level=ProfileIssueLevel.WARNING,
                sheet_name=raw_sheet.sheet_name,
                row_number=header_row_number,
            )
        )
    return SheetProfile(
        sheet_name=raw_sheet.sheet_name,
        sheet_index=base.sheet_index if base else raw_sheet.sheet_index,
        empty=False,
        total_row_count=len(raw_sheet.rows),
        total_column_count=width,
        used_first_row=header_row_number,
        used_last_row=last_row,
        used_first_col=1 if width else None,
        used_last_col=width if width else None,
        header_candidates=base.header_candidates if base else (),
        header_start_row=header_row_number,
        header_end_row=header_row_number,
        header_row=header_row_number,
        header_row_count=1,
        data_first_row=data_first if data_last is not None else None,
        data_last_row=data_last,
        data_first_col=1 if width else None,
        data_last_col=width if width else None,
        data_row_count=max(0, (data_last or header_row_number) - header_row_number),
        fields=tuple(fields),
        issues=tuple(header_issues),
    )


def _mapping_issues(mapping, config: SystemExportConfig) -> list[SystemExportIssue]:
    issues: list[SystemExportIssue] = []
    matched_fields = {
        entry.standard_field_name
        for entry in mapping.entries
        if entry.status is FieldMappingStatus.MATCH
    }
    for definition in config.field_mapping.standard_fields:
        if definition.name not in matched_fields:
            issues.append(
                SystemExportIssue(
                    code="MISSING_FIELD",
                    message=f"未能从系统导出表头唯一确定标准字段“{definition.name}”",
                    sheet_name=config.sheet_name,
                    row_number=config.header_row_number,
                    standard_field_name=definition.name,
                )
            )
    for entry in mapping.entries:
        if entry.status is FieldMappingStatus.MATCH:
            continue
        issues.append(
            SystemExportIssue(
                code=f"FIELD_MAPPING_{entry.status.value}",
                message=entry.reason,
                sheet_name=entry.source_sheet_name,
                row_number=config.header_row_number,
                column_number=entry.source_column_number,
                column_letter=entry.source_column_letter,
                standard_field_name=entry.standard_field_name,
            )
        )
    return issues


def _duplicate_key_issues(
    rows: tuple[SystemExportRow, ...],
    key_fields: tuple[str, ...],
    sheet_name: str,
) -> list[SystemExportIssue]:
    if not key_fields:
        return []
    seen: dict[tuple[tuple[type, Any], ...], int] = {}
    issues: list[SystemExportIssue] = []
    for row in rows:
        if any(is_blank_value(value) for value in row.primary_key_values):
            continue
        key = tuple((type(value), value) for value in row.primary_key_values)
        first = seen.get(key)
        if first is None:
            seen[key] = row.source_row_number
            continue
        issues.append(
            SystemExportIssue(
                code="DUPLICATE_PRIMARY_KEY",
                message=f"主键 {row.primary_key_values!r} 与第 {first} 行重复；两行均保留",
                sheet_name=sheet_name,
                row_number=row.source_row_number,
                standard_field_name=key_fields[0],
            )
        )
    return issues


def _entry_for_audit(entries: tuple[FieldMappingEntry, ...], column_number: int):
    return next((entry for entry in entries if entry.source_column_number == column_number), None)


def _applied_rules(audit) -> tuple[str, ...]:
    if audit.action in {
        CleaningAction.NO_RULE,
        CleaningAction.UNMAPPED,
        CleaningAction.AMBIGUOUS,
        CleaningAction.CONFLICT,
    }:
        return ()
    return (audit.rule_code or f"CLEANING_RULE:{audit.standard_field_name}",)


def _result(source, config, rows, issues, mapping, cleaning):
    counts = Counter(issue.code for issue in issues)
    summary = SystemExportSummary(
        row_count=len(rows),
        issue_count=len(issues),
        error_count=sum(issue.code in {"MISSING_FIELD", "EMPTY_PRIMARY_KEY", "DUPLICATE_PRIMARY_KEY", "CLEANING_FAILED"} for issue in issues),
        warning_count=sum(issue.code not in {"MISSING_FIELD", "EMPTY_PRIMARY_KEY", "DUPLICATE_PRIMARY_KEY", "CLEANING_FAILED"} for issue in issues),
        issue_counts=tuple(sorted(counts.items())),
    )
    return SystemExportWorkbookResult(
        source_file=source.path,
        sheet_name=config.sheet_name,
        header_row_number=config.header_row_number,
        rows=rows,
        issues=issues,
        summary=summary,
        mapping_result=mapping,
        cleaning_result=cleaning,
    )


def _clean_profile(profile, mapping, config, source):
    try:
        return clean_workbook(profile, mapping, config, raw=source)
    except CleaningConfigError as exc:
        raise SystemExportConfigError(str(exc)) from exc


def _standard_field(payload: Any, index: int):
    from excel_qc.field_mapping import StandardFieldDefinition

    if not isinstance(payload, dict):
        raise SystemExportConfigError(f"standard_fields 第 {index} 项必须是对象")
    name = payload.get("name")
    aliases = payload.get("aliases", [])
    if not isinstance(name, str) or not isinstance(aliases, list):
        raise SystemExportConfigError(f"standard_fields 第 {index} 项配置不合法")
    return StandardFieldDefinition(name=name, aliases=tuple(aliases))


def _cleaning_rule(payload: Any, index: int) -> FieldCleaningRule:
    from excel_qc.cleaning import CleaningKind

    if not isinstance(payload, dict):
        raise SystemExportConfigError(f"cleaning_rules 第 {index} 项必须是对象")
    try:
        values = dict(payload)
        values["kind"] = CleaningKind(values.get("kind", "text"))
        for name in (
            "remove_characters", "empty_markers", "date_input_formats", "number_units"
        ):
            if name in values:
                values[name] = tuple(values[name])
        return FieldCleaningRule(**values)
    except (TypeError, ValueError) as exc:
        raise SystemExportConfigError(f"cleaning_rules 第 {index} 项配置不合法：{exc}") from exc


def _validate_config(config: SystemExportConfig) -> None:
    if not isinstance(config, SystemExportConfig):
        raise SystemExportConfigError("config 必须是 SystemExportConfig")
    if not isinstance(config.sheet_name, str) or not config.sheet_name.strip():
        raise SystemExportConfigError("sheet_name 必须是非空字符串")
    if (
        isinstance(config.header_row_number, bool)
        or not isinstance(config.header_row_number, int)
        or config.header_row_number < 1
    ):
        raise SystemExportConfigError("header_row_number 必须是大于 0 的整数")
    if not isinstance(config.field_mapping, FieldMappingConfig):
        raise SystemExportConfigError("field_mapping 必须是 FieldMappingConfig")
    if not config.field_mapping.standard_fields:
        raise SystemExportConfigError("field_mapping 至少需要一个标准字段")
    if not isinstance(config.cleaning, CleaningWorkbookConfig):
        raise SystemExportConfigError("cleaning 必须是 CleaningWorkbookConfig")
    if (
        not isinstance(config.primary_key_fields, tuple)
        or not config.primary_key_fields
        or not all(isinstance(name, str) and name.strip() for name in config.primary_key_fields)
    ):
        raise SystemExportConfigError("primary_key_fields 必须配置一个或多个标准字段名称")
    defined = {field.name for field in config.field_mapping.standard_fields}
    unknown = sorted(set(config.primary_key_fields) - defined)
    if unknown:
        raise SystemExportConfigError("主键字段未定义：" + "、".join(unknown))
    try:
        from excel_qc.field_mapping import map_sheet_fields

        # Public engine performs complete field mapping configuration validation.
        if config.field_mapping.standard_fields:
            map_sheet_fields(
                SheetProfile("", 0, True, None, None, None, None, None, None),
                config.field_mapping,
            )
    except (FieldMappingConfigError, CleaningConfigError) as exc:
        raise SystemExportConfigError(str(exc)) from exc
