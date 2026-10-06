"""P1-09 原始源文档与系统导出数据的确定性回源核验。"""

from __future__ import annotations

import math
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, replace
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from enum import Enum
from pathlib import Path
from typing import Any

from excel_qc.cleaning import (
    CleanCellAudit,
    CleaningAction,
    CleaningConfigError,
    CleaningWorkbookConfig,
    CleaningWorkbookResult,
    FieldCleaningRule,
    clean_workbook,
)
from excel_qc.coordinates import cell_address, column_letter
from excel_qc.field_mapping import (
    FieldMappingConfig,
    FieldMappingConfigError,
    FieldMappingEntry,
    FieldMappingStatus,
    WorkbookFieldMappingResult,
    map_workbook_fields,
    normalize_field_name,
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
from excel_qc.system_export import (
    SystemExportConfig,
    SystemExportFieldValue,
    SystemExportIssue,
    SystemExportRow,
    SystemExportWorkbookResult,
    ingest_system_export,
)
from excel_qc.text import is_blank_value


class SourceVerificationError(Exception):
    """P1-09 回源核验异常基类。"""


class SourceVerificationConfigError(SourceVerificationError):
    """回源核验配置不合法。"""


class SourceVerificationDataError(SourceVerificationError):
    """原始源工作簿无法按配置读取。"""


class VerificationSide(str, Enum):
    ORIGINAL_SOURCE = "ORIGINAL_SOURCE"
    SYSTEM_EXPORT = "SYSTEM_EXPORT"
    BOTH = "BOTH"


class VerificationIssueCategory(str, Enum):
    INPUT = "INPUT"
    MATCHING = "MATCHING"
    COMPARISON = "COMPARISON"


class VerificationRowStatus(str, Enum):
    MATCH = "MATCH"
    MISSING_IN_SYSTEM = "MISSING_IN_SYSTEM"
    EXTRA_IN_SYSTEM = "EXTRA_IN_SYSTEM"
    FIELD_CHANGED = "FIELD_CHANGED"
    DUPLICATE_KEY = "DUPLICATE_KEY"
    UNVERIFIABLE = "UNVERIFIABLE"


class FieldComparisonStatus(str, Enum):
    EQUAL = "EQUAL"
    CHANGED = "CHANGED"
    UNVERIFIABLE = "UNVERIFIABLE"


class FieldValueStatus(str, Enum):
    PRESENT = "PRESENT"
    EMPTY = "EMPTY"
    FIELD_MISSING = "FIELD_MISSING"
    MAPPING_FAILED = "MAPPING_FAILED"
    CLEANING_FAILED = "CLEANING_FAILED"


class ComparisonRuleKind(str, Enum):
    """字段比较口径。转换只在显式选用 NUMBER/DATE 时执行。"""

    EXACT = "EXACT"
    TEXT = "TEXT"
    NUMBER = "NUMBER"
    BOOLEAN = "BOOLEAN"
    DATE = "DATE"


@dataclass(frozen=True)
class OriginalSourceConfig:
    """最初收集的原始源 Excel 配置，不会标记为系统导出。"""

    sheet_name: str
    header_row_number: int
    field_mapping: FieldMappingConfig
    cleaning: CleaningWorkbookConfig = CleaningWorkbookConfig()


@dataclass(frozen=True)
class FieldComparisonRule:
    standard_field: str
    kind: ComparisonRuleKind = ComparisonRuleKind.EXACT


@dataclass(frozen=True)
class SourceVerificationConfig:
    original_source: OriginalSourceConfig
    system_export: SystemExportConfig
    key_fields: tuple[str, ...]
    comparison_fields: tuple[str, ...]
    comparison_rules: tuple[FieldComparisonRule, ...] = ()


@dataclass(frozen=True)
class VerificationLocation:
    side: VerificationSide
    source_file: Path
    sheet_name: str
    row_number: int
    source_field_name: str | None = None
    standard_field_name: str | None = None
    column_number: int | None = None
    column_letter: str | None = None
    cell_address: str | None = None


@dataclass(frozen=True)
class VerificationIssue:
    category: VerificationIssueCategory
    code: str
    message: str
    side: VerificationSide
    standard_field_name: str | None = None
    key_values: tuple[Any, ...] = ()
    locations: tuple[VerificationLocation, ...] = ()
    blocks_verification: bool = True


@dataclass(frozen=True)
class NormalizedFieldValue:
    standard_field_name: str
    original_value: Any
    normalized_value: Any
    status: FieldValueStatus
    applied_rules: tuple[str, ...]
    location: VerificationLocation | None
    reason: str = ""


@dataclass(frozen=True)
class OriginalSourceRow:
    source_file: Path
    sheet_name: str
    source_row_number: int
    fields: tuple[NormalizedFieldValue, ...]

    @property
    def values(self) -> dict[str, Any]:
        return {
            value.standard_field_name: value.normalized_value
            for value in self.fields
            if value.status in {FieldValueStatus.PRESENT, FieldValueStatus.EMPTY}
        }


@dataclass(frozen=True)
class OriginalSourceWorkbookResult:
    source_file: Path
    sheet_name: str
    header_row_number: int
    rows: tuple[OriginalSourceRow, ...]
    issues: tuple[VerificationIssue, ...]
    mapping_result: WorkbookFieldMappingResult | None
    cleaning_result: CleaningWorkbookResult | None
    sheet_available: bool


@dataclass(frozen=True)
class FieldDifference:
    standard_field_name: str
    status: FieldComparisonStatus
    comparison_rule: ComparisonRuleKind
    original_source: NormalizedFieldValue
    system_export: NormalizedFieldValue
    reason: str


@dataclass(frozen=True)
class VerificationRecord:
    status: VerificationRowStatus
    key_fields: tuple[str, ...]
    key_values: tuple[Any, ...]
    original_source_rows: tuple[VerificationLocation, ...]
    system_export_rows: tuple[VerificationLocation, ...]
    field_differences: tuple[FieldDifference, ...] = ()


@dataclass(frozen=True)
class SourceVerificationSummary:
    original_source_input_row_count: int = 0
    system_export_input_row_count: int = 0
    matched_row_count: int = 0
    unchanged_row_count: int = 0
    changed_row_count: int = 0
    missing_in_system_row_count: int = 0
    extra_in_system_row_count: int = 0
    duplicate_key_group_count: int = 0
    unverifiable_row_count: int = 0
    compared_field_count: int = 0
    equal_field_count: int = 0
    changed_field_count: int = 0
    unverifiable_field_count: int = 0
    issue_count: int = 0
    issue_counts: tuple[tuple[str, int], ...] = ()
    verification_complete: bool = False
    data_consistent: bool = False


@dataclass(frozen=True)
class SourceVerificationResult:
    config: SourceVerificationConfig
    original_source: OriginalSourceWorkbookResult
    system_export: SystemExportWorkbookResult
    records: tuple[VerificationRecord, ...]
    issues: tuple[VerificationIssue, ...]
    summary: SourceVerificationSummary


@dataclass(frozen=True)
class _SideDataset:
    side: VerificationSide
    source_file: Path
    sheet_name: str
    rows: tuple[OriginalSourceRow | SystemExportRow, ...]
    mapping_entries: tuple[FieldMappingEntry, ...]
    issues: tuple[VerificationIssue, ...]
    sheet_available: bool
    failed_cells: frozenset[tuple[int, int]] = frozenset()


@dataclass(frozen=True)
class _KeyedRow:
    row: OriginalSourceRow | SystemExportRow
    key_values: tuple[Any, ...]
    signature: tuple[Any, ...]


_NUMBER_TEXT = re.compile(r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$")
_DATE_TEXT = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def verify_source_against_system_export(
    original_source_path: str | Path,
    system_export_path: str | Path,
    config: SourceVerificationConfig,
) -> SourceVerificationResult:
    """按唯一键将最初收集的源文件与 P1-08 系统导出结果进行核验。

    默认逐类型精确比较，不隐式清洗、转型或忽略大小写。任何不可判定字段、
    不唯一键或关键字段映射问题都会阻止“全部一致”的结论。
    """
    _validate_config(config)
    required_fields = tuple(dict.fromkeys((*config.key_fields, *config.comparison_fields)))
    original = _ingest_original_source(original_source_path, config.original_source, frozenset(required_fields))
    system = ingest_system_export(system_export_path, config.system_export)
    source_entries = (
        original.mapping_result.sheet_results[0].entries
        if original.mapping_result and original.mapping_result.sheet_results
        else ()
    )
    system_entries = (
        system.mapping_result.sheet_results[0].entries
        if system.mapping_result.sheet_results
        else ()
    )
    source_dataset = _SideDataset(
        side=VerificationSide.ORIGINAL_SOURCE,
        source_file=original.source_file,
        sheet_name=original.sheet_name,
        rows=original.rows,
        mapping_entries=source_entries,
        issues=original.issues,
        sheet_available=original.sheet_available,
    )
    system_dataset = _system_dataset(system, system_entries, config)
    issues = list(original.issues)
    issues.extend(system_dataset.issues)

    source_schema = _field_schema(source_dataset, config.original_source.field_mapping, required_fields)
    system_schema = _field_schema(system_dataset, config.system_export.field_mapping, required_fields)
    for side_data, schema in ((source_dataset, source_schema), (system_dataset, system_schema)):
        for field_name, status in schema.items():
            if status is FieldValueStatus.FIELD_MISSING:
                issues.append(
                    VerificationIssue(
                        category=VerificationIssueCategory.INPUT,
                        code="FIELD_MISSING",
                        message=f"{side_data.side.value} 未找到已配置的标准字段“{field_name}”",
                        side=side_data.side,
                        standard_field_name=field_name,
                    )
                )
            elif status is FieldValueStatus.MAPPING_FAILED:
                issues.append(
                    VerificationIssue(
                        category=VerificationIssueCategory.INPUT,
                        code="FIELD_MAPPING_UNRESOLVED",
                        message=f"{side_data.side.value} 的标准字段“{field_name}”映射不唯一",
                        side=side_data.side,
                        standard_field_name=field_name,
                    )
                )

    key_schema_valid = source_dataset.sheet_available and system_dataset.sheet_available and all(
        source_schema.get(name) is FieldValueStatus.PRESENT
        and system_schema.get(name) is FieldValueStatus.PRESENT
        for name in config.key_fields
    )
    source_keyed, source_unkeyed, source_key_issues = _key_rows(
        source_dataset, source_schema, config.key_fields
    ) if key_schema_valid else ((), source_dataset.rows, ())
    system_keyed, system_unkeyed, system_key_issues = _key_rows(
        system_dataset, system_schema, config.key_fields
    ) if key_schema_valid else ((), system_dataset.rows, ())
    issues.extend(source_key_issues)
    issues.extend(system_key_issues)

    records: list[VerificationRecord] = []
    if not key_schema_valid:
        for row in source_dataset.rows:
            records.append(_unverifiable_record(row, VerificationSide.ORIGINAL_SOURCE, config.key_fields))
        for row in system_dataset.rows:
            records.append(_unverifiable_record(row, VerificationSide.SYSTEM_EXPORT, config.key_fields))
        if not any(issue.blocks_verification for issue in issues):
            issues.append(
                VerificationIssue(
                    category=VerificationIssueCategory.MATCHING,
                    code="KEY_SCHEMA_UNAVAILABLE",
                    message="至少一侧无法唯一确定全部主键字段，无法进行匹配",
                    side=VerificationSide.BOTH,
                )
            )
    else:
        for row in source_unkeyed:
            records.append(_unverifiable_record(row, VerificationSide.ORIGINAL_SOURCE, config.key_fields))
        for row in system_unkeyed:
            records.append(_unverifiable_record(row, VerificationSide.SYSTEM_EXPORT, config.key_fields))
        source_groups = _group_by_key(source_keyed)
        system_groups = _group_by_key(system_keyed)
        all_signatures = set(source_groups) | set(system_groups)
        signature_order = sorted(all_signatures, key=repr)
        for signature in signature_order:
            source_group = source_groups.get(signature, ())
            system_group = system_groups.get(signature, ())
            exemplar = source_group[0] if source_group else system_group[0]
            key_values = exemplar.key_values
            if len(source_group) > 1 or len(system_group) > 1:
                records.append(
                    VerificationRecord(
                        status=VerificationRowStatus.DUPLICATE_KEY,
                        key_fields=config.key_fields,
                        key_values=key_values,
                        original_source_rows=tuple(_row_location(item.row, VerificationSide.ORIGINAL_SOURCE) for item in source_group),
                        system_export_rows=tuple(_row_location(item.row, VerificationSide.SYSTEM_EXPORT) for item in system_group),
                    )
                )
                locations = tuple(
                    _row_location(item.row, side)
                    for side, group in (
                        (VerificationSide.ORIGINAL_SOURCE, source_group),
                        (VerificationSide.SYSTEM_EXPORT, system_group),
                    )
                    for item in group
                )
                issues.append(
                    VerificationIssue(
                        category=VerificationIssueCategory.MATCHING,
                        code="DUPLICATE_KEY",
                        message="主键在至少一侧重复；关联记录均保留且不继续匹配",
                        side=(
                            VerificationSide.BOTH
                            if source_group and system_group
                            else VerificationSide.ORIGINAL_SOURCE
                            if source_group
                            else VerificationSide.SYSTEM_EXPORT
                        ),
                        key_values=key_values,
                        locations=locations,
                    )
                )
                continue
            if source_group and not system_group:
                if system_unkeyed:
                    row = source_group[0].row
                    records.append(replace(_unverifiable_record(row, VerificationSide.ORIGINAL_SOURCE, config.key_fields), key_values=key_values))
                    issues.append(VerificationIssue(
                        VerificationIssueCategory.MATCHING, "COUNTERPART_KEY_UNAVAILABLE",
                        "系统侧存在无效主键记录，无法确定源记录是否缺失",
                        VerificationSide.BOTH, key_values=key_values,
                        locations=(_row_location(row, VerificationSide.ORIGINAL_SOURCE),),
                    ))
                    continue
                records.append(
                    VerificationRecord(
                        VerificationRowStatus.MISSING_IN_SYSTEM,
                        config.key_fields,
                        key_values,
                        (_row_location(source_group[0].row, VerificationSide.ORIGINAL_SOURCE),),
                        (),
                    )
                )
                continue
            if system_group and not source_group:
                if source_unkeyed:
                    row = system_group[0].row
                    records.append(replace(_unverifiable_record(row, VerificationSide.SYSTEM_EXPORT, config.key_fields), key_values=key_values))
                    issues.append(VerificationIssue(
                        VerificationIssueCategory.MATCHING, "COUNTERPART_KEY_UNAVAILABLE",
                        "源侧存在无效主键记录，无法确定系统记录是否新增",
                        VerificationSide.BOTH, key_values=key_values,
                        locations=(_row_location(row, VerificationSide.SYSTEM_EXPORT),),
                    ))
                    continue
                records.append(
                    VerificationRecord(
                        VerificationRowStatus.EXTRA_IN_SYSTEM,
                        config.key_fields,
                        key_values,
                        (),
                        (_row_location(system_group[0].row, VerificationSide.SYSTEM_EXPORT),),
                    )
                )
                continue
            source_row = source_group[0].row
            system_row = system_group[0].row
            differences: list[FieldDifference] = []
            for field_name in config.comparison_fields:
                source_value = _row_field(
                    source_row,
                    field_name,
                    source_schema.get(field_name, FieldValueStatus.FIELD_MISSING),
                    VerificationSide.ORIGINAL_SOURCE,
                    source_dataset,
                )
                system_value = _row_field(
                    system_row,
                    field_name,
                    system_schema.get(field_name, FieldValueStatus.FIELD_MISSING),
                    VerificationSide.SYSTEM_EXPORT,
                    system_dataset,
                )
                rule = _rule_for(field_name, config.comparison_rules)
                status, reason = _compare_field(source_value, system_value, rule)
                differences.append(
                    FieldDifference(
                        standard_field_name=field_name,
                        status=status,
                        comparison_rule=rule,
                        original_source=source_value,
                        system_export=system_value,
                        reason=reason,
                    )
                )
                if status is FieldComparisonStatus.UNVERIFIABLE:
                    issues.append(
                        VerificationIssue(
                            category=VerificationIssueCategory.COMPARISON,
                            code="FIELD_UNVERIFIABLE",
                            message=f"字段“{field_name}”无法比较：{reason}",
                            side=VerificationSide.BOTH,
                            standard_field_name=field_name,
                            key_values=key_values,
                            locations=tuple(
                                loc for loc in (source_value.location, system_value.location) if loc is not None
                            ),
                        )
                    )
            statuses = {item.status for item in differences}
            if FieldComparisonStatus.UNVERIFIABLE in statuses:
                row_status = VerificationRowStatus.UNVERIFIABLE
            elif FieldComparisonStatus.CHANGED in statuses:
                row_status = VerificationRowStatus.FIELD_CHANGED
            else:
                row_status = VerificationRowStatus.MATCH
            records.append(
                VerificationRecord(
                    status=row_status,
                    key_fields=config.key_fields,
                    key_values=key_values,
                    original_source_rows=(_row_location(source_row, VerificationSide.ORIGINAL_SOURCE),),
                    system_export_rows=(_row_location(system_row, VerificationSide.SYSTEM_EXPORT),),
                    field_differences=tuple(differences),
                )
            )

    records_tuple = tuple(records)
    if not original.rows and not system.rows:
        issues.append(
            VerificationIssue(
                category=VerificationIssueCategory.MATCHING,
                code="NO_DATA_ROWS",
                message="两侧均无数据行，没有可核验记录",
                side=VerificationSide.BOTH,
            )
        )
    issues_tuple = tuple(issues)
    summary = _summary(
        original_row_count=len(original.rows),
        system_row_count=len(system.rows),
        records=records_tuple,
        issues=issues_tuple,
    )
    return SourceVerificationResult(
        config=config,
        original_source=original,
        system_export=system,
        records=records_tuple,
        issues=issues_tuple,
        summary=summary,
    )


def _ingest_original_source(
    path: str | Path,
    config: OriginalSourceConfig,
    relevant_fields: frozenset[str],
) -> OriginalSourceWorkbookResult:
    loaded = load_workbook(path)
    raw_sheet = next((sheet for sheet in loaded.sheets if sheet.sheet_name == config.sheet_name), None)
    if raw_sheet is None:
        issue = VerificationIssue(
            category=VerificationIssueCategory.INPUT,
            code="SHEET_NOT_FOUND",
            message=f"原始源文档中未找到 Sheet“{config.sheet_name}”",
            side=VerificationSide.ORIGINAL_SOURCE,
        )
        return OriginalSourceWorkbookResult(
            loaded.path, config.sheet_name, config.header_row_number, (), (issue,), None, None, False
        )
    profiled = profile_workbook(path)
    if config.header_row_number > len(raw_sheet.rows):
        issue = VerificationIssue(
            category=VerificationIssueCategory.INPUT,
            code="HEADER_ROW_NOT_FOUND",
            message=f"原始源 Sheet“{config.sheet_name}”不存在配置表头行 {config.header_row_number}",
            side=VerificationSide.ORIGINAL_SOURCE,
        )
        return OriginalSourceWorkbookResult(
            loaded.path, config.sheet_name, config.header_row_number, (), (issue,), None, None, False
        )
    sheet_profile = _configured_sheet_profile(raw_sheet, profiled, config.header_row_number)
    workbook_profile = WorkbookProfile(loaded.path, (sheet_profile,), sheet_profile.issues)
    mapping = map_workbook_fields(workbook_profile, config.field_mapping)
    try:
        cleaning = clean_workbook(workbook_profile, mapping, config.cleaning, raw=loaded)
    except CleaningConfigError as exc:
        raise SourceVerificationConfigError(str(exc)) from exc
    mapping_sheet = mapping.sheet_results[0]
    entry_by_column = {entry.source_column_number: entry for entry in mapping_sheet.entries}
    rows: list[OriginalSourceRow] = []
    issues: list[VerificationIssue] = []
    for cleaned_row in cleaning.sheet_results[0].rows:
        fields: list[NormalizedFieldValue] = []
        for audit in cleaned_row.cells:
            entry = entry_by_column.get(audit.source_column_number)
            status = _audit_status(audit.action, entry.status if entry else FieldMappingStatus.UNMATCHED, audit.cleaned_value)
            loc = VerificationLocation(
                side=VerificationSide.ORIGINAL_SOURCE,
                source_file=loaded.path,
                sheet_name=config.sheet_name,
                row_number=audit.row_number,
                source_field_name=audit.source_field_name,
                standard_field_name=audit.standard_field_name,
                column_number=audit.source_column_number,
                column_letter=audit.source_column_letter,
                cell_address=cell_address(audit.row_number, audit.source_column_number - 1),
            )
            fields.append(
                NormalizedFieldValue(
                    standard_field_name=audit.standard_field_name or "",
                    original_value=audit.original_value,
                    normalized_value=audit.cleaned_value,
                    status=status,
                    applied_rules=_audit_rules(audit),
                    location=loc,
                    reason=audit.reason,
                )
            )
            if audit.action is CleaningAction.INVALID:
                issues.append(
                    VerificationIssue(
                        category=VerificationIssueCategory.INPUT,
                        code="CLEANING_FAILED",
                        message=audit.reason,
                        side=VerificationSide.ORIGINAL_SOURCE,
                        standard_field_name=audit.standard_field_name,
                        locations=(loc,),
                        blocks_verification=audit.standard_field_name in relevant_fields,
                    )
                )
        rows.append(OriginalSourceRow(loaded.path, config.sheet_name, cleaned_row.source_row_number, tuple(fields)))
    for entry in mapping_sheet.entries:
        if entry.status is FieldMappingStatus.MATCH:
            continue
        code = f"FIELD_MAPPING_{entry.status.value}"
        issue_location = VerificationLocation(
            side=VerificationSide.ORIGINAL_SOURCE,
            source_file=loaded.path,
            sheet_name=config.sheet_name,
            row_number=config.header_row_number,
            source_field_name=entry.source_field_name,
            standard_field_name=entry.standard_field_name,
            column_number=entry.source_column_number,
            column_letter=entry.source_column_letter,
            cell_address=cell_address(config.header_row_number, entry.source_column_number - 1),
        )
        issues.append(
            VerificationIssue(
                category=VerificationIssueCategory.INPUT,
                code=code,
                message=entry.reason,
                side=VerificationSide.ORIGINAL_SOURCE,
                standard_field_name=entry.standard_field_name,
                locations=(issue_location,),
                blocks_verification=(
                    entry.standard_field_name in relevant_fields
                    or entry.status is FieldMappingStatus.AMBIGUOUS
                    and any(_field_ambiguous(name, config.field_mapping, mapping_sheet.entries) for name in relevant_fields)
                ),
            )
        )
    return OriginalSourceWorkbookResult(
        loaded.path,
        config.sheet_name,
        config.header_row_number,
        tuple(rows),
        tuple(issues),
        mapping,
        cleaning,
        True,
    )


def _configured_sheet_profile(
    raw_sheet: LoadedSheet,
    detected: WorkbookProfile,
    header_row_number: int,
) -> SheetProfile:
    if header_row_number > len(raw_sheet.rows):
        raise SourceVerificationDataError(
            f"配置表头行 {header_row_number} 超出 Sheet“{raw_sheet.sheet_name}”范围"
        )
    base = next((sheet for sheet in detected.sheets if sheet.sheet_name == raw_sheet.sheet_name), None)
    header = raw_sheet.rows[header_row_number - 1]
    width = max(len(header), max((len(row) for row in raw_sheet.rows), default=0))
    last_row = max(
        (
            header_row_number + index + 1
            for index, row in enumerate(raw_sheet.rows[header_row_number:])
            if any(not is_blank_value(value) for value in row)
        ),
        default=header_row_number,
    )
    old_fields = {field.column_index: field for field in (base.fields if base else ())}
    fields: list[FieldProfile] = []
    header_issues: list[ProfileIssue] = []
    for col_index in range(width):
        raw_name = header[col_index] if col_index < len(header) else None
        name = "" if is_blank_value(raw_name) else str(raw_name).strip()
        previous = old_fields.get(col_index)
        if previous is None:
            fields.append(
                FieldProfile(
                    name,
                    col_index,
                    col_index + 1,
                    column_letter(col_index),
                    BasicDataType.OTHER,
                    0,
                    0,
                    (),
                )
            )
        else:
            fields.append(replace(previous, name=name))
        if not name:
            header_issues.append(
                ProfileIssue(
                    "EMPTY_HEADER_CELL",
                    f"第 {column_letter(col_index)} 列表头为空",
                    ProfileIssueLevel.WARNING,
                    raw_sheet.sheet_name,
                    header_row_number,
                    column_letter(col_index),
                )
            )
    names = [field.name.casefold() for field in fields if field.name]
    duplicates = sorted(name for name, count in Counter(names).items() if count > 1)
    if duplicates:
        header_issues.append(
            ProfileIssue(
                "DUPLICATE_HEADER_NAME",
                "配置表头行存在同名列：" + "、".join(duplicates),
                ProfileIssueLevel.WARNING,
                raw_sheet.sheet_name,
                header_row_number,
            )
        )
    data_last = last_row if last_row > header_row_number else None
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
        header_start_row=header_row_number,
        header_end_row=header_row_number,
        header_row=header_row_number,
        header_row_count=1,
        data_first_row=header_row_number + 1 if data_last else None,
        data_last_row=data_last,
        data_first_col=1 if width else None,
        data_last_col=width if width else None,
        data_row_count=(data_last - header_row_number) if data_last else 0,
        fields=tuple(fields),
        issues=tuple(header_issues),
    )


def _system_dataset(
    result: SystemExportWorkbookResult,
    entries: tuple[FieldMappingEntry, ...],
    config: SourceVerificationConfig,
) -> _SideDataset:
    issues: list[VerificationIssue] = []
    relevant = set((*config.key_fields, *config.comparison_fields))
    for issue in result.issues:
        if issue.code in {"EMPTY_PRIMARY_KEY", "DUPLICATE_PRIMARY_KEY"}:
            continue
        blocks = issue.code in {"SHEET_NOT_FOUND", "MISSING_FIELD"} and (
            issue.standard_field_name is None or issue.standard_field_name in relevant
        )
        if issue.code.startswith("FIELD_MAPPING_"):
            blocks = issue.standard_field_name in relevant or (
                issue.code == "FIELD_MAPPING_AMBIGUOUS"
                and any(_field_ambiguous(name, config.system_export.field_mapping, entries) for name in relevant)
            )
        if issue.code == "CLEANING_FAILED":
            blocks = issue.standard_field_name in relevant
        issues.append(
            VerificationIssue(
                category=VerificationIssueCategory.INPUT,
                code=issue.code,
                message=issue.message,
                side=VerificationSide.SYSTEM_EXPORT,
                standard_field_name=issue.standard_field_name,
                locations=_system_issue_locations(result, issue),
                blocks_verification=blocks,
            )
        )
    return _SideDataset(
        VerificationSide.SYSTEM_EXPORT,
        result.source_file,
        result.sheet_name,
        result.rows,
        entries,
        tuple(issues),
        not any(issue.code == "SHEET_NOT_FOUND" for issue in result.issues),
        frozenset(
            (audit.row_number, audit.source_column_number)
            for sheet in result.cleaning_result.sheet_results
            for row in sheet.rows
            for audit in row.cells
            if audit.action is CleaningAction.INVALID
        ),
    )


def _field_schema(
    dataset: _SideDataset,
    mapping: FieldMappingConfig,
    required_fields: tuple[str, ...],
) -> dict[str, FieldValueStatus]:
    schema: dict[str, FieldValueStatus] = {}
    for name in required_fields:
        named_entries = [entry for entry in dataset.mapping_entries if entry.standard_field_name == name]
        if len(named_entries) == 1 and named_entries[0].status is FieldMappingStatus.MATCH:
            schema[name] = FieldValueStatus.PRESENT
            continue
        if any(entry.status in {FieldMappingStatus.CONFLICT, FieldMappingStatus.AMBIGUOUS} for entry in named_entries) or _field_ambiguous(name, mapping, dataset.mapping_entries):
            schema[name] = FieldValueStatus.MAPPING_FAILED
        else:
            schema[name] = FieldValueStatus.FIELD_MISSING
    if not dataset.sheet_available:
        for name in required_fields:
            schema[name] = FieldValueStatus.FIELD_MISSING
    return schema


def _field_ambiguous(
    standard_field: str,
    mapping: FieldMappingConfig,
    entries: tuple[FieldMappingEntry, ...],
) -> bool:
    definition = next((item for item in mapping.standard_fields if item.name == standard_field), None)
    if definition is None:
        return False
    terms = {normalize_field_name(term) for term in (definition.name, *definition.aliases)}
    return any(
        entry.status is FieldMappingStatus.AMBIGUOUS
        and normalize_field_name(entry.source_field_name) in terms
        for entry in entries
    )


def _system_issue_locations(
    result: SystemExportWorkbookResult,
    issue: SystemExportIssue,
) -> tuple[VerificationLocation, ...]:
    if issue.row_number is None:
        return ()
    matching_field = next(
        (
            field
            for row in result.rows
            if row.source_row_number == issue.row_number
            for field in row.fields
            if field.source_column_number == issue.column_number
        ),
        None,
    )
    return (
        VerificationLocation(
            side=VerificationSide.SYSTEM_EXPORT,
            source_file=result.source_file,
            sheet_name=issue.sheet_name or result.sheet_name,
            row_number=issue.row_number,
            source_field_name=matching_field.source_field_name if matching_field else None,
            standard_field_name=issue.standard_field_name,
            column_number=issue.column_number,
            column_letter=issue.column_letter,
            cell_address=(
                cell_address(issue.row_number, issue.column_number - 1)
                if issue.column_number is not None
                else None
            ),
        ),
    )


def _key_rows(
    dataset: _SideDataset,
    schema: dict[str, FieldValueStatus],
    key_fields: tuple[str, ...],
) -> tuple[tuple[_KeyedRow, ...], tuple[Any, ...], tuple[VerificationIssue, ...]]:
    keyed: list[_KeyedRow] = []
    unkeyed: list[Any] = []
    issues: list[VerificationIssue] = []
    for row in dataset.rows:
        values = tuple(_row_field(row, name, schema[name], dataset.side, dataset) for name in key_fields)
        bad = next((value for value in values if value.status not in {FieldValueStatus.PRESENT}), None)
        if bad is not None:
            code = {
                FieldValueStatus.EMPTY: "EMPTY_PRIMARY_KEY",
                FieldValueStatus.FIELD_MISSING: "KEY_FIELD_MISSING",
                FieldValueStatus.MAPPING_FAILED: "KEY_MAPPING_FAILED",
                FieldValueStatus.CLEANING_FAILED: "KEY_CLEANING_FAILED",
            }.get(bad.status, "INVALID_PRIMARY_KEY")
            issues.append(
                VerificationIssue(
                    category=VerificationIssueCategory.MATCHING,
                    code=code,
                    message=f"{dataset.side.value} 主键字段“{bad.standard_field_name}”不可用于匹配",
                    side=dataset.side,
                    standard_field_name=bad.standard_field_name,
                    locations=(bad.location,) if bad.location else (),
                )
            )
            unkeyed.append(row)
            continue
        key_values = tuple(value.normalized_value for value in values)
        if any(is_blank_value(value) for value in key_values):
            bad_value = next(value for value in values if is_blank_value(value.normalized_value))
            issues.append(
                VerificationIssue(
                    VerificationIssueCategory.MATCHING,
                    "EMPTY_PRIMARY_KEY",
                    f"{dataset.side.value} 主键字段“{bad_value.standard_field_name}”为空",
                    dataset.side,
                    bad_value.standard_field_name,
                    locations=(bad_value.location,) if bad_value.location else (),
                )
            )
            unkeyed.append(row)
            continue
        try:
            signature = tuple(_key_atom(value) for value in key_values)
            hash(signature)
        except (TypeError, ValueError):
            issues.append(
                VerificationIssue(
                    VerificationIssueCategory.MATCHING,
                    "INVALID_PRIMARY_KEY",
                    f"{dataset.side.value} 主键包含不支持的值类型",
                    dataset.side,
                    key_values=key_values,
                    locations=(_row_location(row, dataset.side),),
                )
            )
            unkeyed.append(row)
            continue
        keyed.append(_KeyedRow(row, key_values, signature))
    return tuple(keyed), tuple(unkeyed), tuple(issues)


def _key_atom(value: Any) -> tuple[str, Any]:
    type_name = f"{type(value).__module__}.{type(value).__qualname__}"
    if isinstance(value, datetime):
        return type_name, value.isoformat()
    if isinstance(value, date):
        return type_name, value.isoformat()
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError("non-finite number")
        return type_name, value.hex()
    if isinstance(value, Decimal):
        if not value.is_finite():
            raise ValueError("non-finite decimal")
        return type_name, str(value)
    return type_name, value


def _group_by_key(rows: tuple[_KeyedRow, ...]) -> dict[tuple[Any, ...], tuple[_KeyedRow, ...]]:
    groups: dict[tuple[Any, ...], list[_KeyedRow]] = defaultdict(list)
    for row in rows:
        groups[row.signature].append(row)
    return {key: tuple(items) for key, items in groups.items()}


def _row_field(
    row: OriginalSourceRow | SystemExportRow,
    field_name: str,
    schema_status: FieldValueStatus,
    side: VerificationSide,
    dataset: _SideDataset,
) -> NormalizedFieldValue:
    field_values = (
        row.fields
        if isinstance(row, OriginalSourceRow)
        else tuple(_system_field_value(field, row, side, dataset.failed_cells) for field in row.fields)
    )
    value = next(
        (item for item in field_values if item.standard_field_name == field_name),
        None,
    )
    if schema_status is FieldValueStatus.FIELD_MISSING:
        return NormalizedFieldValue(
            field_name, None, None, FieldValueStatus.FIELD_MISSING, (), _row_location(row, side),
            "字段不存在",
        )
    if schema_status is FieldValueStatus.MAPPING_FAILED:
        return NormalizedFieldValue(
            field_name, value.original_value if value else None,
            value.normalized_value if value else None,
            FieldValueStatus.MAPPING_FAILED,
            value.applied_rules if value else (),
            value.location if value else _row_location(row, side),
            "字段映射存在歧义或冲突",
        )
    if value is None:
        return NormalizedFieldValue(
            field_name, None, None, FieldValueStatus.FIELD_MISSING, (), _row_location(row, side),
            "映射结果中没有该字段",
        )
    return value


def _system_field_value(
    field: SystemExportFieldValue,
    row: SystemExportRow,
    side: VerificationSide,
    failed_cells: frozenset[tuple[int, int]],
) -> NormalizedFieldValue:
    if (row.source_row_number, field.source_column_number) in failed_cells:
        status = FieldValueStatus.CLEANING_FAILED
    elif field.mapping_status is FieldMappingStatus.MATCH:
        status = _value_status(field.normalized_value)
    elif field.mapping_status in {FieldMappingStatus.AMBIGUOUS, FieldMappingStatus.CONFLICT}:
        status = FieldValueStatus.MAPPING_FAILED
    else:
        status = FieldValueStatus.FIELD_MISSING
    return NormalizedFieldValue(
        standard_field_name=field.standard_field_name or "",
        original_value=field.original_value,
        normalized_value=field.normalized_value,
        status=status,
        applied_rules=field.applied_rules,
        location=VerificationLocation(
            side=side,
            source_file=row.source_file,
            sheet_name=row.sheet_name,
            row_number=row.source_row_number,
            source_field_name=field.source_field_name,
            standard_field_name=field.standard_field_name,
            column_number=field.source_column_number,
            column_letter=field.source_column_letter,
            cell_address=cell_address(row.source_row_number, field.source_column_number - 1),
        ),
        reason="字段清洗失败" if status is FieldValueStatus.CLEANING_FAILED else "",
    )


def _value_status(value: Any) -> FieldValueStatus:
    if value is None or value == "":
        return FieldValueStatus.EMPTY
    return FieldValueStatus.PRESENT


def _audit_status(
    action: CleaningAction,
    mapping_status: FieldMappingStatus,
    value: Any,
) -> FieldValueStatus:
    if action is CleaningAction.INVALID:
        return FieldValueStatus.CLEANING_FAILED
    if mapping_status in {FieldMappingStatus.AMBIGUOUS, FieldMappingStatus.CONFLICT}:
        return FieldValueStatus.MAPPING_FAILED
    if mapping_status is not FieldMappingStatus.MATCH:
        return FieldValueStatus.FIELD_MISSING
    return _value_status(value)


def _audit_rules(audit: CleanCellAudit) -> tuple[str, ...]:
    if audit.action in {
        CleaningAction.NO_RULE,
        CleaningAction.UNMAPPED,
        CleaningAction.AMBIGUOUS,
        CleaningAction.CONFLICT,
    }:
        return ()
    return (audit.rule_code or f"CLEANING_RULE:{audit.standard_field_name}",)


def _compare_field(
    source: NormalizedFieldValue,
    system: NormalizedFieldValue,
    rule: ComparisonRuleKind,
) -> tuple[FieldComparisonStatus, str]:
    unusable = {
        FieldValueStatus.FIELD_MISSING: "字段不存在",
        FieldValueStatus.MAPPING_FAILED: "字段映射失败",
        FieldValueStatus.CLEANING_FAILED: "字段清洗失败",
    }
    for value in (source, system):
        if value.status in unusable:
            return FieldComparisonStatus.UNVERIFIABLE, f"{value.status.value}：{unusable[value.status]}"
    left = source.normalized_value
    right = system.normalized_value
    if _non_finite(left) or _non_finite(right):
        return FieldComparisonStatus.UNVERIFIABLE, "存在 NaN 或无穷值，无法确定比较结果"
    if left is None or left == "" or right is None or right == "":
        if type(left) is type(right) and left == right:
            return FieldComparisonStatus.EQUAL, "双方空值表示一致（None 与空字符串分别比较）"
        return FieldComparisonStatus.CHANGED, "空值表示不同；None 与空字符串不互相折叠"
    if rule is ComparisonRuleKind.EXACT:
        if type(left) is type(right) and left == right:
            return FieldComparisonStatus.EQUAL, "值和 Python 数据类型均一致"
        return FieldComparisonStatus.CHANGED, "严格比较下值或数据类型不同"
    if rule is ComparisonRuleKind.TEXT:
        if not isinstance(left, str) or not isinstance(right, str):
            return FieldComparisonStatus.UNVERIFIABLE, "TEXT 规则要求双方标准化值均为文本"
        if left == right:
            return FieldComparisonStatus.EQUAL, "文本逐字符一致，未忽略空白或大小写"
        return FieldComparisonStatus.CHANGED, "文本逐字符不同"
    if rule is ComparisonRuleKind.BOOLEAN:
        if not isinstance(left, bool) or not isinstance(right, bool):
            return FieldComparisonStatus.UNVERIFIABLE, "BOOLEAN 规则只接受布尔值，不将 0/1 转为布尔值"
        if left is right:
            return FieldComparisonStatus.EQUAL, "布尔值一致"
        return FieldComparisonStatus.CHANGED, "布尔值不同"
    if rule is ComparisonRuleKind.NUMBER:
        left_number = _decimal_value(left)
        right_number = _decimal_value(right)
        if left_number is None or right_number is None:
            return FieldComparisonStatus.UNVERIFIABLE, "NUMBER 规则要求双方为有限数值或无空白数字文本"
        if left_number == right_number:
            return FieldComparisonStatus.EQUAL, "按显式 NUMBER 规则转换为 Decimal 后一致"
        return FieldComparisonStatus.CHANGED, "按显式 NUMBER 规则转换为 Decimal 后不同"
    if rule is ComparisonRuleKind.DATE:
        left_date = _date_value(left)
        right_date = _date_value(right)
        if left_date is None or right_date is None:
            return FieldComparisonStatus.UNVERIFIABLE, "DATE 规则只接受日期、日期时间或 ISO 日期文本"
        if left_date == right_date:
            return FieldComparisonStatus.EQUAL, "按显式 DATE 规则比较日历日期后一致"
        return FieldComparisonStatus.CHANGED, "按显式 DATE 规则比较日历日期后不同"
    return FieldComparisonStatus.UNVERIFIABLE, f"不支持比较规则 {rule}"


def _decimal_value(value: Any) -> Decimal | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, Decimal):
        number = value
    elif isinstance(value, int):
        number = Decimal(value)
    elif isinstance(value, float):
        if not math.isfinite(value):
            return None
        number = Decimal(str(value))
    elif isinstance(value, str) and _NUMBER_TEXT.fullmatch(value):
        try:
            number = Decimal(value)
        except InvalidOperation:
            return None
    else:
        return None
    return number if number.is_finite() else None


def _date_value(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str) and _DATE_TEXT.fullmatch(value):
        try:
            return date.fromisoformat(value)
        except ValueError:
            return None
    return None


def _non_finite(value: Any) -> bool:
    if isinstance(value, float):
        return not math.isfinite(value)
    if isinstance(value, Decimal):
        return not value.is_finite()
    return False


def _rule_for(field_name: str, rules: tuple[FieldComparisonRule, ...]) -> ComparisonRuleKind:
    return next((rule.kind for rule in rules if rule.standard_field == field_name), ComparisonRuleKind.EXACT)


def _row_location(row: OriginalSourceRow | SystemExportRow, side: VerificationSide) -> VerificationLocation:
    return VerificationLocation(
        side=side,
        source_file=row.source_file,
        sheet_name=row.sheet_name,
        row_number=row.source_row_number,
    )


def _unverifiable_record(
    row: OriginalSourceRow | SystemExportRow,
    side: VerificationSide,
    key_fields: tuple[str, ...],
) -> VerificationRecord:
    location = _row_location(row, side)
    return VerificationRecord(
        status=VerificationRowStatus.UNVERIFIABLE,
        key_fields=key_fields,
        key_values=(),
        original_source_rows=(location,) if side is VerificationSide.ORIGINAL_SOURCE else (),
        system_export_rows=(location,) if side is VerificationSide.SYSTEM_EXPORT else (),
    )


def _summary(
    *,
    original_row_count: int,
    system_row_count: int,
    records: tuple[VerificationRecord, ...],
    issues: tuple[VerificationIssue, ...],
) -> SourceVerificationSummary:
    statuses = Counter(record.status for record in records)
    field_statuses = Counter(
        difference.status for record in records for difference in record.field_differences
    )
    matched = sum(
        len(record.original_source_rows) == 1
        and len(record.system_export_rows) == 1
        and record.status in {
            VerificationRowStatus.MATCH,
            VerificationRowStatus.FIELD_CHANGED,
            VerificationRowStatus.UNVERIFIABLE,
        }
        for record in records
    )
    blocks = any(issue.blocks_verification for issue in issues) or any(
        record.status in {VerificationRowStatus.DUPLICATE_KEY, VerificationRowStatus.UNVERIFIABLE}
        for record in records
    )
    complete = not blocks
    issue_counts = Counter(issue.code for issue in issues)
    return SourceVerificationSummary(
        original_source_input_row_count=original_row_count,
        system_export_input_row_count=system_row_count,
        matched_row_count=matched,
        unchanged_row_count=statuses[VerificationRowStatus.MATCH],
        changed_row_count=statuses[VerificationRowStatus.FIELD_CHANGED],
        missing_in_system_row_count=statuses[VerificationRowStatus.MISSING_IN_SYSTEM],
        extra_in_system_row_count=statuses[VerificationRowStatus.EXTRA_IN_SYSTEM],
        duplicate_key_group_count=statuses[VerificationRowStatus.DUPLICATE_KEY],
        unverifiable_row_count=statuses[VerificationRowStatus.UNVERIFIABLE],
        compared_field_count=field_statuses[FieldComparisonStatus.EQUAL] + field_statuses[FieldComparisonStatus.CHANGED],
        equal_field_count=field_statuses[FieldComparisonStatus.EQUAL],
        changed_field_count=field_statuses[FieldComparisonStatus.CHANGED],
        unverifiable_field_count=field_statuses[FieldComparisonStatus.UNVERIFIABLE],
        issue_count=len(issues),
        issue_counts=tuple(sorted(issue_counts.items())),
        verification_complete=complete,
        data_consistent=complete and bool(records) and all(
            record.status is VerificationRowStatus.MATCH for record in records
        ),
    )


def _validate_config(config: SourceVerificationConfig) -> None:
    if not isinstance(config, SourceVerificationConfig):
        raise SourceVerificationConfigError("config 必须是 SourceVerificationConfig")
    if not isinstance(config.original_source, OriginalSourceConfig):
        raise SourceVerificationConfigError("original_source 必须是 OriginalSourceConfig")
    if not isinstance(config.system_export, SystemExportConfig):
        raise SourceVerificationConfigError("system_export 必须是 SystemExportConfig")
    for name, side in (("原始源", config.original_source), ("系统导出", config.system_export)):
        if not isinstance(side.sheet_name, str) or not side.sheet_name.strip():
            raise SourceVerificationConfigError(f"{name} Sheet 名不能为空")
        if isinstance(side.header_row_number, bool) or not isinstance(side.header_row_number, int) or side.header_row_number < 1:
            raise SourceVerificationConfigError(f"{name}表头行必须为正整数")
        if not isinstance(side.field_mapping, FieldMappingConfig):
            raise SourceVerificationConfigError(f"{name} field_mapping 类型无效")
        if not side.field_mapping.standard_fields:
            raise SourceVerificationConfigError(f"{name}至少配置一个标准字段")
    if not isinstance(config.key_fields, tuple) or not config.key_fields:
        raise SourceVerificationConfigError("key_fields 必须配置一个或多个标准字段")
    if not isinstance(config.comparison_fields, tuple) or not config.comparison_fields:
        raise SourceVerificationConfigError("comparison_fields 必须配置一个或多个字段")
    if any(not isinstance(item, str) or not item.strip() for item in (*config.key_fields, *config.comparison_fields)):
        raise SourceVerificationConfigError("主键和比较字段必须为非空标准字段名称")
    if len(set(config.key_fields)) != len(config.key_fields):
        raise SourceVerificationConfigError("key_fields 不得重复")
    if len(set(config.comparison_fields)) != len(config.comparison_fields):
        raise SourceVerificationConfigError("comparison_fields 不得重复")
    if config.system_export.primary_key_fields != config.key_fields:
        raise SourceVerificationConfigError("system_export.primary_key_fields 必须与 key_fields 完全一致")
    for side_name, definitions in (
        ("original_source", config.original_source.field_mapping.standard_fields),
        ("system_export", config.system_export.field_mapping.standard_fields),
    ):
        known = {definition.name for definition in definitions}
        missing = sorted(set((*config.key_fields, *config.comparison_fields)) - known)
        if missing:
            raise SourceVerificationConfigError(
                f"{side_name} 未配置标准字段：" + "、".join(missing)
            )
    if any(not isinstance(rule, FieldComparisonRule) for rule in config.comparison_rules):
        raise SourceVerificationConfigError("comparison_rules 必须由 FieldComparisonRule 组成")
    rule_fields = [rule.standard_field for rule in config.comparison_rules]
    if len(rule_fields) != len(set(rule_fields)):
        raise SourceVerificationConfigError("同一标准字段只能配置一条比较规则")
    if set(rule_fields) - set(config.comparison_fields):
        raise SourceVerificationConfigError("比较规则只能应用于 comparison_fields")
    if not all(isinstance(rule.kind, ComparisonRuleKind) for rule in config.comparison_rules):
        raise SourceVerificationConfigError("存在不支持的字段比较规则")
    try:
        # Reuse existing public mapping functions for comprehensive mapping config validation.
        from excel_qc.field_mapping import map_sheet_fields

        empty_sheet = SheetProfile("", 0, True, None, None, None, None, None, None)
        map_sheet_fields(empty_sheet, config.original_source.field_mapping)
        map_sheet_fields(empty_sheet, config.system_export.field_mapping)
    except FieldMappingConfigError as exc:
        raise SourceVerificationConfigError(str(exc)) from exc
    if not isinstance(config.original_source.cleaning, CleaningWorkbookConfig):
        raise SourceVerificationConfigError("original_source.cleaning 配置类型无效")
    if not isinstance(config.system_export.cleaning, CleaningWorkbookConfig):
        raise SourceVerificationConfigError("system_export.cleaning 配置类型无效")
    for side_name, cleaning, definitions in (
        ("original_source", config.original_source.cleaning, config.original_source.field_mapping.standard_fields),
        ("system_export", config.system_export.cleaning, config.system_export.field_mapping.standard_fields),
    ):
        if any(not isinstance(rule, FieldCleaningRule) for rule in cleaning.rules):
            raise SourceVerificationConfigError(f"{side_name} 清洗规则类型无效")
        known = {definition.name for definition in definitions}
        unknown = sorted(rule.standard_field for rule in cleaning.rules if rule.standard_field not in known)
        if unknown:
            raise SourceVerificationConfigError(
                f"{side_name} 清洗规则引用未定义标准字段：" + "、".join(unknown)
            )
