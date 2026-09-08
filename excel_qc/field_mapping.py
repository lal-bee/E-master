"""P1-03 字段映射引擎：原始字段名 → 可配置标准字段。

只负责确定性字段映射与结果表达，不包含模糊匹配、AI、数据清洗或
模板生成。标准字段名与别名来自调用方提供的 `FieldMappingConfig`，
核心引擎不硬编码任何设备业务字段。
"""

from __future__ import annotations

import json
import re
import unicodedata
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Any

from excel_qc.profiler import FieldProfile, SheetProfile, WorkbookProfile

_INVISIBLE_WHITESPACE_PATTERN = re.compile(r"[\u200b-\u200d\u2060\ufeff]+")
_PUNCTUATION_TRANSLATION = str.maketrans(
    {
        "（": "(",
        "）": ")",
        "【": "[",
        "】": "]",
        "｛": "{",
        "｝": "}",
        "：": ":",
        "；": ";",
        "，": ",",
        "、": ",",
        "“": '"',
        "”": '"',
        "‘": "'",
        "’": "'",
    }
)


class FieldMappingError(Exception):
    """字段映射领域异常基类。"""


class FieldMappingConfigError(FieldMappingError):
    """字段映射配置或配置文件不合法。"""


class FieldMappingStatus(str, Enum):
    """单条源字段的映射状态。"""

    MATCH = "MATCH"
    UNMATCHED = "UNMATCHED"
    AMBIGUOUS = "AMBIGUOUS"
    CONFLICT = "CONFLICT"


class FieldMatchRule(str, Enum):
    """命中的匹配规则。"""

    EXACT_NAME = "EXACT_NAME"
    ALIAS = "ALIAS"


@dataclass(frozen=True)
class StandardFieldDefinition:
    """一个标准字段及其可配置别名。"""

    name: str
    aliases: tuple[str, ...] = ()


@dataclass(frozen=True)
class FieldMappingConfig:
    """字段映射配置：标准字段列表。"""

    standard_fields: tuple[StandardFieldDefinition, ...]


@dataclass(frozen=True)
class FieldMappingEntry:
    """一条源字段的映射结果，保留完整来源追溯信息。"""

    source_sheet_name: str
    source_column_number: int
    source_column_letter: str
    source_field_name: str
    source_display_name: str
    normalized_source_name: str
    standard_field_name: str | None
    status: FieldMappingStatus
    match_rule: FieldMatchRule | None
    matched_alias: str | None
    reason: str


@dataclass(frozen=True)
class FieldMappingSummary:
    """一个映射结果的统计汇总。"""

    total: int = 0
    match_count: int = 0
    unmatched_count: int = 0
    ambiguous_count: int = 0
    conflict_count: int = 0


@dataclass(frozen=True)
class FieldMappingResult:
    """一个 Sheet 的字段映射结果。"""

    sheet_name: str
    entries: tuple[FieldMappingEntry, ...]
    summary: FieldMappingSummary


@dataclass(frozen=True)
class WorkbookFieldMappingResult:
    """整个 WorkbookProfile 的字段映射结果。"""

    file_path: Path
    sheet_results: tuple[FieldMappingResult, ...]
    summary: FieldMappingSummary


def normalize_field_name(name: str) -> str:
    """对字段名执行确定性归一化，用于精确匹配。

    规则：Unicode NFKC → 去除不可见空白 → 常见标点归一 →
    折叠空白 → 去除首尾空白 → casefold。
    """
    text = unicodedata.normalize("NFKC", name)
    text = _INVISIBLE_WHITESPACE_PATTERN.sub("", text)
    text = text.translate(_PUNCTUATION_TRANSLATION)
    text = " ".join(text.split())
    return text.strip().casefold()


def load_field_mapping_config(path: str | Path) -> FieldMappingConfig:
    """从 JSON 文件加载字段映射配置。

    JSON 结构：
    {
      "standard_fields": [
        {"name": "设备档案.设备名称", "aliases": ["设备名称", "设备名"]}
      ]
    }
    """
    config_path = Path(path)
    if not config_path.exists() or not config_path.is_file():
        raise FieldMappingConfigError(
            f"字段映射配置文件不存在: {config_path}"
        )
    try:
        payload = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise FieldMappingConfigError(
            f"字段映射配置文件无法解析: {config_path}：{exc}"
        ) from exc
    if not isinstance(payload, dict):
        raise FieldMappingConfigError(
            "字段映射配置顶层必须是 JSON 对象"
        )
    raw_fields = payload.get("standard_fields")
    if not isinstance(raw_fields, list):
        raise FieldMappingConfigError(
            "字段映射配置缺少 standard_fields 列表"
        )
    standard_fields: list[StandardFieldDefinition] = []
    for index, raw_field in enumerate(raw_fields, start=1):
        if not isinstance(raw_field, dict):
            raise FieldMappingConfigError(
                f"standard_fields 第 {index} 项必须是对象"
            )
        name = raw_field.get("name")
        if not isinstance(name, str) or not name.strip():
            raise FieldMappingConfigError(
                f"standard_fields 第 {index} 项 name 必须是非空字符串"
            )
        raw_aliases = raw_field.get("aliases", [])
        if not isinstance(raw_aliases, list) or not all(
            isinstance(alias, str) for alias in raw_aliases
        ):
            raise FieldMappingConfigError(
                f"standard_fields 第 {index} 项 aliases 必须是字符串列表"
            )
        aliases = tuple(alias.strip() for alias in raw_aliases)
        if any(not alias for alias in aliases):
            raise FieldMappingConfigError(
                f"standard_fields 第 {index} 项包含空别名"
            )
        standard_fields.append(
            StandardFieldDefinition(name=name.strip(), aliases=aliases)
        )
    config = FieldMappingConfig(tuple(standard_fields))
    _validate_config(config)
    return config


def map_sheet_fields(
    sheet: SheetProfile,
    config: FieldMappingConfig,
) -> FieldMappingResult:
    """把一个 SheetProfile 的字段映射到标准字段。"""
    _validate_config(config)
    lookup = _build_lookup(config)
    preliminaries = [
        _map_single_field(sheet, field, config, lookup)
        for field in sheet.fields
    ]
    entries = _resolve_conflicts(preliminaries)
    summary = _summary(entries)
    return FieldMappingResult(
        sheet_name=sheet.sheet_name,
        entries=tuple(entries),
        summary=summary,
    )


def map_workbook_fields(
    profile: WorkbookProfile,
    config: FieldMappingConfig,
) -> WorkbookFieldMappingResult:
    """把整个 WorkbookProfile 的各 Sheet 字段映射到标准字段。"""
    _validate_config(config)
    sheet_results = tuple(
        map_sheet_fields(sheet, config) for sheet in profile.sheets
    )
    summary = _merge_summaries(
        tuple(result.summary for result in sheet_results)
    )
    return WorkbookFieldMappingResult(
        file_path=profile.path,
        sheet_results=sheet_results,
        summary=summary,
    )


def _validate_config(config: FieldMappingConfig) -> None:
    if not isinstance(config, FieldMappingConfig):
        raise FieldMappingConfigError("config 必须是 FieldMappingConfig")
    seen_names: dict[str, StandardFieldDefinition] = {}
    for definition in config.standard_fields:
        if not isinstance(definition, StandardFieldDefinition):
            raise FieldMappingConfigError(
                "standard_fields 只允许 StandardFieldDefinition"
            )
        if not isinstance(definition.name, str):
            raise FieldMappingConfigError("标准字段名必须是字符串")
        name = definition.name.strip()
        if not name:
            raise FieldMappingConfigError("标准字段名不能为空")
        normalized_name = normalize_field_name(name)
        if normalized_name in seen_names:
            raise FieldMappingConfigError(
                f"标准字段名重复: “{name}”"
            )
        seen_names[normalized_name] = definition
        for alias in definition.aliases:
            if not isinstance(alias, str):
                raise FieldMappingConfigError(
                    f"标准字段“{name}”的别名必须是字符串"
                )
            alias_text = alias.strip()
            if not alias_text:
                raise FieldMappingConfigError(
                    f"标准字段“{name}”包含空别名"
                )


def _map_single_field(
    sheet: SheetProfile,
    field: FieldProfile,
    config: FieldMappingConfig,
    lookup: dict[str, tuple[str, ...]],
) -> FieldMappingEntry:
    source_field_name = field.name
    source_display_name = field.display_name
    if not source_field_name.strip():
        return FieldMappingEntry(
            source_sheet_name=sheet.sheet_name,
            source_column_number=field.column_number,
            source_column_letter=field.column_letter,
            source_field_name=source_field_name,
            source_display_name=source_display_name,
            normalized_source_name="",
            standard_field_name=None,
            status=FieldMappingStatus.UNMATCHED,
            match_rule=None,
            matched_alias=None,
            reason=(
                f"空表头字段（{source_display_name}）不参与标准字段匹配，"
                "需人工配置/确认"
            ),
        )

    normalized_name = normalize_field_name(source_field_name)
    candidate_names = lookup.get(normalized_name, ())
    if not candidate_names:
        return FieldMappingEntry(
            source_sheet_name=sheet.sheet_name,
            source_column_number=field.column_number,
            source_column_letter=field.column_letter,
            source_field_name=source_field_name,
            source_display_name=source_display_name,
            normalized_source_name=normalized_name,
            standard_field_name=None,
            status=FieldMappingStatus.UNMATCHED,
            match_rule=None,
            matched_alias=None,
            reason=(
                f"源字段“{source_field_name}”（归一化：{normalized_name}）"
                "未匹配任何标准字段名或别名"
            ),
        )
    if len(candidate_names) > 1:
        return FieldMappingEntry(
            source_sheet_name=sheet.sheet_name,
            source_column_number=field.column_number,
            source_column_letter=field.column_letter,
            source_field_name=source_field_name,
            source_display_name=source_display_name,
            normalized_source_name=normalized_name,
            standard_field_name=None,
            status=FieldMappingStatus.AMBIGUOUS,
            match_rule=None,
            matched_alias=None,
            reason=(
                f"源字段“{source_field_name}”同时命中多个标准字段："
                + "、".join(candidate_names)
            ),
        )

    standard_name = candidate_names[0]
    rule, matched_alias = _match_rule_for(
        normalized_name=normalized_name,
        standard_name=standard_name,
        config=config,
    )
    return FieldMappingEntry(
        source_sheet_name=sheet.sheet_name,
        source_column_number=field.column_number,
        source_column_letter=field.column_letter,
        source_field_name=source_field_name,
        source_display_name=source_display_name,
        normalized_source_name=normalized_name,
        standard_field_name=standard_name,
        status=FieldMappingStatus.MATCH,
        match_rule=rule,
        matched_alias=matched_alias,
        reason=(
            f"源字段“{source_field_name}”{_rule_text(rule, matched_alias)}"
            f"标准字段“{standard_name}”"
        ),
    )


def _resolve_conflicts(
    entries: list[FieldMappingEntry],
) -> list[FieldMappingEntry]:
    """同一标准字段被多个源字段命中时，全部相关项标记为 CONFLICT。"""
    resolved: list[FieldMappingEntry] = []
    matched_by_standard: dict[str, list[FieldMappingEntry]] = {}
    for entry in entries:
        if (
            entry.status is FieldMappingStatus.MATCH
            and entry.standard_field_name is not None
        ):
            matched_by_standard.setdefault(entry.standard_field_name, []).append(entry)

    for entry in entries:
        if (
            entry.status is FieldMappingStatus.MATCH
            and entry.standard_field_name is not None
            and len(matched_by_standard[entry.standard_field_name]) > 1
        ):
            sources = matched_by_standard[entry.standard_field_name]
            source_labels = "、".join(
                f"{item.source_display_name}"
                f"（{item.source_column_letter}{item.source_column_number}）"
                for item in sources
            )
            resolved.append(
                _replace_entry(
                    entry,
                    status=FieldMappingStatus.CONFLICT,
                    reason=(
                        f"标准字段“{entry.standard_field_name}”同时被"
                        f"多个源字段命中：{source_labels}，需人工确认"
                    ),
                )
            )
        else:
            resolved.append(entry)
    return resolved


def _build_lookup(
    config: FieldMappingConfig,
) -> dict[str, tuple[str, ...]]:
    lookup: dict[str, list[str]] = {}
    for definition in config.standard_fields:
        name = definition.name.strip()
        terms = [name, *definition.aliases]
        for term in terms:
            normalized = normalize_field_name(term)
            if normalized not in lookup:
                lookup[normalized] = []
            if name not in lookup[normalized]:
                lookup[normalized].append(name)
    return {
        normalized: tuple(names)
        for normalized, names in lookup.items()
    }


def _match_rule_for(
    *,
    normalized_name: str,
    standard_name: str,
    config: FieldMappingConfig,
) -> tuple[FieldMatchRule, str | None]:
    definition = _find_definition(config, standard_name)
    if definition is None:
        return FieldMatchRule.EXACT_NAME, None
    if normalize_field_name(definition.name) == normalized_name:
        return FieldMatchRule.EXACT_NAME, None
    for alias in definition.aliases:
        if alias.strip() and normalize_field_name(alias) == normalized_name:
            return FieldMatchRule.ALIAS, alias.strip()
    return FieldMatchRule.EXACT_NAME, None


def _find_definition(
    config: FieldMappingConfig,
    standard_name: str,
) -> StandardFieldDefinition | None:
    for definition in config.standard_fields:
        if definition.name.strip() == standard_name:
            return definition
    return None


def _rule_text(rule: FieldMatchRule | None, alias: str | None) -> str:
    if rule is FieldMatchRule.ALIAS and alias is not None:
        return f"通过别名“{alias}”映射到"
    return "精确匹配到"


def _replace_entry(
    entry: FieldMappingEntry,
    *,
    status: FieldMappingStatus,
    reason: str,
) -> FieldMappingEntry:
    return FieldMappingEntry(
        source_sheet_name=entry.source_sheet_name,
        source_column_number=entry.source_column_number,
        source_column_letter=entry.source_column_letter,
        source_field_name=entry.source_field_name,
        source_display_name=entry.source_display_name,
        normalized_source_name=entry.normalized_source_name,
        standard_field_name=entry.standard_field_name,
        status=status,
        match_rule=entry.match_rule,
        matched_alias=entry.matched_alias,
        reason=reason,
    )


def _summary(entries: list[FieldMappingEntry]) -> FieldMappingSummary:
    counts = {status: 0 for status in FieldMappingStatus}
    for entry in entries:
        counts[entry.status] += 1
    return FieldMappingSummary(
        total=len(entries),
        match_count=counts[FieldMappingStatus.MATCH],
        unmatched_count=counts[FieldMappingStatus.UNMATCHED],
        ambiguous_count=counts[FieldMappingStatus.AMBIGUOUS],
        conflict_count=counts[FieldMappingStatus.CONFLICT],
    )


def _merge_summaries(
    summaries: tuple[FieldMappingSummary, ...],
) -> FieldMappingSummary:
    return FieldMappingSummary(
        total=sum(item.total for item in summaries),
        match_count=sum(item.match_count for item in summaries),
        unmatched_count=sum(item.unmatched_count for item in summaries),
        ambiguous_count=sum(item.ambiguous_count for item in summaries),
        conflict_count=sum(item.conflict_count for item in summaries),
    )
