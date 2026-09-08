"""P1-05 设备类型编码映射引擎。

把 P1-04 清洗后的设备类型字段值，依据人工确认的配置规则库映射为
系统设备类型编码。只做确定性精确匹配，不猜测、不调用 AI、不写 Excel。
"""

from __future__ import annotations

import json
import re
import unicodedata
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Any

from excel_qc.cleaning import (
    CleanCellAudit,
    CleanedSheetResult,
    CleaningAction,
    CleaningWorkbookResult,
)

_INVISIBLE_WHITESPACE_PATTERN = re.compile(r"[\u200b-\u200d\u2060\ufeff]+")


class EquipmentTypeError(Exception):
    """设备类型编码映射领域异常基类。"""


class EquipmentTypeConfigError(EquipmentTypeError):
    """设备类型映射配置不合法。"""


class EquipmentTypeStatus(str, Enum):
    """设备类型编码映射状态。"""

    MATCH = "MATCH"
    UNMATCHED = "UNMATCHED"
    CONFLICT = "CONFLICT"
    INVALID = "INVALID"
    SKIPPED = "SKIPPED"


@dataclass(frozen=True)
class EquipmentTypeRule:
    """一条设备类型编码映射规则。"""

    aliases: tuple[str, ...]
    standard_type_name: str
    system_code: str


@dataclass(frozen=True)
class EquipmentTypeCodeConfig:
    """设备类型编码映射配置。"""

    target_standard_field: str
    rules: tuple[EquipmentTypeRule, ...]
    code_pattern: str | None = None


@dataclass(frozen=True)
class EquipmentTypeMapAudit:
    """一条设备类型编码映射审计。"""

    sheet_name: str
    row_number: int
    source_column_number: int | None
    source_column_letter: str | None
    source_field_name: str
    original_value: Any
    cleaned_value: Any
    matched_standard_type: str | None
    system_code: str | None
    status: EquipmentTypeStatus
    reason: str


@dataclass(frozen=True)
class EquipmentTypeSummary:
    """设备类型映射统计。"""

    total: int = 0
    match_count: int = 0
    unmatched_count: int = 0
    conflict_count: int = 0
    invalid_count: int = 0
    skipped_count: int = 0


@dataclass(frozen=True)
class EquipmentTypeSheetResult:
    """一个 Sheet 的设备类型映射结果。"""

    sheet_name: str
    audits: tuple[EquipmentTypeMapAudit, ...]
    summary: EquipmentTypeSummary


@dataclass(frozen=True)
class EquipmentTypeWorkbookResult:
    """整个工作簿的设备类型映射结果。"""

    source_file: Path
    sheet_results: tuple[EquipmentTypeSheetResult, ...]
    summary: EquipmentTypeSummary


def load_equipment_type_config(path: str | Path) -> EquipmentTypeCodeConfig:
    """从 JSON 文件加载设备类型编码映射配置。

    JSON 结构：
    {
      "target_standard_field": "设备类型名称",
      "code_pattern": "D-[A-Z0-9]{3}",
      "rules": [
        {
          "aliases": ["除尘器", "除尘设备"],
          "standard_type_name": "除尘器",
          "system_code": "D-CCQ"
        }
      ]
    }
    """
    config_path = Path(path)
    if not config_path.exists() or not config_path.is_file():
        raise EquipmentTypeConfigError(
            f"设备类型映射配置文件不存在: {config_path}"
        )
    try:
        payload = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise EquipmentTypeConfigError(
            f"设备类型映射配置文件无法解析: {config_path}：{exc}"
        ) from exc
    if not isinstance(payload, dict):
        raise EquipmentTypeConfigError(
            "设备类型映射配置顶层必须是 JSON 对象"
        )

    target_field = payload.get("target_standard_field")
    if not isinstance(target_field, str) or not target_field.strip():
        raise EquipmentTypeConfigError(
            "target_standard_field 必须是非空字符串"
        )
    raw_rules = payload.get("rules")
    if not isinstance(raw_rules, list):
        raise EquipmentTypeConfigError("设备类型映射配置缺少 rules 列表")

    code_pattern = payload.get("code_pattern")
    if code_pattern is not None and not isinstance(code_pattern, str):
        raise EquipmentTypeConfigError("code_pattern 必须是字符串")

    rules: list[EquipmentTypeRule] = []
    for index, raw_rule in enumerate(raw_rules, start=1):
        if not isinstance(raw_rule, dict):
            raise EquipmentTypeConfigError(
                f"rules 第 {index} 项必须是对象"
            )
        aliases = _string_tuple(
            raw_rule.get("aliases"),
            f"rules 第 {index} 项 aliases",
        )
        standard_type = raw_rule.get("standard_type_name")
        if not isinstance(standard_type, str) or not standard_type.strip():
            raise EquipmentTypeConfigError(
                f"rules 第 {index} 项 standard_type_name 必须是非空字符串"
            )
        system_code = raw_rule.get("system_code")
        if not isinstance(system_code, str) or not system_code.strip():
            raise EquipmentTypeConfigError(
                f"rules 第 {index} 项 system_code 必须是非空字符串"
            )
        rules.append(
            EquipmentTypeRule(
                aliases=aliases,
                standard_type_name=standard_type.strip(),
                system_code=system_code.strip(),
            )
        )
    config = EquipmentTypeCodeConfig(
        target_standard_field=target_field.strip(),
        rules=tuple(rules),
        code_pattern=code_pattern,
    )
    _validate_config(config)
    return config


def map_equipment_type_codes(
    cleaning_result: CleaningWorkbookResult,
    config: EquipmentTypeCodeConfig,
) -> EquipmentTypeWorkbookResult:
    """按配置库把 CleaningWorkbookResult 中的目标字段映射为系统编码。"""
    _validate_config(config)
    sheet_results = tuple(
        _map_sheet(sheet_result, config)
        for sheet_result in cleaning_result.sheet_results
    )
    merged = _merge_summaries(
        tuple(result.summary for result in sheet_results)
    )
    return EquipmentTypeWorkbookResult(
        source_file=cleaning_result.source_file,
        sheet_results=sheet_results,
        summary=merged,
    )


def _map_sheet(
    sheet_result: CleanedSheetResult,
    config: EquipmentTypeCodeConfig,
) -> EquipmentTypeSheetResult:
    audits: list[EquipmentTypeMapAudit] = []
    lookup = _build_lookup(config)
    for row in sheet_result.rows:
        target_cells = [
            cell
            for cell in row.cells
            if cell.standard_field_name == config.target_standard_field
        ]
        if not target_cells:
            audits.append(
                _skipped_audit(
                    sheet_result.sheet_name,
                    row.source_row_number,
                    reason=(
                        f"本 Sheet 未找到目标标准字段"
                        f"“{config.target_standard_field}”，跳过编码映射"
                    ),
                )
            )
            continue

        target_cell = target_cells[0]
        if _is_unresolved_cleaning_action(target_cell.action):
            audits.append(
                _skipped_audit(
                    sheet_result.sheet_name,
                    row.source_row_number,
                    source_cell=target_cell,
                    reason=_unresolved_reason(target_cell),
                )
            )
            continue

        cleaned_value = target_cell.cleaned_value
        if _is_empty(cleaned_value):
            audits.append(
                EquipmentTypeMapAudit(
                    sheet_name=sheet_result.sheet_name,
                    row_number=row.source_row_number,
                    source_column_number=target_cell.source_column_number,
                    source_column_letter=target_cell.source_column_letter,
                    source_field_name=target_cell.source_field_name,
                    original_value=target_cell.original_value,
                    cleaned_value=cleaned_value,
                    matched_standard_type=None,
                    system_code=None,
                    status=EquipmentTypeStatus.UNMATCHED,
                    reason="空值不参与设备类型编码映射",
                )
            )
            continue

        normalized = _normalize_value(_value_to_text(cleaned_value))
        candidates = lookup.get(normalized, ())
        unique = _unique_rules(candidates)
        if not unique:
            audits.append(
                EquipmentTypeMapAudit(
                    sheet_name=sheet_result.sheet_name,
                    row_number=row.source_row_number,
                    source_column_number=target_cell.source_column_number,
                    source_column_letter=target_cell.source_column_letter,
                    source_field_name=target_cell.source_field_name,
                    original_value=target_cell.original_value,
                    cleaned_value=cleaned_value,
                    matched_standard_type=None,
                    system_code=None,
                    status=EquipmentTypeStatus.UNMATCHED,
                    reason=(
                        f"清洗后值“{cleaned_value}”未命中任何"
                        "设备类型映射规则"
                    ),
                )
            )
            continue
        if len(unique) > 1:
            labels = "、".join(
                f"{rule.standard_type_name}:{rule.system_code}"
                for rule in unique
            )
            audits.append(
                EquipmentTypeMapAudit(
                    sheet_name=sheet_result.sheet_name,
                    row_number=row.source_row_number,
                    source_column_number=target_cell.source_column_number,
                    source_column_letter=target_cell.source_column_letter,
                    source_field_name=target_cell.source_field_name,
                    original_value=target_cell.original_value,
                    cleaned_value=cleaned_value,
                    matched_standard_type=None,
                    system_code=None,
                    status=EquipmentTypeStatus.CONFLICT,
                    reason=(
                        f"清洗后值“{cleaned_value}”命中多个规则：{labels}，"
                        "无法自动选择"
                    ),
                )
            )
            continue

        matched_rule = unique[0]
        if config.code_pattern and not re.fullmatch(
            config.code_pattern, matched_rule.system_code
        ):
            audits.append(
                EquipmentTypeMapAudit(
                    sheet_name=sheet_result.sheet_name,
                    row_number=row.source_row_number,
                    source_column_number=target_cell.source_column_number,
                    source_column_letter=target_cell.source_column_letter,
                    source_field_name=target_cell.source_field_name,
                    original_value=target_cell.original_value,
                    cleaned_value=cleaned_value,
                    matched_standard_type=matched_rule.standard_type_name,
                    system_code=matched_rule.system_code,
                    status=EquipmentTypeStatus.INVALID,
                    reason=(
                        f"命中编码“{matched_rule.system_code}”不符合"
                        f"code_pattern“{config.code_pattern}”"
                    ),
                )
            )
            continue
        audits.append(
            EquipmentTypeMapAudit(
                sheet_name=sheet_result.sheet_name,
                row_number=row.source_row_number,
                source_column_number=target_cell.source_column_number,
                source_column_letter=target_cell.source_column_letter,
                source_field_name=target_cell.source_field_name,
                original_value=target_cell.original_value,
                cleaned_value=cleaned_value,
                matched_standard_type=matched_rule.standard_type_name,
                system_code=matched_rule.system_code,
                status=EquipmentTypeStatus.MATCH,
                reason=(
                    f"清洗后值“{cleaned_value}”映射为标准类型"
                    f"“{matched_rule.standard_type_name}”"
                    f"、编码“{matched_rule.system_code}”"
                ),
            )
        )
    return EquipmentTypeSheetResult(
        sheet_name=sheet_result.sheet_name,
        audits=tuple(audits),
        summary=_summary_from_audits(tuple(audits)),
    )


def _build_lookup(
    config: EquipmentTypeCodeConfig,
) -> dict[str, tuple[EquipmentTypeRule, ...]]:
    lookup: dict[str, list[EquipmentTypeRule]] = {}
    for rule in config.rules:
        for alias in rule.aliases:
            normalized = _normalize_value(alias)
            if normalized:
                lookup.setdefault(normalized, []).append(rule)
    return {
        normalized: tuple(rules)
        for normalized, rules in lookup.items()
    }


def _unique_rules(
    rules: tuple[EquipmentTypeRule, ...],
) -> tuple[EquipmentTypeRule, ...]:
    unique: list[EquipmentTypeRule] = []
    seen: set[tuple[str, str]] = set()
    for rule in rules:
        key = (rule.standard_type_name, rule.system_code)
        if key not in seen:
            seen.add(key)
            unique.append(rule)
    return tuple(unique)


def _is_unresolved_cleaning_action(action: CleaningAction) -> bool:
    return action in {
        CleaningAction.UNMAPPED,
        CleaningAction.AMBIGUOUS,
        CleaningAction.CONFLICT,
    }


def _unresolved_reason(cell: CleanCellAudit) -> str:
    if cell.action is CleaningAction.UNMAPPED:
        return "目标字段在字段映射中为 UNMATCHED，跳过编码映射"
    if cell.action is CleaningAction.AMBIGUOUS:
        return "目标字段在字段映射中为 AMBIGUOUS，跳过编码映射"
    if cell.action is CleaningAction.CONFLICT:
        return "目标字段在字段映射中为 CONFLICT，跳过编码映射"
    return "目标字段映射未就绪，跳过编码映射"


def _skipped_audit(
    sheet_name: str,
    row_number: int,
    *,
    source_cell: CleanCellAudit | None = None,
    reason: str,
) -> EquipmentTypeMapAudit:
    if source_cell is None:
        return EquipmentTypeMapAudit(
            sheet_name=sheet_name,
            row_number=row_number,
            source_column_number=None,
            source_column_letter=None,
            source_field_name="",
            original_value=None,
            cleaned_value=None,
            matched_standard_type=None,
            system_code=None,
            status=EquipmentTypeStatus.SKIPPED,
            reason=reason,
        )
    return EquipmentTypeMapAudit(
        sheet_name=sheet_name,
        row_number=row_number,
        source_column_number=source_cell.source_column_number,
        source_column_letter=source_cell.source_column_letter,
        source_field_name=source_cell.source_field_name,
        original_value=source_cell.original_value,
        cleaned_value=source_cell.cleaned_value,
        matched_standard_type=None,
        system_code=None,
        status=EquipmentTypeStatus.SKIPPED,
        reason=reason,
    )


def _validate_config(config: EquipmentTypeCodeConfig) -> None:
    if not isinstance(config, EquipmentTypeCodeConfig):
        raise EquipmentTypeConfigError(
            "config 必须是 EquipmentTypeCodeConfig"
        )
    if not isinstance(config.target_standard_field, str) or not (
        config.target_standard_field.strip()
    ):
        raise EquipmentTypeConfigError(
            "target_standard_field 必须是非空字符串"
        )
    if config.code_pattern is not None:
        try:
            re.compile(config.code_pattern)
        except re.error as exc:
            raise EquipmentTypeConfigError(
                f"code_pattern 不是合法正则: {config.code_pattern}"
            ) from exc

    seen_rules: set[tuple[tuple[str, ...], str, str]] = set()
    for rule in config.rules:
        if not isinstance(rule, EquipmentTypeRule):
            raise EquipmentTypeConfigError(
                "rules 只允许 EquipmentTypeRule"
            )
        if not isinstance(rule.standard_type_name, str) or not (
            rule.standard_type_name.strip()
        ):
            raise EquipmentTypeConfigError(
                "standard_type_name 必须是非空字符串"
            )
        if not isinstance(rule.system_code, str) or not rule.system_code.strip():
            raise EquipmentTypeConfigError("system_code 必须是非空字符串")
        if (
            not rule.aliases
            or not isinstance(rule.aliases, (tuple, list))
            or not all(isinstance(alias, str) and alias.strip() for alias in rule.aliases)
        ):
            raise EquipmentTypeConfigError(
                "aliases 必须是非空字符串元组/列表"
            )
        aliases = tuple(alias.strip() for alias in rule.aliases)
        key = (
            tuple(sorted(set(aliases))),
            rule.standard_type_name.strip(),
            rule.system_code.strip(),
        )
        if key in seen_rules:
            raise EquipmentTypeConfigError(
                "存在重复的设备类型映射规则"
            )
        seen_rules.add(key)


def _normalize_value(value: str) -> str:
    text = unicodedata.normalize("NFKC", value)
    text = _INVISIBLE_WHITESPACE_PATTERN.sub("", text)
    return " ".join(text.split()).strip().casefold()


def _is_empty(value: Any) -> bool:
    if value is None:
        return True
    return isinstance(value, str) and not value.strip()


def _string_tuple(value: Any, label: str) -> tuple[str, ...]:
    if (
        not isinstance(value, list)
        or not value
        or not all(
        isinstance(item, str) and item.strip()
        for item in value
        )
    ):
        raise EquipmentTypeConfigError(f"{label} 必须是非空字符串列表")
    return tuple(item.strip() for item in value)


def _summary_from_audits(
    audits: tuple[EquipmentTypeMapAudit, ...],
) -> EquipmentTypeSummary:
    counts = {status: 0 for status in EquipmentTypeStatus}
    for audit in audits:
        counts[audit.status] += 1
    return EquipmentTypeSummary(
        total=len(audits),
        match_count=counts[EquipmentTypeStatus.MATCH],
        unmatched_count=counts[EquipmentTypeStatus.UNMATCHED],
        conflict_count=counts[EquipmentTypeStatus.CONFLICT],
        invalid_count=counts[EquipmentTypeStatus.INVALID],
        skipped_count=counts[EquipmentTypeStatus.SKIPPED],
    )


def _merge_summaries(
    summaries: tuple[EquipmentTypeSummary, ...],
) -> EquipmentTypeSummary:
    return EquipmentTypeSummary(
        total=sum(item.total for item in summaries),
        match_count=sum(item.match_count for item in summaries),
        unmatched_count=sum(item.unmatched_count for item in summaries),
        conflict_count=sum(item.conflict_count for item in summaries),
        invalid_count=sum(item.invalid_count for item in summaries),
        skipped_count=sum(item.skipped_count for item in summaries),
    )


def _value_to_text(value: Any) -> str:
    """把清洗后的任意简单值转为文本以参与归一化匹配。"""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    return "" if value is None else str(value)
