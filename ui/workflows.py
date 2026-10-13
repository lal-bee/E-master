"""P1-12 薄编排层：只调用领域公共接口，管理输入版本与导出资格。"""
from __future__ import annotations

import json
import math
import os
from dataclasses import dataclass, fields, is_dataclass, replace
from datetime import date, datetime
from decimal import Decimal
from enum import Enum
from pathlib import Path
from threading import Lock
from uuid import uuid4

from excel_qc import (
    CleaningWorkbookResult, EquipmentTypeWorkbookResult, QualityWorkbookResult,
    TemplateWorkbookResult, WorkbookFieldMappingResult, WorkbookProfile,
    ComparisonRuleKind, FieldComparisonRule, OriginalSourceConfig,
    SourceVerificationConfig, build_standard_template, clean_workbook,
    export_standard_template, load_cleaning_config, load_equipment_type_config,
    load_field_mapping_config, load_quality_gate_config, load_system_export_config,
    load_template_config, map_equipment_type_codes, map_workbook_fields,
    profile_configured_sheet, profile_workbook, run_quality_gate,
    verify_source_against_system_export,
)
from ui.source_verification_report import export_source_verification_report


class WorkflowError(Exception):
    """可直接显示给用户的编排错误。"""


ROLES = {
    "prepare": {"source": ".xlsx", "mapping": ".json", "cleaning": ".json",
                "equipment": ".json", "template": ".json", "gate": ".json"},
    "verify": {"source": ".xlsx", "system": ".xlsx", "original_config": ".json",
               "system_config": ".json", "verification": ".json"},
}


def display_value(value):
    """展示模型保留类型、空值与组合边界，不用于计算结论。"""
    if isinstance(value, Enum):
        return value.value
    if isinstance(value, Path):
        return str(value)
    if is_dataclass(value):
        return {f.name: display_value(getattr(value, f.name)) for f in fields(value)}
    if isinstance(value, dict):
        return {str(k): display_value(v) for k, v in value.items()}
    if isinstance(value, (tuple, list)):
        return [display_value(v) for v in value]
    if isinstance(value, (date, datetime, Decimal)):
        return {"type": type(value).__name__, "value": str(value)}
    if isinstance(value, float) and not math.isfinite(value):
        return {"type": "float", "value": repr(value)}
    if isinstance(value, int) and not isinstance(value, bool) and abs(value) > 2**53 - 1:
        return {"type": "int", "value": str(value)}
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return {"type": type(value).__name__, "value": repr(value)}


def read_json(path):
    try:
        payload = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise WorkflowError(f"JSON 格式错误：第 {exc.lineno} 行，第 {exc.colno} 列：{exc.msg}") from exc
    if not isinstance(payload, dict):
        raise WorkflowError("JSON 顶层必须是对象")
    return payload


def _field_list(payload, key):
    value = payload.get(key)
    if not isinstance(value, list) or not value or not all(isinstance(v, str) and v.strip() for v in value):
        raise WorkflowError(f"核验配置 {key} 必须是非空字段名列表")
    return tuple(value)


def load_verification_config(original_path, system_path, verification_path):
    """源侧接入 JSON 复用 P1-08 格式/加载器，转换为原始源身份。"""
    original = load_system_export_config(original_path)
    system = load_system_export_config(system_path)
    payload = read_json(verification_path)
    if set(payload) - {"key_fields", "comparison_fields", "comparison_rules"}:
        raise WorkflowError("核验配置包含未知字段；请检查 key_fields/comparison_fields/comparison_rules")
    raw_rules = payload.get("comparison_rules", [])
    if not isinstance(raw_rules, list):
        raise WorkflowError("comparison_rules 必须是列表")
    rules = []
    for index, item in enumerate(raw_rules, 1):
        if not isinstance(item, dict) or set(item) - {"standard_field", "kind"}:
            raise WorkflowError(f"comparison_rules 第 {index} 项格式错误")
        name = item.get("standard_field")
        if not isinstance(name, str) or not name.strip():
            raise WorkflowError(f"comparison_rules 第 {index} 项缺少 standard_field")
        try:
            rules.append(FieldComparisonRule(name, ComparisonRuleKind(item.get("kind", "EXACT"))))
        except (ValueError, TypeError) as exc:
            raise WorkflowError(f"comparison_rules 第 {index} 项 kind 无效，允许 EXACT/TEXT/NUMBER/DATE/BOOLEAN") from exc
    keys = _field_list(payload, "key_fields")
    if original.primary_key_fields != keys:
        raise WorkflowError("原始源接入配置 primary_key_fields 必须与核验 key_fields 一致")
    return SourceVerificationConfig(
        OriginalSourceConfig(original.sheet_name, original.header_row_number,
                             original.field_mapping, original.cleaning),
        system, keys, _field_list(payload, "comparison_fields"), tuple(rules),
    )


@dataclass(frozen=True)
class PreparationResult:
    profile: WorkbookProfile
    mapping: WorkbookFieldMappingResult
    cleaning: CleaningWorkbookResult
    equipment: EquipmentTypeWorkbookResult
    template: TemplateWorkbookResult
    gate: QualityWorkbookResult


def prepare_source(source, settings, configs):
    profile = profile_configured_sheet(source, settings["sheet"], settings["header"])
    mapping = map_workbook_fields(profile, load_field_mapping_config(configs["mapping"]))
    cleaning = clean_workbook(profile, mapping, load_cleaning_config(configs["cleaning"]))
    equipment = map_equipment_type_codes(cleaning, load_equipment_type_config(configs["equipment"]))
    template_config = load_template_config(configs["template"])
    if template_config.sheets is not None and template_config.sheets != (settings["sheet"],):
        raise WorkflowError("模板配置 sheets 必须仅包含当前所选 Sheet，或省略 sheets")
    template = build_standard_template(profile, mapping, cleaning, equipment,
                                       replace(template_config, sheets=(settings["sheet"],)))
    gate = run_quality_gate(profile, mapping, cleaning, equipment, template,
                            load_quality_gate_config(configs["gate"]))
    return PreparationResult(profile, mapping, cleaning, equipment, template, gate)


def user_message(exc):
    if isinstance(exc, FileNotFoundError):
        return "文件不存在，请重新选择文件。"
    if isinstance(exc, PermissionError):
        return "无法访问文件，请检查权限并关闭正在占用输出文件的 Excel。"
    if isinstance(exc, (WorkflowError, OSError)) or exc.__class__.__module__.startswith("excel_qc"):
        return str(exc)
    return "处理失败，请检查文件和配置是否匹配；若仍失败，请联系维护人员查看本地终端记录。"


class WorkflowSession:
    """上传只创建新副本；输出始终在 exports 下，来源副本不提前释放。"""

    def __init__(self, root: Path):
        self.root = root
        self.files = {}
        self.settings = {}
        self.revision = 0
        self.result_revision = None
        self.result = None
        self.status = "尚未运行"
        self.lock = Lock()

    def invalidate(self):
        self.revision += 1
        self.result_revision = None
        self.result = None
        self.status = "输入已更换，旧结果失效，请重新运行"

    def upload(self, flow, role, name, data):
        self.clear_input(flow, role)
        name = Path(name.replace("\\", "/")).name
        if not name or Path(name).suffix.lower() != ROLES[flow][role]:
            raise WorkflowError(f"此文件角色只接受 {ROLES[flow][role]} 文件")
        if not data:
            raise WorkflowError("所选文件为空")
        if len(data) > 32 * 1024 * 1024:
            raise WorkflowError("文件超过 32 MB，请使用单文件小样本")
        path = self.root / "inputs" / role / uuid4().hex / name
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("xb") as file:
            file.write(data)
        if path.suffix.lower() == ".json":
            preview = read_json(path)
        else:
            preview = display_value(profile_workbook(path))
        self.files[role] = path  # 仅探查 / 解析成功后才启用新输入
        return {"name": name, "preview": preview, "revision": self.revision}

    def clear_input(self, flow, role):
        """选文件时先撤销旧角色，读取/传输失败也不能重新使用旧输入。"""
        if role not in ROLES[flow]:
            raise WorkflowError("未知文件角色")
        self.invalidate()
        self.files.pop(role, None)

    def configure(self, settings):
        if settings != self.settings:
            self.invalidate()
            self.settings = dict(settings)

    def run(self, flow, expected_revision):
        if expected_revision != self.revision:
            raise WorkflowError("输入版本已改变，请重新运行")
        self.result = None
        self.result_revision = None
        self.status = "运行中"
        try:
            missing = set(ROLES[flow]) - self.files.keys()
            if missing:
                raise WorkflowError("文件未选择完整，请检查页面中的 Excel 和 JSON 配置")
            if flow == "prepare":
                result = prepare_source(self.files["source"], self.settings, self.files)
            else:
                config = load_verification_config(self.files["original_config"], self.files["system_config"],
                                                  self.files["verification"])
                result = verify_source_against_system_export(self.files["source"], self.files["system"], config)
            self.result = result
            self.result_revision = self.revision
            self.status = "处理结束"
            return {"result": display_value(result), "revision": self.revision,
                    "exportable": flow == "verify" or result.gate.can_proceed}
        except Exception:
            self.status = "运行失败，请检查提示后重新运行"
            raise

    def export(self, flow, expected_revision):
        if self.result is None or self.result_revision != self.revision or expected_revision != self.revision:
            raise WorkflowError("尚无当前输入的有效结果，请先运行")
        if flow == "prepare" and not self.result.gate.can_proceed:
            raise WorkflowError("质量门禁 FAIL，禁止导出正式导入模板；请先修正问题")
        output = self.root / "exports" / f"{flow}-{uuid4().hex}.xlsx"
        output.parent.mkdir(parents=True, exist_ok=True)
        # 输出名称由服务生成，既有输出不覆盖；也不接受用户提供的磁盘路径。
        for input_path in self.files.values():
            if output.resolve() == input_path.resolve() or (output.exists() and os.path.samefile(output, input_path)):
                raise WorkflowError("导出不得覆盖输入文件")
        if flow == "prepare":
            export_standard_template(self.result.template, output)
        else:
            export_source_verification_report(self.result, output)
        return output
