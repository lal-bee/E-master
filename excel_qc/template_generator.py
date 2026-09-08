"""P1-06 标准模板生成引擎。

把 P1-02 结构探查、P1-03 字段映射、P1-04 数据清洗与 P1-05
设备类型编码映射的结果，确定性装配为内存标准模板，并可导出为
新的 .xlsx 文件。

本模块只做装配与导出：

- 普通字段只消费 P1-04 的 ``cleaned_value``；
- 设备类型编码列只消费 P1-05 的 ``system_code``（且仅 MATCH）；
- 不复算结构、字段映射、清洗或编码；
- 不做 AI、模糊匹配、编码猜测、质量门禁；
- 构建与导出分离，构建不写磁盘，导出绝不覆盖源文件。
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_qc.cleaning import (
    CleaningAction,
    CleaningWorkbookResult,
)
from excel_qc.equipment_type_mapping import (
    EquipmentTypeStatus,
    EquipmentTypeWorkbookResult,
)
from excel_qc.field_mapping import (
    FieldMappingStatus,
    WorkbookFieldMappingResult,
)
from excel_qc.profiler import WorkbookProfile


class TemplateGenerationError(Exception):
    """标准模板生成领域异常基类。"""


class TemplateConfigError(TemplateGenerationError):
    """模板配置或配置文件不合法。"""


class TemplateDataError(TemplateGenerationError):
    """模板装配输入数据与结构不一致。"""


class TemplateIssueSourceStage(str, Enum):
    """TemplateIssue 的来源阶段。"""

    FIELD_MAPPING = "FIELD_MAPPING"
    CLEANING = "CLEANING"
    EQUIPMENT_TYPE_MAPPING = "EQUIPMENT_TYPE_MAPPING"
    TEMPLATE = "TEMPLATE"


@dataclass(frozen=True)
class TemplateColumnConfig:
    """一个模板输出列。

    standard_field 引用 P1-03 映射结果中的标准字段名；普通列从
    P1-04 清洗结果取值。

    code_source_standard_field 不为 None 时，本列是设备类型系统编码列，
    只从 P1-05 结果取值（仅 MATCH 写入 system_code）。
    """

    standard_field: str
    header: str | None = None
    code_source_standard_field: str | None = None

    @property
    def is_code_column(self) -> bool:
        return self.code_source_standard_field is not None

    @property
    def output_header(self) -> str:
        """header 缺省时使用 standard_field 作为表头。"""
        return self.header if self.header is not None else self.standard_field


@dataclass(frozen=True)
class TemplateConfig:
    """标准模板配置。

    sheets 为可选参与输出的源 Sheet 名；缺省时按 WorkbookProfile
    的 Sheet 顺序全部输出。columns 为有序模板列。
    """

    columns: tuple[TemplateColumnConfig, ...] = ()
    sheets: tuple[str, ...] | None = None


@dataclass(frozen=True)
class TemplateCellIssue:
    """一条模板装配问题。

    结构性缺失（如整个 Sheet 缺少某标准字段）时 source_row_number
    为 None；单元格级问题携带源物理行号。
    """

    sheet: str
    source_row_number: int | None
    standard_field: str
    template_column: str
    issue_code: str
    reason: str
    source_stage: TemplateIssueSourceStage


@dataclass(frozen=True)
class TemplateCellTrace:
    """一个输出单元格的装配层来源引用。

    P1-06 不复制 P1-03/P1-04/P1-05 的审计对象，只保留定位与关键值，
    保证每个模板单元格都能沿来源链回查。
    """

    sheet: str
    source_row_number: int | None
    source_column_number: int | None
    source_column_letter: str | None
    source_field_name: str
    standard_field: str
    template_column: str
    original_value: Any
    cleaned_value: Any
    system_code: str | None = None


@dataclass(frozen=True)
class TemplateRow:
    """模板输出行：值与追溯均与模板列顺序对齐。"""

    source_row_number: int
    values: tuple[Any, ...]
    trace: tuple[TemplateCellTrace, ...]


@dataclass(frozen=True)
class TemplateSheetResult:
    """一个输出 Sheet 的模板装配结果。"""

    sheet_name: str
    source_sheet_name: str
    columns: tuple[TemplateColumnConfig, ...]
    rows: tuple[TemplateRow, ...]
    issues: tuple[TemplateCellIssue, ...]


@dataclass(frozen=True)
class TemplateSummary:
    """模板装配统计。"""

    sheet_count: int = 0
    row_count: int = 0
    cell_count: int = 0
    issue_count: int = 0


@dataclass(frozen=True)
class TemplateWorkbookResult:
    """整个源工作簿的标准模板装配结果。"""

    source_file: Path
    sheet_results: tuple[TemplateSheetResult, ...]
    summary: TemplateSummary
    issues: tuple[TemplateCellIssue, ...]


def load_template_config(path: str | Path) -> TemplateConfig:
    """从 JSON 文件加载标准模板配置。

    JSON 结构：
    {
      "sheets": ["设备档案"],          # 可选
      "columns": [
        {"standard_field": "设备名称", "header": "设备名称"},
        {
          "standard_field": "设备类型编码",
          "header": "设备类型编码",
          "code_source_standard_field": "设备类型名称"
        }
      ]
    }

    header 缺省时使用 standard_field。该结构仅描述通用模板引擎配置，
    不代表任何真实业务最终模板。
    """
    config_path = Path(path)
    if not config_path.exists() or not config_path.is_file():
        raise TemplateConfigError(f"模板配置文件不存在: {config_path}")
    try:
        payload = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise TemplateConfigError(
            f"模板配置文件无法解析: {config_path}：{exc}"
        ) from exc
    if not isinstance(payload, dict):
        raise TemplateConfigError("模板配置顶层必须是 JSON 对象")

    sheets: tuple[str, ...] | None = None
    raw_sheets = payload.get("sheets")
    if raw_sheets is not None:
        if not isinstance(raw_sheets, list) or not raw_sheets:
            raise TemplateConfigError(
                "sheets 必须是包含 Sheet 名的非空列表"
            )
        if not all(isinstance(item, str) and item.strip() for item in raw_sheets):
            raise TemplateConfigError("sheets 中的 Sheet 名必须是非空字符串")
        sheets = tuple(item.strip() for item in raw_sheets)
        if len(set(sheets)) != len(sheets):
            raise TemplateConfigError("sheets 中存在重复的 Sheet 名")

    raw_columns = payload.get("columns")
    if not isinstance(raw_columns, list) or not raw_columns:
        raise TemplateConfigError("模板配置缺少非空 columns 列表")

    columns: list[TemplateColumnConfig] = []
    for index, raw_column in enumerate(raw_columns, start=1):
        if not isinstance(raw_column, dict):
            raise TemplateConfigError(f"columns 第 {index} 项必须是对象")
        standard_field = raw_column.get("standard_field")
        if not isinstance(standard_field, str) or not standard_field.strip():
            raise TemplateConfigError(
                f"columns 第 {index} 项 standard_field 必须是非空字符串"
            )

        header = raw_column.get("header")
        if header is not None and (
            not isinstance(header, str) or not header.strip()
        ):
            raise TemplateConfigError(
                f"columns 第 {index} 项 header 必须是非空字符串"
            )

        code_source = raw_column.get("code_source_standard_field")
        if code_source is not None and (
            not isinstance(code_source, str) or not code_source.strip()
        ):
            raise TemplateConfigError(
                f"columns 第 {index} 项 code_source_standard_field 配置不完整，"
                "必须是非空字符串"
            )
        columns.append(
            TemplateColumnConfig(
                standard_field=standard_field.strip(),
                header=header.strip() if isinstance(header, str) else None,
                code_source_standard_field=(
                    code_source.strip()
                    if isinstance(code_source, str)
                    else None
                ),
            )
        )

    config = TemplateConfig(columns=tuple(columns), sheets=sheets)
    _validate_config(config)
    return config


def build_standard_template(
    profile: WorkbookProfile,
    mapping: WorkbookFieldMappingResult,
    cleaning_result: CleaningWorkbookResult,
    equipment_result: EquipmentTypeWorkbookResult,
    config: TemplateConfig,
) -> TemplateWorkbookResult:
    """把同一源文件的四个阶段结果装配为内存标准模板。

    - 只在内存运行，不创建文件、不修改任何输入对象；
    - 严格按 profile 的 Sheet 顺序处理；
    - 每行保留 source_row_number，含内部空行；
    - 普通列使用 P1-04 cleaned_value；编码列仅 MATCH 时写入 system_code。
    """
    _validate_config(config)
    _validate_input_consistency(
        profile=profile,
        mapping=mapping,
        cleaning_result=cleaning_result,
        equipment_result=equipment_result,
    )

    selected_sheet_names = (
        set(config.sheets)
        if config.sheets is not None
        else {sheet.sheet_name for sheet in profile.sheets}
    )
    if not selected_sheet_names:
        raise TemplateDataError("模板配置未选择任何 Sheet")

    mapping_by_sheet = {
        result.sheet_name: result for result in mapping.sheet_results
    }
    cleaning_by_sheet = {
        result.sheet_name: result for result in cleaning_result.sheet_results
    }
    equipment_by_sheet = {
        result.sheet_name: result for result in equipment_result.sheet_results
    }

    sheet_results: list[TemplateSheetResult] = []
    for sheet_profile in profile.sheets:
        if sheet_profile.sheet_name not in selected_sheet_names:
            continue
        sheet_results.append(
            _build_sheet(
                sheet_profile=sheet_profile,
                mapping_sheet=mapping_by_sheet[sheet_profile.sheet_name],
                cleaning_sheet=cleaning_by_sheet[sheet_profile.sheet_name],
                equipment_sheet=equipment_by_sheet[sheet_profile.sheet_name],
                config=config,
            )
        )

    if not sheet_results:
        raise TemplateDataError("没有可生成的标准模板 Sheet")

    issues = tuple(issue for result in sheet_results for issue in result.issues)
    summary = TemplateSummary(
        sheet_count=len(sheet_results),
        row_count=sum(len(result.rows) for result in sheet_results),
        cell_count=sum(
            len(result.rows) * len(result.columns)
            for result in sheet_results
        ),
        issue_count=len(issues),
    )
    return TemplateWorkbookResult(
        source_file=profile.path,
        sheet_results=tuple(sheet_results),
        summary=summary,
        issues=issues,
    )


def export_standard_template(
    result: TemplateWorkbookResult,
    output_path: str | Path,
) -> Path:
    """把内存模板结果导出为新的 .xlsx 文件。

    - 只写表头、列顺序与数据值，可选冻结表头，不做样式美化；
    - 输出路径与源文件相同时抛 TemplateGenerationError；
    - 输出文件父目录不存在时自动创建；
    - 不修改任何源 Excel。
    """
    if not isinstance(result, TemplateWorkbookResult):
        raise TemplateGenerationError(
            "result 必须是 TemplateWorkbookResult"
        )
    if not result.sheet_results:
        raise TemplateGenerationError("没有可导出的模板 Sheet")

    output = Path(output_path)
    if output.resolve() == result.source_file.resolve():
        raise TemplateGenerationError(
            f"标准模板输出不能覆盖源文件: {output}"
        )

    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    workbook.remove(workbook.active)
    for sheet_result in result.sheet_results:
        worksheet = workbook.create_sheet(title=sheet_result.sheet_name)
        _write_sheet(worksheet, sheet_result)
    workbook.save(output)
    return output


def _write_sheet(sheet: Worksheet, result: TemplateSheetResult) -> None:
    headers = [column.output_header for column in result.columns]
    if headers:
        sheet.append(headers)
        sheet.freeze_panes = "A2"
    for row in result.rows:
        sheet.append(list(row.values))


def _validate_config(config: TemplateConfig) -> None:
    if not isinstance(config, TemplateConfig):
        raise TemplateConfigError("config 必须是 TemplateConfig")
    if not isinstance(config.columns, tuple) or not config.columns:
        raise TemplateConfigError("模板配置必须包含至少一列")
    if config.sheets is not None:
        if not isinstance(config.sheets, tuple) or not config.sheets:
            raise TemplateConfigError(
                "sheets 必须是包含 Sheet 名的非空元组"
            )
        if not all(
            isinstance(item, str) and item.strip() for item in config.sheets
        ):
            raise TemplateConfigError(
                "sheets 中的 Sheet 名必须是非空字符串"
            )
        if len(set(config.sheets)) != len(config.sheets):
            raise TemplateConfigError("sheets 中存在重复的 Sheet 名")

    seen_fields: dict[str, int] = {}
    seen_headers: dict[str, int] = {}
    for index, column in enumerate(config.columns, start=1):
        if not isinstance(column, TemplateColumnConfig):
            raise TemplateConfigError(
                f"columns 第 {index} 项必须是 TemplateColumnConfig"
            )
        if not isinstance(column.standard_field, str) or not (
            column.standard_field.strip()
        ):
            raise TemplateConfigError(
                f"columns 第 {index} 项 standard_field 必须是非空字符串"
            )
        header = column.output_header
        if not isinstance(header, str) or not header.strip():
            raise TemplateConfigError(
                f"columns 第 {index} 项 header 必须是非空字符串"
            )
        field_key = _identity_key(column.standard_field)
        header_key = _identity_key(header)
        if field_key in seen_fields:
            raise TemplateConfigError(
                f"模板列 standard_field 重复: "
                f"“{column.standard_field}”（第 {index} 项与第 "
                f"{seen_fields[field_key]} 项）"
            )
        if header_key in seen_headers:
            raise TemplateConfigError(
                f"模板列表头重复: “{header}”（第 {index} 项与第 "
                f"{seen_headers[header_key]} 项）"
            )
        seen_fields[field_key] = index
        seen_headers[header_key] = index
        if column.code_source_standard_field is not None:
            if (
                not isinstance(column.code_source_standard_field, str)
                or not column.code_source_standard_field.strip()
            ):
                raise TemplateConfigError(
                    f"columns 第 {index} 项 code_source_standard_field "
                    "配置不完整，必须是非空字符串"
                )


def _validate_input_consistency(
    *,
    profile: WorkbookProfile,
    mapping: WorkbookFieldMappingResult,
    cleaning_result: CleaningWorkbookResult,
    equipment_result: EquipmentTypeWorkbookResult,
) -> None:
    if not isinstance(profile, WorkbookProfile):
        raise TemplateDataError("profile 必须是 WorkbookProfile")
    if not isinstance(mapping, WorkbookFieldMappingResult):
        raise TemplateDataError(
            "mapping 必须是 WorkbookFieldMappingResult"
        )
    if not isinstance(cleaning_result, CleaningWorkbookResult):
        raise TemplateDataError(
            "cleaning_result 必须是 CleaningWorkbookResult"
        )
    if not isinstance(equipment_result, EquipmentTypeWorkbookResult):
        raise TemplateDataError(
            "equipment_result 必须是 EquipmentTypeWorkbookResult"
        )

    profile_names = tuple(sheet.sheet_name for sheet in profile.sheets)
    mapping_names = tuple(result.sheet_name for result in mapping.sheet_results)
    cleaning_names = tuple(
        result.sheet_name for result in cleaning_result.sheet_results
    )
    equipment_names = tuple(
        result.sheet_name for result in equipment_result.sheet_results
    )
    if not (
        profile_names == mapping_names == cleaning_names == equipment_names
    ):
        raise TemplateDataError(
            "profile/mapping/cleaning/equipment 的 Sheet 名称或顺序不一致"
        )

    source_path = Path(profile.path).resolve()
    for label, other in (
        ("mapping", mapping.file_path),
        ("cleaning_result", cleaning_result.source_file),
        ("equipment_result", equipment_result.source_file),
    ):
        if Path(other).resolve() != source_path:
            raise TemplateDataError(
                f"{label} 与 profile 不属于同一源文件"
            )

    cleaning_by_sheet = {
        result.sheet_name: result for result in cleaning_result.sheet_results
    }
    equipment_by_sheet = {
        result.sheet_name: result for result in equipment_result.sheet_results
    }
    for sheet_profile in profile.sheets:
        expected_rows = _expected_row_numbers(sheet_profile)
        cleaning_sheet = cleaning_by_sheet[sheet_profile.sheet_name]
        equipment_sheet = equipment_by_sheet[sheet_profile.sheet_name]
        cleaning_rows = tuple(
            row.source_row_number for row in cleaning_sheet.rows
        )
        equipment_rows = tuple(
            audit.row_number for audit in equipment_sheet.audits
        )
        if cleaning_rows != expected_rows:
            raise TemplateDataError(
                f"Sheet“{sheet_profile.sheet_name}”清洗结果行覆盖范围不一致"
            )
        if equipment_rows != expected_rows:
            raise TemplateDataError(
                f"Sheet“{sheet_profile.sheet_name}”编码映射结果行覆盖范围不一致"
            )


def _expected_row_numbers(sheet_profile) -> tuple[int, ...]:
    if (
        sheet_profile.data_first_row is None
        or sheet_profile.data_last_row is None
    ):
        return ()
    return tuple(
        range(sheet_profile.data_first_row, sheet_profile.data_last_row + 1)
    )


def _build_sheet(
    *,
    sheet_profile,
    mapping_sheet,
    cleaning_sheet,
    equipment_sheet,
    config: TemplateConfig,
) -> TemplateSheetResult:
    columns = config.columns
    matched_fields = {
        entry.standard_field_name
        for entry in mapping_sheet.entries
        if (
            entry.status is FieldMappingStatus.MATCH
            and entry.standard_field_name
        )
    }
    entry_by_standard_field = {
        entry.standard_field_name: entry
        for entry in mapping_sheet.entries
        if entry.standard_field_name
    }

    issues: list[TemplateCellIssue] = []
    missing_issue_emitted: set[tuple[str, str]] = set()

    if cleaning_sheet.rows:
        for column in columns:
            if column.is_code_column:
                continue
            if column.standard_field in matched_fields:
                continue
            if (cleaning_sheet.sheet_name, column.standard_field) in (
                missing_issue_emitted
            ):
                continue
            entry = entry_by_standard_field.get(column.standard_field)
            if entry is not None and (
                entry.status is FieldMappingStatus.CONFLICT
            ):
                issue_code = "FIELD_MAPPING_CONFLICT"
                reason = entry.reason
            else:
                issue_code = "FIELD_MAPPING_MISSING"
                reason = (
                    f"模板列引用的标准字段“{column.standard_field}”"
                    "在当前 Sheet 没有可用的 MATCH 映射，"
                    "对应单元格输出为空且不做任何猜测"
                )
            issues.append(
                TemplateCellIssue(
                    sheet=cleaning_sheet.sheet_name,
                    source_row_number=None,
                    standard_field=column.standard_field,
                    template_column=column.output_header,
                    issue_code=issue_code,
                    reason=reason,
                    source_stage=TemplateIssueSourceStage.FIELD_MAPPING,
                )
            )
            missing_issue_emitted.add(
                (cleaning_sheet.sheet_name, column.standard_field)
            )

    equipment_by_row = {
        audit.row_number: audit for audit in equipment_sheet.audits
    }
    rows: list[TemplateRow] = []
    for cleaning_row in cleaning_sheet.rows:
        cells_by_standard = {
            cell.standard_field_name: cell
            for cell in cleaning_row.cells
            if cell.standard_field_name
        }
        values: list[Any] = []
        traces: list[TemplateCellTrace] = []
        for column in columns:
            if column.is_code_column:
                value, trace, row_issue = _assemble_code_cell(
                    sheet_name=cleaning_sheet.sheet_name,
                    source_row_number=cleaning_row.source_row_number,
                    column=column,
                    audit=equipment_by_row.get(cleaning_row.source_row_number),
                )
            else:
                value, trace, row_issue = _assemble_clean_cell(
                    sheet_name=cleaning_sheet.sheet_name,
                    source_row_number=cleaning_row.source_row_number,
                    column=column,
                    cell=cells_by_standard.get(column.standard_field),
                    is_matched=column.standard_field in matched_fields,
                )
            values.append(value)
            traces.append(trace)
            if row_issue is not None:
                issues.append(row_issue)
        rows.append(
            TemplateRow(
                source_row_number=cleaning_row.source_row_number,
                values=tuple(values),
                trace=tuple(traces),
            )
        )

    return TemplateSheetResult(
        sheet_name=cleaning_sheet.sheet_name,
        source_sheet_name=cleaning_sheet.sheet_name,
        columns=columns,
        rows=tuple(rows),
        issues=tuple(issues),
    )


def _assemble_clean_cell(
    *,
    sheet_name: str,
    source_row_number: int,
    column: TemplateColumnConfig,
    cell,
    is_matched: bool,
) -> tuple[Any, TemplateCellTrace, TemplateCellIssue | None]:
    template_column = column.output_header
    if not is_matched or cell is None:
        return (
            None,
            _empty_trace(
                sheet=sheet_name,
                source_row_number=source_row_number,
                standard_field=column.standard_field,
                template_column=template_column,
            ),
            None,
        )
    issue = None
    if cell.action is CleaningAction.INVALID:
        issue = TemplateCellIssue(
            sheet=sheet_name,
            source_row_number=source_row_number,
            standard_field=column.standard_field,
            template_column=template_column,
            issue_code="CLEANING_INVALID",
            reason=cell.reason,
            source_stage=TemplateIssueSourceStage.CLEANING,
        )
    trace = TemplateCellTrace(
        sheet=sheet_name,
        source_row_number=source_row_number,
        source_column_number=cell.source_column_number,
        source_column_letter=cell.source_column_letter,
        source_field_name=cell.source_field_name,
        standard_field=column.standard_field,
        template_column=template_column,
        original_value=cell.original_value,
        cleaned_value=cell.cleaned_value,
        system_code=None,
    )
    return cell.cleaned_value, trace, issue


def _assemble_code_cell(
    *,
    sheet_name: str,
    source_row_number: int,
    column: TemplateColumnConfig,
    audit,
) -> tuple[Any, TemplateCellTrace, TemplateCellIssue | None]:
    template_column = column.output_header
    standard_field = column.code_source_standard_field or ""
    if audit is None:
        return (
            None,
            _empty_trace(
                sheet=sheet_name,
                source_row_number=source_row_number,
                standard_field=standard_field,
                template_column=template_column,
            ),
            TemplateCellIssue(
                sheet=sheet_name,
                source_row_number=source_row_number,
                standard_field=standard_field,
                template_column=template_column,
                issue_code="EQUIPMENT_AUDIT_MISSING",
                reason="当前行缺少设备类型编码映射审计记录，编码输出为空",
                source_stage=TemplateIssueSourceStage.EQUIPMENT_TYPE_MAPPING,
            ),
        )

    trace = TemplateCellTrace(
        sheet=sheet_name,
        source_row_number=source_row_number,
        source_column_number=audit.source_column_number,
        source_column_letter=audit.source_column_letter,
        source_field_name=audit.source_field_name,
        standard_field=standard_field,
        template_column=template_column,
        original_value=audit.original_value,
        cleaned_value=audit.cleaned_value,
        system_code=audit.system_code,
    )
    if audit.status is EquipmentTypeStatus.MATCH:
        return audit.system_code, trace, None

    issue_code_by_status = {
        EquipmentTypeStatus.UNMATCHED: "EQUIPMENT_UNMATCHED",
        EquipmentTypeStatus.CONFLICT: "EQUIPMENT_CONFLICT",
        EquipmentTypeStatus.INVALID: "EQUIPMENT_INVALID",
        EquipmentTypeStatus.SKIPPED: "EQUIPMENT_SKIPPED",
    }
    issue = TemplateCellIssue(
        sheet=sheet_name,
        source_row_number=source_row_number,
        standard_field=standard_field,
        template_column=template_column,
        issue_code=issue_code_by_status.get(
            audit.status, "EQUIPMENT_UNRESOLVED"
        ),
        reason=audit.reason,
        source_stage=TemplateIssueSourceStage.EQUIPMENT_TYPE_MAPPING,
    )
    return None, trace, issue


def _empty_trace(
    *,
    sheet: str,
    source_row_number: int | None,
    standard_field: str,
    template_column: str,
) -> TemplateCellTrace:
    return TemplateCellTrace(
        sheet=sheet,
        source_row_number=source_row_number,
        source_column_number=None,
        source_column_letter=None,
        source_field_name="",
        standard_field=standard_field,
        template_column=template_column,
        original_value=None,
        cleaned_value=None,
        system_code=None,
    )


def _identity_key(value: str) -> str:
    return " ".join(value.split())
