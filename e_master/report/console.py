"""控制台摘要渲染。"""

from __future__ import annotations

from e_master.core.models.profile import IssueLevel, WorkbookProfile


def render_console_summary(profile: WorkbookProfile) -> list[str]:
    """生成适合控制台输出的摘要行。"""
    lines = [
        f"文件: {profile.path}",
        f"类型: {profile.file_type} | Sheet 数: {profile.total_sheets} | 问题总数: {len(profile.issues)}",
    ]
    for sheet in profile.sheets:
        header = (
            f"推荐表头第 {sheet.recommended_header.row_number} 行"
            if sheet.recommended_header
            else "未识别表头"
        )
        errors = sum(1 for issue in sheet.issues if issue.level == IssueLevel.ERROR)
        warnings = sum(1 for issue in sheet.issues if issue.level == IssueLevel.WARNING)
        lines.append(
            f"[{sheet.sheet_name}] {header} | 数据行 {sheet.data_row_count} | "
            f"列 {sheet.column_count} | 空单元格 {sheet.empty_cell_count} | "
            f"重复行 {sheet.duplicate_data_rows} | 问题 {len(sheet.issues)} "
            f"(警告 {warnings}，错误 {errors})"
        )
    return lines
