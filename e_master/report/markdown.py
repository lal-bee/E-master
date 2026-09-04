"""把结构画像渲染为 Markdown 报告。"""

from __future__ import annotations

from datetime import datetime

from e_master.core.models.profile import IssueLevel, SheetProfile, WorkbookProfile


def _md(text: str, limit: int = 80) -> str:
    text = str(text).replace("\r", " ").replace("\n", " ").replace("|", "\\|")
    return text if len(text) <= limit else text[: limit - 1] + "…"


def _region(profile: SheetProfile) -> str:
    if profile.used_first_row is None or profile.used_last_col is None:
        return "—"
    return f"第 {profile.used_first_row}–{profile.used_last_row} 行"


def render_markdown(profile: WorkbookProfile, generated_at: datetime | None = None) -> str:
    """渲染为可读的结构报告。"""
    timestamp = (generated_at or datetime.now()).strftime("%Y-%m-%d %H:%M:%S")
    lines: list[str] = [
        "# 源文件结构分析报告",
        "",
        f"- 源文件：`{profile.path.resolve()}`",
        f"- 文件类型：{profile.file_type}",
        f"- Sheet 数量：{profile.total_sheets}",
        f"- 生成时间：{timestamp}",
        f"- 问题总数：{len(profile.issues)}",
        "",
    ]

    if profile.sheets:
        lines.extend(["## Sheet 汇总", ""])
        lines.extend(
            [
                "| Sheet | 数据区域 | 推荐表头 | 数据行 | 列数 | 问题数 |",
                "|---|---|---|---|---|---|",
            ]
        )
        for sheet in profile.sheets:
            header = (
                f"第 {sheet.recommended_header.row_number} 行"
                if sheet.recommended_header
                else "未识别"
            )
            lines.append(
                f"| {_md(sheet.sheet_name)} | {_region(sheet)} | {header} "
                f"| {sheet.data_row_count} | {sheet.column_count} | {len(sheet.issues)} |"
            )
        lines.append("")

    for index, sheet in enumerate(profile.sheets, start=1):
        lines.extend([f"## Sheet {index}：{_md(sheet.sheet_name)}", ""])
        if sheet.used_first_row is None:
            lines.append("该 Sheet 为空。")
            _append_issues(lines, sheet.issues)
            continue

        lines.extend(
            [
                f"- 有内容区域：第 {sheet.used_first_row}–{sheet.used_last_row} 行，"
                f"{_column_range(sheet)} 列",
                f"- 空单元格数（表头下数据区）：{sheet.empty_cell_count}",
                f"- 重复数据行数：{sheet.duplicate_data_rows}",
                "",
            ]
        )
        lines.extend(["### 候选表头", ""])
        if sheet.header_candidates:
            lines.extend(
                [
                    "| 行号 | 置信度 | 依据 | 示例值 |",
                    "|---|---|---|---|",
                ]
            )
            for candidate in sheet.header_candidates:
                mark = "★ 推荐" if candidate == sheet.recommended_header else ""
                lines.append(
                    f"| {candidate.row_number} {mark} | {candidate.score:.3f} "
                    f"| {_md('；'.join(candidate.reasons), 60)} "
                    f"| {_md('，'.join(candidate.sample_values), 60)} |"
                )
        else:
            lines.append("未识别出候选表头行。")
        lines.append("")

        lines.extend(["### 列概况", ""])
        if sheet.columns:
            lines.extend(
                [
                    "| Excel 列 | 字段名 | 非空值 | 样例值 |",
                    "|---|---|---|---|",
                ]
            )
            for column in sheet.columns:
                lines.append(
                    f"| {column.excel_column} | {_md(column.name)} | {column.non_empty_count} "
                    f"| {_md('，'.join(column.sample_values), 60)} |"
                )
        else:
            lines.append("未识别出可靠列名，需人工确认表头。")
        lines.append("")
        _append_issues(lines, sheet.issues)

    _append_issues(lines, profile.issues, title="## 全文件问题")
    return "\n".join(lines) + "\n"


def _column_range(profile: SheetProfile) -> str:
    if profile.used_first_col is None or profile.used_last_col is None:
        return "—"
    from e_master.core.text import column_letter

    return f"{column_letter(profile.used_first_col - 1)}–{column_letter(profile.used_last_col - 1)}"


def _append_issues(lines: list[str], issues: list, title: str = "### 质量问题") -> None:
    if not issues:
        return
    lines.extend([title, ""])
    for issue in issues:
        location = " | ".join(
            part
            for part in (issue.sheet, f"行 {issue.row}" if issue.row else "", f"列 {issue.column}" if issue.column else "")
            if part
        )
        badge = {
            IssueLevel.ERROR: "错误",
            IssueLevel.WARNING: "警告",
            IssueLevel.INFO: "提示",
        }.get(issue.level, issue.level.value)
        prefix = f"[{badge}]" if location else f"[{badge}]"
        lines.append(f"- {prefix} {_md(issue.message)}" + (f"（{location}）" if location else ""))
    lines.append("")
