"""Excel 内容比对与数据质量核验系统 - 结果展示与报告层（P0-09/P0-10）。"""

from __future__ import annotations

from ui.display import (
    ErrorDisplayRow,
    ValidationDisplay,
    build_display,
    error_detail,
    filter_errors,
)
from ui.excel_report import export_validation_report
from ui.html import render_html

__all__ = [
    "ErrorDisplayRow",
    "ValidationDisplay",
    "build_display",
    "error_detail",
    "filter_errors",
    "export_validation_report",
    "render_html",
]
