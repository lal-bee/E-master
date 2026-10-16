"""设备台账 Excel 格式转换。"""

from .converter import ConversionResult, WorkbookError, convert_workbook

__all__ = ["ConversionResult", "WorkbookError", "convert_workbook"]
