"""E-Master 领域异常分类。"""

from __future__ import annotations


class EMasterError(Exception):
    """所有 E-Master 领域异常的基类。"""


class SourceFileNotFoundError(EMasterError):
    """源文件不存在。"""


class UnsupportedFileTypeError(EMasterError):
    """不支持的输入文件类型。"""


class ExcelLoadError(EMasterError):
    """Excel 文件读取失败。"""


class CsvLoadError(EMasterError):
    """CSV 文件读取失败。"""


class AnalysisError(EMasterError):
    """数据结构分析失败。"""
