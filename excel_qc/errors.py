"""Excel 数据质量核验系统的领域异常。"""

from __future__ import annotations


class ExcelQcError(Exception):
    """所有本系统领域异常的基类。"""


class SourceFileNotFoundError(ExcelQcError):
    """源 Excel 文件不存在。"""


class UnsupportedFileTypeError(ExcelQcError):
    """不支持的文件类型。"""


class WorkbookReadError(ExcelQcError):
    """Excel 文件无法读取（损坏、格式非法等）。"""


class SelectionError(ExcelQcError):
    """字段选择配置错误。"""


class UnknownSheetError(SelectionError):
    """选择的 Sheet 不存在。"""


class DuplicateSheetSelectionError(SelectionError):
    """同一 Sheet 被重复加入选择。"""


class UnknownFieldError(SelectionError):
    """选择的字段不存在。"""


class AmbiguousFieldError(SelectionError):
    """字段名存在重复，无法仅凭名称唯一确定列。"""


class EmptySheetSelectionError(SelectionError):
    """空 Sheet 无法选择核验字段。"""


class NoFieldSelectedError(SelectionError):
    """核验范围未就绪：未选择 Sheet，或某 Sheet 未选择任何字段。"""


class StandardDataError(ExcelQcError):
    """标准数据集导入/配置错误。"""


class StandardConfigError(StandardDataError):
    """标准数据集配置不合法。"""


class StandardSheetNotFoundError(StandardDataError):
    """指定的标准数据 Sheet 不存在。"""


class StandardFieldNotFoundError(StandardDataError):
    """指定的匹配字段或标准值字段不存在。"""


class StandardFieldAmbiguousError(StandardDataError):
    """字段名重复，无法仅凭名称确定列。"""


class StandardEmptySheetError(StandardDataError):
    """指定的标准 Sheet 为空。"""


class StandardBlankValueError(StandardDataError):
    """标准数据记录中匹配字段或标准值字段为空。"""


class StandardMappingConflictError(StandardDataError):
    """同一匹配值对应多个不同标准值。"""


class StandardNoDataError(StandardDataError):
    """标准数据集中没有有效数据记录。"""


class ComparisonError(ExcelQcError):
    """标准映射比对错误。"""


class ComparisonConfigError(ComparisonError):
    """映射比对配置不合法（Sheet 未选择、字段未选择等）。"""


class ComparisonBlankValueError(ComparisonError):
    """业务数据行中匹配字段或标准值字段为空，无法执行映射比对。"""


class FormatConfigError(ExcelQcError):
    """格式规则配置不合法（Sheet/字段未选择、规则重复等）。"""
