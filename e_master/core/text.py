"""文本处理与单元格值归一化工具。"""

from __future__ import annotations

import math
from datetime import date, datetime, time
from typing import Any

import pandas as pd


def cell_to_text(value: Any) -> str:
    """把单元格值转成可稳定比较的文本；空值统一为 ''。"""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (datetime, date, time)):
        return str(value)
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return ""
        if value.is_integer():
            return str(int(value))
        return format(value, ".15g")
    if isinstance(value, int):
        return str(value)
    return str(value)


def is_blank_text(value: Any) -> bool:
    """判断单元格是否为空（支持空字符串、NaN、None）。"""
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return isinstance(value, str) and value.strip() == ""


def normalize_key(value: Any) -> str:
    """生成用于去重/比较的归一化文本。"""
    if value is None:
        return ""
    return str(value).strip().lower()


def looks_numeric(text: str) -> bool:
    """判断一段文本是否整体可解析为数字（用于表头异常提示）。"""
    try:
        float(text.strip().replace(",", ""))
    except (TypeError, ValueError):
        return False
    return True


def column_letter(index: int) -> str:
    """把 0 起始列下标转成 Excel 列字母：0 -> A，25 -> Z，26 -> AA。"""
    letters = ""
    index += 1
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters
