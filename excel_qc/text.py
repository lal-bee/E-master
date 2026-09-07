"""单元格值文本化与空白判断工具。"""

from __future__ import annotations

import math
from datetime import date, datetime, time
from decimal import Decimal
from typing import Any


def is_blank_value(value: Any) -> bool:
    """判断单元格值是否为空（None、NaN、空白字符串）。"""
    if value is None:
        return True
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return True
    return isinstance(value, str) and value.strip() == ""


def cell_text(value: Any) -> str:
    """把单元格值转为可稳定比较与展示的文本；空值统一为空字符串。"""
    if is_blank_value(value):
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (datetime, date, time)):
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else format(value, ".15g")
    if isinstance(value, Decimal):
        return str(value)
    return str(value).strip()
