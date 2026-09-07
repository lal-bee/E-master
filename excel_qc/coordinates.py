"""Excel 坐标工具。"""

from __future__ import annotations


def column_letter(column_index: int) -> str:
    """把 0 起始列下标转成 Excel 列字母：0 -> A，25 -> Z，26 -> AA。"""
    if column_index < 0:
        raise ValueError(f"列下标不能为负: {column_index}")
    letters = ""
    index = column_index + 1
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def cell_address(row_number: int, column_index: int) -> str:
    """把 Excel 物理行号（1 起始）与 0 起始列下标转成单元格地址，如 G125。"""
    if row_number < 1:
        raise ValueError(f"行号必须从 1 开始: {row_number}")
    return f"{column_letter(column_index)}{row_number}"
