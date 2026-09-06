"""Pytest 共享夹具。"""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.helpers.workbooks import (
    build_deep_header_workbook,
    build_merged_workbook,
    build_mixed_type_workbook,
    build_trailing_note_workbook,
    build_workbook,
    build_xlsm_workbook,
)


@pytest.fixture
def messy_workbook(tmp_path: Path) -> Path:
    """带标题行、重复数据、空单元格的混乱样例。"""
    return build_workbook(
        tmp_path / "messy.xlsx",
        {
            "设备台账": [
                ["2024年设备主数据台账", "", "", "", ""],
                ["设备编号", "设备名称", "规格型号", "启用日期", "备注"],
                ["EQ-001", "空压机", "GA-75", "2024-01-05", "进口"],
                ["EQ-002", "冷水机", "LS-100", "2024/2/10", "国产"],
                ["EQ-002", "冷水机", "LS-100", "2024/2/10", "国产"],
                ["", "备用机", "", "2024-03-01", ""],
            ],
            "说明": [
                ["本文件由设备管理部提供"],
                ["字段口径以 2024 版台账为准"],
            ],
        },
    )


@pytest.fixture
def empty_workbook(tmp_path: Path) -> Path:
    """含一个完全空 Sheet 的样例。"""
    return build_workbook(tmp_path / "empty.xlsx", {"空表": []})


@pytest.fixture
def trailing_note_workbook(tmp_path: Path) -> Path:
    """数据区末尾混入说明行的样例。"""
    return build_trailing_note_workbook(tmp_path / "trailing-note.xlsx")


@pytest.fixture
def merged_workbook(tmp_path: Path) -> Path:
    """含合并单元格标题行的样例。"""
    return build_merged_workbook(tmp_path / "merged.xlsx")


@pytest.fixture
def xlsm_workbook(tmp_path: Path) -> Path:
    """最小 xlsm 样例。"""
    return build_xlsm_workbook(tmp_path / "macro-enabled.xlsm")


@pytest.fixture
def mixed_type_workbook(tmp_path: Path) -> Path:
    """同一列混合数字、文本与空值的样例。"""
    return build_mixed_type_workbook(tmp_path / "mixed-types.xlsx")


@pytest.fixture
def deep_header_workbook(tmp_path: Path) -> Path:
    """真实表头位于第 17 行、超出 15 行扫描窗口的样例。"""
    return build_deep_header_workbook(tmp_path / "deep-header.xlsx")
