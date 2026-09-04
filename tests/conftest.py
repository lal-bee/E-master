"""Pytest 共享夹具。"""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.helpers.workbooks import build_workbook


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
