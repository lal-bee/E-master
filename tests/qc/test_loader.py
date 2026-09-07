"""Excel 文件导入（P0-01）测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_qc.errors import (
    SourceFileNotFoundError,
    UnsupportedFileTypeError,
    WorkbookReadError,
)
from excel_qc.loader import load_workbook
from tests.qc.helpers import build_workbook


def test_loads_xlsx_sheets_and_keeps_physical_rows(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "source.xlsx",
        {
            "设备档案": [
                ["设备编码", "设备名称"],
                ["EQ-001", "空压机"],
                ["EQ-002", "冷水机"],
            ],
            "说明": [],
        },
    )

    raw = load_workbook(path)

    assert raw.file_type == "xlsx"
    assert raw.sheet_count == 2
    assert raw.sheet_names == ("设备档案", "说明")
    assert raw.sheets[0].rows[0] == ("设备编码", "设备名称")
    assert raw.sheets[0].rows[1] == ("EQ-001", "空压机")
    # 物理行顺序保留：Excel 第 3 行仍是第 2 条数据
    assert raw.sheets[0].rows[2] == ("EQ-002", "冷水机")


def test_load_missing_file_raises(tmp_path: Path) -> None:
    with pytest.raises(SourceFileNotFoundError):
        load_workbook(tmp_path / "not-exists.xlsx")


def test_load_unsupported_type_raises(tmp_path: Path) -> None:
    path = tmp_path / "data.csv"
    path.write_text("a,b\n1,2\n", encoding="utf-8")
    with pytest.raises(UnsupportedFileTypeError):
        load_workbook(path)


def test_load_corrupted_xlsx_raises_domain_error(tmp_path: Path) -> None:
    path = tmp_path / "broken.xlsx"
    path.write_bytes(b"this is not a valid excel zip package")

    with pytest.raises(WorkbookReadError) as exc_info:
        load_workbook(path)

    assert "无法读取" in str(exc_info.value)


def test_load_does_not_modify_source_file(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "readonly.xlsx",
        {"Sheet1": [["设备编码"], ["EQ-001"]]},
    )
    before = path.read_bytes()

    load_workbook(path)

    assert path.read_bytes() == before
