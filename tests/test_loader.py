"""ExcelLoader 单元测试。"""

from __future__ import annotations

import pytest

from e_master.core.errors import SourceFileNotFoundError, UnsupportedFileTypeError
from e_master.core.loader.loader import ExcelLoader
from tests.helpers.workbooks import build_workbook, write_csv


def test_load_xlsx_preserves_sheets_and_physical_rows(messy_workbook):
    raw = ExcelLoader().load(messy_workbook)
    assert raw.file_type == "xlsx"
    assert [sheet.sheet_name for sheet in raw.sheets] == ["设备台账", "说明"]

    first = raw.sheets[0]
    # Excel 第 2 行是表头，第 3 行起是数据
    assert first.data.iloc[1, 0] == "设备编号"
    assert first.data.iloc[2, 0] == "EQ-001"
    assert first.data.iloc[1, 3] == "启用日期"


def test_load_xlsx_empty_sheet(empty_workbook):
    raw = ExcelLoader().load(empty_workbook)
    assert len(raw.sheets) == 1
    assert raw.sheets[0].sheet_name == "空表"


def test_load_utf8_sig_csv(tmp_path):
    path = write_csv(
        tmp_path / "data.csv",
        [["设备编号", "设备名称"], ["EQ-001", "空压机"]],
        encoding="utf-8-sig",
    )
    raw = ExcelLoader().load(path)
    assert raw.file_type == "csv"
    assert raw.sheets[0].data.iloc[0, 0] == "设备编号"
    assert raw.sheets[0].data.iloc[1, 1] == "空压机"


def test_load_gbk_csv(tmp_path):
    path = write_csv(
        tmp_path / "gbk.csv",
        [["设备编号", "备注"], ["EQ-001", "国产"]],
        encoding="gb18030",
    )
    raw = ExcelLoader().load(path)
    assert raw.sheets[0].data.iloc[0, 0] == "设备编号"


def test_load_missing_file_raises(tmp_path):
    with pytest.raises(SourceFileNotFoundError):
        ExcelLoader().load(tmp_path / "not-exists.xlsx")


def test_load_unsupported_type_raises(tmp_path):
    path = tmp_path / "data.pdf"
    path.write_text("not excel", encoding="utf-8")
    with pytest.raises(UnsupportedFileTypeError):
        ExcelLoader().load(path)
