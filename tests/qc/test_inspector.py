"""Sheet 与列名识别（P0-02、P0-03）测试。"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from excel_qc.inspector import inspect_workbook
from excel_qc.models import BasicDataType
from tests.qc.helpers import build_typed_workbook, build_workbook


def test_identifies_sheets_and_fields_on_clean_header(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "clean.xlsx",
        {
            "设备档案": [
                ["设备编码", "设备名称", "投用日期"],
                ["EQ-001", "空压机", datetime(2024, 1, 5)],
                ["EQ-002", "冷水机", datetime(2024, 2, 10)],
            ],
            "单位信息": [
                ["单位编码", "单位名称"],
                ["U-01", "一车间"],
            ],
        },
    )

    structure = inspect_workbook(path)

    assert structure.sheet_count == 2
    assert structure.sheet_names == ["设备档案", "单位信息"]
    sheet = structure.sheets[0]
    assert sheet.empty is False
    assert sheet.header_row == 1
    assert sheet.data_row_count == 2
    assert sheet.field_count == 3
    assert [field.excel_column for field in sheet.fields] == ["A", "B", "C"]
    assert [field.name for field in sheet.fields] == ["设备编码", "设备名称", "投用日期"]
    assert sheet.fields[0].data_type == BasicDataType.TEXT
    assert sheet.fields[1].data_type == BasicDataType.TEXT
    assert sheet.fields[2].data_type == BasicDataType.DATETIME


def test_skips_title_row_and_uses_first_valid_header(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "messy.xlsx",
        {
            "设备档案": [
                ["2024年设备主数据台账"],
                ["设备编码", "设备名称", "投用日期"],
                ["EQ-001", "空压机", "2024-01-05"],
                ["EQ-002", "冷水机", "2024-02-10"],
            ]
        },
    )

    sheet = inspect_workbook(path).sheets[0]

    assert sheet.used_first_row == 1
    assert sheet.used_last_row == 4
    assert sheet.header_row == 2
    assert sheet.data_row_count == 2
    assert [field.name for field in sheet.fields] == ["设备编码", "设备名称", "投用日期"]


def test_marks_empty_sheet(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "empty.xlsx",
        {
            "空表": [],
            "有效表": [["设备编码"], ["EQ-001"]],
        },
    )

    structure = inspect_workbook(path)
    empty_sheet = structure.sheets[0]
    valid_sheet = structure.sheets[1]

    assert empty_sheet.empty is True
    assert empty_sheet.header_row is None
    assert empty_sheet.data_row_count == 0
    assert empty_sheet.field_count == 0
    assert any(issue.code == "EMPTY_SHEET" for issue in empty_sheet.issues)
    assert valid_sheet.empty is False
    assert valid_sheet.header_row == 1


def test_single_column_sheet_uses_first_content_row(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "single-col.xlsx",
        {"设备列表": [["设备编码"], ["EQ-001"], ["EQ-002"]]},
    )

    sheet = inspect_workbook(path).sheets[0]

    assert sheet.header_row == 1
    assert sheet.data_row_count == 2
    assert sheet.fields[0].name == "设备编码"
    assert sheet.fields[0].data_type == BasicDataType.TEXT


def test_blank_header_cell_uses_placeholder(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "blank-header.xlsx",
        {"Sheet1": [["设备编码", "", "投用日期"], ["EQ-001", "备件", "2024-01-01"]]},
    )

    sheet = inspect_workbook(path).sheets[0]
    field_b = sheet.fields[1]

    assert field_b.excel_column == "B"
    assert field_b.name == ""
    assert field_b.display_name == "列B"
    assert any(issue.code == "BLANK_HEADER_CELL" for issue in sheet.issues)


def test_duplicate_header_names_reported(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "dup-header.xlsx",
        {"Sheet1": [["名称", "名称"], ["a", "b"]]},
    )

    sheet = inspect_workbook(path).sheets[0]

    assert any(issue.code == "DUPLICATE_HEADER_NAME" for issue in sheet.issues)


def test_no_text_header_returns_no_header_issue(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "numeric-only.xlsx",
        {"Sheet1": [[1, 2, 3], [4, 5, 6]]},
    )

    sheet = inspect_workbook(path).sheets[0]

    assert sheet.header_row is None
    assert sheet.field_count == 0
    assert any(issue.code == "NO_HEADER_FOUND" for issue in sheet.issues)


def test_basic_data_type_inference(tmp_path: Path) -> None:
    path = build_typed_workbook(tmp_path / "typed.xlsx")

    sheet = inspect_workbook(path).sheets[0]
    types = {field.name: field.data_type for field in sheet.fields}

    assert types["设备编码"] == BasicDataType.TEXT
    assert types["设备名称"] == BasicDataType.TEXT
    assert types["数量"] == BasicDataType.INTEGER
    assert types["单价"] == BasicDataType.NUMBER
    assert types["投用日期"] == BasicDataType.DATETIME
    assert types["在用"] == BasicDataType.BOOLEAN
    assert types["备注"] == BasicDataType.BLANK
    assert sheet.fields[0].non_empty_count == 2


def test_data_row_count_ignores_empty_intermediate_rows(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "sparse.xlsx",
        {
            "Sheet1": [
                ["设备编码", "设备名称"],
                ["EQ-001", "空压机"],
                [],
                ["EQ-002", "冷水机"],
            ]
        },
    )

    sheet = inspect_workbook(path).sheets[0]

    assert sheet.header_row == 1
    assert sheet.data_row_count == 2
