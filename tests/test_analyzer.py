"""结构探查单元测试。"""

from __future__ import annotations

from e_master.core.analyzer.analyzer import WorkbookAnalyzer
from e_master.core.loader.loader import ExcelLoader
from tests.helpers.workbooks import build_workbook


def _analyze(path):
    raw = ExcelLoader().load(path)
    return WorkbookAnalyzer().analyze(raw)


def test_detects_header_and_data_region(messy_workbook):
    profile = _analyze(messy_workbook)
    sheet = profile.sheets[0]

    assert sheet.recommended_header is not None
    assert sheet.recommended_header.row_number == 2
    assert sheet.data_start_row == 3
    assert sheet.data_row_count == 4
    assert sheet.column_count == 5
    assert [column.name for column in sheet.columns] == [
        "设备编号",
        "设备名称",
        "规格型号",
        "启用日期",
        "备注",
    ]


def test_reports_quality_issues(messy_workbook):
    sheet = _analyze(messy_workbook).sheets[0]
    codes = {issue.code for issue in sheet.issues}

    assert "CONTENT_ABOVE_HEADER" in codes
    assert "DUPLICATE_DATA_ROWS" in codes
    assert sheet.duplicate_data_rows == 1
    # 数据行第 6 行（1 行空单元格 × 3 列）
    assert sheet.empty_cell_count == 3
    assert sheet.columns[0].non_empty_count == 3


def test_empty_sheet_reported(empty_workbook):
    sheet = _analyze(empty_workbook).sheets[0]
    assert {issue.code for issue in sheet.issues} == {"EMPTY_SHEET"}


def test_duplicate_header_names_detected(tmp_path):
    path = build_workbook(
        tmp_path / "dup-header.xlsx",
        {
            "Sheet1": [
                ["名称", "名称", "数量"],
                ["a", "b", 1],
                ["c", "d", 2],
            ]
        },
    )
    codes = {issue.code for issue in _analyze(path).sheets[0].issues}
    assert "DUPLICATE_HEADER_NAME" in codes


def test_header_at_first_row(tmp_path):
    path = build_workbook(
        tmp_path / "clean.xlsx",
        {
            "Sheet1": [
                ["设备编号", "启用日期"],
                ["EQ-001", "2024-01-01"],
            ]
        },
    )
    sheet = _analyze(path).sheets[0]
    assert sheet.recommended_header.row_number == 1
    assert sheet.data_start_row == 2
