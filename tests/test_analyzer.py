"""结构探查单元测试。"""

from __future__ import annotations

from e_master.core.analyzer.analyzer import WorkbookAnalyzer
from e_master.core.loader.loader import ExcelLoader
from e_master.report.markdown import render_markdown
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


def test_trailing_note_excluded_from_data_statistics(trailing_note_workbook):
    sheet = _analyze(trailing_note_workbook).sheets[0]

    # 说明行在第 7 行（第 1 行标题、第 2 行表头、第 3–6 行数据）
    assert sheet.recommended_header.row_number == 2
    assert sheet.data_row_count == 4

    # 说明行不能污染列样例值；紧邻其上方的稀疏数据行仍应保留为数据
    column_samples = [value for column in sheet.columns for value in column.sample_values]
    assert not any("说明：" in value for value in column_samples)
    columns_by_name = {column.name: column for column in sheet.columns}
    assert columns_by_name["设备名称"].non_empty_count == 4

    # 数据区只有稀疏行产生空白，说明行不计入空值/重复等统计
    assert sheet.empty_cell_count == 4
    assert sheet.duplicate_data_rows == 1


def test_trailing_note_reported_with_physical_row(trailing_note_workbook):
    profile = _analyze(trailing_note_workbook)
    sheet = profile.sheets[0]

    issue = next(issue for issue in sheet.issues if issue.code == "TRAILING_CONTENT")
    assert issue.row == 7
    assert issue.details["row_numbers"] == [7]
    assert issue.details["row_count"] == 1
    assert "原始行号：7" in issue.message

    # 报告渲染后仍能看到说明行的问题提示与物理行号
    report = render_markdown(profile)
    assert "疑似说明/备注内容" in report
    assert "原始行号：7" in report
    assert "行 7" in report


def test_trailing_note_content_not_lost_by_loader_or_analysis(trailing_note_workbook):
    raw = ExcelLoader().load(trailing_note_workbook)
    # 物理第 7 行内容在原始数据中完整保留
    assert "说明：本台账由设备管理部提供" in raw.sheets[0].data.iloc[6, 0]

    profile = _analyze(trailing_note_workbook)
    # 分析结果仍保留整张 Sheet 的物理使用范围（含说明行），只是不把说明计入数据统计
    assert profile.sheets[0].used_last_row == 7


def test_deep_header_beyond_scan_window_reports_truncation(deep_header_workbook):
    profile = _analyze(deep_header_workbook)
    sheet = profile.sheets[0]

    # 原有流程不崩溃，窗口内推荐结果保持不变
    assert sheet.recommended_header is not None
    assert sheet.recommended_header.row_number == 1
    assert sheet.data_row_count > 0

    issue = next(issue for issue in sheet.issues if issue.code == "HEADER_SCAN_TRUNCATED")
    assert issue.level.value == "warning"
    assert issue.row == 17
    assert "15 行" in issue.message
    assert "更完整" in issue.message
    assert "人工确认" in issue.message

    details = issue.details
    assert details["scan_limit"] == 15
    assert details["scan_window"] == {"start_row": 1, "end_row": 15}
    assert details["scanned_rows"] == 15
    assert details["recommended_header_row"] == 1
    assert details["best_complete_row"] == 17
    assert 17 in details["first_more_complete_rows"]
    assert details["more_complete_row_count"] >= 1


def test_normal_many_data_rows_no_scan_truncation_warning(tmp_path):
    rows = [["设备编号", "设备名称", "数量"]]
    rows.extend(
        [f"EQ-{index:03d}", f"设备{index}", index]
        for index in range(1, 26)
    )
    path = build_workbook(tmp_path / "normal-many-rows.xlsx", {"Sheet1": rows})

    sheet = _analyze(path).sheets[0]
    assert sheet.recommended_header.row_number == 1
    assert not any(issue.code == "HEADER_SCAN_TRUNCATED" for issue in sheet.issues)


def test_merged_cells_analyzed_without_crash(merged_workbook):
    sheet = _analyze(merged_workbook).sheets[0]

    # 合并标题行不会影响表头与数据区识别，物理行号仍正确
    assert sheet.recommended_header.row_number == 2
    assert sheet.data_start_row == 3
    assert sheet.data_row_count == 2
    assert [column.name for column in sheet.columns] == [
        "设备编号",
        "设备名称",
        "规格型号",
        "启用日期",
        "备注",
    ]
    codes = {issue.code for issue in sheet.issues}
    assert "CONTENT_ABOVE_HEADER" in codes


def test_xlsm_analyzed_like_xlsx(xlsm_workbook):
    profile = _analyze(xlsm_workbook)
    sheet = profile.sheets[0]

    assert profile.file_type == "xlsm"
    assert len(profile.sheets) == 1
    assert sheet.recommended_header.row_number == 1
    assert sheet.data_start_row == 2
    assert sheet.data_row_count == 1
    assert [column.name for column in sheet.columns] == ["设备编号", "设备名称"]


def test_mixed_type_column_loaded_and_analyzed(mixed_type_workbook):
    raw = ExcelLoader().load(mixed_type_workbook)
    first = raw.sheets[0]
    # 数字/文本统一按原始文本保留，空值统一为空字符串
    assert first.data.iloc[1, 0] == "1"
    assert first.data.iloc[2, 0] == "EQ-002"
    assert first.data.iloc[3, 0] == ""
    assert first.data.iloc[4, 0] == "3.5"

    sheet = _analyze(mixed_type_workbook).sheets[0]
    assert sheet.recommended_header.row_number == 1
    assert sheet.data_row_count == 4
    assert sheet.empty_cell_count == 3
    columns_by_name = {column.name: column for column in sheet.columns}
    assert columns_by_name["设备编号"].non_empty_count == 3
    assert columns_by_name["设备编号"].sample_values == ("1", "EQ-002", "3.5")
