"""P1-02 Excel 结构探查增强测试。"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from excel_qc.models import BasicDataType
from excel_qc.profiler import (
    ProfileIssueLevel,
    WorkbookProfile,
    profile_workbook,
)
from tests.qc.helpers import (
    build_workbook,
    build_workbook_with_metadata,
)


def _sheet(profile: WorkbookProfile, index: int = 0):
    return profile.sheets[index]


def _codes(sheet) -> list[str]:
    return [issue.code for issue in sheet.issues]


def test_single_sheet_normal_workbook(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "single.xlsx",
        {
            "设备档案": [
                ["设备编码", "设备名称"],
                ["EQ-001", "空压机"],
                ["EQ-002", "冷水机"],
            ]
        },
    )

    profile = profile_workbook(path)

    assert profile.sheet_count == 1
    assert profile.sheet_names == ("设备档案",)
    sheet = _sheet(profile)
    assert sheet.sheet_number == 1
    assert sheet.empty is False
    assert sheet.total_row_count == 3
    assert sheet.total_column_count == 2
    assert sheet.used_first_row == 1
    assert sheet.used_last_row == 3
    assert sheet.header_row == 1
    assert sheet.data_first_row == 2
    assert sheet.data_last_row == 3
    assert sheet.data_first_col == 1
    assert sheet.data_last_col == 2
    assert sheet.data_row_count == 2
    assert [field.name for field in sheet.fields] == ["设备编码", "设备名称"]
    assert sheet.field_count == 2


def test_multi_sheet_keeps_order_index_and_dimensions(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "multi.xlsx",
        {
            "设备档案": [["设备编码"], ["EQ-001"]],
            "技术参数": [["参数名称"], ["额定功率"]],
            "空表": [],
        },
    )

    profile = profile_workbook(path)

    assert profile.sheet_count == 3
    assert profile.sheet_names == ("设备档案", "技术参数", "空表")
    first = _sheet(profile, 0)
    second = _sheet(profile, 1)
    third = _sheet(profile, 2)
    assert (first.sheet_index, first.sheet_number) == (0, 1)
    assert (second.sheet_index, second.sheet_number) == (1, 2)
    assert (third.sheet_index, third.sheet_number) == (2, 3)
    assert first.total_row_count == 2
    assert second.total_row_count == 2
    assert third.empty is True


def test_empty_sheet_is_reported(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "empty.xlsx",
        {
            "空表": [],
            "有效表": [["设备编码"], ["EQ-001"]],
        },
    )

    sheet = _sheet(profile_workbook(path), 0)

    assert sheet.empty is True
    assert sheet.header_row is None
    assert sheet.data_row_count == 0
    assert sheet.fields == ()
    assert "EMPTY_SHEET" in _codes(sheet)


def test_header_not_in_first_row(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "header-row2.xlsx",
        {
            "Sheet1": [
                [],
                ["设备编码", "设备名称"],
                ["EQ-001", "空压机"],
            ]
        },
    )

    sheet = _sheet(profile_workbook(path))

    assert sheet.header_row == 2
    assert sheet.data_first_row == 3
    assert sheet.data_row_count == 1
    assert [field.name for field in sheet.fields] == ["设备编码", "设备名称"]


def test_title_row_before_header(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "title.xlsx",
        {
            "设备档案": [
                ["2024年设备主数据台账"],
                ["设备编码", "设备名称"],
                ["EQ-001", "空压机"],
            ]
        },
    )

    sheet = _sheet(profile_workbook(path))

    assert sheet.used_first_row == 1
    assert sheet.header_row == 2
    assert sheet.data_first_row == 3
    assert "CONTENT_ABOVE_HEADER" in _codes(sheet)


def test_empty_intermediate_row_does_not_count(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "sparse.xlsx",
        {
            "Sheet1": [
                ["设备编码", "设备名称"],
                ["EQ-001", "空压机"],
                [None, None],
                ["EQ-002", "冷水机"],
            ]
        },
    )

    sheet = _sheet(profile_workbook(path))

    assert sheet.data_row_count == 2
    assert sheet.data_first_row == 2
    assert sheet.data_last_row == 4


def test_duplicate_header_names_reported(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "dup-header.xlsx",
        {"Sheet1": [["名称", "名称"], ["a", "b"]]},
    )

    sheet = _sheet(profile_workbook(path))

    assert sheet.header_row == 1
    assert len(sheet.fields) == 2
    assert "DUPLICATE_HEADER_NAME" in _codes(sheet)


def test_empty_header_cell_uses_placeholder(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "blank-header.xlsx",
        {"Sheet1": [["设备编码", "", "启用日期"], ["EQ-001", "备件", "2024-01-01"]]},
    )

    sheet = _sheet(profile_workbook(path))
    field_b = sheet.fields[1]

    assert field_b.name == ""
    assert field_b.display_name == "列B"
    assert field_b.column_letter == "B"
    assert "EMPTY_HEADER_CELL" in _codes(sheet)


def test_no_header_detected_for_numeric_rows(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "no-header.xlsx",
        {"Sheet1": [[1, 2], [3, 4]]},
    )

    sheet = _sheet(profile_workbook(path))

    assert sheet.header_row is None
    assert sheet.header_start_row is None
    assert sheet.fields == ()
    assert "NO_HEADER_FOUND" in _codes(sheet)


def test_header_without_data_rows_reports_no_data_region(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "no-data.xlsx",
        {"Sheet1": [["设备编码", "设备名称"]]},
    )

    sheet = _sheet(profile_workbook(path))

    assert sheet.header_row == 1
    assert sheet.data_first_row is None
    assert sheet.data_last_row is None
    assert sheet.data_row_count == 0
    assert "NO_DATA_REGION" in _codes(sheet)


def test_date_field_features(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "date.xlsx",
        {
            "Sheet1": [
                ["投用日期"],
                [datetime(2024, 1, 5)],
                [datetime(2024, 2, 10)],
            ]
        },
    )

    field = _sheet(profile_workbook(path)).fields[0]

    assert field.data_type is BasicDataType.DATETIME
    assert field.non_empty_count == 2
    assert field.empty_count == 0
    assert field.suspected_date is True
    assert field.suspected_text is False


def test_number_field_features(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "number.xlsx",
        {"Sheet1": [["数量"], [100], [200]]},
    )

    field = _sheet(profile_workbook(path)).fields[0]

    assert field.data_type is BasicDataType.INTEGER
    assert field.suspected_number is True
    assert field.suspected_text is False


def test_text_field_features(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "text.xlsx",
        {"Sheet1": [["设备名称"], ["空压机"], ["冷水机"]]},
    )

    field = _sheet(profile_workbook(path)).fields[0]

    assert field.data_type is BasicDataType.TEXT
    assert field.suspected_text is True
    assert field.suspected_number is False
    assert field.suspected_date is False


def test_mixed_type_field_is_reported(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "mixed.xlsx",
        {"Sheet1": [["列一"], [1], ["EQ-002"], [2.5]]},
    )

    field = _sheet(profile_workbook(path)).fields[0]

    assert field.data_type is BasicDataType.MIXED
    assert field.non_empty_count == 3
    assert field.suspected_text is False
    assert field.suspected_number is False


def test_merged_title_and_cells_are_recorded(tmp_path: Path) -> None:
    path = build_workbook_with_metadata(
        tmp_path / "merged.xlsx",
        {
            "设备档案": [
                ["2024年设备主数据台账", "", ""],
                ["设备编号", "设备名称", "启用日期"],
                ["EQ-001", "空压机", datetime(2024, 1, 5)],
            ]
        },
        merged_ranges={"设备档案": ["A1:C1"]},
    )

    sheet = _sheet(profile_workbook(path))

    assert sheet.merged_cell_ranges == ("A1:C1",)
    assert sheet.header_row == 2
    assert sheet.data_first_row == 3
    assert sheet.fields[0].name == "设备编号"
    assert "MERGED_CELL_RANGES" in _codes(sheet)


def test_multi_row_header_is_detected(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "multi-header.xlsx",
        {
            "设备档案": [
                ["基本信息", "", "技术参数", ""],
                ["设备编号", "设备名称", "参数编码", "备注"],
                ["EQ-001", "空压机", "P-001", "进口"],
                ["EQ-002", "冷水机", "P-002", "国产"],
            ]
        },
    )

    sheet = _sheet(profile_workbook(path))

    assert sheet.multi_row_header is True
    assert sheet.header_start_row == 1
    assert sheet.header_end_row == 2
    assert sheet.header_row == 2
    assert sheet.data_first_row == 3
    assert [field.name for field in sheet.fields] == [
        "设备编号",
        "设备名称",
        "参数编码",
        "备注",
    ]
    assert "MULTI_ROW_HEADER" in _codes(sheet)


def test_hidden_rows_and_columns_are_recorded(tmp_path: Path) -> None:
    path = build_workbook_with_metadata(
        tmp_path / "hidden.xlsx",
        {
            "设备档案": [
                ["设备编码", "设备名称", "数量"],
                ["EQ-001", "空压机", 1],
                ["EQ-002", "冷水机", 2],
            ]
        },
        hidden_rows={"设备档案": [3]},
        hidden_columns={"设备档案": ["B"]},
    )

    sheet = _sheet(profile_workbook(path))

    assert sheet.hidden_row_numbers == (3,)
    assert sheet.hidden_column_letters == ("B",)
    assert "HIDDEN_ROWS" in _codes(sheet)
    assert "HIDDEN_COLUMNS" in _codes(sheet)
    assert [field.column_letter for field in sheet.fields] == ["A", "B", "C"]


def test_trailing_note_row_is_excluded_conservatively(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "trailing.xlsx",
        {
            "设备档案": [
                ["设备编码", "设备名称"],
                ["EQ-001", "空压机"],
                ["EQ-002", "冷水机"],
                ["说明：本表由设备管理部提供，字段口径以 2024 版台账为准", "", ""],
            ]
        },
    )

    sheet = _sheet(profile_workbook(path))

    assert sheet.data_row_count == 2
    assert sheet.data_first_row == 2
    assert sheet.data_last_row == 3
    assert sheet.used_last_row == 4
    assert "TRAILING_CONTENT" in _codes(sheet)
    trailing = next(
        issue for issue in sheet.issues if issue.code == "TRAILING_CONTENT"
    )
    assert trailing.row_number == 4
    assert trailing.level is ProfileIssueLevel.INFO


def test_profile_does_not_modify_source_file(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "readonly.xlsx",
        {
            "设备档案": [
                ["设备编码", "设备名称"],
                ["EQ-001", "空压机"],
            ]
        },
    )
    before = path.read_bytes()

    profile_workbook(path)

    assert path.read_bytes() == before
