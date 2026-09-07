"""字段选择（P0-04）测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_qc.errors import (
    AmbiguousFieldError,
    DuplicateSheetSelectionError,
    EmptySheetSelectionError,
    NoFieldSelectedError,
    UnknownFieldError,
    UnknownSheetError,
)
from excel_qc.inspector import inspect_workbook
from excel_qc.models import FieldSelection, FieldSelectionRequest
from excel_qc.selection import select_fields, validate_selection
from tests.qc.helpers import build_workbook


def _build_multi_sheet_workbook(path: Path) -> Path:
    return build_workbook(
        path,
        {
            "设备档案": [
                ["设备编码", "设备名称", "投用日期"],
                ["EQ-001", "空压机", "2024-01-05"],
                ["EQ-002", "冷水机", "2024-02-10"],
            ],
            "单位信息": [
                ["单位编码", "单位名称"],
                ["U-01", "一车间"],
            ],
            "空表": [],
        },
    )


def test_selects_sheets_and_fields_for_verification(tmp_path: Path) -> None:
    structure = inspect_workbook(_build_multi_sheet_workbook(tmp_path / "multi.xlsx"))

    selection = select_fields(
        structure,
        [
            FieldSelectionRequest("设备档案", ("投用日期", "设备编码")),
            FieldSelectionRequest("单位信息", ("单位编码", "单位名称")),
        ],
    )

    assert selection.sheet_count == 2
    assert selection.sheet_names == ["设备档案", "单位信息"]
    first = selection.selections[0]
    assert first.header_row == 1
    assert [field.excel_column for field in first.selected_fields] == ["C", "A"]
    assert [field.name for field in first.selected_fields] == ["投用日期", "设备编码"]
    assert [field.column_index for field in first.selected_fields] == [2, 0]


def test_unselected_sheet_and_fields_are_excluded(tmp_path: Path) -> None:
    structure = inspect_workbook(_build_multi_sheet_workbook(tmp_path / "multi.xlsx"))

    selection = select_fields(
        structure,
        [FieldSelectionRequest("设备档案", ("设备编码",))],
    )

    assert selection.sheet_count == 1
    assert selection.sheet_names == ["设备档案"]
    sheet = selection.selections[0]
    assert sheet.field_count == 1
    assert sheet.selected_fields[0].name == "设备编码"
    # 设备名称、投用日期未被选中
    assert [field.excel_column for field in sheet.selected_fields] == ["A"]


def test_empty_selection_list_is_allowed(tmp_path: Path) -> None:
    structure = inspect_workbook(_build_multi_sheet_workbook(tmp_path / "multi.xlsx"))

    selection = select_fields(structure, [])

    assert selection.sheet_count == 0
    assert selection.selections == ()


def test_sheet_without_fields_is_represented_and_rejected_by_validation(
    tmp_path: Path,
) -> None:
    structure = inspect_workbook(_build_multi_sheet_workbook(tmp_path / "multi.xlsx"))

    selection = select_fields(structure, [FieldSelectionRequest("设备档案", ())])

    sheet = selection.selections[0]
    assert sheet.sheet_name == "设备档案"
    assert sheet.selected_fields == ()
    assert sheet.field_count == 0
    assert any(issue.code == "NO_FIELD_SELECTED" for issue in selection.issues)

    # 模型可以表达该状态；只有进入后续核验前才作为错误拒绝
    with pytest.raises(NoFieldSelectedError):
        validate_selection(selection)


def test_validate_selection_rejects_empty_scope(tmp_path: Path) -> None:
    structure = inspect_workbook(_build_multi_sheet_workbook(tmp_path / "multi.xlsx"))

    with pytest.raises(NoFieldSelectedError):
        validate_selection(select_fields(structure, []))


def test_valid_selection_passes_validation(tmp_path: Path) -> None:
    structure = inspect_workbook(_build_multi_sheet_workbook(tmp_path / "multi.xlsx"))

    selection = select_fields(
        structure,
        [FieldSelectionRequest("设备档案", ("设备编码",))],
    )

    validate_selection(selection)  # 不应抛出


def test_field_selection_keeps_explicit_coordinates_and_original_structure(
    tmp_path: Path,
) -> None:
    path = build_workbook(
        tmp_path / "coordinates.xlsx",
        {
            "Sheet1": [
                ["字段一", "字段二", "字段三", "字段四", "字段五", "字段六", "字段七"],
                ["a", "b", "c", "d", "e", "f", "g"],
            ]
        },
    )
    structure = inspect_workbook(path)

    selection = select_fields(
        structure,
        [FieldSelectionRequest("Sheet1", ("G",))],
    )

    sheet = selection.selections[0]
    field = sheet.selected_fields[0]
    assert isinstance(field, FieldSelection)
    assert field.name == "字段七"
    assert field.excel_column == "G"
    assert field.excel_column_letter == "G"
    assert field.excel_column_number == 7
    assert field.column_index == 6
    assert sheet.header_row == 1
    assert sheet.structure is structure.sheets[0]
    assert selection.path == path


def test_selection_does_not_modify_source_file(tmp_path: Path) -> None:
    path = _build_multi_sheet_workbook(tmp_path / "readonly-selection.xlsx")
    before = path.read_bytes()
    structure = inspect_workbook(path)

    select_fields(structure, [FieldSelectionRequest("设备档案", ("设备编码",))])

    assert path.read_bytes() == before


def test_empty_sheet_cannot_be_selected(tmp_path: Path) -> None:
    structure = inspect_workbook(_build_multi_sheet_workbook(tmp_path / "multi.xlsx"))

    with pytest.raises(EmptySheetSelectionError):
        select_fields(
            structure,
            [FieldSelectionRequest("空表", ("设备编码",))],
        )


def test_unknown_sheet_raises(tmp_path: Path) -> None:
    structure = inspect_workbook(_build_multi_sheet_workbook(tmp_path / "multi.xlsx"))

    with pytest.raises(UnknownSheetError):
        select_fields(
            structure,
            [FieldSelectionRequest("不存在的Sheet", ("设备编码",))],
        )


def test_unknown_field_raises(tmp_path: Path) -> None:
    structure = inspect_workbook(_build_multi_sheet_workbook(tmp_path / "multi.xlsx"))

    with pytest.raises(UnknownFieldError):
        select_fields(
            structure,
            [FieldSelectionRequest("设备档案", ("不存在的字段",))],
        )


def test_duplicate_sheet_request_raises(tmp_path: Path) -> None:
    structure = inspect_workbook(_build_multi_sheet_workbook(tmp_path / "multi.xlsx"))

    with pytest.raises(DuplicateSheetSelectionError):
        select_fields(
            structure,
            [
                FieldSelectionRequest("设备档案", ("设备编码",)),
                FieldSelectionRequest("设备档案", ("设备名称",)),
            ],
        )


def test_select_by_column_letter_and_blank_header_placeholder(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "blank-header.xlsx",
        {"Sheet1": [["设备编码", "", "投用日期"], ["EQ-001", "备件", "2024-01-01"]]},
    )
    structure = inspect_workbook(path)

    selection = select_fields(
        structure,
        [FieldSelectionRequest("Sheet1", ("B", "c"))],
    )

    fields = selection.selections[0].selected_fields
    assert [field.excel_column for field in fields] == ["B", "C"]
    assert fields[0].display_name == "列B"
    assert fields[0].name == ""
    assert fields[1].name == "投用日期"


def test_duplicate_header_name_requires_column_letter(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "dup-name.xlsx",
        {
            "Sheet1": [
                ["名称", "名称", "数量"],
                ["a", "b", 1],
            ]
        },
    )
    structure = inspect_workbook(path)

    with pytest.raises(AmbiguousFieldError):
        select_fields(structure, [FieldSelectionRequest("Sheet1", ("名称",))])

    selection = select_fields(
        structure,
        [FieldSelectionRequest("Sheet1", ("B", "C"))],
    )
    fields = selection.selections[0].selected_fields
    assert [field.name for field in fields] == ["名称", "数量"]
    assert [field.excel_column for field in fields] == ["B", "C"]


def test_repeated_field_in_one_request_is_idempotent(tmp_path: Path) -> None:
    structure = inspect_workbook(_build_multi_sheet_workbook(tmp_path / "multi.xlsx"))

    selection = select_fields(
        structure,
        [FieldSelectionRequest("设备档案", ("设备编码", "设备编码", "设备名称"))],
    )

    fields = selection.selections[0].selected_fields
    assert [field.excel_column for field in fields] == ["A", "B"]
