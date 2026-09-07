"""标准 Excel 导入（P0-05）测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_qc.errors import (
    SourceFileNotFoundError,
    StandardBlankValueError,
    StandardConfigError,
    StandardEmptySheetError,
    StandardFieldNotFoundError,
    StandardMappingConflictError,
    StandardNoDataError,
    StandardSheetNotFoundError,
    WorkbookReadError,
)
from excel_qc.models import (
    StandardDataset,
    StandardDatasetConfig,
    StandardRecord,
)
from excel_qc.standard import import_standard_dataset
from tests.qc.helpers import build_workbook


def _standard_config(
    sheet_name: str = "标准映射",
    match_field: str = "原值",
    standard_field: str = "标准值",
    dataset_name: str = "测试标准集",
) -> StandardDatasetConfig:
    return StandardDatasetConfig(
        dataset_name=dataset_name,
        sheet_name=sheet_name,
        match_field=match_field,
        standard_field=standard_field,
    )


def _build_standard_workbook(path: Path) -> Path:
    return build_workbook(
        path,
        {
            "标准映射": [
                ["原值", "标准值"],
                ["除尘器", "D-CCQ"],
                ["离心泵", "D-LXB"],
                ["螺杆泵", "D-LGB"],
            ],
            "说明": [["本表仅供测试"]],
        },
    )


def test_normal_import_and_sheet_selection(tmp_path: Path) -> None:
    path = _build_standard_workbook(tmp_path / "standard.xlsx")

    dataset = import_standard_dataset(path, _standard_config())

    assert isinstance(dataset, StandardDataset)
    assert dataset.dataset_name == "测试标准集"
    assert dataset.source_file == path
    assert dataset.sheet_name == "标准映射"
    assert dataset.header_row == 1
    assert dataset.record_count == 3


def test_columns_are_recognized(tmp_path: Path) -> None:
    path = _build_standard_workbook(tmp_path / "standard.xlsx")

    dataset = import_standard_dataset(path, _standard_config())

    match = dataset.match_column
    standard = dataset.standard_column
    assert match.name == "原值"
    assert match.excel_column == "A"
    assert match.excel_column_letter == "A"
    assert match.excel_column_number == 1
    assert match.column_index == 0
    assert standard.name == "标准值"
    assert standard.excel_column_letter == "B"
    assert standard.excel_column_number == 2
    assert standard.column_index == 1


def test_match_field_and_standard_field_can_be_column_letters(tmp_path: Path) -> None:
    path = _build_standard_workbook(tmp_path / "standard.xlsx")

    dataset = import_standard_dataset(
        path,
        _standard_config(match_field="A", standard_field="B"),
    )

    assert dataset.match_column.name == "原值"
    assert dataset.standard_column.name == "标准值"


def test_standard_records_are_read_with_row_numbers(tmp_path: Path) -> None:
    path = _build_standard_workbook(tmp_path / "standard.xlsx")

    dataset = import_standard_dataset(path, _standard_config())

    assert [(record.row_number, record.match_value, record.standard_value)
            for record in dataset.records] == [
        (2, "除尘器", "D-CCQ"),
        (3, "离心泵", "D-LXB"),
        (4, "螺杆泵", "D-LGB"),
    ]


def test_empty_match_value_rejected_with_location(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "blank-match.xlsx",
        {
            "标准映射": [
                ["原值", "标准值"],
                ["", "D-CCQ"],
            ]
        },
    )

    with pytest.raises(StandardBlankValueError) as exc_info:
        import_standard_dataset(path, _standard_config())

    message = str(exc_info.value)
    assert "匹配字段为空" in message
    assert "第 2 行" in message
    assert "A" in message


def test_empty_standard_value_rejected_with_location(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "blank-standard.xlsx",
        {
            "标准映射": [
                ["原值", "标准值"],
                ["除尘器", ""],
            ]
        },
    )

    with pytest.raises(StandardBlankValueError) as exc_info:
        import_standard_dataset(path, _standard_config())

    message = str(exc_info.value)
    assert "标准值字段为空" in message
    assert "第 2 行" in message
    assert "B" in message


def test_empty_data_rows_are_skipped_and_counted(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "blank-row.xlsx",
        {
            "标准映射": [
                ["原值", "标准值"],
                ["除尘器", "D-CCQ"],
                ["", ""],
                ["离心泵", "D-LXB"],
            ]
        },
    )

    dataset = import_standard_dataset(path, _standard_config())

    assert dataset.record_count == 2
    assert dataset.skipped_empty_row_count == 1
    assert [record.match_value for record in dataset.records] == ["除尘器", "离心泵"]


def test_empty_standard_sheet_rejected(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "empty-sheet.xlsx",
        {"标准映射": [], "其他": [["值"]]},
    )

    with pytest.raises(StandardEmptySheetError):
        import_standard_dataset(path, _standard_config())


def test_sheet_without_data_rejected(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "no-data.xlsx",
        {"标准映射": [["原值", "标准值"]]},
    )

    with pytest.raises(StandardNoDataError):
        import_standard_dataset(path, _standard_config())


def test_duplicate_match_with_same_standard_value_is_counted(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "dup-same.xlsx",
        {
            "标准映射": [
                ["原值", "标准值"],
                ["除尘器", "D-CCQ"],
                ["除尘器", "D-CCQ"],
                ["离心泵", "D-LXB"],
            ]
        },
    )

    dataset = import_standard_dataset(path, _standard_config())

    assert dataset.record_count == 2
    assert dataset.duplicate_same_value_count == 1
    assert [record.row_number for record in dataset.records] == [2, 4]


def test_conflicting_duplicate_match_value_rejected(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "dup-conflict.xlsx",
        {
            "标准映射": [
                ["原值", "标准值"],
                ["除尘器", "D-CCQ"],
                ["除尘器", "D-XXX"],
            ]
        },
    )

    with pytest.raises(StandardMappingConflictError) as exc_info:
        import_standard_dataset(path, _standard_config())

    message = str(exc_info.value)
    assert "除尘器" in message
    assert "D-CCQ" in message
    assert "D-XXX" in message
    assert "第 2 行" in message
    assert "第 3 行" in message
    assert "A" in message


def test_invalid_sheet_rejected(tmp_path: Path) -> None:
    path = _build_standard_workbook(tmp_path / "standard.xlsx")

    with pytest.raises(StandardSheetNotFoundError):
        import_standard_dataset(path, _standard_config(sheet_name="不存在的Sheet"))


def test_invalid_field_rejected(tmp_path: Path) -> None:
    path = _build_standard_workbook(tmp_path / "standard.xlsx")

    with pytest.raises(StandardFieldNotFoundError):
        import_standard_dataset(path, _standard_config(match_field="不存在的字段"))


def test_same_column_for_match_and_standard_rejected(tmp_path: Path) -> None:
    path = _build_standard_workbook(tmp_path / "standard.xlsx")

    with pytest.raises(StandardConfigError):
        import_standard_dataset(path, _standard_config(match_field="A", standard_field="A"))


def test_missing_standard_file_rejected(tmp_path: Path) -> None:
    with pytest.raises(SourceFileNotFoundError):
        import_standard_dataset(
            tmp_path / "missing.xlsx",
            _standard_config(),
        )


def test_corrupted_standard_file_rejected(tmp_path: Path) -> None:
    path = tmp_path / "broken.xlsx"
    path.write_bytes(b"this is not a valid excel file")

    with pytest.raises(WorkbookReadError):
        import_standard_dataset(path, _standard_config())


def test_import_does_not_modify_standard_file(tmp_path: Path) -> None:
    path = _build_standard_workbook(tmp_path / "readonly-standard.xlsx")
    before = path.read_bytes()

    import_standard_dataset(path, _standard_config())

    assert path.read_bytes() == before


def test_standard_dataset_can_be_read_by_downstream_consumer(tmp_path: Path) -> None:
    path = _build_standard_workbook(tmp_path / "standard.xlsx")
    dataset = import_standard_dataset(path, _standard_config())

    mapping = {record.match_value: record.standard_value for record in dataset.records}
    assert isinstance(next(iter(dataset.records)), StandardRecord)
    assert mapping == {
        "除尘器": "D-CCQ",
        "离心泵": "D-LXB",
        "螺杆泵": "D-LGB",
    }
