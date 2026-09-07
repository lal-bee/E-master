"""标准映射比对（P0-06）测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_qc.comparison import compare_standard_mapping
from excel_qc.errors import (
    ComparisonBlankValueError,
    ComparisonConfigError,
    StandardMappingConflictError,
)
from excel_qc.inspector import inspect_workbook
from excel_qc.models import (
    ComparisonStatus,
    FieldSelectionRequest,
    MappingCheckConfig,
    StandardColumn,
    StandardDataset,
    StandardDatasetConfig,
    StandardRecord,
)
from excel_qc.selection import select_fields
from excel_qc.standard import import_standard_dataset
from tests.qc.helpers import build_workbook


def _build_business_workbook(
    path: Path,
    rows: list[list[str]],
    *,
    sheet_name: str = "设备档案",
    headers: tuple[str, str] = ("名称", "编码"),
) -> Path:
    return build_workbook(
        path,
        {sheet_name: [list(headers), *rows]},
    )


def _select_business(
    path: Path,
    *,
    sheet_name: str = "设备档案",
    match_field: str = "名称",
    standard_field: str = "编码",
):
    structure = inspect_workbook(path)
    return select_fields(
        structure,
        [FieldSelectionRequest(sheet_name, (match_field, standard_field))],
    )


def _build_standard_dataset(path: Path) -> StandardDataset:
    return import_standard_dataset(
        path,
        StandardDatasetConfig(
            dataset_name="测试标准集",
            sheet_name="标准映射",
            match_field="原值",
            standard_field="标准值",
        ),
    )


def _build_standard_workbook(path: Path, rows: list[list[str]]) -> Path:
    return build_workbook(
        path,
        {"标准映射": [["原值", "标准值"], *rows]},
    )


def test_single_data_row_pass_and_records_context(tmp_path: Path) -> None:
    business = _build_business_workbook(
        tmp_path / "business.xlsx",
        [["除尘器", "D-CCQ"]],
    )
    standard = _build_standard_workbook(
        tmp_path / "standard.xlsx",
        [["除尘器", "D-CCQ"]],
    )
    selection = _select_business(business)

    result = compare_standard_mapping(
        selection,
        MappingCheckConfig("设备档案", "名称", "编码"),
        _build_standard_dataset(standard),
    )

    assert result.summary.total_rows == 1
    assert result.summary.pass_count == 1
    row = result.rows[0]
    assert row.status is ComparisonStatus.PASS
    assert row.sheet_name == "设备档案"
    assert row.row_number == 2
    assert row.match_field_name == "名称"
    assert row.standard_field_name == "编码"
    assert row.match_value == "除尘器"
    assert row.actual_value == "D-CCQ"
    assert row.expected_value == "D-CCQ"
    assert row.match_column_letter == "A"
    assert row.match_column_number == 1
    assert row.standard_column_letter == "B"
    assert row.standard_column_number == 2


def test_multiple_rows_all_pass(tmp_path: Path) -> None:
    business = _build_business_workbook(
        tmp_path / "business.xlsx",
        [["除尘器", "D-CCQ"], ["离心泵", "D-LXB"]],
    )
    standard = _build_standard_workbook(
        tmp_path / "standard.xlsx",
        [["除尘器", "D-CCQ"], ["离心泵", "D-LXB"]],
    )

    result = compare_standard_mapping(
        _select_business(business),
        MappingCheckConfig("设备档案", "名称", "编码"),
        _build_standard_dataset(standard),
    )

    assert result.summary.pass_count == 2
    assert result.summary.total_rows == 2
    assert all(row.status is ComparisonStatus.PASS for row in result.rows)


def test_missing_match_value_returns_data_not_found(tmp_path: Path) -> None:
    business = _build_business_workbook(
        tmp_path / "business.xlsx",
        [["未知设备", "D-XXX"]],
    )
    standard = _build_standard_workbook(
        tmp_path / "standard.xlsx",
        [["除尘器", "D-CCQ"]],
    )

    result = compare_standard_mapping(
        _select_business(business),
        MappingCheckConfig("设备档案", "名称", "编码"),
        _build_standard_dataset(standard),
    )

    row = result.rows[0]
    assert row.status is ComparisonStatus.DATA_NOT_FOUND
    assert row.match_value == "未知设备"
    assert row.actual_value == "D-XXX"
    assert row.expected_value == ""
    assert "不存在于当前标准数据集" in row.reason


def test_wrong_standard_value_returns_data_error(tmp_path: Path) -> None:
    business = _build_business_workbook(
        tmp_path / "business.xlsx",
        [["除尘器", "D-CQJ"]],
    )
    standard = _build_standard_workbook(
        tmp_path / "standard.xlsx",
        [["除尘器", "D-CCQ"]],
    )

    result = compare_standard_mapping(
        _select_business(business),
        MappingCheckConfig("设备档案", "名称", "编码"),
        _build_standard_dataset(standard),
    )

    row = result.rows[0]
    assert row.status is ComparisonStatus.DATA_ERROR
    assert row.actual_value == "D-CQJ"
    assert row.expected_value == "D-CCQ"
    assert "对应的标准值为" in row.reason


def test_mixed_results_are_returned_per_row(tmp_path: Path) -> None:
    business = _build_business_workbook(
        tmp_path / "business.xlsx",
        [
            ["除尘器", "D-CCQ"],
            ["离心泵", "D-LXB"],
            ["未知设备", "D-XXX"],
            ["螺杆泵", "D-WRONG"],
        ],
    )
    standard = _build_standard_workbook(
        tmp_path / "standard.xlsx",
        [["除尘器", "D-CCQ"], ["离心泵", "D-LXB"], ["螺杆泵", "D-LGB"]],
    )

    result = compare_standard_mapping(
        _select_business(business),
        MappingCheckConfig("设备档案", "名称", "编码"),
        _build_standard_dataset(standard),
    )

    assert [row.status for row in result.rows] == [
        ComparisonStatus.PASS,
        ComparisonStatus.PASS,
        ComparisonStatus.DATA_NOT_FOUND,
        ComparisonStatus.DATA_ERROR,
    ]
    assert result.summary.pass_count == 2
    assert result.summary.data_not_found_count == 1
    assert result.summary.data_error_count == 1
    assert result.summary.total_rows == 4
    assert [row.row_number for row in result.rows] == [2, 3, 4, 5]


def test_engine_reads_p0_04_selection_and_p0_05_dataset(tmp_path: Path) -> None:
    business = _build_business_workbook(
        tmp_path / "business.xlsx",
        [["除尘器", "D-CCQ"]],
    )
    standard_path = _build_standard_workbook(
        tmp_path / "standard.xlsx",
        [["除尘器", "D-CCQ"]],
    )
    selection = _select_business(business)
    dataset = _build_standard_dataset(standard_path)

    result = compare_standard_mapping(
        selection,
        MappingCheckConfig("设备档案", "名称", "编码"),
        dataset,
    )

    assert result.business_file == business
    assert result.dataset_name == dataset.dataset_name
    assert result.match_column.name == "名称"
    assert result.standard_column.name == "编码"
    assert result.rows[0].status is ComparisonStatus.PASS


def test_generic_business_fields_are_not_hardcoded(tmp_path: Path) -> None:
    business = build_workbook(
        tmp_path / "generic-business.xlsx",
        {
            "业务数据": [
                ["左侧关键词", "右侧编码"],
                ["甲", "V-A"],
            ]
        },
    )
    standard = _build_standard_workbook(
        tmp_path / "standard.xlsx",
        [["甲", "V-A"]],
    )
    selection = _select_business(
        business,
        sheet_name="业务数据",
        match_field="左侧关键词",
        standard_field="右侧编码",
    )

    result = compare_standard_mapping(
        selection,
        MappingCheckConfig("业务数据", "左侧关键词", "右侧编码"),
        _build_standard_dataset(standard),
    )

    assert result.rows[0].status is ComparisonStatus.PASS


def test_mapping_conflict_in_dataset_is_not_silently_accepted(
    tmp_path: Path,
) -> None:
    business = _build_business_workbook(
        tmp_path / "business.xlsx",
        [["除尘器", "D-CCQ"]],
    )
    selection = _select_business(business)
    conflicting_dataset = StandardDataset(
        dataset_name="冲突标准集",
        source_file=Path("conflict.xlsx"),
        sheet_name="标准映射",
        header_row=1,
        match_column=StandardColumn("原值", "A", 1, 0),
        standard_column=StandardColumn("标准值", "B", 2, 1),
        records=(
            StandardRecord(2, "除尘器", "D-CCQ"),
            StandardRecord(3, "除尘器", "D-XXX"),
        ),
    )

    with pytest.raises(StandardMappingConflictError):
        compare_standard_mapping(
            selection,
            MappingCheckConfig("设备档案", "名称", "编码"),
            conflicting_dataset,
        )


def test_empty_business_match_value_raises(tmp_path: Path) -> None:
    business = _build_business_workbook(
        tmp_path / "business.xlsx",
        [["", "D-CCQ"]],
    )
    standard = _build_standard_workbook(
        tmp_path / "standard.xlsx",
        [["除尘器", "D-CCQ"]],
    )

    with pytest.raises(ComparisonBlankValueError) as exc_info:
        compare_standard_mapping(
            _select_business(business),
            MappingCheckConfig("设备档案", "名称", "编码"),
            _build_standard_dataset(standard),
        )

    assert "第 2 行" in str(exc_info.value)
    assert "匹配字段为空" in str(exc_info.value)


def test_empty_business_standard_value_raises(tmp_path: Path) -> None:
    business = _build_business_workbook(
        tmp_path / "business.xlsx",
        [["除尘器", ""]],
    )
    standard = _build_standard_workbook(
        tmp_path / "standard.xlsx",
        [["除尘器", "D-CCQ"]],
    )

    with pytest.raises(ComparisonBlankValueError) as exc_info:
        compare_standard_mapping(
            _select_business(business),
            MappingCheckConfig("设备档案", "名称", "编码"),
            _build_standard_dataset(standard),
        )

    assert "第 2 行" in str(exc_info.value)
    assert "标准值字段为空" in str(exc_info.value)


def test_fully_empty_business_rows_are_skipped(tmp_path: Path) -> None:
    business = _build_business_workbook(
        tmp_path / "business.xlsx",
        [
            ["除尘器", "D-CCQ"],
            ["", ""],
            ["离心泵", "D-LXB"],
        ],
    )
    standard = _build_standard_workbook(
        tmp_path / "standard.xlsx",
        [["除尘器", "D-CCQ"], ["离心泵", "D-LXB"]],
    )

    result = compare_standard_mapping(
        _select_business(business),
        MappingCheckConfig("设备档案", "名称", "编码"),
        _build_standard_dataset(standard),
    )

    assert result.skipped_empty_row_count == 1
    assert result.summary.total_rows == 2
    assert [row.row_number for row in result.rows] == [2, 4]


def test_compare_field_must_be_in_selection(tmp_path: Path) -> None:
    business = _build_business_workbook(
        tmp_path / "business.xlsx",
        [["除尘器", "D-CCQ"]],
    )
    standard = _build_standard_workbook(
        tmp_path / "standard.xlsx",
        [["除尘器", "D-CCQ"]],
    )
    structure = inspect_workbook(business)
    # 只选择“名称”，未选择“编码”
    selection = select_fields(
        structure,
        [FieldSelectionRequest("设备档案", ("名称",))],
    )

    with pytest.raises(ComparisonConfigError):
        compare_standard_mapping(
            selection,
            MappingCheckConfig("设备档案", "名称", "编码"),
            _build_standard_dataset(standard),
        )


def test_comparison_does_not_modify_business_or_standard_files(
    tmp_path: Path,
) -> None:
    business_path = _build_business_workbook(
        tmp_path / "business.xlsx",
        [["除尘器", "D-CCQ"]],
    )
    standard_path = _build_standard_workbook(
        tmp_path / "standard.xlsx",
        [["除尘器", "D-CCQ"]],
    )
    business_before = business_path.read_bytes()
    standard_before = standard_path.read_bytes()

    compare_standard_mapping(
        _select_business(business_path),
        MappingCheckConfig("设备档案", "名称", "编码"),
        _build_standard_dataset(standard_path),
    )

    assert business_path.read_bytes() == business_before
    assert standard_path.read_bytes() == standard_before
