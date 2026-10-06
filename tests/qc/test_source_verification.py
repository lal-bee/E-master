"""P1-09 回源核验：均使用临时生成的原始源和系统导出 Excel。"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
import subprocess
import sys
from copy import deepcopy

import pytest
from openpyxl import Workbook

from excel_qc import (
    CleaningKind,
    CleaningWorkbookConfig,
    ComparisonRuleKind,
    FieldCleaningRule,
    FieldComparisonRule,
    FieldComparisonStatus,
    FieldMappingConfig,
    FieldValueStatus,
    NormalizedFieldValue,
    OriginalSourceConfig,
    SourceVerificationConfig,
    SourceVerificationConfigError,
    StandardFieldDefinition,
    SystemExportConfig,
    SystemExportDataError,
    VerificationRowStatus,
    VerificationSide,
    verify_source_against_system_export,
)
from excel_qc.source_verification import _compare_field
import excel_qc.source_verification as verification_module


def _excel(path: Path, rows: list[list[object]], sheet: str = "Data") -> Path:
    book = Workbook()
    tab = book.active
    tab.title = sheet
    for row in rows:
        tab.append(row)
    book.save(path)
    return path


def _config(
    *,
    keys: tuple[str, ...] = ("id",),
    compare: tuple[str, ...] = ("name",),
    source_aliases: dict[str, tuple[str, ...]] | None = None,
    system_aliases: dict[str, tuple[str, ...]] | None = None,
    source_cleaning: CleaningWorkbookConfig = CleaningWorkbookConfig(),
    system_cleaning: CleaningWorkbookConfig = CleaningWorkbookConfig(),
    rules: tuple[FieldComparisonRule, ...] = (),
    source_header: int = 1,
    system_header: int = 1,
) -> SourceVerificationConfig:
    source_aliases = source_aliases or {"id": ("源编号",), "name": ("源名称",)}
    system_aliases = system_aliases or {"id": ("系统编号",), "name": ("系统名称",)}
    source_mapping = FieldMappingConfig(tuple(StandardFieldDefinition(name, aliases) for name, aliases in source_aliases.items()))
    system_mapping = FieldMappingConfig(tuple(StandardFieldDefinition(name, aliases) for name, aliases in system_aliases.items()))
    return SourceVerificationConfig(
        OriginalSourceConfig("Data", source_header, source_mapping, source_cleaning),
        SystemExportConfig("Data", system_header, system_mapping, system_cleaning, keys),
        keys,
        compare,
        rules,
    )


def _run(tmp_path: Path, source: list[list[object]], system: list[list[object]], config: SourceVerificationConfig | None = None):
    left = _excel(tmp_path / "original.xlsx", source)
    right = _excel(tmp_path / "system.xlsx", system)
    before = (left.read_bytes(), right.read_bytes())
    result = verify_source_against_system_export(left, right, config or _config())
    assert (left.read_bytes(), right.read_bytes()) == before
    return result


def test_end_to_end_match_and_provenance_with_late_headers(tmp_path: Path) -> None:
    config = _config(source_header=2, system_header=3,
                     source_cleaning=CleaningWorkbookConfig((FieldCleaningRule("name", rule_code="TRIM_SRC"),)),
                     system_cleaning=CleaningWorkbookConfig((FieldCleaningRule("name", rule_code="TRIM_SYS"),)))
    left = [["说明"], ["源编号", "源名称"], ["0007", "  水泵  "]]
    right = [["导出"], ["说明"], ["系统编号", "系统名称"], ["0007", "水泵"]]
    result = _run(tmp_path, left, right, config)
    assert [item.status for item in result.records] == [VerificationRowStatus.MATCH]
    assert result.summary.verification_complete
    assert (result.summary.original_source_input_row_count, result.summary.system_export_input_row_count) == (1, 1)
    field = result.records[0].field_differences[0]
    assert (field.original_source.original_value, field.original_source.normalized_value, field.original_source.applied_rules) == ("  水泵  ", "水泵", ("TRIM_SRC",))
    assert field.system_export.applied_rules == ("TRIM_SYS",)
    assert (field.original_source.location.row_number, field.original_source.location.column_number, field.original_source.location.cell_address) == (3, 2, "B3")
    assert (field.system_export.location.row_number, field.system_export.location.column_number, field.system_export.location.cell_address) == (4, 2, "B4")
    assert result.original_source.source_file.name == "original.xlsx"
    assert result.system_export.source_file.name == "system.xlsx"


def test_missing_extra_and_multi_field_change_count(tmp_path: Path) -> None:
    config = _config(compare=("name", "flag"), source_aliases={"id": ("源编号",), "name": ("源名称",), "flag": ("源启用",)},
                     system_aliases={"id": ("系统编号",), "name": ("系统名称",), "flag": ("系统启用",)})
    result = _run(tmp_path,
                  [["源编号", "源名称", "源启用"], ["A", "甲", True], ["B", "乙", False]],
                  [["系统编号", "系统名称", "系统启用"], ["A", "改", False], ["C", "丙", True]], config)
    assert {item.status for item in result.records} == {VerificationRowStatus.FIELD_CHANGED, VerificationRowStatus.MISSING_IN_SYSTEM, VerificationRowStatus.EXTRA_IN_SYSTEM}
    assert (result.summary.changed_row_count, result.summary.changed_field_count) == (1, 2)
    assert (result.summary.missing_in_system_row_count, result.summary.extra_in_system_row_count) == (1, 1)
    assert result.summary.matched_row_count == 1
    assert result.summary.verification_complete and not result.summary.data_consistent
    assert sum(len(record.original_source_rows) for record in result.records) == result.summary.original_source_input_row_count
    assert sum(len(record.system_export_rows) for record in result.records) == result.summary.system_export_input_row_count


def test_composite_key_preserves_boundaries_and_type(tmp_path: Path) -> None:
    config = _config(keys=("id", "part"), source_aliases={"id": ("源编号",), "part": ("源分段",), "name": ("源名称",)},
                     system_aliases={"id": ("系统编号",), "part": ("系统分段",), "name": ("系统名称",)})
    result = _run(tmp_path,
                  [["源编号", "源分段", "源名称"], ["ab", "c", "甲"], ["a", "bc", "乙"], [0, "z", "数字"]],
                  [["系统编号", "系统分段", "系统名称"], ["a", "bc", "乙"], ["ab", "c", "甲"], [False, "z", "布尔"]], config)
    assert result.summary.unchanged_row_count == 2
    assert result.summary.missing_in_system_row_count == 1
    assert result.summary.extra_in_system_row_count == 1


@pytest.mark.parametrize("source_dupes,system_dupes", [(True, False), (False, True), (True, True)])
def test_duplicate_key_conflict_retains_both_sides(tmp_path: Path, source_dupes: bool, system_dupes: bool) -> None:
    source = [["源编号", "源名称"], ["A", "甲"]]
    system = [["系统编号", "系统名称"], ["A", "甲"]]
    if source_dupes:
        source.append(["A", "再甲"])
    if system_dupes:
        system.append(["A", "再甲"])
    result = _run(tmp_path, source, system)
    assert len(result.records) == 1
    record = result.records[0]
    assert record.status is VerificationRowStatus.DUPLICATE_KEY
    assert (len(record.original_source_rows), len(record.system_export_rows)) == (len(source) - 1, len(system) - 1)
    assert result.summary.duplicate_key_group_count == 1
    assert not result.summary.verification_complete
    assert any(issue.code == "DUPLICATE_KEY" and len(issue.locations) == len(source) + len(system) - 2 for issue in result.issues)


def test_empty_key_and_cleaning_failure_keep_rows(tmp_path: Path) -> None:
    config = _config(source_cleaning=CleaningWorkbookConfig((FieldCleaningRule("name", kind=CleaningKind.DATE, date_input_formats=("%Y-%m-%d",), date_output_format="%Y-%m-%d", rule_code="DATE_SRC"),)),
                     system_cleaning=CleaningWorkbookConfig((FieldCleaningRule("name", kind=CleaningKind.DATE, date_input_formats=("%Y-%m-%d",), date_output_format="%Y-%m-%d", rule_code="DATE_SYS"),)))
    result = _run(tmp_path,
                  [["源编号", "源名称"], [None, "坏日期"], ["A", "坏日期"]],
                  [["系统编号", "系统名称"], ["A", "坏日期"]], config)
    assert len(result.original_source.rows) == 2
    assert any(item.status is VerificationRowStatus.UNVERIFIABLE and item.original_source_rows[0].row_number == 2 for item in result.records)
    pair = next(item for item in result.records if item.key_values == ("A",))
    assert pair.status is VerificationRowStatus.UNVERIFIABLE
    difference = pair.field_differences[0]
    assert difference.status is FieldComparisonStatus.UNVERIFIABLE
    assert difference.original_source.status is FieldValueStatus.CLEANING_FAILED
    assert difference.system_export.status is FieldValueStatus.CLEANING_FAILED
    assert difference.original_source.original_value == "坏日期"
    assert any(issue.code == "EMPTY_PRIMARY_KEY" for issue in result.issues)
    assert len([issue for issue in result.issues if issue.code == "CLEANING_FAILED"]) >= 2
    assert {issue.side for issue in result.issues if issue.code == "CLEANING_FAILED"} == {
        VerificationSide.ORIGINAL_SOURCE, VerificationSide.SYSTEM_EXPORT,
    }
    assert not result.summary.data_consistent


def test_missing_field_and_ambiguous_mapping_are_unverifiable(tmp_path: Path) -> None:
    missing = _run(tmp_path, [["源编号"], ["A"]], [["系统编号", "系统名称"], ["A", None]])
    assert missing.records[0].status is VerificationRowStatus.UNVERIFIABLE
    assert missing.records[0].field_differences[0].original_source.status is FieldValueStatus.FIELD_MISSING
    assert any(issue.code == "FIELD_MISSING" for issue in missing.issues)
    ambiguous_config = _config(source_aliases={"id": ("源编号",), "name": ("别名",), "other": ("别名",)})
    ambiguous = _run(tmp_path, [["源编号", "别名"], ["A", "甲"]], [["系统编号", "系统名称"], ["A", "甲"]], ambiguous_config)
    assert ambiguous.records[0].status is VerificationRowStatus.UNVERIFIABLE
    assert ambiguous.records[0].field_differences[0].original_source.status is FieldValueStatus.MAPPING_FAILED
    assert any(issue.code == "FIELD_MAPPING_AMBIGUOUS" for issue in ambiguous.issues)


def test_none_empty_zero_false_leading_zero_and_dates(tmp_path: Path) -> None:
    fields = {"id": ("源编号",), "blank": ("源空",), "number": ("源数",), "flag": ("源布尔",), "day": ("源日期",)}
    other = {"id": ("系统编号",), "blank": ("系统空",), "number": ("系统数",), "flag": ("系统布尔",), "day": ("系统日期",)}
    config = _config(compare=("blank", "number", "flag", "day"), source_aliases=fields, system_aliases=other,
                     rules=(FieldComparisonRule("day", ComparisonRuleKind.DATE),))
    result = _run(tmp_path,
                  [["源编号", "源空", "源数", "源布尔", "源日期"], ["0007", None, 0, False, datetime(2025, 1, 2)]],
                  [["系统编号", "系统空", "系统数", "系统布尔", "系统日期"], ["0007", "", False, 0, "2025-01-02"]], config)
    record = result.records[0]
    assert record.key_values == ("0007",)
    assert record.status is VerificationRowStatus.FIELD_CHANGED
    statuses = {difference.standard_field_name: difference.status for difference in record.field_differences}
    assert statuses == {"blank": FieldComparisonStatus.EQUAL, "number": FieldComparisonStatus.CHANGED,
                        "flag": FieldComparisonStatus.CHANGED, "day": FieldComparisonStatus.EQUAL}
    assert result.summary.changed_field_count == 2


def test_empty_datasets_and_one_side_empty(tmp_path: Path) -> None:
    empty = _run(tmp_path, [["源编号", "源名称"]], [["系统编号", "系统名称"]])
    assert empty.records == ()
    assert not empty.summary.verification_complete
    assert not empty.summary.data_consistent
    assert any(issue.code == "NO_DATA_ROWS" for issue in empty.issues)
    one = _run(tmp_path, [["源编号", "源名称"], ["A", "甲"]], [["系统编号", "系统名称"]])
    assert one.summary.missing_in_system_row_count == 1
    assert one.summary.system_export_input_row_count == 0
    assert one.summary.verification_complete
    assert not one.summary.data_consistent
    other = _run(tmp_path, [["源编号", "源名称"]], [["系统编号", "系统名称"], ["A", "甲"]])
    assert other.summary.extra_in_system_row_count == 1
    assert other.summary.original_source_input_row_count == 0
    assert other.summary.verification_complete


def test_invalid_config_and_missing_key_schema(tmp_path: Path) -> None:
    config = _config()
    with pytest.raises(SourceVerificationConfigError, match="完全一致"):
        bad = SourceVerificationConfig(config.original_source, config.system_export, ("name",), ("name",))
        _run(tmp_path, [["源编号", "源名称"]], [["系统编号", "系统名称"]], bad)
    result = _run(tmp_path, [["源名称"], ["甲"]], [["系统编号", "系统名称"], ["A", "甲"]])
    assert all(item.status is VerificationRowStatus.UNVERIFIABLE for item in result.records)
    assert not result.summary.verification_complete
    assert result.summary.missing_in_system_row_count == 0
    assert result.summary.extra_in_system_row_count == 0


def test_exact_value_types_and_explicit_number_rule() -> None:
    def field(value):
        return NormalizedFieldValue("v", value, value, FieldValueStatus.EMPTY if value is None or value == "" else FieldValueStatus.PRESENT, (), None)

    assert _compare_field(field(None), field(""), ComparisonRuleKind.EXACT)[0] is FieldComparisonStatus.CHANGED
    assert _compare_field(field(None), field(None), ComparisonRuleKind.EXACT)[0] is FieldComparisonStatus.EQUAL
    assert _compare_field(field(False), field(0), ComparisonRuleKind.EXACT)[0] is FieldComparisonStatus.CHANGED
    assert _compare_field(field(True), field(1), ComparisonRuleKind.EXACT)[0] is FieldComparisonStatus.CHANGED
    assert _compare_field(field("01"), field(1), ComparisonRuleKind.EXACT)[0] is FieldComparisonStatus.CHANGED
    assert _compare_field(field(" 甲"), field("甲"), ComparisonRuleKind.TEXT)[0] is FieldComparisonStatus.CHANGED
    assert _compare_field(field("1"), field(1), ComparisonRuleKind.NUMBER)[0] is FieldComparisonStatus.EQUAL
    assert _compare_field(field(False), field(0), ComparisonRuleKind.NUMBER)[0] is FieldComparisonStatus.UNVERIFIABLE


def test_duplicate_header_and_failed_key_do_not_claim_match(tmp_path: Path) -> None:
    duplicated = _run(tmp_path, [["源编号", "源名称", "源名称"], ["A", "甲", "乙"]],
                      [["系统编号", "系统名称"], ["A", "甲"]])
    assert duplicated.records[0].status is VerificationRowStatus.UNVERIFIABLE
    assert duplicated.records[0].field_differences[0].original_source.status is FieldValueStatus.MAPPING_FAILED
    assert not duplicated.summary.verification_complete

    config = _config(source_cleaning=CleaningWorkbookConfig((FieldCleaningRule("id", kind=CleaningKind.DATE,
                      date_input_formats=("%Y-%m-%d",), date_output_format="%Y-%m-%d"),)))
    failed = _run(tmp_path, [["源编号", "源名称"], ["bad", "甲"]],
                  [["系统编号", "系统名称"], ["bad", "甲"]], config)
    assert all(record.status is VerificationRowStatus.UNVERIFIABLE for record in failed.records)
    assert any(issue.code == "KEY_CLEANING_FAILED" for issue in failed.issues)


def test_system_result_is_not_mutated_and_demo_runs(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    actual_ingest = verification_module.ingest_system_export
    snapshots = []

    def observed(path, config):
        result = actual_ingest(path, config)
        snapshots.append((result, deepcopy(result)))
        return result

    monkeypatch.setattr(verification_module, "ingest_system_export", observed)
    _run(tmp_path, [["源编号", "源名称"], ["A", "甲"]],
         [["系统编号", "系统名称"], ["A", "甲"]])
    assert snapshots[0][0] == snapshots[0][1]
    demo = Path(__file__).resolve().parents[2] / "examples" / "source_verification_demo.py"
    completed = subprocess.run([sys.executable, str(demo)], cwd=demo.parents[1], capture_output=True, text=True, check=True)
    assert "MATCH ('0007',)" in completed.stdout
    assert "verification_complete: True" in completed.stdout


def test_missing_sheet_is_input_issue_and_keeps_other_side(tmp_path: Path) -> None:
    base = _config()
    source = OriginalSourceConfig("不存在", 1, base.original_source.field_mapping)
    config = SourceVerificationConfig(source, base.system_export, base.key_fields, base.comparison_fields)
    result = _run(tmp_path, [["源编号", "源名称"], ["A", "甲"]],
                  [["系统编号", "系统名称"], ["A", "甲"]], config)
    assert len(result.system_export.rows) == 1
    assert result.original_source.rows == ()
    assert result.records[0].status is VerificationRowStatus.UNVERIFIABLE
    assert any(issue.code == "SHEET_NOT_FOUND" for issue in result.issues)
    assert not result.summary.verification_complete


def test_each_side_can_use_different_sheet_and_header(tmp_path: Path) -> None:
    config = _config(source_header=2, system_header=1)
    config = SourceVerificationConfig(
        OriginalSourceConfig("原始", 2, config.original_source.field_mapping),
        SystemExportConfig("导出", 1, config.system_export.field_mapping, primary_key_fields=("id",)),
        config.key_fields,
        config.comparison_fields,
    )
    source = _excel(tmp_path / "source.xlsx", [["说明"], ["源编号", "源名称"], ["A", "甲"]], "原始")
    system = _excel(tmp_path / "export.xlsx", [["系统编号", "系统名称"], ["A", "甲"]], "导出")
    result = verify_source_against_system_export(source, system, config)
    assert result.records[0].status is VerificationRowStatus.MATCH
    assert result.records[0].original_source_rows[0].row_number == 3
    assert result.records[0].system_export_rows[0].row_number == 2


def test_invalid_key_does_not_suppress_valid_matches(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        [["源编号", "源名称"], [None, "未知"], ["A", "同值"], ["B", "旧值"]],
        [["系统编号", "系统名称"], ["A", "同值"], ["B", "新值"], ["C", "可能是未知"]],
    )
    status_by_key = {record.key_values: record.status for record in result.records if record.key_values}
    assert status_by_key == {
        ("A",): VerificationRowStatus.MATCH,
        ("B",): VerificationRowStatus.FIELD_CHANGED,
        ("C",): VerificationRowStatus.UNVERIFIABLE,
    }
    assert (result.summary.matched_row_count, result.summary.unchanged_row_count,
            result.summary.changed_row_count, result.summary.unverifiable_row_count) == (2, 1, 1, 2)
    assert result.summary.missing_in_system_row_count == result.summary.extra_in_system_row_count == 0
    assert any(issue.code == "COUNTERPART_KEY_UNAVAILABLE" and issue.key_values == ("C",) for issue in result.issues)
    assert sum(len(record.original_source_rows) for record in result.records) == 3
    assert sum(len(record.system_export_rows) for record in result.records) == 3
    assert not result.summary.verification_complete and not result.summary.data_consistent


def test_duplicate_group_does_not_suppress_unique_match_or_lose_rows(tmp_path: Path) -> None:
    result = _run(
        tmp_path,
        [["源编号", "源名称"], ["D", "甲"], ["D", "乙"], ["U", "同值"]],
        [["系统编号", "系统名称"], ["D", "甲"], ["U", "同值"]],
    )
    by_key = {record.key_values: record for record in result.records}
    assert by_key[("D",)].status is VerificationRowStatus.DUPLICATE_KEY
    assert (len(by_key[("D",)].original_source_rows), len(by_key[("D",)].system_export_rows)) == (2, 1)
    assert by_key[("U",)].status is VerificationRowStatus.MATCH
    assert (result.summary.duplicate_key_group_count, result.summary.matched_row_count) == (1, 1)
    assert result.summary.missing_in_system_row_count == result.summary.extra_in_system_row_count == 0
    assert sorted(location.row_number for record in result.records for location in record.original_source_rows) == [2, 3, 4]
    assert sorted(location.row_number for record in result.records for location in record.system_export_rows) == [2, 3]
    assert result.summary.issue_count == len(result.issues)


def test_system_ingestion_failure_is_not_an_empty_dataset(tmp_path: Path) -> None:
    base = _config()
    config = SourceVerificationConfig(
        base.original_source,
        SystemExportConfig("不存在", 1, base.system_export.field_mapping, primary_key_fields=("id",)),
        base.key_fields,
        base.comparison_fields,
    )
    failed = _run(tmp_path, [["源编号", "源名称"], ["A", "甲"]],
                  [["系统编号", "系统名称"], ["A", "甲"]], config)
    assert failed.summary.missing_in_system_row_count == 0
    assert failed.summary.unverifiable_row_count == 1
    assert not failed.summary.verification_complete
    assert any(issue.code == "SHEET_NOT_FOUND" and issue.side is VerificationSide.SYSTEM_EXPORT for issue in failed.issues)


def test_empty_sheet_and_all_invalid_keys_are_not_success(tmp_path: Path) -> None:
    source = _excel(tmp_path / "empty-source.xlsx", [])
    system = _excel(tmp_path / "system.xlsx", [["系统编号", "系统名称"], ["A", "甲"]])
    empty_source = verify_source_against_system_export(source, system, _config())
    assert empty_source.summary.original_source_input_row_count == 0
    assert empty_source.summary.missing_in_system_row_count == empty_source.summary.extra_in_system_row_count == 0
    assert not empty_source.summary.verification_complete
    assert any(issue.code == "HEADER_ROW_NOT_FOUND" for issue in empty_source.issues)

    empty_system = _excel(tmp_path / "empty-system.xlsx", [])
    with pytest.raises(SystemExportDataError, match="表头行"):
        verify_source_against_system_export(system, empty_system, _config())

    invalid = _run(tmp_path,
                   [["源编号", "源名称"], [None, "甲"], [None, "乙"]],
                   [["系统编号", "系统名称"], [None, "甲"]])
    assert len(invalid.records) == 3
    assert all(record.status is VerificationRowStatus.UNVERIFIABLE for record in invalid.records)
    assert invalid.summary.unverifiable_row_count == 3
    assert not invalid.summary.verification_complete


def test_changed_field_is_retained_alongside_unverifiable_field(tmp_path: Path) -> None:
    config = _config(
        compare=("name", "day"),
        source_aliases={"id": ("源编号",), "name": ("源名称",), "day": ("源日期",)},
        system_aliases={"id": ("系统编号",), "name": ("系统名称",), "day": ("系统日期",)},
        source_cleaning=CleaningWorkbookConfig((FieldCleaningRule("day", kind=CleaningKind.DATE,
                           date_input_formats=("%Y-%m-%d",), date_output_format="%Y-%m-%d"),)),
        system_cleaning=CleaningWorkbookConfig((FieldCleaningRule("day", kind=CleaningKind.DATE,
                           date_input_formats=("%Y-%m-%d",), date_output_format="%Y-%m-%d"),)),
    )
    result = _run(tmp_path,
                  [["源编号", "源名称", "源日期"], ["A", "旧", "无效日期"]],
                  [["系统编号", "系统名称", "系统日期"], ["A", "新", "无效日期"]], config)
    record = result.records[0]
    assert record.status is VerificationRowStatus.UNVERIFIABLE
    assert {item.status for item in record.field_differences} == {
        FieldComparisonStatus.CHANGED, FieldComparisonStatus.UNVERIFIABLE,
    }
    assert (result.summary.matched_row_count, result.summary.changed_row_count,
            result.summary.changed_field_count, result.summary.unverifiable_field_count) == (1, 0, 1, 1)
    assert not result.summary.verification_complete and not result.summary.data_consistent
