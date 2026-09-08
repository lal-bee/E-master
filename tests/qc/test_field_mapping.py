"""P1-03 字段映射引擎测试。"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from excel_qc.field_mapping import (
    FieldMappingConfig,
    FieldMappingConfigError,
    FieldMappingStatus,
    FieldMatchRule,
    StandardFieldDefinition,
    load_field_mapping_config,
    map_sheet_fields,
    map_workbook_fields,
    normalize_field_name,
)
from excel_qc.models import BasicDataType
from excel_qc.profiler import (
    FieldProfile,
    SheetProfile,
    WorkbookProfile,
    profile_workbook,
)
from tests.qc.helpers import build_workbook


def _field(
    name: str,
    column_number: int,
    column_letter: str,
) -> FieldProfile:
    return FieldProfile(
        name=name,
        column_index=column_number - 1,
        column_number=column_number,
        column_letter=column_letter,
        data_type=BasicDataType.TEXT,
        non_empty_count=1,
        empty_count=0,
        sample_values=("样例",),
    )


def _sheet(
    name: str,
    fields: tuple[FieldProfile, ...],
    *,
    sheet_index: int = 0,
) -> SheetProfile:
    return SheetProfile(
        sheet_name=name,
        sheet_index=sheet_index,
        empty=not fields,
        total_row_count=2 if fields else 0,
        total_column_count=len(fields) or 1,
        used_first_row=1 if fields else None,
        used_last_row=2 if fields else None,
        used_first_col=1 if fields else None,
        used_last_col=len(fields) if fields else None,
        fields=fields,
    )


def _profile(
    sheets: tuple[SheetProfile, ...],
    *,
    path: Path = Path("业务文件.xlsx"),
) -> WorkbookProfile:
    return WorkbookProfile(path=path, sheets=sheets)


def test_exact_field_name_match() -> None:
    config = FieldMappingConfig(
        (
            StandardFieldDefinition(
                name="设备档案.设备名称",
                aliases=("设备名称", "设备名"),
            ),
        )
    )
    sheet = _sheet("设备档案", (_field("设备档案.设备名称", 1, "A"),))

    result = map_sheet_fields(sheet, config)

    entry = result.entries[0]
    assert entry.status is FieldMappingStatus.MATCH
    assert entry.standard_field_name == "设备档案.设备名称"
    assert entry.match_rule is FieldMatchRule.EXACT_NAME
    assert entry.matched_alias is None


def test_alias_match() -> None:
    config = FieldMappingConfig(
        (
            StandardFieldDefinition(
                name="设备档案.设备名称",
                aliases=("设备名称", "设备名"),
            ),
        )
    )
    sheet = _sheet("设备档案", (_field("设备名", 1, "A"),))

    entry = map_sheet_fields(sheet, config).entries[0]

    assert entry.status is FieldMappingStatus.MATCH
    assert entry.standard_field_name == "设备档案.设备名称"
    assert entry.match_rule is FieldMatchRule.ALIAS
    assert entry.matched_alias == "设备名"


def test_leading_and_trailing_spaces_are_ignored() -> None:
    config = FieldMappingConfig(
        (StandardFieldDefinition(name="设备名称", aliases=("设备名称",)),)
    )
    sheet = _sheet("Sheet1", (_field("  设备名称  ", 1, "A"),))

    entry = map_sheet_fields(sheet, config).entries[0]

    assert entry.status is FieldMappingStatus.MATCH
    assert entry.normalized_source_name == "设备名称"


def test_case_insensitive_match() -> None:
    config = FieldMappingConfig(
        (StandardFieldDefinition(name="EQUIPMENT_NAME"),)
    )
    sheet = _sheet("Sheet1", (_field("equipment_name", 1, "A"),))

    entry = map_sheet_fields(sheet, config).entries[0]

    assert entry.status is FieldMappingStatus.MATCH
    assert entry.standard_field_name == "EQUIPMENT_NAME"
    assert entry.normalized_source_name == "equipment_name"


def test_full_width_to_half_width_match() -> None:
    config = FieldMappingConfig(
        (StandardFieldDefinition(name="EQUIPMENT_NAME"),)
    )
    sheet = _sheet("Sheet1", (_field("ＥＱＵＩＰＭＥＮＴ＿ＮＡＭＥ", 1, "A"),))

    entry = map_sheet_fields(sheet, config).entries[0]

    assert entry.status is FieldMappingStatus.MATCH
    assert entry.normalized_source_name == "equipment_name"


def test_chinese_and_ascii_parentheses_are_unified() -> None:
    config = FieldMappingConfig(
        (
            StandardFieldDefinition(
                name="设备档案.设备名称",
                aliases=("设备名称(中文)",),
            ),
        )
    )
    sheet = _sheet("设备档案", (_field("设备名称（中文）", 1, "A"),))

    entry = map_sheet_fields(sheet, config).entries[0]

    assert entry.status is FieldMappingStatus.MATCH
    assert entry.match_rule is FieldMatchRule.ALIAS
    assert entry.normalized_source_name == "设备名称(中文)"


def test_invisible_and_continuous_whitespace_are_normalized() -> None:
    config = FieldMappingConfig(
        (
            StandardFieldDefinition(
                name="设备名称",
                aliases=("设备名称 测试",),
            ),
        )
    )
    source_name = "设备\u200b名称 \u3000 测试"
    sheet = _sheet("Sheet1", (_field(source_name, 1, "A"),))

    entry = map_sheet_fields(sheet, config).entries[0]

    assert entry.status is FieldMappingStatus.MATCH
    assert entry.normalized_source_name == "设备名称 测试"


def test_unmatched_field_is_reported() -> None:
    config = FieldMappingConfig(
        (StandardFieldDefinition(name="设备名称"),)
    )
    sheet = _sheet("Sheet1", (_field("完全未知的字段", 1, "A"),))

    entry = map_sheet_fields(sheet, config).entries[0]

    assert entry.status is FieldMappingStatus.UNMATCHED
    assert entry.standard_field_name is None
    assert "未匹配" in entry.reason


def test_ambiguous_field_is_not_randomly_chosen() -> None:
    config = FieldMappingConfig(
        (
            StandardFieldDefinition(name="A", aliases=("COMMON",)),
            StandardFieldDefinition(name="B", aliases=("COMMON",)),
        )
    )
    sheet = _sheet("Sheet1", (_field("COMMON", 1, "A"),))

    entry = map_sheet_fields(sheet, config).entries[0]

    assert entry.status is FieldMappingStatus.AMBIGUOUS
    assert entry.standard_field_name is None
    assert "A" in entry.reason
    assert "B" in entry.reason


def test_conflicting_sources_are_both_marked() -> None:
    config = FieldMappingConfig(
        (
            StandardFieldDefinition(
                name="设备档案.设备名称",
                aliases=("设备名称", "名称"),
            ),
        )
    )
    sheet = _sheet(
        "设备档案",
        (
            _field("设备名称", 1, "A"),
            _field("名称", 2, "B"),
        ),
    )

    result = map_sheet_fields(sheet, config)

    assert [entry.status for entry in result.entries] == [
        FieldMappingStatus.CONFLICT,
        FieldMappingStatus.CONFLICT,
    ]
    assert result.summary.match_count == 0
    assert result.summary.conflict_count == 2
    assert all(entry.standard_field_name == "设备档案.设备名称" for entry in result.entries)


def test_empty_header_placeholder_is_not_matched() -> None:
    config = FieldMappingConfig(
        (
            StandardFieldDefinition(
                name="列B",
            ),
        )
    )
    blank_field = _field("", 2, "B")
    sheet = _sheet("Sheet1", (blank_field,))

    entry = map_sheet_fields(sheet, config).entries[0]

    assert entry.status is FieldMappingStatus.UNMATCHED
    assert entry.source_field_name == ""
    assert entry.source_display_name == "列B"
    assert entry.normalized_source_name == ""
    assert entry.standard_field_name is None


def test_source_sheet_name_is_preserved() -> None:
    config = FieldMappingConfig(
        (StandardFieldDefinition(name="设备名称"),)
    )
    sheet = _sheet("技术参数", (_field("设备名称", 1, "A"),))

    result = map_sheet_fields(sheet, config)

    assert result.sheet_name == "技术参数"
    assert result.entries[0].source_sheet_name == "技术参数"


def test_source_column_and_letter_are_preserved() -> None:
    config = FieldMappingConfig(
        (StandardFieldDefinition(name="设备名称"),)
    )
    sheet = _sheet("Sheet1", (_field("设备名称", 7, "G"),))

    entry = map_sheet_fields(sheet, config).entries[0]

    assert entry.source_column_number == 7
    assert entry.source_column_letter == "G"


def test_original_field_name_is_preserved() -> None:
    config = FieldMappingConfig(
        (
            StandardFieldDefinition(
                name="设备名称",
                aliases=("设备名称（中文）",),
            ),
        )
    )
    sheet = _sheet("Sheet1", (_field("设备名称（中文）", 3, "C"),))

    entry = map_sheet_fields(sheet, config).entries[0]

    assert entry.source_field_name == "设备名称（中文）"
    assert entry.normalized_source_name == "设备名称(中文)"


def test_matching_rule_is_alias_for_alias_hit() -> None:
    config = FieldMappingConfig(
        (
            StandardFieldDefinition(
                name="设备名称",
                aliases=("设备名",),
            ),
        )
    )
    sheet = _sheet("Sheet1", (_field("设备名", 1, "A"),))

    entry = map_sheet_fields(sheet, config).entries[0]

    assert entry.match_rule is FieldMatchRule.ALIAS
    assert entry.matched_alias == "设备名"
    assert "别名" in entry.reason


def test_reason_is_present_for_all_statuses() -> None:
    config = FieldMappingConfig(
        (
            StandardFieldDefinition(
                name="设备名称",
                aliases=("名称",),
            ),
            StandardFieldDefinition(name="其他字段"),
        )
    )
    sheet = _sheet(
        "Sheet1",
        (
            _field("设备名称", 1, "A"),
            _field("未知", 2, "B"),
            _field("", 3, "C"),
        ),
    )

    result = map_sheet_fields(sheet, config)

    assert [entry.status for entry in result.entries] == [
        FieldMappingStatus.MATCH,
        FieldMappingStatus.UNMATCHED,
        FieldMappingStatus.UNMATCHED,
    ]
    assert all(entry.reason for entry in result.entries)


def test_json_configuration_loads_and_maps(tmp_path: Path) -> None:
    config_path = tmp_path / "field-mapping.json"
    config_path.write_text(
        json.dumps(
            {
                "standard_fields": [
                    {
                        "name": "设备档案.设备名称",
                        "aliases": ["设备名称", "设备名"],
                    },
                    {
                        "name": "设备档案.设备编码",
                        "aliases": ["设备编号"],
                    },
                ]
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    config = load_field_mapping_config(config_path)
    sheet = _sheet("Sheet1", (_field("设备编号", 1, "A"),))

    entry = map_sheet_fields(sheet, config).entries[0]

    assert entry.status is FieldMappingStatus.MATCH
    assert entry.standard_field_name == "设备档案.设备编码"


def test_normalize_field_name_public_utility() -> None:
    assert normalize_field_name("  设备名称（中文）  ") == "设备名称(中文)"
    assert normalize_field_name("ＡＢＣ　１２３") == "abc 123"
    assert normalize_field_name("设备\u200b名称") == "设备名称"


@pytest.mark.parametrize(
    ("payload", "expected_text"),
    [
        ("{bad json", "无法解析"),
        ("[]", "顶层必须是 JSON 对象"),
        ('{"names": []}', "standard_fields"),
        ('{"standard_fields": [{"name": ""}]}', "name 必须是非空字符串"),
        ('{"standard_fields": [{"name": "设备名称", "aliases": "bad"}]}', "字符串列表"),
        ('{"standard_fields": [{"name": "设备名称", "aliases": ["", "设备名"]}]}', "空别名"),
    ],
)
def test_invalid_json_config_raises_clear_error(
    tmp_path: Path,
    payload: str,
    expected_text: str,
) -> None:
    config_path = tmp_path / "bad-config.json"
    config_path.write_text(payload, encoding="utf-8")

    with pytest.raises(FieldMappingConfigError) as exc_info:
        load_field_mapping_config(config_path)

    assert expected_text in str(exc_info.value)


def test_missing_json_config_raises(tmp_path: Path) -> None:
    with pytest.raises(FieldMappingConfigError) as exc_info:
        load_field_mapping_config(tmp_path / "missing.json")

    assert "不存在" in str(exc_info.value)


def test_duplicate_standard_field_name_raises() -> None:
    config = FieldMappingConfig(
        (
            StandardFieldDefinition(name="设备名称"),
            StandardFieldDefinition(name=" 设备名称 "),
        )
    )

    with pytest.raises(FieldMappingConfigError) as exc_info:
        map_sheet_fields(_sheet("Sheet1", ()), config)

    assert "重复" in str(exc_info.value)


def test_empty_sheet_maps_to_empty_result() -> None:
    config = FieldMappingConfig(
        (StandardFieldDefinition(name="设备名称"),)
    )
    sheet = _sheet("空表", ())
    workbook = _profile((sheet,))

    result = map_workbook_fields(workbook, config)

    assert len(result.sheet_results) == 1
    assert result.sheet_results[0].entries == ()
    assert result.sheet_results[0].summary.total == 0
    assert result.summary.total == 0


def test_multi_sheet_mapping_result() -> None:
    config = FieldMappingConfig(
        (StandardFieldDefinition(name="设备名称"),)
    )
    first = _sheet("设备档案", (_field("设备名称", 1, "A"),))
    second = _sheet("技术参数", (_field("未知字段", 1, "A"),))
    workbook = _profile((first, second))

    result = map_workbook_fields(workbook, config)

    assert len(result.sheet_results) == 2
    assert result.sheet_results[0].summary.match_count == 1
    assert result.sheet_results[1].summary.unmatched_count == 1
    assert result.summary.total == 2
    assert result.summary.match_count == 1
    assert result.summary.unmatched_count == 1


def test_similar_but_not_identical_name_is_not_auto_matched() -> None:
    config = FieldMappingConfig(
        (StandardFieldDefinition(name="设备编码"),)
    )
    sheet = _sheet("Sheet1", (_field("设备编码ID", 1, "A"),))

    entry = map_sheet_fields(sheet, config).entries[0]

    assert entry.status is FieldMappingStatus.UNMATCHED
    assert entry.standard_field_name is None


def test_mapping_does_not_modify_source_file(tmp_path: Path) -> None:
    path = build_workbook(
        tmp_path / "source.xlsx",
        {
            "设备档案": [
                ["设备名称", "设备编码"],
                ["空压机", "EQ-001"],
            ]
        },
    )
    config = FieldMappingConfig(
        (
            StandardFieldDefinition(name="设备名称"),
            StandardFieldDefinition(name="设备编码"),
        )
    )
    before = path.read_bytes()
    profile = profile_workbook(path)

    map_workbook_fields(profile, config)

    assert path.read_bytes() == before
