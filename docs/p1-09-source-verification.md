# P1-09 回源核验引擎

状态：技术验收通过，用户已确认 Git 封版。真实业务样本验证尚未完成；P1-10 尚未开始。

## 范围与入口

`excel_qc.verify_source_against_system_export(original_source_path, system_export_path, config)`
返回纯内存 `SourceVerificationResult`。第一个文件必须是最初收集的原始源文档；
整理后的导入模板不能替代它。第二个文件经 P1-08 `ingest_system_export` 接入。
源侧以现有 `load_workbook`、`profile_workbook`、`map_workbook_fields`、
`clean_workbook` 组成薄适配层，身份固定为 `ORIGINAL_SOURCE`。两侧分别配置
Sheet、物理表头行、字段别名和清洗规则；字段最后映射到相同的标准字段。
文件均只读，未调用 P1-10 报告接口，也未修改 P1-08 的公共模型。

配置由 `OriginalSourceConfig`、`SystemExportConfig`、`SourceVerificationConfig`
组成。`key_fields` 支持单字段或有序组合，须与系统导出配置的主键完全一致。
`comparison_fields` 明确列出要比较的标准字段；未列入的字段不参与比较。
每个比较字段可选 `FieldComparisonRule`。无规则时为 `EXACT`。

## 标准化与比较

清洗只执行两侧显式配置的 P1-04 规则。`NormalizedFieldValue` 同时保留原始值、
标准化值、规则、状态和来源坐标。`VerificationLocation` 记录文件、Sheet、物理行、
原列名、标准字段名、列号、列字母和单元格地址。行级来源保留文件、Sheet 和物理行。
字段状态区分 `PRESENT`、`EMPTY`、`FIELD_MISSING`、`MAPPING_FAILED`、
`CLEANING_FAILED`。字段不存在或两侧都清洗失败不会判为一致。

`EXACT` 要求 Python 值及类型均一致；`None` 只与 `None` 相同，空字符串只与空字符串
相同，二者不折叠；`0` 与 `False`、`1` 与 `True` 不相同。`TEXT` 逐字符比较，
不隐式去空格或忽略大小写。`BOOLEAN` 只接受布尔值。`NUMBER` 是显式规则，
允许有限数值和无空白数字文本转成 `Decimal` 比较，不接受布尔值。
`DATE` 是显式规则，接受日期、日期时间及 `YYYY-MM-DD` 文本，按日历日期比较；
日期时间的时分秒在此规则下不参与比较。未指定规则不会转型。
文本编码的前导零保留；Excel 数值单元格已丢失的前导零不补齐。

## 匹配、分类及无法核验

主键组成项使用带类型的结构化元组，不拼接字符串。两侧均有效且唯一时才一对一
匹配。任一侧出现重复键，关联的双方所有记录组成一条 `DUPLICATE_KEY` 冲突记录，
不继续产生普通匹配、缺失或新增结论。空键、无效键、关键字段不存在、映射歧义和
清洗失败保留输入行和问题；一侧有无效键时，另一侧未配对的键也标为
`UNVERIFIABLE`，避免误报普通缺失或新增。匹配后的字段若有不可判定项，该行
为 `UNVERIFIABLE`，同时保留逐字段状态和原因。

可判定分类为 `MATCH`、`FIELD_CHANGED`、`MISSING_IN_SYSTEM`、
`EXTRA_IN_SYSTEM`；另有 `DUPLICATE_KEY`、`UNVERIFIABLE`。
仅凭源侧缺失与系统侧新增无法可靠推断 `KEY_CHANGED`，本阶段不自动判定，
也不使用名称相似度猜测。

## 统计口径

`SourceVerificationSummary` 分别统计两侧输入行、唯一键匹配行、完全一致行、
变化行、系统缺失行、系统新增行、重复键冲突组、不可核验行，以及已比较、
一致、变化、不可核验字段数。一行多个字段变化只计一条变化行。
`matched_row_count` 包括唯一键匹配后字段不可判定的行，后者同时计入
`unverifiable_row_count`。重复键按冲突组统计，不把冲突记录计作一致。
每条输入记录恰好落入一条结果记录的对应侧来源列表；匹配对数、两侧输入行数、
冲突组数、问题数和字段数是不同单位，不能直接相加。
`verification_complete` 表示配置范围内的核验是否可判定；已确认的字段变化、
缺失或新增仍可使其为真。`data_consistent` 仅在核验完成且全部记录均为 `MATCH`
时为真。若同一匹配对同时有变化字段与不可判定字段，行级为 `UNVERIFIABLE`，
变化字段仍计入 `changed_field_count`，但不计入 `changed_row_count`。
阻断问题或不可核验记录令 `verification_complete=False`；两侧都无数据行时
返回 `NO_DATA_ROWS`，两个布尔值均为假。没有含义不明确的总体准确率。
数据异常通过 `VerificationIssue` 的输入、匹配、比较类别追溯；配置无效抛出
`SourceVerificationConfigError`。缺失 Sheet 是输入问题；配置表头行超出空 Sheet
范围时，源侧返回 `HEADER_ROW_NOT_FOUND`，系统侧沿用 P1-08 的
`SystemExportDataError`，不会将接入失败伪装为正常空数据集。仅支持 `.xlsx`、
每侧一次一个 Sheet、单行表头；不支持的文件由现有 loader 明确报错。

## 演示与验证

演示脚本自行生成两份临时 Excel，并标注配置仅用于演示：

```powershell
.\.venv\Scripts\python.exe examples\source_verification_demo.py
```

调用核心接口：

```python
from excel_qc import verify_source_against_system_export
from examples.source_verification_demo import demo_config  # 演示配置

result = verify_source_against_system_export("original.xlsx", "export.xlsx", demo_config())
for record in result.records:
    print(record.status.value, record.key_values, record.field_differences)
print(result.summary.verification_complete, result.summary.data_consistent)
```

开发前完整回归基线：387 passed。技术验收开始前复测专项 16 passed、
完整回归 403 passed。修复并补充测试后的最终结果见下方验收记录。
`.venv\Scripts\python.exe -m compileall -q .` 和
`git diff --check` 均通过（Git 对 LF/CRLF 给出换行提示，无空白错误）。

| 验收项 | 对应专项测试与断言 |
|---|---|
| 唯一键匹配与异常并存 | `test_invalid_key_does_not_suppress_valid_matches`：正常 `MATCH`、`FIELD_CHANGED` 与不可判定记录共存；未配对键有原因 |
| 重复键和逐行归属 | `test_duplicate_group_does_not_suppress_unique_match_or_lose_rows` 与 `test_duplicate_key_conflict_retains_both_sides`：冲突组保留两侧全部行，旁侧唯一键正常匹配，输入物理行恰好出现一次 |
| 空数据与接入失败 | `test_empty_datasets_and_one_side_empty`、`test_missing_sheet_is_input_issue_and_keeps_other_side`、`test_system_ingestion_failure_is_not_an_empty_dataset`、`test_empty_sheet_and_all_invalid_keys_are_not_success` |
| 比较与统计 | `test_missing_extra_and_multi_field_change_count`、`test_changed_field_is_retained_alongside_unverifiable_field`、`test_exact_value_types_and_explicit_number_rule`、`test_none_empty_zero_false_leading_zero_and_dates` |
| 失败追溯及只读 | `test_empty_key_and_cleaning_failure_keep_rows`、`test_end_to_end_match_and_provenance_with_late_headers`、`test_system_result_is_not_mutated_and_demo_runs`；`_run` 对每对输入文件断言处理前后字节一致 |
| 配置与示例 | `test_invalid_config_and_missing_key_schema`、`test_missing_field_and_ambiguous_mapping_are_unverifiable`、`test_each_side_can_use_different_sheet_and_header`；示例脚本实际运行 |

开发中发现并修复：系统侧清洗失败起初未关联到字段状态；
一侧无效主键时，另一侧未配对记录可能被误报为普通缺失/新增；
双侧空数据原本可能被视为已完成。现分别标记 `CLEANING_FAILED`、
`COUNTERPART_KEY_UNAVAILABLE`、`NO_DATA_ROWS`，并有专项断言。
本轮验收发现完成状态与数据一致性没有独立字段，已增加
`SourceVerificationSummary.data_consistent`；同时补足异常键与正常键并存、
记录归属、空 Sheet、单侧接入失败和已确认差异与不可判定共存的断言。

最终实测：P1-09 专项 **21 passed**；完整回归 **408 passed**；编译检查、
`git diff --check` 和临时 Excel 演示脚本运行通过。以上为本轮实际执行结果。

真实业务原始源与导出样本尚未提供，因此两侧 Sheet、字段别名、主键含义、
清洗口径、日期与数字字段的比较规则均待样本确认。本阶段不输出 Excel 差异报告。
