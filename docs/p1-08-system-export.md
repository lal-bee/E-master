# P1-08 系统导出数据接入

P1-08 将业务系统导出的 `.xlsx` 作为第二个只读数据源接入，为后续 P1-09
提供标准化行记录和从系统导出单元格到处理值的追溯。回源核验基准仍是最初收集的
原始源文档；本阶段不读取该原始源文档，也不执行两数据源匹配或字段差异比较。

## 接入流程

`ingest_system_export(path, config)` 使用既有 `loader` 读取工作簿，按配置选取
Sheet 和物理表头行，建立供既有 `field_mapping` 使用的结构，然后调用 P1-03
字段映射和 P1-04 清洗引擎。映射歧义、未映射字段、缺失字段、空主键、重复主键和
清洗失败均以 `SystemExportIssue` 明确返回；数据行始终保留，不按主键覆盖或去重。

`SystemExportWorkbookResult` 包含系统导出文件、Sheet、表头行、逐行数据、问题、
统计，以及底层 `mapping_result` 和 `cleaning_result`。每项标准字段值保留原字段名、
列号/列字母、标准字段名、原始值、标准化值、应用规则和映射状态。每行保留物理行号、
单字段或组合主键字段及主键值。未配置清洗规则的值维持 loader 读到的类型和值；文本
单元格中的前导零保持原样。Excel 数值单元格若已丢失前导零，本模块不会推断或补齐。

## 配置与最小调用

`examples/system_export_demo.json` 和 `examples/system_export_demo.py` 是演示用途，
没有适配或验证任何真实业务系统导出模板。按真实文件确认配置后运行：

```powershell
python examples/system_export_demo.py .\system-export.xlsx
```

也可直接在 Python 中构造 `SystemExportConfig`，使用 `FieldMappingConfig`、
`StandardFieldDefinition` 和 `CleaningWorkbookConfig` 配置字段及规则。Sheet 名、
表头行、字段别名和主键均由配置指定，没有字段猜测逻辑。

## 当前限制

- 仅支持 `.xlsx`，一次调用接入一个配置指定的 Sheet；
- 多行表头、合并表头、自动 Sheet 选择和自动主键推断不在本阶段范围；
- 所有 `standard_fields` 都按期望字段处理，未唯一映射时产生 `MISSING_FIELD`；
- 真实系统导出的字段、主键语义、清洗规则和空值口径待业务样本确认；
- 未实现 P1-09 的双数据源匹配、差异比较与分类，也未实现报告、AI、UI 或批量处理。

## 技术验收记录（2026-10-06）

**结论：P1-08 技术验收通过。** 全量回归、专项检查、编译和差异格式检查均通过；
真实业务系统导出样本尚未提供，因此真实模板适配、业务主键语义及清洗口径尚未验证。
本次验收停在 P1-08，不启动 P1-09。

| 检查项 | 依据与结论 |
|---|---|
| 复用既有读取、探查、映射与清洗 | `ingest_system_export` 调用 `load_workbook`、`profile_workbook`、`map_workbook_fields`、`clean_workbook`；`test_temporary_excel_end_to_end_ingestion_integration` 覆盖完整接入链路。 |
| 接口、JSON 配置与脚本一致 | `test_json_config_loader_and_downstream_result_shape` 验证配置字段和结果；`test_minimal_example_runs_with_temporary_demo_workbook` 用临时 Excel 实际启动示例脚本。 |
| 原值、处理值、规则及坐标 | `test_ingests_configured_sheet_header_mapping_and_cleaning` 检查表头在第 2 行、数据在第 3 行和 B 列来源；`test_cleaning_failure_keeps_raw_value_and_source_coordinates` 检查 B2 坐标与失败原值。 |
| 映射问题和记录保留 | `test_missing_fields_and_ambiguous_mapping_are_reported_without_guessing` 与 `test_same_source_alias_for_two_standard_fields_is_ambiguous` 检查冲突/歧义/缺失字段；结果不输出猜测值。 |
| 空值、0、False 与编码前导零 | `test_preserves_zero_false_blank_and_text_leading_zero` 通过实际临时 `.xlsx` 检查 None、0、False 和文本 `0007`；`test_none_and_empty_string_remain_distinguishable_in_result` 验证 loader 返回的 None 与空字符串在接入结果中保持区分。 |
| 空键、重复键、组合键 | `test_empty_duplicate_and_composite_keys_are_retained_and_reported` 断言空键和重复行均保留；`test_composite_keys_do_not_collide_by_string_concatenation` 检查 `("ab", "c")` 与 `("a", "bc")` 不被误判重复。 |
| 空表与不支持的输入 | `test_empty_sheet_header_only_and_no_effective_data_are_explicit` 验证空 Sheet 明确报错、仅表头/无有效数据返回零行；`test_non_xlsx_input_is_rejected_explicitly` 验证 `.xls` 被 loader 明确拒绝。 |
| 只读及既有接口兼容 | `test_input_workbook_is_unchanged_and_empty_intermediate_rows_are_preserved` 比较处理前后文件 SHA-256；完整回归覆盖旧公共接口。 |
| PROJECT_SPEC 范围 | 本次仅在 `PROJECT_SPEC.md` 追加 P1-08 补充约定；原 V1.0 业务边界未被改写。 |

本轮验收发现并修复：最小示例脚本按文档命令启动时无法导入仓库根目录的
`excel_qc` 包；脚本现在将项目根目录加入模块路径，并由临时 Excel 子进程测试验证。
同时补足了组合主键碰撞、空表/仅表头、非 `.xlsx` 输入和 None/空字符串传递的验收断言。

实际验证结果：

- P1-08 专项：16 passed；
- 完整回归：387 passed；
- `.venv\Scripts\python.exe -m compileall -q .`：通过；
- `git diff --check`：通过（Git 有 LF/CRLF 转换提示，无空白错误）。

真实系统的 Sheet、表头、字段别名、字段完整性、单/组合主键口径、日期/数值清洗规则、
业务空值标记及异常等级仍待真实业务样本和业务规则确认。空白单元格的具体表示受
Excel 文件内容及 openpyxl 读取结果影响；接入层保留 loader 提供的原始 Python 值，不做
空值猜测。文本形式的编码前导零会保留，Excel 数值单元格中已丢失的前导零无法恢复。
