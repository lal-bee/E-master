# 当前项目状态

> 更新时间：2026-10-06。状态按仓库和本轮实测记录。

## 版本与阶段

- 正式基线 V1.0（P0-01～P0-10）保持不变。
- P1-08 系统导出数据接入已技术验收、提交并推送；基线提交为
  `7cef5a061a6850bc628db955efe8e6b7445ff7ae`，开发 P1-09 前
  `HEAD` 与 `origin/master` 一致，工作区干净。
- P1-09 回源核验引擎：技术验收通过，用户已确认 Git 封版；真实业务样本验证尚未完成。
- P1-10 尚未开始。
- P1-07 提交 `d8a311e` 在 Git 历史中；其用户验收状态仍以阶段记录为准，
  不因提交事实自行改写。

## P1-09 实现

`excel_qc/source_verification.py` 提供
`verify_source_against_system_export(original_source_path, system_export_path, config)`。
原始源侧以现有 loader、profiler、field_mapping、cleaning 标准化，系统侧直接
复用 P1-08 公共接入接口。两侧分别配置 Sheet、单行物理表头、字段映射和清洗，
按配置的单字段或组合主键匹配，仅比较明确配置的标准字段。

结果包括两侧标准化行、原始值与处理值、规则、文件/Sheet/物理单元格坐标、
行级分类、字段差异、问题和分项统计。空键、重复键、关键字段缺失或歧义、
清洗失败都会保留记录并阻止不可靠的“全部一致”结论。两侧文件只读；不输出报告。
演示脚本为 `examples/source_verification_demo.py`，自行创建临时 Excel，
配置仅供演示。接口和统计口径见 `docs/p1-09-source-verification.md`。

## 测试与限制

- 开发前完整回归：`.venv\Scripts\python.exe -m pytest -q`，387 passed。
- P1-09 专项：21 passed；完整回归：408 passed；
  `.venv\Scripts\python.exe -m compileall -q .` 通过；
  `git diff --check` 通过。详见阶段设计文档。
- 仅支持 `.xlsx`，每侧一次指定一个 Sheet、单行表头；不推断主键变化，
  不补回 Excel 已丢失的文本前导零。业务字段、规则、主键及真实样本待确认。
- 回源核验基准必须是最初收集的原始源文档，整理后的导入模板不能替代。
