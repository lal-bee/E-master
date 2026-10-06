# 当前项目状态

> 更新时间：2026-10-06。本文状态按当前仓库、Git 和本轮实测更新。

## 1. 版本与 Git

- 正式基线：V1.0（P0-01～P0-10）；P1-02～P1-06 的既有阶段记录见 Git 历史；
- P1-07 导入前质量门禁：代码提交 `d8a311e`，技术实现完成，验收仍待用户确认；
- P1-08 系统导出数据接入：技术验收通过；真实业务样本验证未完成；
- 下一阶段：P1-09 回源核验引擎，尚未开始；
- 当前分支：`master`。实现前 `HEAD` 为 `84b64fd`，且与 `origin/master` 相同；
- 本轮不执行 commit、push、reset、revert 或清理文件；工作区改动均为本轮 P1-08。

## 2. P1-08 范围与接口

新增 `excel_qc/system_export.py`。入口为 `ingest_system_export(path, config)`，
配置包含系统导出 Sheet 名、物理表头行、标准字段及别名、复用的 P1-04 清洗规则、
单字段或组合主键。Excel 只读，接入复用 P0 loader、P1-02 profiler、P1-03
field_mapping 和 P1-04 cleaning。

返回 `SystemExportWorkbookResult`，包含源文件、Sheet、表头坐标、逐行标准化数据、
逐字段来源追溯、问题明细、统计汇总，并提供底层 mapping/cleaning 结果供后续模块消费。
来源追溯包括物理行号、原字段名、列号/列字母、标准字段名、原始值、标准化值、
应用规则及映射状态。问题包含 Sheet/表头/字段映射异常、缺失字段、清洗失败、
空主键和重复主键。异常记录不丢弃，重复键不覆盖、不去重。

演示配置与调用脚本：`examples/system_export_demo.json`、
`examples/system_export_demo.py`。示例仅作接口演示，没有声称适配真实系统导出模板。

限制：当前支持 `.xlsx` 和每次调用一个指定 Sheet；表头按单行配置；不推断主键，
不补 Excel 数值单元格中已丢失的编码前导零。文本编码的前导零会保留。真实系统字段、
Sheet、主键含义、空值和清洗口径仍需真实导出样本确认。

## 3. 测试与检查

- 实现前基线：项目虚拟环境 `.venv\Scripts\python.exe -m pytest -q`，实测
  `371 passed`；
- P1-08 专项：`tests/qc/test_system_export.py`，16 项通过；
- 本轮完整回归：`387 passed`（371 项基线 + 16 项 P1-08）；
- `.venv\Scripts\python.exe -m compileall -q .`：通过；
- `git diff --check`：通过；有 Git 关于 LF/CRLF 的提示，无空白错误。

## 4. 阶段边界

- 本轮只实现 P1-08；
- P1-09 的双数据源匹配、字段差异比较和差异分类尚未开始；
- P1-10 报告、AI 辅助、UI 整合和批量任务系统尚未开始；
- 回源核验的正确基准是最初收集的原始源文档，整理后的导入模板不能替代它。
- 技术验收结论记录于 `docs/p1-08-system-export.md`；真实系统模板、字段和业务规则
  仍待样本验证，未因此启动 P1-09。

## 5. 历史提交状态更正

旧状态曾记录 P1-05/P1-06 暂未推送及 P1-07 未提交。当前 Git 显示 `HEAD` 与
`origin/master` 同为 `84b64fd`，P1-05、P1-06 和 P1-07 对应提交均在其历史中；
因此不再保留“当前尚未推送/尚未提交”的旧说法。提交事实不改变各阶段是否已被用户验收。
