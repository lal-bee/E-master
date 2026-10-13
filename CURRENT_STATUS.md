# 当前项目状态

> 更新时间：2026-10-13。状态按仓库和本轮实测记录。

## 版本与阶段

- 正式基线 V1.0（P0-01～P0-10）保持不变。
- P1-08 系统导出数据接入已技术验收、提交并推送；基线提交为
  `7cef5a061a6850bc628db955efe8e6b7445ff7ae`，开发 P1-09 前
  `HEAD` 与 `origin/master` 一致，工作区干净。
- P1-09 回源核验引擎：技术验收通过，用户已确认 Git 封版；真实业务样本验证尚未完成。
- P1-10 核验报告：实现及技术验证完成，已提交 `b5bd943`，待用户验收；真实业务样本验证尚未完成。
- 本轮开发前本地 `HEAD` 与已存储的 `origin/master` 均为 `b5bd943`，工作区和暂存区干净；未执行 fetch。
- P1-11 AI 辅助能力：TODO，用户明确暂缓，保留规划，不计为已完成。
- P1-12 UI 整合：当前阶段，技术验收通过，用户界面体验确认及用户验收待办；未提交、未推送。
- P1-13 批量处理、P1-14 最终验收尚未开始。UI 使用确定性规则，不依赖 AI。
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

## P1-10 报告

`ui/source_verification_report.py` 提供
`export_source_verification_report(result, output_path, overwrite=False)`，
仅消费 P1-09 内存结果，导出“核验概要”“差异统计”“差异明细”“异常记录”
四个 Sheet。完成状态与数据一致性分开表达；差异字段和问题来源可追溯。
默认拒绝覆盖，任何情况下都禁止覆盖两侧输入文件，写入失败保留已有报告。
演示脚本和具体口径见 `examples/source_verification_report_demo.py`、
`docs/p1-10-verification-report.md`。开发前完整回归为 408 passed。
本阶段专项 16 passed、完整回归 424 passed；全项目编译、
`git diff --check` 及临时 Excel 演示运行通过。详见阶段文档。

## P1-12 UI 整合

独立入口 `.\.venv\Scripts\python.exe -m ui.local_app`，浏览器访问
`http://127.0.0.1:8765`。纯标准库本地 HTTP + HTML/JavaScript，不新增依赖。
原始资料整理编排 P1-02～P1-07；导入后回源核验编排 P1-09/P1-10（复用 P1-08）。
`profile_configured_sheet` 复用已有物理表头适配逻辑，不重写结构或领域算法。
门禁、结果失效和导出在前后端均有约束；上传副本保留在独立临时会话目录。
两流程均有合成演示和 JSON 配置，明确标注未适配真实业务模板。

开发前完整回归实测 **424 passed**，P1-10 示例实际运行通过。
新增专项 **29 passed**；完整回归 **453 passed**；全项目编译、JavaScript
语法检查与 `git diff --check` 通过。Windows Edge 无头浏览器实际操作通过：
文件上传、两流程、门禁 PASS/FAIL、筛选与明细、两份下载、表头及输入变更失效。
下载文件已用 openpyxl 回读核对。内置 Browser 工具连接失败，使用本机 Edge
自动化完成；人工文件选择对话框、浏览器另存为覆盖选择及真实业务样本未验证。
详情与截图见 `docs/p1-12-ui-integration.md`。不将自动化技术验证记为用户验收。

### 本轮技术验收

验收开始时工作区已有上述 P1-12 未提交改动，完整回归记录为 453 passed。
发现并修复：选文件读取/解码失败时服务端可能沿用旧角色；表头未失焦时
服务端旧结果尚可导出；超出 JavaScript 安全整数范围的值在界面传输中丢精度。
文件选择先撤销旧角色，解析成功才启用新文件；表头首次输入同时撤销服务端
结果，范围未确认时禁用运行；大整数以带类型的十进制文本展示。

P1-12 专项 **32 passed**，联合 P1-10 报告专项 **48 passed**，完整回归
**456 passed**；编译、产品及验证脚本 JavaScript 语法、`git diff --check` 通过。
Edge 自动化验证运行/下载锁定、重复点击、未失焦失效、FileReader 失败、
Sheet/配置更换、失败清除结果、HTML 转义、两个真实标签页隔离及各自下载。
实际下载一个模板和四份报告，全部回读通过；四 Sheet 统计、差异、异常及来源
与各页当次结果一致，公式单元格数为 0，保留前导零、长编码、0、False。

P1-12 技术验收通过；UI 实际调用的 P1-10 报告接入复核通过。
用户界面体验确认、人工文件对话框、Excel 桌面阅读及真实业务样本仍待办。
P1-10、P1-12 均未标记用户已验收。本轮无 commit/push，停在 P1-12。
