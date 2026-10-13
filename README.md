# Excel 内容比对与数据质量核验系统

面向设备主数据及其他结构化 Excel 数据的数据质量核验工具：

导入 Excel → 识别 Sheet 与字段 → 选择核验范围 → 导入标准数据 →
规则比对 → 错误分类与定位 → 生成核验报告。

> 原“Excel 格式转换工具”方向已终止；本仓库不再以格式转换为核心功能。

## 当前进度：P0-01 ~ P0-10，V1.1 P1-12 UI 整合

- Excel 文件导入：`.xlsx` 只读读取、文件合法性检查、读取异常处理；
- Sheet 识别：Sheet 名称与数量、空 Sheet、有效行列范围、数据行数；
- 列名识别：首个有效表头、Excel 列号、字段名、基础数据类型；
- 字段选择：用户自由选择参与核验的 Sheet 与字段，输出可直接被后续
  标准数据导入/比对使用的核验范围。
- 标准 Excel 导入：用户指定标准数据 Sheet、匹配字段与标准值字段，
  生成可被后续比对使用的标准数据集。
- 标准映射比对：基于 P0-04 选择范围与 P0-05 标准数据集，逐行输出
  PASS / DATA_NOT_FOUND / DATA_ERROR。
- 格式错误检查：基于 P0-04 已选字段执行通用日期/数字/整数/文本/布尔/
  必填/长度规则，输出 FORMAT_ERROR 与完整单元格定位。
- 统一错误定位：将 P0-06/P0-07 结果统一为 CellLocation +
  ValidationIssue/ValidationResult，供后续展示与报告消费。
- 核验结果展示：展示层只消费 ValidationResult，提供统计、错误分类、
  明细、筛选与详情（纯标准库静态 HTML）。
- Excel 核验报告：只消费 ValidationResult，导出含“核验汇总”与
  “错误明细”的工作簿（.xlsx），不读取源 Excel、不重复核验。
- P1-08 系统导出数据接入：按配置读取指定 `.xlsx` 的 Sheet 和物理表头行，
  复用字段映射与清洗规则，返回标准化行、主键、来源单元格追溯、问题和统计；
  缺失/歧义字段、空/重复主键和清洗失败均保留为问题，不丢行、不去重。
  回源核验基准仍为最初收集的原始源文档。
- P1-09 回源核验引擎：原始源文档与系统导出分别映射、清洗后，按配置主键
  唯一匹配，仅比较指定字段；保留两侧原值、处理值、物理坐标、问题和统计。
  空键、重复键、映射歧义及清洗失败阻止不可靠的整体完成结论。
- P1-10 核验报告：只消费 P1-09 结果，导出核验概要、差异统计、差异明细、
  异常记录四个中文 Sheet；默认拒绝覆盖，保留字段变化和问题来源。

P1-08 技术验收通过；P1-09 技术验收通过，用户已确认 Git 封版。
P1-10 实现及技术验证完成，已提交 `b5bd943`，待用户验收。
P1-11 AI 辅助能力为 **TODO，暂缓**，保留原规划；当前 P1-12 UI 整合
技术验收通过，用户界面体验确认及用户验收待办。P1-13 批量处理和 P1-14 最终验收尚未开始。
真实业务样本验证未完成。两侧字段、Sheet、主键和清洗/比较口径
仍需业务样本确认。演示配置不代表已适配真实系统。

核心代码位于 [excel_qc/](excel_qc/)，测试位于 [tests/qc/](tests/qc/)。

## 本机中文操作界面（P1-12）

使用项目已有虚拟环境启动，无需新增 UI 或 AI 依赖：

```powershell
Set-Location 'D:\Python_code\E-Master'
.\.venv\Scripts\python.exe -m ui.local_app
```

在浏览器打开 [本机工作台](http://127.0.0.1:8765)。端口占用时追加
`--port 8766`，然后访问对应端口；终端 Ctrl+C 停止。

- **原始资料整理**：选择 Excel 与五份 JSON → 选择 Sheet / 物理表头 →
  映射、清洗、编码、门禁预览 → PASS 后导出标准模板。
- **导入后回源核验**：选择最初原始源与系统实际导出 Excel，以及三份 JSON →
  确认源文件角色 → 核验、筛选、追溯 → 导出四 Sheet 问题报告。
- 点击“加载合成演示”可直接体验；另有“加载门禁失败演示”。
  **演示配置，尚未适配真实业务模板。** 系统导入仍由用户在外部完成。
- 更换输入立即使旧结果失效。门禁 FAIL 禁止正式模板导出；核验有差异或
  无法判定问题仍允许导出报告。文件仅上传本机回环服务，原始文件只读。

配置、操作与限制见 [P1-12 UI 整合说明](docs/p1-12-ui-integration.md)。
可单独生成合成 Excel / JSON，供文件选择器加载：

```powershell
.\.venv\Scripts\python.exe examples\ui_demo.py --output .\demo-data
```

每次创建新子目录，不覆盖已有文件。上传快照和服务输出保留于终端打印的临时
会话目录，页面刷新需重新选择；导出通过浏览器下载，应用不接受任意输出路径。

## 测试

```powershell
python -m pytest
```

P1-09 开发前实测：387 passed；本轮专项 21 passed、完整回归 408 passed。
编译检查及 `git diff --check` 通过。详情见
[P1-09 回源核验说明](docs/p1-09-source-verification.md) 与 [CURRENT_STATUS.md](CURRENT_STATUS.md)。

P1-10 开发前完整回归实测 408 passed；本阶段专项 16 passed、完整回归
424 passed，编译及 `git diff --check` 通过。细节见
[P1-10 核验报告说明](docs/p1-10-verification-report.md)。

P1-12 开发前本地实测 **424 passed**，新增专项 **29 passed**，完整回归
**453 passed**；编译、JS 语法与 `git diff --check` 通过。Windows Edge 无头浏览器
实际页面操作通过，包含真实 file input、两流程下载、FAIL 阻断及输入失效；
人工文件对话框与真实业务样本仍待验证。

本次 P1-12 技术验收：专项 **32 passed**，联合 P1-10 报告专项 **48 passed**，
完整回归 **456 passed**；编译、JavaScript 语法、`git diff --check` 通过。
浏览器实测两个标签页隔离、运行与下载锁定、表头未失焦及读取失败后的失效、
HTML 文本转义。一个模板和四份报告实际下载回读通过，报告四 Sheet 与当次 UI
结果一致，前导零、长编码、0、False 和公式样式文本保留，公式单元格数为 0。
技术验收通过不代表用户界面体验确认或真实业务样本验证完成。

## P1-10 最小示例

```python
from excel_qc import verify_source_against_system_export
from ui import export_source_verification_report

result = verify_source_against_system_export(source_path, export_path, config)
report_path = export_source_verification_report(result, "verification-report.xlsx")
```

临时生成两份演示 Excel、核验并导出报告：

```powershell
.\.venv\Scripts\python.exe examples\source_verification_report_demo.py
```

## P1-09 最小示例

演示脚本自行生成两份临时 Excel；配置仅用于演示：

```powershell
.\.venv\Scripts\python.exe examples\source_verification_demo.py
```

```python
from excel_qc import verify_source_against_system_export
from examples.source_verification_demo import demo_config  # 演示配置

result = verify_source_against_system_export("original.xlsx", "export.xlsx", demo_config())
for record in result.records:
    print(record.status.value, record.key_values, record.field_differences)
print(result.summary.verification_complete, result.summary.data_consistent)
```

## P1-08 最小示例

演示配置和脚本： [system_export_demo.json](examples/system_export_demo.json) 与
[system_export_demo.py](examples/system_export_demo.py)。先确认演示配置与输入文件结构
一致，再运行：

```powershell
python examples/system_export_demo.py .\system-export.xlsx
```

代码也可直接调用：

```python
from excel_qc import ingest_system_export, load_system_export_config

config = load_system_export_config("examples/system_export_demo.json")
result = ingest_system_export("system-export.xlsx", config)
for row in result.rows:
    print(row.source_row_number, row.primary_key_value, row.values)
for issue in result.issues:
    print(issue.code, issue.row_number, issue.message)
```

配置和返回值细节见 [P1-08 接入说明](docs/p1-08-system-export.md)。

## 文档

- 根目录规范文件：
  - [AGENTS.md](AGENTS.md)：Codex/AI 开发约束；
  - [PROJECT_SPEC.md](PROJECT_SPEC.md)：V1.0 需求与架构基准；
  - [PROJECT.md](PROJECT.md)：当前开发进度与阶段状态；
- 当前 V1.0 设计文档：[P0-01 ~ P0-10 最小实现说明](docs/qc-p0-design.md)；
- 旧项目历史文档（保留，不属于 V1.0 新链路）：
  [docs/architecture.md](docs/architecture.md)、
  [docs/development-plan.md](docs/development-plan.md)。

## 说明

- 旧项目代码（`e_master/`、旧 `tests/`、旧文档）按约定暂不清理、不迁移、
  不重构，继续原样保留。
- 根目录 `main.py` 与 `test_deepseek.py` 为旧项目遗留入口/连通性脚本，
  V1.0 新链路不使用，继续保留。
- 当前停在 P1-12 等待用户验收；AI 为 TODO，P1-13/P1-14 未开始。
