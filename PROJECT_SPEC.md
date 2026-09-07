# PROJECT_SPEC.md

> 文档类型：项目规格说明（需求和架构基准）
> 项目名称：Excel 内容比对与数据质量核验系统
> 当前版本：V1.0（P0-01 ~ P0-10）
> 状态：V1.0 功能已完成，141 passed
> 说明：后续开发 V1.1 或更高版本前，应以本文件为需求和架构基准；
> 新增范围必须先更新本文件并经用户确认，再进入实现。

## 1. 项目背景

设备主数据等结构化 Excel 数据在维护过程中常见以下问题：

1. 同一数据来自不同部门/系统，字段口径不统一；
2. 手工整理文件结构混乱，表头位置、空行、说明行不确定；
3. 标准编码与业务文件中的实际值不一致；
4. 日期、数字、长度、必填等格式错误需要逐单元格定位；
5. 问题数据只给汇总数字时，实施人员无法直接回到 Excel 修改。

因此需要一套“可解释、可定位、只读原文件”的核验工具。

## 2. 项目目标

V1.0 目标：

1. 对 `.xlsx` 业务文件完成 Sheet 与字段结构识别；
2. 由用户选择参与核验的 Sheet 与字段，不依赖业务字段名硬编码；
3. 由用户把标准映射 Excel 导入为通用标准数据集；
4. 支持标准映射比对与基础格式检查两类确定性核验；
5. 所有错误统一为带完整定位信息的 `ValidationResult`；
6. 输出可供人工复核的静态 HTML 展示与 Excel 核验报告；
7. 全程不修改用户原始 Excel。

## 3. 核心业务流程

```text
业务 Excel（.xlsx）
  → 文件只读导入
  → Sheet 识别（空 Sheet、内容范围、表头、数据行数）
  → 列名识别（字段名、Excel 列、基础类型）
  → 用户选择 Sheet/字段（WorkbookSelection）
  → 标准 Excel 导入（StandardDataset，由用户配置匹配字段/标准值字段）
  → P0-06 标准映射比对（PASS/DATA_NOT_FOUND/DATA_ERROR）
  → P0-07 格式检查（PASS/FORMAT_ERROR）
  → P0-08 统一 ValidationResult
  → P0-09 静态 HTML 展示
  → P0-10 Excel 核验报告
```

## 4. V1.0 功能范围

V1.0 = P0-01 ~ P0-10，全部已完成：

| 阶段 | 内容 | 状态 |
|---|---|---|
| P0-01 | Excel 文件导入 | 已完成 |
| P0-02 | Sheet 识别 | 已完成 |
| P0-03 | 列名识别 | 已完成 |
| P0-04 | 字段选择 | 已完成 |
| P0-05 | 标准 Excel 导入 | 已完成 |
| P0-06 | 标准映射比对 | 已完成 |
| P0-07 | 格式错误检查 | 已完成 |
| P0-08 | 统一错误定位 | 已完成 |
| P0-09 | 核验结果展示 | 已完成 |
| P0-10 | Excel 核验报告 | 已完成 |

## 5. P0-01 ~ P0-10 说明

### 5.1 P0-01 Excel 文件导入

- 仅接受 `.xlsx`；
- 文件不存在、类型不支持、文件损坏分别抛出明确领域异常；
- 使用 openpyxl 只读模式，保留 Excel 物理行顺序；
- 不修改源文件。

### 5.2 P0-02 Sheet 识别

- 输出 Sheet 名称、数量、空 Sheet、内容范围、首个有效表头行与数据行数；
- 多列表格要求表头行为文本型；单列表格取首个有内容行；
- 输出 `EMPTY_SHEET`、`NO_HEADER_FOUND`、`BLANK_HEADER_CELL`、
  `DUPLICATE_HEADER_NAME` 等提示。

### 5.3 P0-03 列名识别

- 输出字段名、Excel 列字母/列号、基础数据类型、非空数量与样例值；
- 整数与小数混存统一归为“数值”；
- 空表头使用“列X”占位名。

### 5.4 P0-04 字段选择

- 输入：结构识别结果 + `FieldSelectionRequest`；
- 输出：`WorkbookSelection`，包含各被选 Sheet、表头物理行号、唯一列坐标与
  原始结构快照；
- 字段名、空表头占位名、Excel 列字母均可作为选择标识；
- 重复字段名必须改用列字母；同一列重复选择幂等；
- 空 Sheet 不能参与选择；
- 后续比对/格式引擎只使用该选择范围，不重新猜测表头。

### 5.5 P0-05 标准 Excel 导入

- 输入：标准 Excel 路径 + `StandardDatasetConfig`；
- 输出：通用 `StandardDataset`（数据集名、来源、Sheet、表头、匹配列、标准值列、
  记录列表）；
- 同一匹配值对应相同标准值：保留首条并计数；
- 同一匹配值对应不同标准值：抛出 `StandardMappingConflictError`；
- 匹配字段或标准值字段为空时抛出可定位的领域异常；
- 不绑定设备类型等具体业务概念。

### 5.6 P0-06 标准映射比对

- 输入：`WorkbookSelection` + `MappingCheckConfig` + `StandardDataset`；
- 一次调用针对一个业务 Sheet 的两个已选字段；
- 逐行输出 `PASS` / `DATA_NOT_FOUND` / `DATA_ERROR`；
- 匹配值不在标准数据集：`DATA_NOT_FOUND`；
- 匹配值在标准数据集但实际标准值不一致：`DATA_ERROR`；
- 空数据行跳过；非空行中匹配字段或标准值字段为空时抛出领域异常；
- 不改写业务文件或标准文件。

### 5.7 P0-07 格式错误检查

- 输入：`WorkbookSelection` + 用户格式规则 `FieldFormatRule`；
- 支持 `DATE`、`DATETIME`、`NUMBER`、`INTEGER`、`TEXT`、`BOOLEAN`、
  `required`、`min_length`、`max_length`、可选 `date_format`；
- 错误编码：`FORMAT_001` 日期、`FORMAT_002` 数字、`FORMAT_003` 整数、
  `FORMAT_004` 文本长度、`FORMAT_005` 必填为空、`FORMAT_006` 布尔；
- 只检查 P0-04 已选字段；
- 不改写业务文件。

### 5.8 P0-08 统一错误定位

- 提供 `CellLocation`、`ValidationIssue`、`ValidationResult`、
  `ValidationSummary`；
- `comparison_result_to_validation_result()` 适配 P0-06；
- `format_check_result_to_validation_result()` 适配 P0-07；
- issues 按“Sheet 原始顺序 → 物理行号 → 列号”稳定排序；
- 原有 `ComparisonResult`/`FormatCheckResult` 保持兼容。

### 5.9 P0-09 核验结果展示

- `ui/display.py`：`ValidationResult` → `ValidationDisplay` + 筛选；
- `ui/html.py`：生成无第三方依赖的静态 HTML 页面；
- 展示统计、错误分类、明细、筛选与详情；
- 不执行 Excel 解析、比对或格式校验。

### 5.10 P0-10 Excel 核验报告

- `ui/excel_report.py`：`export_validation_report(ValidationResult, output_path)`；
- 输出固定两个 Sheet：`核验汇总` 与 `错误明细`；
- 错误明细保留原始 issue 顺序及完整定位信息；
- 不读取业务 Excel，不重复执行核验；
- 输出路径等于源文件时拒绝写入。

## 6. 核心数据模型

```text
WorkbookStructure / SheetStructure / ColumnField
    → 结构识别结果

WorkbookSelection / SheetSelection / FieldSelection
    → 用户确定的核验范围与固定坐标

StandardDatasetConfig / StandardDataset / StandardRecord
    → 标准数据集的配置与只读快照

MappingCheckConfig / ComparisonResult / ComparisonRowResult / ComparisonSummary
    → P0-06 结果

FieldFormatRule / FormatCheckResult / FormatIssue / FormatSummary
    → P0-07 结果

CellLocation / ValidationIssue / ValidationResult / ValidationSummary
    → P0-08 统一结果（UI 与报告的唯一输入契约）
```

## 7. ValidationResult 说明

`ValidationResult` 是 P0-09/P0-10 的唯一输入：

```text
ValidationResult
├── file                    # 被核验业务文件
├── summary
│    ├── total_checks       # 总检查数
│    ├── pass_count         # 通过数
│    ├── error_count        # 错误数
│    ├── format_error_count
│    ├── data_not_found_count
│    ├── data_error_count
│    └── skipped_empty_row_count
└── issues                  # ValidationIssue 集合（可跨多个 Sheet）
     ├── location           # CellLocation
     ├── error_type / error_code
     ├── actual_value / expected_value / reason
     └── source             # COMPARISON / FORMAT_CHECK
```

规则：

- UI 与 Excel 报告不修改 `ValidationResult`；
- 排序只由 P0-08 负责，展示与报告沿用既定顺序；
- 当前代码提供单次比对/格式检查到 `ValidationResult` 的适配器；
- 若需把多次比对、多个 Sheet 合并成一个报告，由调用方组装
  `ValidationResult`（V1.0 未提供自动合并工具函数）。

## 8. 错误分类

| error_type | source | 错误编码 | 含义 |
|---|---|---|---|
| `FORMAT_ERROR` | `FORMAT_CHECK` | `FORMAT_001` ~ `FORMAT_006` | 日期/数字/整数/长度/必填/布尔格式错误 |
| `DATA_NOT_FOUND` | `COMPARISON` | `DATA_001` | 匹配值不存在于标准数据集 |
| `DATA_ERROR` | `COMPARISON` | `DATA_002` | 匹配值存在，但标准值不一致 |

通过行不进入 issues，只累计到 summary。

## 9. 错误定位

每条 `ValidationIssue` 至少保留：

| 定位字段 | 含义 |
|---|---|
| `source_file` | 来源业务文件 |
| `sheet_name` | Sheet 名 |
| `row_number` | Excel 物理行号（1 起始） |
| `column_number` / `column_letter` | Excel 1 起始列号与列字母 |
| `column_name` | 字段名（空表头为占位名） |
| `cell_reference` | 由列字母 + 物理行号生成的单元格地址 |
| `actual_value` / `expected_value` | 实际值 / 期望值 |
| `error_type` / `error_code` / `reason` / `source` | 分类与解释 |

定位来自 P0-06/P0-07 已记录的原始坐标，不重新猜测或扫描表头。

## 10. UI 职责

P0-09 UI 职责：

- 展示文件、统计卡与错误分类；
- 展示错误明细，并支持按错误类型、Sheet、字段、错误编码筛选；
- 提供单条错误详情；
- 只消费 `ValidationResult`；
- 不执行任何 Excel 核验逻辑。

## 11. Excel 报告职责

P0-10 Excel 报告职责：

- `核验汇总`：核验文件、结论、总检查/通过/错误、三类错误与跳过空行；
- `错误明细`：序号、文件、Sheet、行号、列号、列字母、字段名、单元格、
  错误类型、错误编码、实际值、期望值、错误原因、来源；
- 报告保留 `ValidationResult.issues` 的原始顺序；
- 只消费 `ValidationResult`，不读取业务 Excel；
- 不覆盖源文件，不自动修复错误。

## 12. 非功能要求

1. 原始 Excel 只读：文件导入、识别、选择、比对、格式检查、报告均不修改源文件；
2. 通用性：核心引擎不硬编码设备类型、设备编码、单位等业务规则；
3. 可解释性：每条错误都有错误编码、原因、实际值、期望值与完整单元格定位；
4. 确定性：核验规则无随机性；除人工配置外不依赖 AI；
5. 可测试性：模块间通过数据契约交互，无全局状态；测试可离线运行；
6. 向后兼容：后续版本不得破坏既有公共模型、函数、错误语义与测试；
7. 技术约束：Python + openpyxl；当前新引擎只读 `.xlsx`；
8. 阶段化：功能按阶段交付，每阶段必须测试与文档同步。

## 13. V1.0 明确不包含的功能

- 新引擎的 CSV、XLS、XLSM 导入（旧 `e_master/` 能力不属于 V1.0 新链路）；
- 复杂多级/合并表头的自动处理，以及无法唯一判定的歧义表头；
- 新核验链路中的尾注行自动识别与排除（旧项目分析器能力未迁移）；
- 自动修改/修复业务 Excel；
- 双业务文件内容比对；
- AI 解释、AI 映射建议、AI 自动修复；
- 历史任务、任务持久化、数据库、登录权限、在线协作；
- 批量目录处理；
- 新引擎的统一 CLI/Web 端到端入口（当前为模块与函数级能力）；
- 多结果自动合并工具（调用方需自行组装 `ValidationResult`）；
- 行级上下文数据访问（“显示整行数据”能力）；
- 实时监控、调度、队列。

## 14. 后续 V1.1/V2 候选方向

以下仅为候选方向，未经用户明确指令不得实施：

### V1.1 候选

1. 核验任务编排层：一个任务配置同时覆盖多个 Sheet/多组映射/多组格式规则；
2. 多结果合并：把多次 P0-06/P0-07 结果合并为一个统一 `ValidationResult`；
3. 校验任务配置与持久化（JSON/YAML），复用同格式文件；
4. 端到端 CLI 或最小 Web API；
5. 行级上下文与“整行导出”能力；
6. 真实设备主数据样本回归集与金标准结果；
7. 更稳健的表头/尾注处理与人工确认流程；
8. 报告阅读体验增强（长文本换行、分组、数据来源追溯）。

### V2 候选

1. FastAPI 服务化；
2. Vue 3 + Element Plus 前端；
3. SQLite 历史任务与审计；
4. 登录权限与多人协作；
5. AI 辅助解释/映射候选（仅建议，不直接改主数据）。
