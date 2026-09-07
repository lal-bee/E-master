# P0-01 ~ P0-10 最小实现说明

> 项目：Excel 内容比对与数据质量核验系统
> 阶段：P0-01 Excel 文件导入、P0-02 Sheet 识别、P0-03 列名识别、
> P0-04 字段选择、P0-05 标准 Excel 导入、P0-06 标准映射比对、
> P0-07 格式错误检查、P0-08 统一错误定位、P0-09 核验结果展示、
> P0-10 Excel 核验报告
> 状态：已实现并通过测试（141 passed，含旧项目回归 26 项）

## 1. 代码结构

```text
excel_qc/
├── __init__.py      # 包入口：inspect_workbook / load_workbook / select_fields
├── errors.py        # 领域异常：文件读取与字段选择异常
├── coordinates.py   # Excel 列字母与单元格地址工具
├── models.py        # 结构识别与字段选择的数据契约
├── loader.py        # P0-01：.xlsx 只读导入与合法性检查
├── inspector.py     # P0-02 / P0-03：Sheet 与列名识别
├── selection.py     # P0-04：字段选择解析与选择结果校验
├── standard.py      # P0-05：标准 Excel 导入与标准数据集构建
├── comparison.py    # P0-06：标准映射比对引擎
├── format_check.py  # P0-07：格式错误检查
├── validation.py    # P0-08：统一错误定位与适配
└── text.py          # 单元格值文本化与空白判断

tests/qc/            # 新项目测试（115 项）
```

```text
ui/
├── __init__.py      # 结果展示与报告入口
├── display.py       # P0-09：ValidationResult → 展示数据模型 + 筛选
├── html.py          # P0-09：纯标准库静态 HTML 页面
└── excel_report.py  # P0-10：ValidationResult → .xlsx 核验报告
```

## 2. P0-01 Excel 文件导入

- 仅接受 `.xlsx`；文件不存在、类型不支持分别抛出明确异常。
- 使用 openpyxl 只读模式读取，不修改用户原始文件（有字节级回归测试）。
- 损坏或无法解析的文件统一转换为 `WorkbookReadError`，不泄漏底层异常。
- 原始内容按 Excel 物理行顺序保留，供后续行级/单元格级定位使用。

## 3. P0-02 Sheet 识别

每个 Sheet 输出：

- Sheet 名称与数量；
- 是否有内容（空 Sheet 标记为 `empty=True`，并产生 `EMPTY_SHEET`）；
- 有内容的行范围与列范围（Excel 物理行号/列号，1 起始）；
- 首个有效表头行号；
- 表头下方的非空数据行数。

约定：数据行数指“表头下方、任一核验列中存在内容的行”，完全空行不计数。

## 4. P0-03 列名识别

默认取“首个有效表头行”作为字段名来源：

- 多列表格：表头行至少有 2 个非空单元格，且这些值均为文本型（排除纯数字标题）；
- 单列表格：无法区分标题与表头，取首个有内容行；
- 未找到有效表头时输出 `NO_HEADER_FOUND`，字段名需人工确认。

每个字段输出：

- Excel 列字母（A、B、...、AA）；
- 字段名（表头为空时使用“列A”等占位名并提示）；
- 基础数据类型（文本/整数/数值/日期时间/布尔/空/混合等）；
- 非空值数量与前 3 个样例值。

识别提示项：`EMPTY_SHEET`、`NO_HEADER_FOUND`、`BLANK_HEADER_CELL`、
`DUPLICATE_HEADER_NAME`。

## 5. P0-04 字段选择数据模型

选择模型由三层组成，职责如下：

```text
WorkbookSelection
├── path                 # 当前 Excel 文件
├── selections           # 被选中的 Sheet
│    └── SheetSelection
│         ├── sheet_name
│         ├── header_row          # 表头物理行号（1 起始）
│         ├── structure           # 该 Sheet 的原始结构识别结果
│         └── selected_fields
│              └── FieldSelection # 每个被选中的字段
│                   ├── name                  # 字段名
│                   ├── excel_column_letter   # Excel 列字母（如 G）
│                   ├── excel_column_number   # Excel 列号（1 起始，如 7）
│                   ├── column_index          # 0 起始列下标
│                   └── data_type             # 基础数据类型
└── issues              # 选择过程提示（如 NO_FIELD_SELECTED）
```

说明：

- `FieldSelectionRequest` 是用户请求输入，一个请求对应一个 Sheet；
- `FieldSelection` 是解析完成后的稳定字段快照，后续模块直接读取即可，
  不需要再扫描 Excel 或重新猜测表头；
- `SheetSelection.structure` 保留 Sheet 原始识别结果，保证“Sheet 被选中”
  始终可追溯到原始结构；
- `WorkbookSelection` 保存文件路径、被选 Sheet、已选字段及选择提示，
  是 P0-05/P0-06 的统一输入契约。

## 6. P0-04 选择逻辑

- 未出现在请求中的 Sheet 不进入 `WorkbookSelection`；
- 每个 Sheet 内未出现的字段不进入该 Sheet 的 `selected_fields`；
- 字段标识支持字段名（如“设备编码”）、空表头占位名（如“列B”）和
  Excel 列字母（如“B”）；
- 重复字段名必须改用 Excel 列字母选择，避免歧义；
- 空 Sheet 不能参与选择；
- 未知 Sheet、未知字段、同一 Sheet 重复请求均抛出明确异常；
- 同一列在一条请求中重复出现时按幂等处理；
- “Sheet 已选中但未选择字段”是可表达状态：Sheet 保留在
  `WorkbookSelection` 中，`selected_fields` 为空，并产生
  `NO_FIELD_SELECTED` 提示；
- `validate_selection()` 在进入后续核验前校验该状态：未选择 Sheet 或
  某 Sheet 未选择字段时抛出 `NoFieldSelectedError`；
- 选择逻辑只依赖结构识别结果，不硬编码任何业务字段名。

## 7. 与 P0-05 / P0-06 的接口关系

- P0-05 标准 Excel 导入只需读取 `WorkbookSelection.path` 与被选 Sheet 的
  结构信息，即可确定要加载的标准数据范围；
- P0-06 比对引擎遍历 `WorkbookSelection.selections`，从
  `SheetSelection.header_row` 与 `FieldSelection.column_index` 直接读取
  对应列，不再重复猜测表头或列位置；
- 错误定位时，`FieldSelection.excel_column_letter`、
  `excel_column_number` 与 Sheet 物理行号组合即可得到原始单元格坐标。

## 8. P0-05 标准 Excel 导入

### 8.1 功能目标

把用户指定的标准代码表 Excel（如设备类型映射、状态映射、单位映射等）只读导入
为统一的标准数据集，供 P0-06 比对引擎读取。标准数据集与业务字段解耦，
不硬编码任何“设备类型”“设备编码”等字段名。

### 8.2 数据模型

```text
StandardDatasetConfig
├── dataset_name        # 数据集名称
├── sheet_name          # 使用哪个 Sheet
├── match_field         # 哪一列作为匹配字段（字段名或 Excel 列字母）
└── standard_field      # 哪一列作为标准值字段（字段名或 Excel 列字母）

StandardDataset
├── dataset_name
├── source_file         # 来源标准 Excel
├── sheet_name
├── header_row          # 表头物理行号
├── match_column        # StandardColumn：列名/列字母/列号/0 起始下标
├── standard_column     # StandardColumn：列名/列字母/列号/0 起始下标
├── records             # StandardRecord 列表
├── skipped_empty_row_count
└── duplicate_same_value_count

StandardRecord
├── row_number          # 标准 Excel 物理行号（1 起始）
├── match_value         # 规范文本
└── standard_value      # 规范文本
```

标准 Excel 的 Sheet 与表头识别复用 P0-02/P0-03 的 inspector，字段配置在导入时
解析为唯一列；匹配字段和标准值字段不允许是同一列。

### 8.3 重复值处理

- 同一匹配值对应不同标准值：抛出 `StandardMappingConflictError`，异常信息包含
  Sheet、冲突所在的两个物理行号、匹配列字母与两个标准值，不静默覆盖；
- 同一匹配值对应相同标准值：保留首次出现的记录，`duplicate_same_value_count`
  累加，不重复写入 `records`。

### 8.4 空值处理

- 匹配字段为空：抛出 `StandardBlankValueError`，定位到 Sheet、行、Excel 列；
- 标准值字段为空：抛出 `StandardBlankValueError`，定位到 Sheet、行、Excel 列；
- 整行为空：跳过该行，累计到 `skipped_empty_row_count`，不产生记录；
- 标准 Sheet 为空：抛出 `StandardEmptySheetError`；
- 有表头但没有有效数据行：抛出 `StandardNoDataError`。

单元格值统一通过 `cell_text()` 转为规范文本，空值统一为空字符串。

### 8.5 异常处理

- 标准文件不存在/损坏：沿用 P0-01 loader 的 `SourceFileNotFoundError`、
  `WorkbookReadError`；
- Sheet 不存在：`StandardSheetNotFoundError`；
- 字段不存在：`StandardFieldNotFoundError`；
- 字段名重复：`StandardFieldAmbiguousError`，要求改用 Excel 列字母；
- 配置不合法（空数据集名、同列等）：`StandardConfigError`。

### 8.6 与 P0-04 的关系

- P0-04 回答“业务 Excel 中哪些 Sheet/字段参与核验”，结果是 `WorkbookSelection`；
- P0-05 回答“后续核验使用哪个标准数据集、如何取得标准值”，结果是
  `StandardDataset`；
- 两者职责分离，P0-05 不修改 P0-04 的模型与选择语义。

### 8.7 为 P0-06 提供的数据契约

P0-06 可通过 `StandardDataset.records` 直接建立“匹配值 → 标准值”映射；
每条记录带有原始物理行号，列坐标固定在 `match_column`/`standard_column`，
未来错误报告可追溯到标准文件来源。

## 9. P0-06 标准映射比对

### 9.1 功能目标

基于 P0-04 的 `WorkbookSelection` 与 P0-05 的 `StandardDataset`，
对业务 Excel 中用户选中的两个字段执行“标准映射关系校验”，逐行输出
`PASS` / `DATA_NOT_FOUND` / `DATA_ERROR`，不实现格式校验、错误报告或 UI。

### 9.2 比对输入与流程

```text
WorkbookSelection（业务字段与坐标）
      +
MappingCheckConfig（业务 Sheet、匹配字段、标准值字段）
      +
StandardDataset（匹配值 → 标准值）
      ↓
compare_standard_mapping()
      ↓
ComparisonResult（逐行结果 + 汇总）
```

业务 Sheet、匹配字段和标准值字段必须已包含在 P0-04 的
`SheetSelection.selected_fields` 中；引擎按 `FieldSelection` 的列坐标
直接读取原始业务 Excel，不再重新扫描或猜测表头。

### 9.3 结果规则

- `PASS`：匹配值存在于标准数据集，且业务实际标准值与期望值一致；
- `DATA_NOT_FOUND`：匹配值不存在于当前标准数据集，期望值为空；
- `DATA_ERROR`：匹配值存在，但业务实际标准值与标准数据集中的标准值不一致。

本阶段不产生 `FORMAT_ERROR`，该类型属于 P0-07。

### 9.4 结果模型

```text
ComparisonResult
├── business_file
├── sheet_name
├── dataset_name
├── match_column / standard_column（FieldSelection 坐标）
├── rows（ComparisonRowResult 列表）
├── summary（ComparisonSummary）
└── skipped_empty_row_count

ComparisonRowResult
├── sheet_name
├── row_number
├── match_field_name / match_value
├── standard_field_name / actual_value / expected_value
├── status（PASS / DATA_NOT_FOUND / DATA_ERROR）
├── reason
└── 两列各自 Excel 列字母与列号

ComparisonSummary
├── total_rows
├── pass_count
├── data_not_found_count
└── data_error_count
```

每条结果均保留 Sheet、业务物理行号、字段名、列字母/列号、实际值与期望值，
为 P0-08 错误定位保留全部坐标信息。

### 9.5 空业务值处理

- 业务整行为空：跳过并累计到 `skipped_empty_row_count`；
- 非空行中匹配字段为空：抛出 `ComparisonBlankValueError`，定位到行与匹配列；
- 非空行中标准值字段为空：抛出 `ComparisonBlankValueError`，定位到行与标准值列。

空值不会被当作 `FORMAT_ERROR`，也不与 P0-05 空值规则冲突。

### 9.6 重复/冲突标准值

P0-06 直接由 `StandardDataset.records` 构建唯一映射；若数据集在构建或
外部构造时仍存在同一匹配值对应多个不同标准值，比较引擎立即抛出
`StandardMappingConflictError`，不静默选择第一个。

### 9.7 与 P0-04 / P0-05 的关系

- P0-04 提供“业务 Excel 中哪些 Sheet/字段参与核验”；
- P0-05 提供“标准数据集及匹配字段/标准值字段”；
- P0-06 只消费两者，不重新解析输入，也不修改业务/标准 Excel 文件。

## 10. P0-07 格式错误检查

### 10.1 功能目标与职责边界

P0-07 只判断“业务 Excel 单元格格式是否符合通用格式规则”，不判断业务值
是否正确。P0-06 负责 PASS/DATA_NOT_FOUND/DATA_ERROR，P0-07 负责
PASS/FORMAT_ERROR；两者结果模型独立，后续 P0-08 可并存读取，P0-06 语义
未被修改。

P0-07 只检查 P0-04 `WorkbookSelection` 中已选中的 Sheet 和字段，未选择字段
不产生任何格式错误。

### 10.2 格式规则模型

```text
FieldFormatRule
├── sheet_name
├── field            # 字段名或 Excel 列字母（必须已被 P0-04 选中）
├── value_type       # DATE/DATETIME/NUMBER/INTEGER/TEXT/BOOLEAN
├── required
├── min_length
├── max_length
└── date_format      # 可选，仅 DATE/DATETIME 使用
```

规则不绑定具体业务字段名，完全由用户按文件配置。

### 10.3 格式错误编码

| 编码 | 含义 |
|---|---|
| FORMAT_001 | 日期/日期时间格式错误 |
| FORMAT_002 | 数字格式错误 |
| FORMAT_003 | 整数格式错误 |
| FORMAT_004 | 文本长度错误 |
| FORMAT_005 | 必填字段为空 |
| FORMAT_006 | 布尔值格式错误 |

所有错误统一为 `error_type="FORMAT_ERROR"`。

### 10.4 日期/数字/文本/布尔处理规则

- 日期：Excel 日期单元格经 openpyxl 返回 `datetime`；纯日期规则接受无时间
  部分的对象，日期时间规则同时接受纯日期与带时间值。文本日期支持
  `2026-08-01`、`2026/08/01`、`2026.08.01`，日期时间额外支持
  `yyyy-MM-dd HH:mm[:ss]` 形式；非法日期如 `2026-13-45` 判定 FORMAT_001。
  未按日期格式存储的纯数字序列值不猜测为日期；
- 整数：接受 `int`、整数型浮点（如 100.0）与规范整数文本（如 `"100"`）；
  100.5 判定 FORMAT_003；
- 数字：接受 int/float/Decimal 与可解析为有限数值的文本；
- 文本：检查长度（按去除首尾空白后的字符数），超长/过短产生 FORMAT_004；
- 布尔：接受 Excel 布尔值与文本 TRUE/FALSE（不区分大小写），其他值产生
  FORMAT_006。

### 10.5 空值处理

- 完全空白行：跳过，计入 `skipped_empty_row_count`，不产生错误；
- 必填字段为空（None、空字符串、空白字符串）：FORMAT_005；
- 非必填字段为空：通过，不产生错误。

### 10.6 错误定位字段

`FormatIssue` 为每条格式错误保留：Sheet、Excel 物理行号、字段名、列字母、
列号、单元格地址、实际值、实际数据类型、期望格式、错误编码与错误原因，
供 P0-08 直接使用，无需重新扫描 Excel。

### 10.7 与 P0-04 / P0-06 的关系

- 检查范围完全来自 P0-04，不重新定义字段选择；
- 与 P0-06 结果并列存在，P0-07 不修改 `ComparisonResult`；
- 输入为“P0-04 选择结果 + 用户格式规则”，输出为
  `FormatCheckResult`（`FormatIssue` 列表 + `FormatSummary`）。

## 11. P0-08 统一错误定位

### 11.1 公共模型

```text
CellLocation
├── source_file       # 文件名/来源文件
├── sheet_name
├── row_number        # Excel 真实物理行号（1 起始）
├── column_number     # Excel 1 起始列号
├── column_letter     # Excel 列字母
├── column_name       # 字段名
└── cell_reference    # 由列字母 + 物理行号生成，如 G125

ValidationIssue
├── location          # CellLocation
├── error_type        # FORMAT_ERROR / DATA_NOT_FOUND / DATA_ERROR
├── error_code
├── actual_value
├── expected_value
├── reason
└── source            # COMPARISON / FORMAT_CHECK

ValidationResult
├── file
├── summary（ValidationSummary：总检查、通过、错误、分类型错误、跳过的空行）
└── issues（ValidationIssue 集合，支持多个 Sheet）
```

### 11.2 适配层

- `comparison_result_to_validation_result()`：把 P0-06 `ComparisonResult`
  转为 `ValidationResult`；PASS 不计入 issues，DATA_NOT_FOUND/DATA_ERROR
  分别映射为 DATA_001/DATA_002，source=COMPARISON；
- `format_check_result_to_validation_result()`：把 P0-07
  `FormatCheckResult` 转为 `ValidationResult`；错误保留 FORMAT_xxx 编码，
  source=FORMAT_CHECK；
- 原有 `ComparisonResult`/`FormatCheckResult` 及其测试保持兼容，不做删除
  或破坏性重构。

### 11.3 定位与排序规则

- 定位数据直接来自 P0-06/P0-07 已记录坐标，不重新扫描 Excel；
- 行号一律使用 Excel 真实物理行号（表头在第 5 行时首条数据为第 6 行）；
- 列同时保留列号、列字母与字段名；
- `cell_reference` 统一由定位模型生成，不依赖调用方拼接；
- issues 按“Sheet 原始顺序 → 物理行号 → 列号”稳定排序。

### 11.4 为 P0-09 提供的数据接口

P0-09/P0-10 可直接消费 `ValidationResult`：按文件/Sheet 汇总、
展示错误明细与单元格，无需再次读取业务 Excel 定位。

## 12. P0-09 核验结果展示

### 12.1 UI 技术方案

当前项目未引入任何 Web/UI 依赖。P0-09 采用纯 Python 标准库：

- `ui/display.py`：`build_display(ValidationResult)` 把统一结果映射为展示
  数据模型，并提供筛选函数；
- `ui/html.py`：`render_html(ValidationDisplay)` 生成无第三方依赖、可本地
  打开的静态 HTML 页面；
- 不引入浏览器自动化测试框架，展示层测试直接针对展示数据模型与筛选逻辑。

### 12.2 UI 职责

- 展示核验文件、总检查数、通过数、错误数与跳过的空行数；
- 展示错误分类统计：格式错误、数据不存在、数据错误；
- 展示错误明细表：序号/S/行/列/字段/单元格/错误类型/错误编码/实际值/
  期望值/错误原因/来源；
- 支持按错误类型、Sheet、字段、错误编码筛选；
- 点击错误行显示完整错误详情；
- 不执行 Excel 解析、标准数据导入、数据比对或格式校验。

### 12.3 ValidationResult 消费方式

展示层唯一输入为 `ValidationResult`：

```text
ValidationResult
    ↓ build_display()
ValidationDisplay（展示数据模型）
    ↓ filter_errors() / error_detail()
表格与详情
    ↓ render_html()
静态 HTML 页面
```

筛选只返回新的展示行集合，不修改底层 `ValidationResult`；
排序直接沿用 `ValidationResult` 中“Sheet 原始顺序 → 物理行号 → 列号”
的既定顺序，不在 UI 层重新计算。

### 12.4 数据上下文边界

P0-08/P0-09 当前不提供可靠的整行上下文数据访问能力；为保持定位准确，
展示层不扫描 Excel、不按内容猜测行号，数据上下文功能留待架构提供上下文
读取契约后再实现。

## 13. P0-10 Excel 核验报告

### 13.1 技术方案与职责

- 新增 `ui/excel_report.py`，提供
  `export_validation_report(result: ValidationResult, output_path) -> Path`；
- 报告层唯一输入为 P0-08 的 `ValidationResult`，直接读取
  `file / summary / issues`，不打开、不修改被核验的源 Excel；
- 不依赖、不重复实现 P0-06/P0-07 的比对与格式检查逻辑；
- 不依赖 P0-09 的 HTML 展示数据模型，二者并列消费 `ValidationResult`。

### 13.2 输出工作簿

工作簿固定包含两个 Sheet：

```text
核验汇总
├── 核验文件
├── 结论（全部通过 / 发现 N 条错误）
├── 总检查数 / 通过数 / 错误数
├── 格式错误 / 数据不存在 / 数据错误
└── 跳过空行

错误明细
├── 序号
├── 文件 / Sheet / 行号 / 列号 / 列字母 / 字段名 / 单元格
├── 错误类型 / 错误编码
└── 实际值 / 期望值 / 错误原因 / 来源
```

- 错误明细按 `ValidationResult.issues` 既定顺序写入，不在报告层重新排序；
- 无错误时错误明细 Sheet 仍保留表头；
- 表头加粗、冻结首行，错误明细启用筛选，不引入复杂报表能力；
- 输出路径的父目录不存在时自动创建；
- 输出路径与源文件相同时抛出 `ValueError`，防止覆盖用户原始数据；
- 输出文件已存在且路径不同于源文件时，按调用方指定路径覆盖。

### 13.3 与 P0-08/P0-09 的关系

```text
ValidationResult
    ├─→ build_display() + render_html()   # P0-09：静态 HTML
    └─→ export_validation_report()        # P0-10：.xlsx 核验报告
```

定位、分类与排序规则均来自 P0-08；报告层只做序列化输出。

## 14. 关键约定与边界

- Excel 中的日期单元格经 openpyxl 读回后一般为 `datetime`，本阶段统一显示为
  “日期时间”，更精细的格式校验留待后续阶段。
- 整数与小数混存的列统一归类为“数值”，避免误报“混合”。
- 表头位于标题/说明行之后的文件已覆盖测试；更深层或歧义表头（例如前几行均为
  多列文本说明）不在此阶段自动猜测，后续可由用户选择/调整表头行。
- P0-07 不实现 UI、Excel 报告、双文件比对、AI 解释、自动修复、历史任务、
  数据库、P0-08 统一错误定位与 P0-09/P0-10 结果展示/报告导出。
- P0-08 不实现 UI、Web 页面、Excel 报告、双文件比对、自动修复、AI 分析、
  历史任务、数据库、P0-09 结果展示与 P0-10 报告导出。
- P0-09 不实现 Excel 报告导出（P0-10）、自动修改/修复 Excel、AI 解释、
  双文件比对、历史任务、登录权限、数据库持久化、在线协作与复杂前端框架。
- 旧项目（`e_master/`、旧 `tests/`、旧文档）按用户约定暂不清理、不迁移、
  不重构，测试继续通过。
