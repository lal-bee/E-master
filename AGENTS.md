# AGENTS.md

> 用途：为开发智能体（Codex 等）固化本项目的工作上下文、模块边界与行为约束。
> 本文件不是业务需求文档；需求与架构基准见 `PROJECT_SPEC.md`。

## 1. 项目名称

Excel 内容比对与数据质量核验系统

## 2. 当前版本

V1.0（P0-01 ~ P0-10，Excel 核验报告阶段已完成）

## 3. 项目核心目标

把用户提供的业务 Excel 与用户指定的标准数据集进行确定性核验：

1. 结构识别与字段选择不依赖业务硬编码；
2. 对标准映射、基础格式执行可解释的规则核验；
3. 每条错误都可定位到 Sheet、物理行号、列号、列字母与单元格；
4. 核验结果统一为 `ValidationResult`，再交由展示层与 Excel 报告消费；
5. 全程只读原始 Excel，不修改、不覆盖用户数据。

## 4. 核心架构

```text
Excel 文件（.xlsx）
    ↓ P0-01 loader
Sheet/列识别
    ↓ P0-02/P0-03 inspector
字段选择范围
    ↓ P0-04 selection
标准数据集（用户配置）
    ↓ P0-05 standard
标准映射比对 / 格式检查
    ↓ P0-06 comparison / P0-07 format_check
统一错误定位 ValidationResult
    ↓ P0-08 validation
UI 展示 / Excel 核验报告
    ↓ P0-09 ui / P0-10 ui.excel_report
```

分层规则：

- `excel_qc/`：核心领域层，负责导入、识别、选择、标准数据、比对、格式检查、统一结果；
- `ui/`：结果展示与报告层，只消费 `ValidationResult`；
- 上层不得重新实现下层核验逻辑；
- 业务规则与核心代码分离，核心引擎不硬编码具体业务概念。

## 5. P0-01 ~ P0-10 完成状态

- P0-01 Excel 文件导入：已完成
- P0-02 Sheet 识别：已完成
- P0-03 列名识别：已完成
- P0-04 字段选择：已完成
- P0-05 标准 Excel 导入：已完成
- P0-06 标准映射比对：已完成
- P0-07 格式错误检查：已完成
- P0-08 统一错误定位：已完成
- P0-09 核验结果展示：已完成
- P0-10 Excel 核验报告：已完成

V1.0 已形成模块级完整链路：导入 → 识别 → 选择 → 标准数据集 →
比对/格式检查 → `ValidationResult` → UI 展示 → Excel 报告。

## 6. 当前测试状态

- 总测试：141 passed；
- 其中 `tests/qc/` 新项目测试 115 项，旧代码回归 26 项；
- 运行方式：项目虚拟环境执行 `python -m pytest -q`；
- 编译检查：`python -m compileall .` 通过。

## 7. 核心模块职责

| 模块 | 职责 |
|---|---|
| `excel_qc/loader.py` | P0-01：`.xlsx` 只读导入与合法性检查，保留物理行顺序 |
| `excel_qc/inspector.py` | P0-02/P0-03：Sheet、表头、字段与基础类型识别 |
| `excel_qc/selection.py` | P0-04：把用户 Sheet/字段选择解析为固定坐标范围 |
| `excel_qc/standard.py` | P0-05：把标准 Excel 导入为通用 `StandardDataset` |
| `excel_qc/comparison.py` | P0-06：按 `WorkbookSelection` 与 `StandardDataset` 做逐行映射比对 |
| `excel_qc/format_check.py` | P0-07：按用户规则执行字段格式检查 |
| `excel_qc/validation.py` | P0-08：把比对/格式结果统一为 `ValidationResult` |
| `excel_qc/models.py` | P0 数据契约：结构、选择、标准、比对、格式、定位与统一结果模型 |
| `ui/display.py` | P0-09：`ValidationResult` → 展示数据模型与筛选 |
| `ui/html.py` | P0-09：静态 HTML 页面渲染 |
| `ui/excel_report.py` | P0-10：`ValidationResult` → `.xlsx` 核验报告 |

## 8. 开发原则

1. 原始 Excel 只读，任何阶段不得改写用户业务文件或标准文件；
2. 业务规则（设备类型、编码、单位等）不硬编码到核心引擎；
3. UI 层不得重新执行核验逻辑，只消费 `ValidationResult`；
4. Excel 报告层只消费 `ValidationResult`，不读取业务 Excel，不重复执行
   P0-06/P0-07；
5. P0-06 只做标准映射比对，P0-07 只做格式检查，两者结果由 P0-08 并存统一；
6. 错误定位依赖已记录的 Sheet、物理行号、列号与列字母，不重新猜测表头；
7. 新功能必须保持对既有公共模型、函数与测试的向后兼容；
8. 开发必须分阶段，每个阶段先确定最小范围，再实现、补测试、更新文档；
9. 每阶段必须有测试并通过全部回归，才可视为完成；
10. 未经用户明确指令不得进入下一阶段，不擅自扩大或改写需求。

## 9. 禁止事项

- 禁止删除、迁移或重构旧项目 `e_master/`、旧 `tests/`、旧文档；
- 禁止修改旧项目业务代码；
- 禁止修改用户原始 Excel，禁止把报告输出路径指向源文件；
- 禁止在核心引擎中硬编码设备类型、设备编码等业务规则；
- 禁止在 UI/报告层重新实现导入、比对、格式检查逻辑；
- 禁止未经明确指令进入 P0-11/V1.1/V2.0 等后续阶段；
- 禁止“顺手优化”审查中发现的非阻塞问题；
- 禁止执行 `git commit`、`git push`，除非用户明确要求；
- 禁止删除 `.git` 以外的历史文件，禁止执行宽范围破坏性操作。

## 10. 后续开发方式

1. 以 `PROJECT_SPEC.md` 为需求与架构基准，以 `PROJECT.md`/`README.md` 为状态入口；
2. 启动新阶段前，先读取本项目文件和相关设计文档，再向用户确认阶段目标与最小范围；
3. 每个阶段在核心代码完成后补充 `tests/qc/` 测试，并运行
   `python -m pytest -q` 与 `python -m compileall .`；
4. 文档与代码同步更新；
5. 修改行为边界时先更新 `PROJECT_SPEC.md` 与设计文档，再实施；
6. V1.1 候选方向只有在用户明确要求后才可进入，不能由智能体自行开始。
