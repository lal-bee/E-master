# CURRENT_STATUS.md

> 用途：记录项目当前真实状态，供后续阶段启动前核对。
> 更新时间：2026-09-08

## 1. 当前版本

- 正式基线：V1.0（P0-01 ~ P0-10，已完成）
- 开发上下文：V1.1（`PROJECT_PLAN_V1.1.md` 已建立；P1-04 已完成，待验收）
- 已进入的功能开发：P1-02 Excel 结构探查增强（已完成、已验收）；
  P1-03 字段映射引擎（已完成、已验收）；
  P1-04 数据清洗引擎（已完成，待验收）

## 2. 当前 Git 状态

- 当前分支：`master`
- 当前 commit：`7a6b0e3 feat: complete P1-02 and P1-03 field mapping foundation`
- 与 `origin/master`：`git status` 显示已同步（ahead 0）
- V1.0 标签：`v1.0.0` 本地存在
- 工作区：P1-02/P1-03 已 commit/push；P1-04 改动尚未 commit；
  `.env`、`.venv/`、缓存均已被忽略，未进入 Git
- 未执行任何 push

## 3. 当前测试结果

使用项目虚拟环境 `.venv\Scripts\python.exe -m pytest -q`：

```text
224 passed
```

其中：V1.0 原测试 141 项保持通过；P1-02 新增 19 项；P1-03 新增 30 项；
P1-04 新增 34 项。

## 4. V1.0 状态

P0-01 ~ P0-10 全部完成：

```text
导入 → Sheet/列识别 → 字段选择 → 标准数据集 → 标准映射比对/格式检查
→ ValidationResult → UI 展示 → Excel 核验报告
```

核心代码位于：

- `excel_qc/`：loader、inspector、selection、standard、comparison、
  format_check、validation、models、errors、coordinates、text
- `ui/`：display、html、excel_report

相关文档：

- `PROJECT_SPEC.md`：V1.0 需求与架构基准
- `docs/qc-p0-design.md`：P0-01 ~ P0-10 设计说明
- `PROJECT.md`、`README.md`：项目状态与说明

## 5. V1.1 当前阶段

- V1.1 计划：`PROJECT_PLAN_V1.1.md`
- 已完成阶段：V1.0 P0-01 ~ P0-10（基线）；
  V1.1 P1-02 Excel 结构探查增强（已完成、已验收）；
  V1.1 P1-03 字段映射引擎（已完成、已验收、已 commit/push）；
  V1.1 P1-04 数据清洗引擎（已完成，待验收）
- 当前阶段：P1-04：已完成，待验收
- 下一阶段：P1-05：设备类型编码映射

## 5.1 P1-02 实现记录

实际修改文件：

- `excel_qc/loader.py`：新增 `include_sheet_metadata=False` 可选参数，
  默认保持 V1.0 只读模式；可获取 Sheet 索引、物理总行列数、合并范围、
  隐藏行/列；
- `excel_qc/profiler.py`：新增 `profile_workbook()` 及
  `WorkbookProfile/SheetProfile/FieldProfile/HeaderCandidate/ProfileIssue`；
- `excel_qc/__init__.py`：仅新增 P1-02 公共导出；
- `tests/qc/helpers.py`：新增合并/隐藏行列测试工作簿构造函数；
- `tests/qc/test_structure_profile.py`：新增 19 项 P1-02 测试。

验证结果：

- P1-02 新测试：19 passed
- 完整测试：160 passed
- compileall：通过

尚未 commit、尚未 push。

## 5.2 P1-03 实现记录

实际修改文件：

- `excel_qc/field_mapping.py`：新增字段映射引擎与配置加载；
- `excel_qc/__init__.py`：新增 P1-03 公共导出；
- `tests/qc/test_field_mapping.py`：新增 30 项 P1-03 测试。

验证结果：

- P1-03 新测试：30 passed
- 完整测试：190 passed
- compileall：通过

已 commit：`7a6b0e3`，已 push 至 `origin/master`。

## 5.3 P1-04 实现记录

实际修改文件：

- `excel_qc/cleaning.py`：新增数据清洗引擎；
- `excel_qc/__init__.py`：新增 P1-04 公共导出；
- `tests/qc/test_cleaning.py`：新增 34 项 P1-04 测试。

验证结果：

- P1-04 新测试：34 passed
- 完整测试：224 passed
- compileall：通过

尚未 commit、尚未 push。

## 6. 与 `PROJECT_PLAN_V1.1.md` 的基线核对结果

一致性确认：

- V1.0 的 P0-01 ~ P0-10 模块、141 项测试与提交基线存在；
- 计划中列出的后续 V1.1 模块（transformer/verification/rules/ai 等）当前尚不存在，
  与“尚未进入 P1-04”一致；
- 旧项目 `e_master/`、旧 `tests/`、旧文档仍保留。

需要记录的不一致/口径差异：

1. `PROJECT_PLAN_V1.1.md` 描述 V1.0 包含“重复数据检查”，但新 V1.0
   `excel_qc/` 仅覆盖重复字段名、标准映射重复值/冲突，不包含通用的
   数据行重复检测；该能力只存在于旧 `e_master/` 分析器中；
2. 计划书中的部分“V1.0 已完成”表述同时混合了旧项目能力；进入 P1 前应以
   实际代码为准，必要时在对应 P1 阶段补充确认。

## 7. 当前开发边界

- 禁止修改 V1.0 业务逻辑、公共 API、测试逻辑；
- 禁止删除旧项目 `e_master/`、旧 `tests/`、旧文档；
- 禁止未经用户明确指令开发 P1-05 或任何后续 V1.1 功能；
- 禁止在未获明确要求时 push 远程仓库。
