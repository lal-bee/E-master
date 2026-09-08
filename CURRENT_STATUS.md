# CURRENT_STATUS.md

> 用途：记录项目当前真实状态，供后续阶段启动前核对。
> 更新时间：2026-09-08

## 1. 当前版本

- 正式基线：V1.0（P0-01 ~ P0-10，已完成）
- 开发上下文：V1.1（`PROJECT_PLAN_V1.1.md` 已建立）
- P1-02 Excel 结构探查增强：已完成、已验收
- P1-03 字段映射引擎：已完成、已验收，已 commit/push（7a6b0e3）
- P1-04 数据清洗引擎：已完成、已验收，已 commit/push（36c6a9a）
- P1-05 设备类型编码映射：已完成、已 commit（本地 a2db5c5），暂未 push
- P1-06 标准模板生成：已完成技术实现，待验收

## 2. 当前 Git 状态

- 当前分支：`master`
- 当前本地 commit：`a2db5c5 feat: complete P1-05 equipment type code mapping`
- 相对 `origin/master`：ahead 1，P1-05 尚未 push
- P1-05 暂未 push 原因：GitHub HTTPS 443 网络连接失败；
  不因 push 未完成而阻塞后续开发，也未执行 force push/reset/revert
- V1.0 标签：`v1.0.0` 本地存在
- 工作区：包含 P1-06 未提交改动（新增 `excel_qc/template_generator.py`、
  `tests/qc/test_template_generator.py`，修改 `excel_qc/__init__.py`、
  `AGENTS.md`、`CURRENT_STATUS.md`、`PROJECT.md`）
- 本阶段未执行 commit、push

## 3. 当前测试结果

使用项目虚拟环境 `.venv\Scripts\python.exe -m pytest -q`：

```text
299 passed
```

其中：V1.0 原测试 141 项保持通过；P1-02 新增 19 项；P1-03 新增 30 项；
P1-04 新增 34 项；P1-05 新增 30 项；P1-06 新增 45 项。

`python -m compileall .` 通过；`git diff --check` 无错误。

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
  V1.1 P1-04 数据清洗引擎（已完成、已验收、已 commit/push）；
  V1.1 P1-05 设备类型编码映射（已完成、已 commit，暂未 push）；
  V1.1 P1-06 标准模板生成（已完成技术实现，待验收）
- 当前阶段：P1-06：已完成技术实现，待验收
- 下一阶段：P1-07：导入前质量门禁（尚未开始）

## 5.1 P1-02 ~ P1-05 实现记录

- P1-02：`excel_qc/loader.py`、`excel_qc/profiler.py`、
  `excel_qc/__init__.py`、`tests/qc/test_structure_profile.py`；19 项测试
- P1-03：`excel_qc/field_mapping.py`、`excel_qc/__init__.py`、
  `tests/qc/test_field_mapping.py`；30 项测试；commit `7a6b0e3`
- P1-04：`excel_qc/cleaning.py`、`excel_qc/__init__.py`、
  `tests/qc/test_cleaning.py`；34 项测试；commit `36c6a9a`
- P1-05：`excel_qc/equipment_type_mapping.py`、`excel_qc/__init__.py`、
  `tests/qc/test_equipment_type_mapping.py`；30 项测试；commit `a2db5c5`

## 5.2 P1-06 实现记录

实际修改文件：

- `excel_qc/template_generator.py`：新增 P1-06 标准模板生成引擎；
- `excel_qc/__init__.py`：新增 P1-06 公共导出；
- `tests/qc/test_template_generator.py`：新增 45 项 P1-06 测试；
- 文档：AGENTS.md、CURRENT_STATUS.md、PROJECT.md 同步更新。

实现范围：

- 模板配置模型与 JSON 加载（TemplateColumnConfig/TemplateConfig）；
- 内存模板装配（build_standard_template）：按 P1-02 Sheet 顺序，
  普通字段消费 P1-04 cleaned_value，编码列仅 P1-05 MATCH 写入 system_code；
- 来源追溯 TemplateCellTrace、问题清单 TemplateCellIssue；
- .xlsx 导出（export_standard_template）：仅表头、列顺序与数据值，
  输出路径禁止等于源文件，构建与导出分离；
- 未实现：P1-07 质量门禁、真实业务模板字段固化、AI/模糊匹配/编码猜测。

验证结果：

- P1-06 新测试：45 passed
- 完整测试：299 passed
- compileall：通过
- git diff --check：通过
- 尚未 commit、尚未 push

## 6. 与 `PROJECT_PLAN_V1.1.md` 的基线核对结果

一致性确认：

- V1.0 的 P0-01 ~ P0-10 模块、141 项测试与提交基线存在；
- P1-02 ~ P1-05 均已按计划完成并进入 P1-06；
- P1-07/P1-08/P1-09/P1-10 等后续模块尚未实现；
- 旧项目 `e_master/`、旧 `tests/`、旧文档仍保留。

需要记录的口径说明：

1. `PROJECT_PLAN_V1.1.md` 描述 V1.0 包含“重复数据检查”，但新 V1.0
   `excel_qc/` 仅覆盖重复字段名、标准映射重复值/冲突，不包含通用的
   数据行重复检测；该能力只存在于旧 `e_master/` 分析器中；
2. 计划书中的部分“V1.0 已完成”表述混合了旧项目能力；实际以代码为准；
3. P1-06 模板列配置目前仅使用测试用最小示例，不代表真实业务最终模板；
   真实系统导入模板列集合/顺序待后续业务确认。

## 7. 当前开发边界

- 禁止修改 V1.0 业务逻辑、公共 API、测试逻辑；
- 禁止删除旧项目 `e_master/`、旧 `tests/`、旧文档；
- 禁止未经用户明确指令开发 P1-07 或任何后续 V1.1 功能；
- P1-06 完成验收前不执行 commit/push；
- 未经用户明确要求不 push 远程仓库（P1-05 仍未 push）。
