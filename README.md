# Excel 内容比对与数据质量核验系统

面向设备主数据及其他结构化 Excel 数据的数据质量核验工具：

导入 Excel → 识别 Sheet 与字段 → 选择核验范围 → 导入标准数据 →
规则比对 → 错误分类与定位 → 生成核验报告。

> 原“Excel 格式转换工具”方向已终止；本仓库不再以格式转换为核心功能。

## 当前进度：P0-01 ~ P0-10

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

核心代码位于 [excel_qc/](excel_qc/)，测试位于 [tests/qc/](tests/qc/)。

## 测试

```powershell
python -m pytest
```

当前结果：141 passed（新阶段 115 项 + 旧代码回归 26 项）。

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
- 后续阶段未开始，等待确认。
