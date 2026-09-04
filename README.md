# E-Master

基于 Python 和 AI 能力的 Excel 自动化转换工具：将格式混乱、字段不统一的 Excel 转换为符合系统导入要求的标准文件。

## 当前能力（P1）

- 读取 xlsx / xlsm / csv 源文件（多 Sheet、类型保真、保留原始行号）；
- 输出源文件结构分析报告：候选表头、数据区域、列概况、空值/重复行等基础质量检查；
- 支持 `inspect` 命令，报告输出到控制台与 Markdown 文件。

## 快速开始

```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt -r requirements-dev.txt

# 检查源文件结构
python -m e_master inspect 示例文件.xlsx

# 指定报告输出目录
python -m e_master inspect 示例文件.xlsx --output-dir reports
```

运行测试：

```powershell
python -m pytest
```

## 文档

- [架构设计](docs/architecture.md)
- [开发阶段计划](docs/development-plan.md)

## 开发原则

1. 模块化设计，功能独立测试；
2. API Key 只从 `.env` 读取，禁止硬编码；
3. 每完成一个模块提交 Git。
