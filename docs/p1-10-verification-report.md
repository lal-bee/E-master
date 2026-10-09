# P1-10 回源核验报告

状态：实现及技术验证完成，待用户验收。真实业务样本验证尚未完成；P1-11 尚未开始。

## 接口与运行

`ui.export_source_verification_report(result, output_path, *, overwrite=False)`
只消费 P1-09 的 `SourceVerificationResult`，返回 `.xlsx` 输出路径。报告层
不读取两侧业务 Excel，不重做映射、清洗、主键匹配或字段比较，也不修改输入结果。
P0-10 `export_validation_report` 保持原样。

```python
from excel_qc import verify_source_against_system_export
from ui import export_source_verification_report

result = verify_source_against_system_export(source_path, export_path, config)
report_path = export_source_verification_report(result, "verification-report.xlsx")
```

完整演示只使用临时文件，自行生成两份演示 Excel，含一致、变化、缺失、新增、
重复键和空主键，不保存或提交二进制报告：

```powershell
.\.venv\Scripts\python.exe examples\source_verification_report_demo.py
```

演示配置不表示已经适配真实业务模板。

## Sheet 与状态映射

| Sheet | 内容与单位 |
|---|---|
| 核验概要 | 两侧文件、报告生成时间、完成状态、数据一致性结论、输入行及主要统计、限制提示。P1-09 没有核验时间字段，因此明确写“P1-09 结果未提供”，不编造时间或批次。 |
| 差异统计 | 直接使用 P1-09 `summary`；每项注明输入行、匹配对、结果记录、冲突组、字段项或问题对象。问题编码分项来自 `issue_counts`。 |
| 差异明细 | `MATCH` 默认不展开；`FIELD_CHANGED` 与 `UNVERIFIABLE` 中变化或不可判定的字段逐项展开；缺失、新增、重复键等无字段差异的记录使用行级明细。保留双方原值、标准化值、类型、规则、原因、文件/Sheet/物理坐标。 |
| 异常记录 | 每条 P1-09 问题至少一行；多个来源位置逐个展开，包括重复键的全部关联记录。无位置的问题仍单独保留一行。 |

概要中的“核验完成状态”来自 `verification_complete`：真为“已完成”，假为
“未完成”。数据结论先看完成状态：未完成为“无法判定”；已完成且
`data_consistent=True` 为“一致”；已完成且 `data_consistent=False`
为“不一致”。两项原始布尔值另外保留。差异、异常可同时存在：同一匹配对
的已确认变化字段留在差异明细，不可判定问题留在异常记录；报告不改动引擎统计。

统计表中的 `MATCH`、`FIELD_CHANGED`、`MISSING_IN_SYSTEM`、
`EXTRA_IN_SYSTEM`、`DUPLICATE_KEY`、`UNVERIFIABLE` 均保留原码与中文名称。
`matched_row_count` 是一对一匹配对数；重复键是冲突组数；
`changed_field_count` 是字段项数。`UNVERIFIABLE` 可以与匹配对数交叉，
问题数按 P1-09 问题对象统计，异常 Sheet 因多位置展开后行数可能更大。
这些单位不能直接相加，报告不输出“准确率”。

## 值与写入安全

主键采用含组成项类型的 JSON 文本表示，保留组合边界、前导零和长整数。
原值、标准化值统一按文本写入并另列类型：`None` 显示 `None`，空字符串
显示 `""`，布尔值为 `True/False`，数字和 `Decimal` 采用十进制文本，
日期/时间采用 ISO 格式。未知值类型明确报错。所有文本强制写为 Excel
字符串；即使以 `=`、`+`、`-`、`@` 开头，也不成为公式。

输出只接受 `.xlsx`。默认拒绝覆盖已有文件；`overwrite=True` 也不得覆盖
原始源或系统导出文件。保护检查结合规范化路径及同一文件判断，可识别
相对路径、符号链接和现存硬链接。先在输出目录写入临时文件，成功后以
独占硬链接创建新报告，或在显式覆盖时原子替换；失败只删除本次临时文件，
保留原报告；默认创建方式要求输出文件系统支持同目录硬链接。
使用 `openpyxl` 写入；非法控制字符、超过 32767 字符的单元格
或超过 1048576 行的 Sheet 明确报错，不截断、不丢行。

四个 Sheet 使用中文表头、加粗表头、冻结首行；明细 Sheet 有筛选与适度状态
着色。空明细只有表头，不插入占位业务行。不使用公式计算结论。

## 验证与待确认

开发前完整回归实测 408 passed。本阶段专项
`tests/qc/test_source_verification_report.py` **16 passed**；完整回归
**424 passed**。`.venv\Scripts\python.exe -m compileall -q .` 和
`git diff --check` 通过；端到端临时 Excel 演示脚本实际运行通过。
专项覆盖四 Sheet 回读、完成/一致状态、统计单位、多字段变化、重复键来源、
字段缺失与空值、差异与异常并存、长文本与前导零、公式前缀文本、
容量与非法字符报错、覆盖保护、失败清理及输入结果不变。

真实原始源和系统导出样本仍未验证；真实字段别名、清洗规则、主键含义及
报表使用人员的列展示习惯待样本确认。P1-11 AI、UI 整合和批量处理未开始。
