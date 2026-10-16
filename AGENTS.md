# 当前项目开发约束

需求基准：`PROJECT_SPEC.md`（2026-10-14 设备台账 Excel 格式转换工具）。
使用说明：`README.md`；状态：`CURRENT_STATUS.md`；验收：`docs/conversion-acceptance.md`。
历史 V1.0/V1.1、旧阶段与兼容要求已由当前基准取代。

- 仅实现单文件 `.xlsx` 上传、整本预检查、固定转换规则、统计与下载。
- 不增加登录、数据库、批量、任务中心、AI、JSON 配置、多流程及模板管理。
- `excel_qc/rules.py` 管理固定规则和映射读取；`excel_qc/converter.py` 扫描与改写 OOXML 副本。
- `ui/local_app.py` 仅绑定 `127.0.0.1`，`ui/web/` 提供当前单页流程。
- 原始业务文件和映射配置始终只读；不覆盖用户原件，不推断业务编码。
- 不使用工作簿整体重存替换 OOXML 副本，避免丢失未知部件、图片、格式和辅助表。
- 先收集整本阻断问题，通过后才生成结果；问题定位包括工作表、行号、单元格、字段及原值。
- 只转换指定六个精确表头，目标公式、公式范围和合并单元格必须明确阻断。
- 测试使用合成样本及临时目录；针对规则、原件哈希、OOXML 保真和实际浏览器闭环验证。
- 不擅自修改 `docs/` 中的业务 Excel 或映射内容。
- 禁止 commit/push，不进入历史后续阶段。

验证命令：

```powershell
.\.venv\Scripts\python.exe -m pytest -q
.\.venv\Scripts\python.exe -m compileall -q excel_qc ui tests
node tests/browser_smoke.mjs
```
