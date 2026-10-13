"""生成可供 P1-12 界面加载的合成 Excel；演示配置尚未适配真实业务模板。"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from uuid import uuid4

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

CONFIG_ROOT = Path(__file__).with_name("ui_demo_configs")


def create_demo(root: Path):
    """每次使用新子目录，不覆盖已有文件；返回两流程文件角色及 A 默认设置。"""
    root = root / ("演示数据-" + uuid4().hex[:10])
    root.mkdir(parents=True)

    def write(name, header, rows):
        path = root / name
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "演示数据"
        sheet.append(["合成演示资料，尚未适配真实业务模板"])
        sheet.append(header)
        for row in rows:
            sheet.append(row)
        workbook.create_sheet("空白Sheet")
        workbook.save(path)
        workbook.close()
        return path

    good = write("整理用原始资料.xlsx", ["源编号", "源名称", "源类型"],
                 [("0007", " 水泵甲 ", "泵"), ("0008", "风机乙", "风机")])
    bad = write("门禁失败原始资料.xlsx", ["源编号", "源名称", "源类型"],
                [("0007", None, "未知类型"), ("0007", "重复编号", "泵")])
    source = write("最初原始源文档.xlsx", ["源编号", "源名称", "源数量"],
                   [("MATCH", "同值", 1), ("CHANGED", "旧值", 2), ("MISSING", "仅源", 3),
                    ("DUP", "甲", 1), ("DUP", "乙", 1), (None, "无主键", 1)])
    system = write("系统实际导出.xlsx", ["系统编号", "系统名称", "系统数量"],
                   [("MATCH", "同值", 1), ("CHANGED", "新值", "错误数值"),
                    ("EXTRA", "仅系统", 3), ("DUP", "甲", 1)])

    def config(name):
        target = root / (name + ".json")
        with target.open("x", encoding="utf-8") as file:
            file.write((CONFIG_ROOT / target.name).read_text(encoding="utf-8"))
        return target

    return {
        "prepare": {"source": good, **{role: config(role) for role in
                    ("mapping", "cleaning", "equipment", "template", "gate")}},
        "verify": {"source": source, "system": system, **{role: config(role) for role in
                   ("original_config", "system_config", "verification")}},
        "fail_source": bad,
        "settings": {"sheet": "演示数据", "header": 2},
    }


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=Path("demo-data"))
    args = parser.parse_args()
    print(json.dumps(create_demo(args.output), default=str, ensure_ascii=True, indent=2))
