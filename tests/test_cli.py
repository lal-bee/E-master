"""CLI 集成测试。"""

from __future__ import annotations

from e_master.cli import main
from tests.helpers.workbooks import build_workbook


def test_inspect_writes_report(tmp_path, capsys):
    source = build_workbook(
        tmp_path / "report.xlsx",
        {
            "设备台账": [
                ["设备编号", "设备名称"],
                ["EQ-001", "空压机"],
            ]
        },
    )
    code = main(["inspect", str(source)])
    output = capsys.readouterr().out

    assert code == 0
    assert "报告已生成" in output
    report = source.with_suffix(".inspect.md")
    assert report.exists()
    content = report.read_text(encoding="utf-8")
    assert "设备台账" in content
    assert "EQ-001" in content
    assert "候选表头" in content


def test_inspect_no_file_flag(tmp_path, capsys):
    source = build_workbook(
        tmp_path / "no-file.xlsx",
        {"Sheet1": [["字段A"], ["值1"]]},
    )
    code = main(["inspect", str(source), "--no-file"])
    assert code == 0
    assert not source.with_suffix(".inspect.md").exists()


def test_inspect_missing_file_returns_error(tmp_path, capsys):
    code = main(["inspect", str(tmp_path / "missing.xlsx")])
    assert code == 2
    assert "错误" in capsys.readouterr().err
