"""E-Master 命令行入口。"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from e_master import __version__
from e_master.core.errors import EMasterError
from e_master.core.pipeline import run_inspection
from e_master.report.console import render_console_summary
from e_master.report.markdown import render_markdown

logger = logging.getLogger("e_master")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="e-master",
        description="E-Master：Excel 结构检查与自动转换工具",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    parser.add_argument("-v", "--verbose", action="store_true", help="输出调试日志")
    subparsers = parser.add_subparsers(dest="command", required=True)

    inspect_parser = subparsers.add_parser(
        "inspect",
        help="读取 Excel/CSV 并输出源文件结构分析报告",
    )
    inspect_parser.add_argument("file", type=Path, help="源 Excel/CSV 文件路径")
    inspect_parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="报告输出目录（默认输出到源文件同目录）",
    )
    inspect_parser.add_argument(
        "--no-file",
        action="store_true",
        help="只输出到控制台，不写报告文件",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    if args.command == "inspect":
        return _run_inspect(args)
    parser.error(f"未知命令: {args.command}")
    return 2


def _run_inspect(args: argparse.Namespace) -> int:
    source = Path(args.file)
    try:
        profile = run_inspection(source)
        for line in render_console_summary(profile):
            print(line)

        report_path: Path | None = None
        if not args.no_file:
            output_dir = Path(args.output_dir) if args.output_dir else profile.path.parent
            output_dir.mkdir(parents=True, exist_ok=True)
            report_path = output_dir / f"{profile.path.stem}.inspect.md"
            report_path.write_text(render_markdown(profile), encoding="utf-8")
            print(f"\n报告已生成: {report_path}")
        return 0
    except EMasterError as exc:
        print(f"错误: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:  # noqa: BLE001
        logger.exception("inspect 执行失败")
        print(f"意外错误: {exc}", file=sys.stderr)
        return 2
