"""P1 检查流水线：读取 -> 结构分析。"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from e_master.core.analyzer.analyzer import WorkbookAnalyzer
from e_master.core.loader.loader import ExcelLoader
from e_master.core.models.profile import WorkbookProfile
from e_master.core.pipeline.context import JobContext
from e_master.core.pipeline.stage import Stage


@dataclass
class LoaderStage(Stage):
    """读取源文件。"""

    name: str = "loader"

    def run(self, context: JobContext) -> None:
        context.raw_workbook = ExcelLoader().load(context.source_path)


@dataclass
class AnalyzerStage(Stage):
    """对原始内容做结构探查。"""

    name: str = "analyzer"

    def run(self, context: JobContext) -> None:
        if context.raw_workbook is None:
            raise RuntimeError("AnalyzerStage 需要先执行 LoaderStage")
        context.workbook_profile = WorkbookAnalyzer().analyze(context.raw_workbook)


def run_inspection(source_path: str | Path) -> WorkbookProfile:
    """执行一次完整的结构检查，返回结构画像。"""
    context = JobContext(source_path=Path(source_path))
    for stage in (LoaderStage(), AnalyzerStage()):
        stage.run(context)
    assert context.workbook_profile is not None
    return context.workbook_profile
