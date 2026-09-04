"""流水线上下文：各阶段共享的任务状态。"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from e_master.core.loader.loader import RawWorkbook
from e_master.core.models.profile import WorkbookProfile


@dataclass
class JobContext:
    """一次转换/检查任务的上下文，后续阶段继续在此扩展。"""

    source_path: Path
    raw_workbook: RawWorkbook | None = None
    workbook_profile: WorkbookProfile | None = None
    output_dir: Path | None = field(default=None)
