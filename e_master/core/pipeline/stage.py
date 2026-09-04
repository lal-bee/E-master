"""流水线阶段基类。"""

from __future__ import annotations

from abc import ABC, abstractmethod

from e_master.core.pipeline.context import JobContext


class Stage(ABC):
    """一个可独立测试的处理阶段。"""

    name: str = "stage"

    @abstractmethod
    def run(self, context: JobContext) -> None:
        """在上下文中执行本阶段。"""
