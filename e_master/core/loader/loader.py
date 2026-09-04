"""读取源文件，统一转成可追溯的原始文本表格。"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd

from e_master.core.errors import (
    CsvLoadError,
    EMasterError,
    ExcelLoadError,
    SourceFileNotFoundError,
    UnsupportedFileTypeError,
)
from e_master.core.text import cell_to_text


@dataclass
class RawSheet:
    """一个 Sheet 的原始内容。

    data 未做表头识别；Excel 物理第 N 行对应 data 的第 N-1 行（0 起始下标），
    保证行级溯源可用。
    """

    sheet_name: str
    data: pd.DataFrame


@dataclass
class RawWorkbook:
    """一个源文件的原始内容。"""

    path: Path
    file_type: str
    sheets: list[RawSheet]


class ExcelLoader:
    """源文件加载器，当前支持 xlsx/xlsm/csv。"""

    EXCEL_SUFFIXES = {".xlsx", ".xlsm"}
    CSV_SUFFIXES = {".csv"}

    def load(self, path: str | Path) -> RawWorkbook:
        source = Path(path)
        if not source.exists():
            raise SourceFileNotFoundError(f"源文件不存在: {source}")
        suffix = source.suffix.lower()
        if suffix in self.EXCEL_SUFFIXES:
            return self._load_excel(source)
        if suffix in self.CSV_SUFFIXES:
            return self._load_csv(source)
        raise UnsupportedFileTypeError(
            f"不支持的源文件类型 '{suffix}'，目前支持: "
            + ", ".join(sorted(self.EXCEL_SUFFIXES | self.CSV_SUFFIXES))
        )

    def _load_excel(self, path: Path) -> RawWorkbook:
        try:
            with pd.ExcelFile(path, engine="openpyxl") as book:
                sheets: list[RawSheet] = []
                for sheet_name in book.sheet_names:
                    frame = pd.read_excel(
                        book,
                        sheet_name=sheet_name,
                        header=None,
                        dtype=str,
                        keep_default_na=False,
                    )
                    frame = frame.map(cell_to_text)
                    sheets.append(RawSheet(sheet_name=sheet_name, data=frame))
        except EMasterError:
            raise
        except Exception as exc:  # noqa: BLE001 - 包装为领域异常
            raise ExcelLoadError(f"无法读取 Excel 文件 {path}: {exc}") from exc
        return RawWorkbook(path=path, file_type=path.suffix.lower().lstrip("."), sheets=sheets)

    def _load_csv(self, path: Path) -> RawWorkbook:
        last_error: Exception | None = None
        for encoding in ("utf-8-sig", "gb18030", "utf-8"):
            try:
                frame = self._read_csv_with_encoding(path, encoding)
            except UnicodeDecodeError as exc:
                last_error = exc
                continue
            except pd.errors.ParserError as exc:
                last_error = exc
                continue
            frame = frame.map(cell_to_text)
            sheets = [RawSheet(sheet_name=path.name, data=frame)]
            return RawWorkbook(path=path, file_type="csv", sheets=sheets)
        raise CsvLoadError(f"无法读取 CSV 文件 {path}（编码识别失败）: {last_error}")

    @staticmethod
    def _read_csv_with_encoding(path: Path, encoding: str) -> pd.DataFrame:
        separator = ExcelLoader._detect_separator(path, encoding)
        return pd.read_csv(
            path,
            header=None,
            dtype=str,
            keep_default_na=False,
            encoding=encoding,
            sep=separator,
            on_bad_lines="error",
        )

    @staticmethod
    def _detect_separator(path: Path, encoding: str) -> str:
        """用前 64KB 内容粗判分隔符：tab > 分号 > 逗号。"""
        try:
            with path.open("r", encoding=encoding, errors="strict", newline="") as handle:
                sample = handle.read(65536)
        except UnicodeDecodeError:
            raise
        counts = {candidate: sample.count(candidate) for candidate in ("\t", ";", ",")}
        if counts["\t"] > counts[";"] and counts["\t"] > counts[","]:
            return "\t"
        if counts[";"] > counts[","]:
            return ";"
        return ","
