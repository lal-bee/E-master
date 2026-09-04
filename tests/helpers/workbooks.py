"""在测试中动态生成 Excel/CSV 样例，避免把二进制文件提交进仓库。"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook


def build_workbook(path: Path, sheets: dict[str, list[list[Any]]]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    book.remove(book.active)
    for sheet_name, rows in sheets.items():
        worksheet = book.create_sheet(sheet_name)
        for row in rows:
            worksheet.append(list(row))
    book.save(path)
    return path


def write_csv(
    path: Path,
    rows: list[list[Any]],
    encoding: str = "utf-8-sig",
    delimiter: str = ",",
) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = "\r\n".join(delimiter.join(str(cell) for cell in row) for row in rows)
    path.write_text(payload, encoding=encoding, newline="")
    return path
