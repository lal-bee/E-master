"""演示 P1-08：使用演示配置接入一份结构相符的 .xlsx 文件。"""

from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from excel_qc import ingest_system_export, load_system_export_config


def main() -> int:
    if len(sys.argv) != 2:
        print("用法: python examples/system_export_demo.py <system-export.xlsx>")
        return 2
    config = load_system_export_config(
        PROJECT_ROOT / "examples" / "system_export_demo.json"
    )
    result = ingest_system_export(Path(sys.argv[1]), config)
    print(f"文件: {result.source_file}")
    print(f"Sheet: {result.sheet_name}; 数据行: {result.summary.row_count}")
    for row in result.rows:
        print(
            f"第 {row.source_row_number} 行 | 主键={row.primary_key_value!r} "
            f"| 数据={row.values!r}"
        )
    for issue in result.issues:
        print(
            f"问题 {issue.code}: Sheet={issue.sheet_name}, "
            f"行={issue.row_number}, 列={issue.column_letter}: {issue.message}"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
