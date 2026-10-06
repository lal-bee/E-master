"""P1-09 演示配置与临时 Excel 示例；不代表真实业务模板已适配。"""

from __future__ import annotations

import sys
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from excel_qc import (  # noqa: E402
    FieldMappingConfig,
    OriginalSourceConfig,
    SourceVerificationConfig,
    StandardFieldDefinition,
    SystemExportConfig,
    verify_source_against_system_export,
)


def _write(path: Path, headers: tuple[str, str], values: tuple[str, str]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "演示数据"
    sheet.append(headers)
    sheet.append(values)
    workbook.save(path)


def demo_config() -> SourceVerificationConfig:
    """仅用于演示字段映射；实际业务必须按原始样本分别确认两侧配置。"""
    return SourceVerificationConfig(
        original_source=OriginalSourceConfig(
            "演示数据", 1,
            FieldMappingConfig((
                StandardFieldDefinition("record_id", ("源编号",)),
                StandardFieldDefinition("label", ("源名称",)),
            )),
        ),
        system_export=SystemExportConfig(
            "演示数据", 1,
            FieldMappingConfig((
                StandardFieldDefinition("record_id", ("系统编号",)),
                StandardFieldDefinition("label", ("系统名称",)),
            )),
            primary_key_fields=("record_id",),
        ),
        key_fields=("record_id",),
        comparison_fields=("label",),
    )


def main() -> None:
    with TemporaryDirectory(prefix="p1-09-demo-") as directory:
        root = Path(directory)
        source = root / "original-source.xlsx"
        system = root / "system-export.xlsx"
        _write(source, ("源编号", "源名称"), ("0007", "水泵"))
        _write(system, ("系统编号", "系统名称"), ("0007", "水泵"))
        result = verify_source_against_system_export(source, system, demo_config())
        for record in result.records:
            print(record.status.value, record.key_values)
        print("verification_complete:", result.summary.verification_complete)
        print("matched_row_count:", result.summary.matched_row_count)


if __name__ == "__main__":
    main()
