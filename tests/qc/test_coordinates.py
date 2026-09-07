"""Excel 坐标工具测试。"""

from __future__ import annotations

import pytest

from excel_qc.coordinates import cell_address, column_letter


@pytest.mark.parametrize(
    ("index", "expected"),
    [
        (0, "A"),
        (25, "Z"),
        (26, "AA"),
        (51, "AZ"),
        (701, "ZZ"),
        (702, "AAA"),
    ],
)
def test_column_letter(index: int, expected: str) -> None:
    assert column_letter(index) == expected


def test_column_letter_rejects_negative() -> None:
    with pytest.raises(ValueError):
        column_letter(-1)


def test_cell_address() -> None:
    assert cell_address(125, 6) == "G125"
