from __future__ import annotations

from openpyxl.worksheet.cell_range import CellRange

from plan_to_report.excel_io import _split_merged_cells


def test_vertical_merge_keeps_value_on_last_row_only():
    values = [[None, None], [None, None], [None, "活动机制"]]
    notes = _split_merged_cells(values, [CellRange("B1:B3")])

    assert values[0][1] is None
    assert values[1][1] is None
    assert values[2][1] == "活动机制"
    assert notes == []
