from __future__ import annotations

import pandas as pd
from PySide6.QtWidgets import QComboBox, QTableWidget

from .excel_io import ParsedSheet, table_from_range
from .plan_config import SheetRange


def load_role_table(
    parsed_sheets: list[ParsedSheet],
    sheet_table: QTableWidget,
    sheet_ranges: dict[int, SheetRange],
    role: str,
    direction: str = "row",
) -> pd.DataFrame | None:
    for row_index, sheet in enumerate(parsed_sheets):
        combo = sheet_table.cellWidget(row_index, 5)
        if not isinstance(combo, QComboBox):
            continue
        if combo.currentText().strip() != role:
            continue

        sheet_range = sheet_ranges.get(row_index, SheetRange())
        start_cell = sheet_range.start_cell or "A1"
        end_cell = sheet_range.end_cell.strip() or _last_cell(sheet.data)
        return table_from_range(sheet.data, start_cell, end_cell, direction)
    return None


def _last_cell(dataframe: pd.DataFrame) -> str:
    from openpyxl.utils.cell import get_column_letter

    if dataframe.empty:
        return "A1"
    return f"{get_column_letter(max(1, len(dataframe.columns)))}{max(1, len(dataframe))}"
