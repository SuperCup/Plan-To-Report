from __future__ import annotations

import pandas as pd
from openpyxl.utils.cell import get_column_letter
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..plan_config import SheetRange

_PREVIEW_MAX_COL_WIDTH = 420


class SheetExcelPreview(QWidget):
    """原始 Excel 网格预览，支持点击选取起始/结束单元格。"""

    range_changed = Signal(str, str)  # start_cell, end_cell

    _MAX_ROWS = 300
    _MAX_COLS = 80

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._dataframe: pd.DataFrame | None = None
        self._start_cell = "A1"
        self._end_cell = ""
        self._pick_mode = "start"

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        toolbar = QHBoxLayout()
        self.info_label = QLabel("请在上方 Sheet 列表中选择一行以预览内容。")
        self.start_label = QLabel("起始：A1")
        self.end_label = QLabel("结束：（自动）")
        self.set_start_btn = QPushButton("点击设起始")
        self.set_end_btn = QPushButton("点击设结束")
        self.set_start_btn.setCheckable(True)
        self.set_end_btn.setCheckable(True)
        self.set_start_btn.setChecked(True)
        self.set_start_btn.clicked.connect(lambda: self._set_pick_mode("start"))
        self.set_end_btn.clicked.connect(lambda: self._set_pick_mode("end"))
        self.clear_btn = QPushButton("重置为 A1")
        self.clear_btn.clicked.connect(self._reset_range)

        toolbar.addWidget(self.info_label, stretch=1)
        toolbar.addWidget(self.start_label)
        toolbar.addWidget(self.end_label)
        toolbar.addWidget(self.set_start_btn)
        toolbar.addWidget(self.set_end_btn)
        toolbar.addWidget(self.clear_btn)
        layout.addLayout(toolbar)

        self.grid = QTableWidget(0, 0)
        self.grid.setAlternatingRowColors(True)
        configure_table_full_text(self.grid)
        self.grid.cellClicked.connect(self._on_cell_clicked)
        layout.addWidget(self.grid, stretch=1)

        hint = QLabel("提示：先点「点击设起始/结束」，再单击单元格；结束留空则取 sheet 末行末列。")
        hint.setStyleSheet("color: #666;")
        layout.addWidget(hint)

    def load_sheet(self, file_name: str, sheet_name: str, dataframe: pd.DataFrame, sheet_range: SheetRange | None = None) -> None:
        self._dataframe = dataframe
        if sheet_range:
            self._start_cell = sheet_range.start_cell or "A1"
            self._end_cell = sheet_range.end_cell or ""
        else:
            self._start_cell = "A1"
            self._end_cell = ""

        self.info_label.setText(f"{file_name} / {sheet_name}")
        self._render_grid()
        self._update_range_labels()
        self.range_changed.emit(self._start_cell, self._end_cell)

    def clear(self) -> None:
        self._dataframe = None
        self.grid.setRowCount(0)
        self.grid.setColumnCount(0)
        self.info_label.setText("请在上方 Sheet 列表中选择一行以预览内容。")

    def current_range(self) -> SheetRange:
        return SheetRange(start_cell=self._start_cell, end_cell=self._end_cell)

    def set_range(self, start_cell: str, end_cell: str) -> None:
        self._start_cell = start_cell or "A1"
        self._end_cell = end_cell or ""
        self._render_grid()
        self._update_range_labels()
        self.range_changed.emit(self._start_cell, self._end_cell)

    def _set_pick_mode(self, mode: str) -> None:
        self._pick_mode = mode
        self.set_start_btn.setChecked(mode == "start")
        self.set_end_btn.setChecked(mode == "end")

    def _reset_range(self) -> None:
        self.set_range("A1", "")

    def _on_cell_clicked(self, row: int, column: int) -> None:
        if self._dataframe is None:
            return
        cell = f"{get_column_letter(column + 1)}{row + 1}"
        if self._pick_mode == "start":
            self._start_cell = cell
        else:
            self._end_cell = cell
        self._render_grid()
        self._update_range_labels()
        self.range_changed.emit(self._start_cell, self._end_cell)

    def _render_grid(self) -> None:
        if self._dataframe is None or self._dataframe.empty:
            self.grid.setRowCount(0)
            self.grid.setColumnCount(0)
            return

        row_count = min(len(self._dataframe), self._MAX_ROWS)
        col_count = min(len(self._dataframe.columns), self._MAX_COLS)
        self.grid.setRowCount(row_count)
        self.grid.setColumnCount(col_count)

        self.grid.setVerticalHeaderLabels([str(i + 1) for i in range(row_count)])
        self.grid.setHorizontalHeaderLabels([get_column_letter(i + 1) for i in range(col_count)])

        start_row, start_col = _parse_cell(self._start_cell)
        end_row, end_col = (_parse_cell(self._end_cell) if self._end_cell else (None, None))

        for row_index in range(row_count):
            for column_index in range(col_count):
                value = self._dataframe.iloc[row_index, column_index]
                text = "" if pd.isna(value) else str(value).strip()
                item = QTableWidgetItem(text)
                if text:
                    item.setToolTip(text)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                excel_row = row_index + 1
                excel_col = column_index + 1
                if excel_row == start_row and excel_col == start_col:
                    item.setBackground(QColor(198, 239, 206))
                    item.setToolTip("起始位置")
                elif end_row and excel_row == end_row and excel_col == end_col:
                    item.setBackground(QColor(255, 235, 156))
                    item.setToolTip("结束位置")
                self.grid.setItem(row_index, column_index, item)

        autosize_table_columns(self.grid, max_width=_PREVIEW_MAX_COL_WIDTH)

    def _update_range_labels(self) -> None:
        self.start_label.setText(f"起始：{self._start_cell}")
        end_text = self._end_cell if self._end_cell else "（自动）"
        self.end_label.setText(f"结束：{end_text}")


def configure_table_full_text(table: QTableWidget) -> None:
    table.setWordWrap(True)
    table.setTextElideMode(Qt.ElideNone)
    table.setHorizontalScrollMode(QTableWidget.ScrollPerPixel)
    table.setVerticalScrollMode(QTableWidget.ScrollPerPixel)
    header = table.horizontalHeader()
    header.setSectionResizeMode(QHeaderView.Interactive)
    header.setStretchLastSection(False)
    header.setMaximumSectionSize(2000)


def autosize_table_columns(table: QTableWidget, max_width: int = 420) -> None:
    table.resizeColumnsToContents()
    for column_index in range(table.columnCount()):
        width = min(table.columnWidth(column_index), max_width)
        table.setColumnWidth(column_index, max(width, 48))
    table.resizeRowsToContents()


def _parse_cell(cell: str) -> tuple[int, int]:
    from openpyxl.utils.cell import coordinate_to_tuple

    row, col = coordinate_to_tuple(cell.upper())
    return row, col
