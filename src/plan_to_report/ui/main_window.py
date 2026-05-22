from __future__ import annotations

from pathlib import Path

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)
from openpyxl.utils.cell import get_column_letter

from ..engine import ConversionEngine
from ..excel_io import ParsedSheet, parse_workbook_sheets, table_from_range, write_outputs
from ..models import ConversionIssue, ConversionResult, ConversionTemplate
from ..app_settings import load_app_settings, save_app_settings, settings_path
from ..plan_config import AiSettings, SheetRange
from ..plan_pipeline import PlanConversionInput, output_tables, run_plan_conversion
from ..sheet_inputs import load_role_table
from ..template_loader import list_templates, load_template
from .excel_preview import SheetExcelPreview, autosize_table_columns, configure_table_full_text
from .plan_wizard import PlanWizardWidget


class MainWindow(QMainWindow):
    def __init__(self, project_root: Path) -> None:
        super().__init__()
        self.project_root = project_root
        self.template_dir = project_root / "templates"
        self.output_dir = project_root / "output"
        self.current_template: ConversionTemplate | None = None
        self.current_result: ConversionResult | None = None
        self.parsed_sheets: list[ParsedSheet] = []
        self.sheet_ranges: dict[int, SheetRange] = {}
        self._preview_row_index: int | None = None
        self.app_settings = load_app_settings(project_root)

        self.setWindowTitle("规划转提报工具 - 批量Sheet版")
        self.resize(1400, 860)

        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)

        layout.addWidget(self._build_template_bar())
        layout.addWidget(self._build_upload_panel())

        splitter = QSplitter(Qt.Vertical)
        self.excel_preview = SheetExcelPreview()
        self.excel_preview.range_changed.connect(self._on_preview_range_changed)
        preview_group = QGroupBox("Excel 内容预览（点击单元格选取执行区域）")
        preview_layout = QVBoxLayout(preview_group)
        preview_layout.addWidget(self.excel_preview)
        splitter.addWidget(preview_group)

        self.result_tabs = QTabWidget()
        self.result_tabs.setVisible(False)
        splitter.addWidget(self.result_tabs)
        splitter.setStretchFactor(0, 4)
        splitter.setStretchFactor(1, 2)
        layout.addWidget(splitter, stretch=1)

        layout.addWidget(self._build_execution_panel())

        action_layout = QHBoxLayout()
        self.run_button = QPushButton("执行转换")
        self.run_button.clicked.connect(self.run_conversion)
        self.export_button = QPushButton("导出结果")
        self.export_button.clicked.connect(self.export_outputs)
        self.export_button.setEnabled(False)
        action_layout.addWidget(self.run_button)
        action_layout.addWidget(self.export_button)
        action_layout.addStretch()
        layout.addLayout(action_layout)

        self.status = QLabel("请选择转换模板。")
        layout.addWidget(self.status)

        self.reload_templates()

    def _build_template_bar(self) -> QWidget:
        panel = QWidget()
        layout = QHBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)

        self.template_combo = QComboBox()
        self.template_combo.currentIndexChanged.connect(self.load_selected_template)
        reload_button = QPushButton("刷新模板")
        reload_button.clicked.connect(self.reload_templates)

        layout.addWidget(QLabel("转换模板"))
        layout.addWidget(self.template_combo, stretch=1)
        layout.addWidget(reload_button)
        return panel

    def _build_upload_panel(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)

        button_layout = QHBoxLayout()
        upload_button = QPushButton("批量上传 Excel")
        upload_button.clicked.connect(self.choose_excels)
        preview_button = QPushButton("预览选中 Sheet")
        preview_button.clicked.connect(self._preview_selected_sheet)
        self.upload_status = QLabel("未上传文件。")
        button_layout.addWidget(upload_button)
        button_layout.addWidget(preview_button)
        button_layout.addWidget(self.upload_status, stretch=1)
        layout.addLayout(button_layout)

        self.sheet_table = QTableWidget(0, 8)
        self.sheet_table.setHorizontalHeaderLabels(
            ["文件", "Sheet", "行数", "列数", "合并单元格提示", "Sheet 类型", "起始", "结束"]
        )
        self.sheet_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.sheet_table.setSelectionMode(QTableWidget.SingleSelection)
        configure_table_full_text(self.sheet_table)
        self.sheet_table.itemSelectionChanged.connect(self._on_sheet_selection_changed)
        self.sheet_table.setMaximumHeight(200)
        layout.addWidget(self.sheet_table)
        return panel

    def _build_execution_panel(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)

        self.execution_stack = QWidget()
        stack_layout = QVBoxLayout(self.execution_stack)
        stack_layout.setContentsMargins(0, 0, 0, 0)

        self.legacy_execution_panel = QWidget()
        legacy_layout = QFormLayout(self.legacy_execution_panel)
        self.direction_combo = QComboBox()
        self.direction_combo.addItem("逐行执行：首行为字段名，后续每行一条记录", "row")
        self.direction_combo.addItem("逐列执行：首列为字段名，后续每列一条记录", "column")
        legacy_layout.addRow("执行顺序：", self.direction_combo)
        legacy_hint = QLabel("非规划表 Sheet：在预览区选取区域后，按逐行/逐列解析。")
        legacy_hint.setWordWrap(True)
        legacy_hint.setStyleSheet("color: #666;")
        legacy_layout.addRow(legacy_hint)

        self.plan_wizard = PlanWizardWidget(
            project_root=self.project_root,
            on_save_ai=self._persist_ai_settings,
        )
        self.plan_wizard.load_ai_settings(
            self.app_settings.deepseek,
            settings_path(self.project_root),
        )

        stack_layout.addWidget(self.legacy_execution_panel)
        stack_layout.addWidget(self.plan_wizard)
        self.plan_wizard.hide()

        layout.addWidget(self.execution_stack)
        return panel

    def reload_templates(self) -> None:
        self.template_combo.blockSignals(True)
        self.template_combo.clear()
        for path in list_templates(self.template_dir):
            self.template_combo.addItem(path.stem, str(path))
        self.template_combo.blockSignals(False)

        if self.template_combo.count() == 0:
            self.status.setText("未找到模板，请在 templates 目录添加 .json 模板。")
            return
        self.template_combo.setCurrentIndex(0)
        self.load_selected_template()

    def load_selected_template(self) -> None:
        path = self.template_combo.currentData()
        if not path:
            return

        try:
            self.current_template = load_template(path)
        except Exception as exc:
            QMessageBox.critical(self, "模板错误", str(exc))
            return

        self._render_sheet_table()
        self.result_tabs.clear()
        self.result_tabs.setVisible(False)
        self.current_result = None
        self.export_button.setEnabled(False)
        self.status.setText(f"已加载模板：{self.current_template.name} v{self.current_template.version}")

    def choose_excels(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "选择一个或多个 Excel 文件",
            str(self.project_root),
            "Excel Files (*.xlsx *.xlsm)",
        )
        if not paths:
            return

        parsed: list[ParsedSheet] = []
        errors: list[str] = []
        for path in paths:
            try:
                parsed.extend(parse_workbook_sheets(path))
            except Exception as exc:
                errors.append(f"{Path(path).name}: {exc}")

        self.parsed_sheets = parsed
        self.sheet_ranges.clear()
        self._preview_row_index = None
        self._render_sheet_table()
        self.excel_preview.clear()
        self.upload_status.setText(f"已上传 {len(paths)} 个文件，解析到 {len(parsed)} 个 sheet。请选择一行预览并选取执行区域。")
        if errors:
            QMessageBox.warning(self, "部分文件解析失败", "\n".join(errors))
        if parsed:
            self.sheet_table.selectRow(0)
            self._preview_sheet_at_row(0)

    def _render_sheet_table(self) -> None:
        if not hasattr(self, "sheet_table"):
            return

        self.sheet_table.setRowCount(len(self.parsed_sheets))
        role_options = self._role_options()
        for row_index, sheet in enumerate(self.parsed_sheets):
            sheet_range = self.sheet_ranges.get(row_index, SheetRange())
            values = [
                sheet.file_path.name,
                sheet.sheet_name,
                str(len(sheet.data)),
                str(len(sheet.data.columns)),
                "; ".join(sheet.merge_notes),
                "",
                sheet_range.start_cell,
                sheet_range.end_cell or "（自动）",
            ]
            for column_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                if value:
                    item.setToolTip(value)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.sheet_table.setItem(row_index, column_index, item)

            combo = QComboBox()
            combo.addItems(role_options)
            combo.currentTextChanged.connect(lambda _text, idx=row_index: self._on_role_changed(idx))
            self.sheet_table.setCellWidget(row_index, 5, combo)

        autosize_table_columns(self.sheet_table, max_width=360)

    def _on_role_changed(self, row_index: int) -> None:
        combo = self.sheet_table.cellWidget(row_index, 5)
        if isinstance(combo, QComboBox) and combo.currentText().strip() == "规划表":
            self.sheet_table.selectRow(row_index)
            self._preview_sheet_at_row(row_index)
            self.status.setText("已指定为规划表：请按下方四步向导配置后执行转换。")
        self._update_execution_mode(row_index)

    def _on_sheet_selection_changed(self) -> None:
        rows = self.sheet_table.selectionModel().selectedRows()
        if not rows:
            return
        self._preview_sheet_at_row(rows[0].row())

    def _preview_selected_sheet(self) -> None:
        rows = self.sheet_table.selectionModel().selectedRows()
        if not rows:
            QMessageBox.information(self, "提示", "请先在 Sheet 列表中选择一行。")
            return
        self._preview_sheet_at_row(rows[0].row())

    def _preview_sheet_at_row(self, row_index: int) -> None:
        if row_index < 0 or row_index >= len(self.parsed_sheets):
            return

        self._preview_row_index = row_index
        sheet = self.parsed_sheets[row_index]
        sheet_range = self.sheet_ranges.get(row_index, SheetRange())
        self.excel_preview.load_sheet(
            sheet.file_path.name,
            sheet.sheet_name,
            sheet.data,
            sheet_range,
        )
        self._update_execution_mode(row_index)
        if self._is_plan_sheet_row(row_index):
            self.plan_wizard.refresh_from_sheet(
                sheet.data,
                sheet_range.start_cell,
                sheet_range.end_cell,
            )

    def _on_preview_range_changed(self, start_cell: str, end_cell: str) -> None:
        if self._preview_row_index is None:
            return

        self.sheet_ranges[self._preview_row_index] = SheetRange(start_cell=start_cell, end_cell=end_cell)
        start_item = self.sheet_table.item(self._preview_row_index, 6)
        end_item = self.sheet_table.item(self._preview_row_index, 7)
        if start_item:
            start_item.setText(start_cell)
        if end_item:
            end_item.setText(end_cell or "（自动）")

        if self._is_plan_sheet_row(self._preview_row_index):
            sheet = self.parsed_sheets[self._preview_row_index]
            self.plan_wizard.refresh_from_sheet(sheet.data, start_cell, end_cell)

    def _role_options(self) -> list[str]:
        roles = [
            "",
            "规划表",
            "商品清单",
            "商品匹配逻辑",
            "预算分配表",
            "预算匹配逻辑",
            "门店清单",
            "门店匹配逻辑",
        ]
        if self.current_template:
            roles.extend(spec.key for spec in self.current_template.inputs)
        return list(dict.fromkeys(roles))

    def run_conversion(self) -> None:
        if not self.current_template:
            QMessageBox.warning(self, "提示", "请先选择模板。")
            return
        if not self.parsed_sheets:
            QMessageBox.warning(self, "提示", "请先批量上传 Excel。")
            return

        plan_index = self._find_plan_sheet_index()
        try:
            if plan_index is not None:
                self.current_result = self._run_plan_conversion(plan_index)
            else:
                input_tables = self._collect_typed_tables()
                engine = ConversionEngine(self.current_template)
                self.current_result = engine.run_tables(input_tables)
        except Exception as exc:
            QMessageBox.critical(self, "执行失败", str(exc))
            return

        self._render_result_preview(self.current_result)
        self.export_button.setEnabled(True)
        issue_count = len(self.current_result.issues)
        self.status.setText(f"转换完成，生成 {len(self.current_result.tables)} 个结果表，发现 {issue_count} 条提示/异常。")

    def _persist_ai_settings(self, ai_settings: AiSettings) -> None:
        self.app_settings.deepseek = ai_settings
        path = save_app_settings(self.project_root, self.app_settings)
        self.plan_wizard.load_ai_settings(ai_settings, path)
        self.status.setText(f"已保存 DeepSeek 配置：{path}")

    def _run_plan_conversion(self, plan_index: int) -> ConversionResult:
        ready, message = self.plan_wizard.is_ready()
        if not ready:
            raise ValueError(message)

        self._persist_ai_settings(self.plan_wizard.get_ai_settings())

        sheet = self.parsed_sheets[plan_index]
        config = self.plan_wizard.get_config()
        logic_table = load_role_table(
            self.parsed_sheets,
            self.sheet_table,
            self.sheet_ranges,
            "商品匹配逻辑",
        )
        product_table = load_role_table(
            self.parsed_sheets,
            self.sheet_table,
            self.sheet_ranges,
            "商品清单",
        )
        pipeline_output = run_plan_conversion(
            PlanConversionInput(
                plan_raw=sheet.data,
                config=config,
                logic_table=logic_table,
                product_table=product_table,
                project_root=self.project_root,
            )
        )
        return ConversionResult(
            tables=output_tables(pipeline_output),
            issues=pipeline_output.issues,
        )

    def _find_plan_sheet_index(self) -> int | None:
        if self._preview_row_index is not None and self._is_plan_sheet_row(self._preview_row_index):
            return self._preview_row_index
        for row_index, _sheet in enumerate(self.parsed_sheets):
            if self._is_plan_sheet_row(row_index):
                return row_index
        return None

    def _is_plan_sheet_row(self, row_index: int) -> bool:
        combo = self.sheet_table.cellWidget(row_index, 5)
        return isinstance(combo, QComboBox) and combo.currentText().strip() == "规划表"

    def _update_execution_mode(self, row_index: int) -> None:
        is_plan = self._is_plan_sheet_row(row_index)
        self.plan_wizard.setVisible(is_plan)
        self.legacy_execution_panel.setVisible(not is_plan)

    def _collect_typed_tables(self) -> dict[str, pd.DataFrame]:
        input_tables: dict[str, pd.DataFrame] = {}
        direction = self.direction_combo.currentData()

        for row_index, sheet in enumerate(self.parsed_sheets):
            combo = self.sheet_table.cellWidget(row_index, 5)
            if not isinstance(combo, QComboBox):
                continue
            role = combo.currentText().strip()
            if not role or role == "规划表":
                continue

            sheet_range = self.sheet_ranges.get(row_index, SheetRange())
            start_cell = sheet_range.start_cell or "A1"
            end_cell = sheet_range.end_cell.strip() or _last_cell(sheet.data)
            input_tables[role] = table_from_range(sheet.data, start_cell, end_cell, direction)

        if not input_tables:
            raise ValueError("请至少为一个非规划表 sheet 指定类型。")
        return input_tables

    def _render_result_preview(self, result: ConversionResult) -> None:
        self.result_tabs.clear()
        for name, table in result.tables.items():
            self.result_tabs.addTab(_dataframe_to_table(table), name)
        self.result_tabs.setVisible(True)

    def export_outputs(self) -> None:
        if not self.current_template or not self.current_result:
            return

        target_dir = QFileDialog.getExistingDirectory(self, "选择导出目录", str(self.output_dir))
        if not target_dir:
            return

        file_names = {output.key: output.file_name for output in self.current_template.outputs}
        file_names["异常清单"] = "异常清单.xlsx"
        try:
            written = write_outputs(target_dir, self.current_result.tables, file_names)
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", str(exc))
            return

        QMessageBox.information(self, "导出完成", "\n".join(str(path) for path in written))


def _dataframe_to_table(dataframe: pd.DataFrame) -> QTableWidget:
    max_rows = min(len(dataframe), 500)
    table = QTableWidget(max_rows, len(dataframe.columns))
    table.setHorizontalHeaderLabels([str(column) for column in dataframe.columns])
    table.setAlternatingRowColors(True)
    configure_table_full_text(table)

    for row_index in range(max_rows):
        for column_index, column in enumerate(dataframe.columns):
            value = dataframe.iloc[row_index][column]
            text = "" if pd.isna(value) else str(value)
            item = QTableWidgetItem(text)
            if text:
                item.setToolTip(text)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            table.setItem(row_index, column_index, item)

    autosize_table_columns(table)
    return table


def _last_cell(dataframe: pd.DataFrame) -> str:
    if dataframe.empty:
        return "A1"
    return f"{get_column_letter(max(1, len(dataframe.columns)))}{max(1, len(dataframe))}"


def run_app(project_root: Path) -> int:
    app = QApplication([])
    window = MainWindow(project_root)
    window.show()
    return app.exec()
