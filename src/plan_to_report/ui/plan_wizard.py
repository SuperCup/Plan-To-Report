from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QSpinBox,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..plan_config import (
    LOGIC_FIELD_OPTIONS,
    ActivityColumnOption,
    AiSettings,
    FieldRowMapping,
    PlanSheetConfig,
)
from ..plan_parser import infer_config_from_range
from .excel_preview import autosize_table_columns, configure_table_full_text


class PlanWizardWidget(QWidget):
    """规划表四步配置：框选区域 → 字段映射 → 勾选活动列 → 确认商品区。"""

    def __init__(
        self,
        parent: QWidget | None = None,
        project_root: Path | None = None,
        on_save_ai: Callable[[AiSettings], None] | None = None,
    ) -> None:
        super().__init__(parent)
        self._project_root = project_root
        self._on_save_ai = on_save_ai
        self._config = PlanSheetConfig()
        self._raw: pd.DataFrame | None = None
        self._start_cell = "D2"
        self._end_cell = ""

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        header = QHBoxLayout()
        self.step_label = QLabel("步骤 1/4：在上方预览区框选起始与结束单元格")
        self.step_label.setStyleSheet("font-weight: bold;")
        header.addWidget(self.step_label, stretch=1)
        self.prev_btn = QPushButton("上一步")
        self.next_btn = QPushButton("下一步")
        self.prev_btn.clicked.connect(self._go_prev)
        self.next_btn.clicked.connect(self._go_next)
        header.addWidget(self.prev_btn)
        header.addWidget(self.next_btn)
        layout.addLayout(header)

        self.stack = QStackedWidget()
        layout.addWidget(self.stack)

        self.stack.addWidget(self._build_step_range())
        self.stack.addWidget(self._build_step_fields())
        self.stack.addWidget(self._build_step_columns())
        self.stack.addWidget(self._build_step_product())

        self._step_index = 0
        self._update_step_ui()

    def _build_step_range(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        self.range_hint = QLabel(
            "① 在 Excel 预览中点击设置「起始」（建议 D2）与「结束」（含商品矩阵末行末列）。\n"
            "② 设置完成后点击「下一步」进入字段映射。"
        )
        self.range_hint.setWordWrap(True)
        layout.addWidget(self.range_hint)
        self.range_summary = QLabel("当前选区：未设置")
        layout.addWidget(self.range_summary)
        layout.addStretch()
        return page

    def _build_step_fields(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.addWidget(QLabel("② 将 D 列各行映射到活动参数字段（仅映射行会参与解析）："))
        self.field_table = QTableWidget(0, 3)
        self.field_table.setHorizontalHeaderLabels(["Excel 行", "D 列文本", "映射字段"])
        configure_table_full_text(self.field_table)
        layout.addWidget(self.field_table)
        refresh_btn = QPushButton("根据当前选区重新识别字段行")
        refresh_btn.clicked.connect(self._refresh_field_table)
        layout.addWidget(refresh_btn)
        return page

    def _build_step_columns(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        bar = QHBoxLayout()
        bar.addWidget(QLabel("③ 勾选需要生成汇总表的活动列："))
        select_all = QPushButton("全选")
        select_none = QPushButton("全不选")
        select_all.clicked.connect(lambda: self._set_all_columns(True))
        select_none.clicked.connect(lambda: self._set_all_columns(False))
        bar.addWidget(select_all)
        bar.addWidget(select_none)
        bar.addStretch()
        layout.addLayout(bar)

        self.column_table = QTableWidget(0, 3)
        self.column_table.setHorizontalHeaderLabels(["生成", "列", "摘要"])
        configure_table_full_text(self.column_table)
        layout.addWidget(self.column_table)
        return page

    def _build_step_product(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        group = QGroupBox("④ 商品区参数（矩阵勾选解析）")
        form = QFormLayout(group)
        self.product_row_spin = QSpinBox()
        self.product_row_spin.setRange(1, 5000)
        self.product_row_spin.setValue(11)
        self.brand_col_spin = QSpinBox()
        self.brand_col_spin.setRange(1, 100)
        self.brand_col_spin.setValue(3)
        self.spec_col_spin = QSpinBox()
        self.spec_col_spin.setRange(1, 100)
        self.spec_col_spin.setValue(4)
        form.addRow("商品区起始行：", self.product_row_spin)
        form.addRow("子品牌列（C=3）：", self.brand_col_spin)
        form.addRow("规格列（D=4）：", self.spec_col_spin)
        layout.addWidget(group)

        ai_group = QGroupBox("DeepSeek 配置（可保存到 config/app_settings.json）")
        ai_form = QFormLayout(ai_group)
        self.ai_enable_check = QCheckBox("启用 DeepSeek 辅助（机制拆分 / 歧义口味）")
        self.ai_key_edit = QLineEdit()
        self.ai_key_edit.setPlaceholderText("在此填写 API Key，或保存后下次自动加载")
        self.ai_key_edit.setEchoMode(QLineEdit.Password)
        self.ai_model_edit = QLineEdit("deepseek-chat")
        self.ai_base_url_edit = QLineEdit("https://api.deepseek.com")
        self.ai_config_hint = QLabel("")
        self.ai_config_hint.setStyleSheet("color: #666;")
        self.ai_config_hint.setWordWrap(True)
        self.save_ai_btn = QPushButton("保存 DeepSeek 配置")
        self.save_ai_btn.clicked.connect(self._save_ai_settings_clicked)
        self.show_key_btn = QPushButton("显示 Key")
        self.show_key_btn.setCheckable(True)
        self.show_key_btn.toggled.connect(self._toggle_key_visibility)
        key_row = QHBoxLayout()
        key_row.addWidget(self.ai_key_edit, stretch=1)
        key_row.addWidget(self.show_key_btn)
        ai_form.addRow(self.ai_enable_check)
        ai_form.addRow("API Key：", key_row)
        ai_form.addRow("模型：", self.ai_model_edit)
        ai_form.addRow("接口地址：", self.ai_base_url_edit)
        ai_form.addRow(self.save_ai_btn)
        ai_form.addRow(self.ai_config_hint)
        layout.addWidget(ai_group)

        note = QLabel(
            "请同时上传「商品匹配逻辑」「商品清单」并指定 Sheet 类型。\n"
            "口味缩写可在 templates/口味别名.json 中维护。\n"
            "配置完成后点击「执行转换」生成活动汇总表与活动对应 UPC 表。"
        )
        note.setWordWrap(True)
        note.setStyleSheet("color: #666;")
        layout.addWidget(note)
        layout.addStretch()
        return page

    def refresh_from_sheet(self, raw: pd.DataFrame, start_cell: str, end_cell: str) -> None:
        self._raw = raw
        self._start_cell = start_cell or "D2"
        self._end_cell = end_cell or ""
        if raw is None or raw.empty:
            self.range_summary.setText("当前选区：数据为空")
            return

        self._config = infer_config_from_range(raw, self._start_cell, self._end_cell)
        end_text = self._config.end_cell or "（自动）"
        self.range_summary.setText(
            f"当前选区：{self._start_cell} — {end_text}；"
            f"活动区 {self._config.activity_row_start}-{self._config.activity_row_end} 行；"
            f"活动列自 {self._config.first_activity_col} 列起"
        )
        self.product_row_spin.setValue(self._config.product_row_start)
        self.brand_col_spin.setValue(self._config.brand_col)
        self.spec_col_spin.setValue(self._config.spec_col)
        if self._step_index >= 1:
            self._refresh_field_table()
        if self._step_index >= 2:
            self._refresh_column_table()

    def get_config(self) -> PlanSheetConfig:
        self._sync_config_from_ui()
        return self._config

    def is_ready(self) -> tuple[bool, str]:
        config = self.get_config()
        if not config.selected_column_indices():
            return False, "请至少勾选一个活动列（步骤 3）。"
        if not [m for m in config.field_mappings if m.field_key]:
            return False, "请至少映射一个活动字段行（步骤 2）。"
        return True, ""

    def _sync_config_from_ui(self) -> None:
        self._config.start_cell = self._start_cell
        self._config.end_cell = self._end_cell
        self._config.product_row_start = self.product_row_spin.value()
        self._config.brand_col = self.brand_col_spin.value()
        self._config.spec_col = self.spec_col_spin.value()

        mappings: list[FieldRowMapping] = []
        for row in range(self.field_table.rowCount()):
            row_item = self.field_table.item(row, 0)
            label_item = self.field_table.item(row, 1)
            combo = self.field_table.cellWidget(row, 2)
            if not row_item or not label_item or not isinstance(combo, QComboBox):
                continue
            mappings.append(
                FieldRowMapping(
                    excel_row=int(row_item.text()),
                    label_text=label_item.text(),
                    field_key=combo.currentText().strip(),
                )
            )
        if mappings:
            self._config.field_mappings = mappings

        columns: list[ActivityColumnOption] = []
        for row in range(self.column_table.rowCount()):
            check = self.column_table.cellWidget(row, 0)
            col_item = self.column_table.item(row, 1)
            preview_item = self.column_table.item(row, 2)
            if not isinstance(check, QCheckBox) or not col_item or not preview_item:
                continue
            letter = col_item.text()
            col_index = self._column_letter_to_index(letter)
            columns.append(
                ActivityColumnOption(
                    excel_col=col_index,
                    column_letter=letter,
                    preview=preview_item.text(),
                    selected=check.isChecked(),
                )
            )
        if columns:
            self._config.activity_columns = columns

        self._config.ai_settings = self.get_ai_settings()

    def load_ai_settings(self, settings: AiSettings, config_path: Path | None = None) -> None:
        self.ai_enable_check.setChecked(settings.enabled)
        self.ai_key_edit.setText(settings.api_key)
        self.ai_model_edit.setText(settings.model or "deepseek-chat")
        self.ai_base_url_edit.setText(settings.base_url or "https://api.deepseek.com")
        self._config.ai_settings = settings
        if config_path and config_path.exists():
            self.ai_config_hint.setText(f"配置文件：{config_path}")
        else:
            self.ai_config_hint.setText(
                "未找到配置文件，可填写后点「保存」；也可复制 config/app_settings.example.json"
            )

    def get_ai_settings(self) -> AiSettings:
        return AiSettings(
            enabled=self.ai_enable_check.isChecked(),
            api_key=self.ai_key_edit.text().strip(),
            model=self.ai_model_edit.text().strip() or "deepseek-chat",
            base_url=self.ai_base_url_edit.text().strip() or "https://api.deepseek.com",
        )

    def _save_ai_settings_clicked(self) -> None:
        if self._on_save_ai:
            self._on_save_ai(self.get_ai_settings())

    def _toggle_key_visibility(self, visible: bool) -> None:
        self.ai_key_edit.setEchoMode(QLineEdit.Normal if visible else QLineEdit.Password)
        self.show_key_btn.setText("隐藏 Key" if visible else "显示 Key")

    def _refresh_field_table(self) -> None:
        if self._raw is None:
            return
        mappings = self._config.field_mappings
        self.field_table.setRowCount(len(mappings))
        for row_index, mapping in enumerate(mappings):
            row_item = QTableWidgetItem(str(mapping.excel_row))
            row_item.setFlags(row_item.flags() & ~Qt.ItemIsEditable)
            label_item = QTableWidgetItem(mapping.label_text)
            label_item.setFlags(label_item.flags() & ~Qt.ItemIsEditable)
            self.field_table.setItem(row_index, 0, row_item)
            self.field_table.setItem(row_index, 1, label_item)

            combo = QComboBox()
            combo.addItems(LOGIC_FIELD_OPTIONS)
            if mapping.field_key in LOGIC_FIELD_OPTIONS:
                combo.setCurrentText(mapping.field_key)
            self.field_table.setCellWidget(row_index, 2, combo)

        autosize_table_columns(self.field_table, max_width=360)

    def _refresh_column_table(self) -> None:
        if self._raw is None:
            return
        self._config = infer_config_from_range(self._raw, self._start_cell, self._end_cell)
        if self._config.field_mappings:
            self._sync_field_mappings_from_table()
            self._config.field_mappings = [
                FieldRowMapping(m.excel_row, m.label_text, m.field_key)
                for m in self._read_field_mappings_from_table()
            ] or self._config.field_mappings
        columns = self._config.activity_columns
        self.column_table.setRowCount(len(columns))
        for row_index, column in enumerate(columns):
            check = QCheckBox()
            check.setChecked(column.selected)
            self.column_table.setCellWidget(row_index, 0, check)
            col_item = QTableWidgetItem(column.column_letter)
            col_item.setFlags(col_item.flags() & ~Qt.ItemIsEditable)
            preview_item = QTableWidgetItem(column.preview)
            preview_item.setFlags(preview_item.flags() & ~Qt.ItemIsEditable)
            self.column_table.setItem(row_index, 1, col_item)
            self.column_table.setItem(row_index, 2, preview_item)
        autosize_table_columns(self.column_table, max_width=480)

    def _read_field_mappings_from_table(self) -> list[FieldRowMapping]:
        mappings: list[FieldRowMapping] = []
        for row in range(self.field_table.rowCount()):
            row_item = self.field_table.item(row, 0)
            label_item = self.field_table.item(row, 1)
            combo = self.field_table.cellWidget(row, 2)
            if not row_item or not label_item or not isinstance(combo, QComboBox):
                continue
            mappings.append(
                FieldRowMapping(
                    excel_row=int(row_item.text()),
                    label_text=label_item.text(),
                    field_key=combo.currentText().strip(),
                )
            )
        return mappings

    def _sync_field_mappings_from_table(self) -> None:
        mappings = self._read_field_mappings_from_table()
        if mappings:
            self._config.field_mappings = mappings

    def _set_all_columns(self, selected: bool) -> None:
        for row in range(self.column_table.rowCount()):
            check = self.column_table.cellWidget(row, 0)
            if isinstance(check, QCheckBox):
                check.setChecked(selected)

    def _go_prev(self) -> None:
        if self._step_index > 0:
            self._step_index -= 1
            self._update_step_ui()

    def _go_next(self) -> None:
        if self._step_index == 0 and (self._raw is None or self._raw.empty):
            self.step_label.setText("请先在预览区加载规划表并设置选区。")
            return
        if self._step_index == 1:
            self._sync_field_mappings_from_table()
            self._refresh_column_table()
        if self._step_index < 3:
            self._step_index += 1
            self._update_step_ui()

    def _update_step_ui(self) -> None:
        self.stack.setCurrentIndex(self._step_index)
        self.prev_btn.setEnabled(self._step_index > 0)
        self.next_btn.setEnabled(self._step_index < 3)
        titles = [
            "步骤 1/4：在上方预览区框选起始与结束单元格",
            "步骤 2/4：映射活动参数字段行（D 列）",
            "步骤 3/4：勾选需要生成的活动列",
            "步骤 4/4：确认商品区参数",
        ]
        self.step_label.setText(titles[self._step_index])
        if self._step_index == 1 and self._raw is not None:
            self._refresh_field_table()
        if self._step_index == 2 and self._raw is not None:
            self._sync_field_mappings_from_table()
            self._refresh_column_table()

    @staticmethod
    def _column_letter_to_index(letter: str) -> int:
        from openpyxl.utils.cell import column_index_from_string

        return column_index_from_string(letter)
