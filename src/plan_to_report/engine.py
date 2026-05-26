from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd

from .excel_io import WorkbookData, get_sheet, read_workbook
from .models import ConversionIssue, ConversionResult, ConversionTemplate, FieldSpec, OutputSpec


InputFiles = dict[str, str | Path]
InputTables = dict[str, pd.DataFrame]


class ConversionEngine:
    def __init__(self, template: ConversionTemplate) -> None:
        self.template = template
        self.issues: list[ConversionIssue] = []

    def run(self, input_files: InputFiles) -> ConversionResult:
        workbooks = self._load_workbooks(input_files)
        tables: dict[str, pd.DataFrame] = {}

        for output in self.template.outputs:
            tables[output.key] = self._build_output(output, workbooks)

        issue_table = self._issue_table()
        if not issue_table.empty:
            tables["异常清单"] = issue_table

        return ConversionResult(tables=tables, issues=self.issues)

    def run_tables(self, input_tables: InputTables) -> ConversionResult:
        tables: dict[str, pd.DataFrame] = {}

        for output in self.template.outputs:
            tables[output.key] = self._build_output_from_tables(output, input_tables)

        issue_table = self._issue_table()
        if not issue_table.empty:
            tables["异常清单"] = issue_table

        return ConversionResult(tables=tables, issues=self.issues)

    def _load_workbooks(self, input_files: InputFiles) -> dict[str, WorkbookData]:
        workbooks: dict[str, WorkbookData] = {}
        for spec in self.template.inputs:
            path = input_files.get(spec.key)
            if not path:
                if spec.required:
                    self._add_issue("error", "输入文件", None, spec.key, f"缺少输入文件：{spec.label}")
                continue
            workbooks[spec.key] = read_workbook(path)
        return workbooks

    def _build_output_from_tables(self, output: OutputSpec, input_tables: InputTables) -> pd.DataFrame:
        primary_table = input_tables.get(output.primary_input)
        if primary_table is None:
            self._add_issue("error", output.label, None, None, f"找不到主 sheet 类型：{output.primary_input}")
            return pd.DataFrame(columns=[field.name for field in output.fields])

        rows: list[dict[str, Any]] = []
        for index, row in primary_table.iterrows():
            row_number = int(index) + 2
            if self._should_skip_row(output, row, row_number):
                continue

            generated: dict[str, Any] = {}
            for field in output.fields:
                generated[field.name] = self._resolve_field_from_tables(field, row, input_tables, output, row_number)
                if field.required and _is_blank(generated[field.name]):
                    self._add_issue("warning", output.label, row_number, field.name, "必填字段为空。")
            rows.append(generated)

        return pd.DataFrame(rows, columns=[field.name for field in output.fields])

    def _build_output(self, output: OutputSpec, workbooks: dict[str, WorkbookData]) -> pd.DataFrame:
        workbook = workbooks.get(output.primary_input)
        if workbook is None:
            self._add_issue("error", output.label, None, None, f"找不到主输入：{output.primary_input}")
            return pd.DataFrame(columns=[field.name for field in output.fields])

        try:
            _, primary_table = get_sheet(workbook, output.primary_sheet)
        except KeyError as exc:
            self._add_issue("error", output.label, None, None, str(exc))
            return pd.DataFrame(columns=[field.name for field in output.fields])

        rows: list[dict[str, Any]] = []
        for index, row in primary_table.iterrows():
            row_number = int(index) + 2
            if self._should_skip_row(output, row, row_number):
                continue

            generated: dict[str, Any] = {}
            for field in output.fields:
                generated[field.name] = self._resolve_field(field, row, workbooks, output, row_number)
                if field.required and _is_blank(generated[field.name]):
                    self._add_issue("warning", output.label, row_number, field.name, "必填字段为空。")
            rows.append(generated)

        return pd.DataFrame(rows, columns=[field.name for field in output.fields])

    def _should_skip_row(self, output: OutputSpec, row: pd.Series, row_number: int) -> bool:
        for rule in output.filters:
            matched = self._match_condition(row, rule)
            if matched and rule.get("action", "skip") == "skip":
                reason = rule.get("reason", "命中过滤规则，跳过生成。")
                self._add_issue("info", output.label, row_number, None, reason)
                return True
        return False

    def _resolve_field(
        self,
        field: FieldSpec,
        row: pd.Series,
        workbooks: dict[str, WorkbookData],
        output: OutputSpec,
        row_number: int,
    ) -> Any:
        source = field.source
        source_type = source.get("type")

        if source_type == "direct":
            return row.get(source.get("column"))

        if source_type == "constant":
            return source.get("value")

        if source_type == "concat":
            parts = [self._resolve_inline(part, row, workbooks, output, row_number) for part in source.get("parts", [])]
            separator = source.get("separator", "")
            return separator.join("" if _is_blank(part) else str(part) for part in parts)

        if source_type == "regex_extract":
            text = row.get(source.get("column"))
            if _is_blank(text):
                return None
            pattern = source.get("pattern", "")
            try:
                match = re.search(pattern, str(text))
            except re.error as exc:
                self._add_issue("error", output.label, row_number, field.name, f"正则规则错误：{exc}")
                return source.get("default")
            if not match:
                self._add_issue("warning", output.label, row_number, field.name, "文本未匹配正则规则。")
                return source.get("default")
            group = int(source.get("group", 1))
            return match.group(group)

        if source_type == "lookup":
            return self._lookup_value(source, row, workbooks, output, row_number, field.name)

        if source_type == "manual":
            self._add_issue("warning", output.label, row_number, field.name, "字段需要人工补充。")
            return source.get("placeholder", "")

        self._add_issue("warning", output.label, row_number, field.name, f"未知字段来源类型：{source_type}")
        return None

    def _resolve_inline(
        self,
        source: dict[str, Any],
        row: pd.Series,
        workbooks: dict[str, WorkbookData],
        output: OutputSpec,
        row_number: int,
    ) -> Any:
        return self._resolve_field(FieldSpec(name="inline", source=source), row, workbooks, output, row_number)

    def _resolve_field_from_tables(
        self,
        field: FieldSpec,
        row: pd.Series,
        input_tables: InputTables,
        output: OutputSpec,
        row_number: int,
    ) -> Any:
        source = field.source
        source_type = source.get("type")

        if source_type == "lookup":
            return self._lookup_value_from_tables(source, row, input_tables, output, row_number, field.name)

        return self._resolve_field(field, row, {}, output, row_number)

    def _lookup_value(
        self,
        source: dict[str, Any],
        row: pd.Series,
        workbooks: dict[str, WorkbookData],
        output: OutputSpec,
        row_number: int,
        field_name: str,
    ) -> Any:
        lookup_input = source.get("input")
        workbook = workbooks.get(lookup_input)
        if workbook is None:
            self._add_issue("error", output.label, row_number, field_name, f"找不到查表输入：{lookup_input}")
            return None

        try:
            _, lookup_table = get_sheet(workbook, source.get("sheet"))
        except KeyError as exc:
            self._add_issue("error", output.label, row_number, field_name, str(exc))
            return None

        left_column = source.get("left_column")
        right_column = source.get("right_column")
        return_column = source.get("return_column")
        left_value = row.get(left_column)

        if _is_blank(left_value):
            self._add_issue("warning", output.label, row_number, field_name, f"查表字段为空：{left_column}")
            return None

        if right_column not in lookup_table.columns or return_column not in lookup_table.columns:
            self._add_issue("error", output.label, row_number, field_name, "查表字段或返回字段不存在。")
            return None

        matches = lookup_table[lookup_table[right_column].astype(str) == str(left_value)]
        if matches.empty:
            self._add_issue("warning", output.label, row_number, field_name, f"未匹配到：{left_value}")
            return source.get("default")
        if len(matches) > 1:
            self._add_issue("warning", output.label, row_number, field_name, f"匹配到多条，仅取第一条：{left_value}")
        return matches.iloc[0][return_column]

    def _lookup_value_from_tables(
        self,
        source: dict[str, Any],
        row: pd.Series,
        input_tables: InputTables,
        output: OutputSpec,
        row_number: int,
        field_name: str,
    ) -> Any:
        lookup_input = source.get("input")
        lookup_table = input_tables.get(lookup_input)
        if lookup_table is None:
            self._add_issue("error", output.label, row_number, field_name, f"找不到查表 sheet 类型：{lookup_input}")
            return None

        left_column = source.get("left_column")
        right_column = source.get("right_column")
        return_column = source.get("return_column")
        left_value = row.get(left_column)

        if _is_blank(left_value):
            self._add_issue("warning", output.label, row_number, field_name, f"查表字段为空：{left_column}")
            return None

        if right_column not in lookup_table.columns or return_column not in lookup_table.columns:
            self._add_issue("error", output.label, row_number, field_name, "查表字段或返回字段不存在。")
            return None

        matches = lookup_table[lookup_table[right_column].astype(str) == str(left_value)]
        if matches.empty:
            self._add_issue("warning", output.label, row_number, field_name, f"未匹配到：{left_value}")
            return source.get("default")
        if len(matches) > 1:
            self._add_issue("warning", output.label, row_number, field_name, f"匹配到多条，仅取第一条：{left_value}")
        return matches.iloc[0][return_column]

    def _match_condition(self, row: pd.Series, rule: dict[str, Any]) -> bool:
        column = rule.get("column")
        value = row.get(column)
        expected = rule.get("value")
        op = rule.get("operator", "equals")

        if op == "empty":
            return _is_blank(value)
        if op == "not_empty":
            return not _is_blank(value)
        if _is_blank(value):
            return False

        actual = str(value)
        if op == "equals":
            return actual == str(expected)
        if op == "not_equals":
            return actual != str(expected)
        if op == "contains":
            return str(expected) in actual
        if op == "not_contains":
            return str(expected) not in actual
        if op == "regex":
            return re.search(str(expected), actual) is not None
        return False

    def _issue_table(self) -> pd.DataFrame:
        return pd.DataFrame(
            [
                {
                    "级别": issue.level,
                    "输出表": issue.output,
                    "源数据行": issue.row_number,
                    "字段": issue.field,
                    "说明": issue.message,
                }
                for issue in self.issues
            ]
        )

    def _add_issue(self, level: str, output: str, row_number: int | None, field: str | None, message: str) -> None:
        self.issues.append(
            ConversionIssue(
                level=level,
                output=output,
                row_number=row_number,
                field=field,
                message=message,
            )
        )


def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, pd.Series):
        return value.dropna().empty
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False
