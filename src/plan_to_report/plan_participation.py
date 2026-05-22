from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd

from .deepseek_assist import DeepseekAssistError, assist_parse_participation_cell
from .flavor_aliases import expand_flavor_token, load_flavor_aliases
from .plan_config import PlanSheetConfig

CHECKMARK_VALUES = {"√", "✓", "☑", "✔", "Y", "y", "1", "是", "yes", "YES"}


@dataclass
class ProductRow:
    excel_row: int
    brand: str
    spec: str
    row_flavor: str | None


@dataclass
class ProductParticipation:
    column_index: int
    excel_row: int
    brand: str
    spec: str
    flavor_hint: str | None = None


def collect_participations(
    raw: pd.DataFrame,
    config: PlanSheetConfig,
    column_index: int,
    project_root: Path | None = None,
) -> list[ProductParticipation]:
    end_row = _resolve_end_row(raw, config)
    product_rows = build_product_rows(raw, config, end_row)
    if not product_rows:
        return []

    aliases = load_flavor_aliases(project_root)
    known_flavors = [row.row_flavor for row in product_rows if row.row_flavor]
    participations: list[ProductParticipation] = []
    seen: set[tuple[int, str | None]] = set()

    for anchor in product_rows:
        cell_text = _cell_text(raw, anchor.excel_row, column_index)
        if not cell_text:
            continue

        targets: list[ProductRow] = []
        if _is_checkmark(cell_text):
            targets = _resolve_checkmark_targets(product_rows, anchor)
        elif _is_packaging_block(cell_text):
            targets = _resolve_packaging_block(
                product_rows,
                anchor,
                cell_text,
                aliases,
                known_flavors,
            )
        elif _is_multi_flavor_text(cell_text):
            tokens = _split_flavor_tokens(cell_text)
            targets = _resolve_flavor_tokens(
                product_rows,
                anchor,
                tokens,
                aliases,
                known_flavors,
            )
            if not targets or (len(targets) == 1 and targets[0].excel_row == anchor.excel_row):
                targets = _try_ai_flavor_resolve(
                    product_rows,
                    anchor,
                    cell_text,
                    known_flavors,
                    config,
                ) or targets
        else:
            targets = _resolve_flavor_tokens(
                product_rows,
                anchor,
                [cell_text],
                aliases,
                known_flavors,
            )

        for target in targets:
            key = (target.excel_row, target.row_flavor)
            if key in seen:
                continue
            seen.add(key)
            participations.append(
                ProductParticipation(
                    column_index=column_index,
                    excel_row=target.excel_row,
                    brand=target.brand,
                    spec=target.spec,
                    flavor_hint=target.row_flavor,
                )
            )

    return participations


def build_product_rows(raw: pd.DataFrame, config: PlanSheetConfig, end_row: int) -> list[ProductRow]:
    rows: list[ProductRow] = []
    for excel_row in range(config.product_row_start, end_row + 1):
        spec = _resolve_spec(raw, config, excel_row)
        if not spec or spec == "规格":
            continue
        brand = _resolve_brand(raw, config, excel_row)
        if not brand:
            continue
        c_text = _cell_text(raw, excel_row, config.brand_col)
        row_flavor = None
        if c_text and c_text != brand and c_text not in {"品牌", "子品牌"}:
            row_flavor = c_text
        rows.append(ProductRow(excel_row=excel_row, brand=brand, spec=spec, row_flavor=row_flavor))
    return rows


def _resolve_checkmark_targets(product_rows: list[ProductRow], anchor: ProductRow) -> list[ProductRow]:
    if anchor.row_flavor:
        return [anchor]

    block = [row for row in product_rows if row.brand == anchor.brand and row.spec == anchor.spec]
    flavor_rows = [row for row in block if row.row_flavor]
    if flavor_rows:
        return flavor_rows
    return [anchor]


def _is_packaging_block(text: str) -> bool:
    lines = [line.strip() for line in text.replace("\r\n", "\n").split("\n") if line.strip()]
    if len(lines) < 2:
        return False
    spec_like = sum(1 for line in lines[1:] if _looks_like_spec_line(line))
    return spec_like >= 1


def _looks_like_spec_line(line: str) -> bool:
    keywords = ("桶", "袋", "包", "盒", "碗", "杯")
    return any(keyword in line for keyword in keywords)


def _resolve_packaging_block(
    product_rows: list[ProductRow],
    anchor: ProductRow,
    cell_text: str,
    aliases: dict[str, str],
    known_flavors: list[str],
) -> list[ProductRow]:
    lines = [line.strip() for line in cell_text.replace("\r\n", "\n").split("\n") if line.strip()]
    flavor_line = lines[0]
    spec_lines = lines[1:]
    flavor_tokens = _split_flavor_tokens(flavor_line)

    matched: list[ProductRow] = []
    seen: set[int] = set()
    block = [row for row in product_rows if row.brand == anchor.brand]

    for spec_line in spec_lines:
        for row in block:
            if row.excel_row in seen:
                continue
            if spec_line not in row.spec and row.spec not in spec_line:
                continue
            if row.row_flavor:
                if any(_flavor_equals(row.row_flavor, token) for token in flavor_tokens):
                    matched.append(row)
                    seen.add(row.excel_row)
            else:
                expanded_targets = _resolve_flavor_tokens(
                    product_rows,
                    row,
                    flavor_tokens,
                    aliases,
                    known_flavors,
                )
                for item in expanded_targets:
                    if item.excel_row not in seen:
                        matched.append(item)
                        seen.add(item.excel_row)

    if matched:
        return matched
    return _resolve_flavor_tokens(product_rows, anchor, flavor_tokens, aliases, known_flavors)


def _try_ai_flavor_resolve(
    product_rows: list[ProductRow],
    anchor: ProductRow,
    cell_text: str,
    known_flavors: list[str],
    config: PlanSheetConfig,
) -> list[ProductRow]:
    if not config.ai_settings.enabled:
        return []
    try:
        flavors = assist_parse_participation_cell(
            cell_text,
            anchor.brand,
            anchor.spec,
            known_flavors,
            config.ai_settings,
        )
    except DeepseekAssistError:
        return []
    if not flavors:
        return []
    return _resolve_flavor_tokens(product_rows, anchor, flavors, {}, known_flavors)


def _resolve_flavor_tokens(
    product_rows: list[ProductRow],
    anchor: ProductRow,
    tokens: list[str],
    aliases: dict[str, str],
    known_flavors: list[str],
) -> list[ProductRow]:
    block = [row for row in product_rows if row.brand == anchor.brand and row.spec == anchor.spec]
    if not block:
        block = [row for row in product_rows if row.brand == anchor.brand]

    matched_rows: list[ProductRow] = []
    seen_rows: set[int] = set()

    for token in tokens:
        expanded = expand_flavor_token(token, aliases, known_flavors)
        for candidate in expanded:
            for row in block:
                if row.excel_row in seen_rows:
                    continue
                if row.row_flavor and _flavor_equals(row.row_flavor, candidate):
                    matched_rows.append(row)
                    seen_rows.add(row.excel_row)
                elif not row.row_flavor and candidate in row.brand:
                    matched_rows.append(row)
                    seen_rows.add(row.excel_row)

    if matched_rows:
        return matched_rows

    return [anchor]


def _flavor_equals(left: str, right: str) -> bool:
    left = left.strip()
    right = right.strip()
    return left == right or left in right or right in left


def _resolve_brand(raw: pd.DataFrame, config: PlanSheetConfig, row: int) -> str:
    """子品牌组：取 C 列中「非口味行」的末次赋值（该行 D 列无规格时视为组名）。"""
    brand = ""
    scan_start = max(1, config.product_row_start - 30)
    for current in range(scan_start, row + 1):
        value = _cell_text(raw, current, config.brand_col)
        if not value or value in {"品牌", "子品牌"}:
            continue
        spec_here = _cell_text(raw, current, config.spec_col)
        if spec_here and spec_here != "规格":
            continue
        brand = value
    return brand


def _resolve_spec(raw: pd.DataFrame, config: PlanSheetConfig, row: int) -> str:
    spec = ""
    scan_start = max(1, config.product_row_start - 30)
    for current in range(scan_start, row + 1):
        value = _cell_text(raw, current, config.spec_col)
        if value and value != "规格":
            spec = value
    return spec


def _split_flavor_tokens(text: str) -> list[str]:
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    parts = re.split(r"[/、\n]+", normalized)
    return [part.strip() for part in parts if part.strip()]


def _is_checkmark(text: str) -> bool:
    return text.strip() in CHECKMARK_VALUES


def _is_multi_flavor_text(text: str) -> bool:
    if "/" in text or "\n" in text or "、" in text:
        if not _is_checkmark(text) and len(text) <= 120:
            return True
    return False


def _resolve_end_row(raw: pd.DataFrame, config: PlanSheetConfig) -> int:
    from openpyxl.utils.cell import coordinate_to_tuple

    if config.end_cell and config.end_cell.strip():
        end_row, _end_col = coordinate_to_tuple(config.end_cell.upper())
        return end_row
    return len(raw)


def _cell_text(raw: pd.DataFrame, row: int, col: int) -> str:
    if row < 1 or col < 1:
        return ""
    row_index = row - 1
    col_index = col - 1
    if row_index >= len(raw) or col_index >= len(raw.columns):
        return ""
    value = raw.iat[row_index, col_index]
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()
