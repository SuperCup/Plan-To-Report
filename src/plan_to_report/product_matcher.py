from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

import pandas as pd

from .models import ConversionIssue
from .flavor_aliases import expand_flavor_token
from .plan_participation import ProductParticipation

LOGIC_BRAND_COLUMNS = ("品牌",)
LOGIC_SPEC_COLUMNS = ("规格",)
LOGIC_RULE_COLUMNS = ("选品逻辑", "选品逻辑说明", "逻辑")

PRODUCT_BARCODE_COLUMNS = ("商品条形码", "UPC条形码", "条形码")
PRODUCT_BRAND_COLUMNS = ("品牌名称",)
PRODUCT_FLAVOR_COLUMNS = ("口味名称", "口味")
PRODUCT_SPEC_COLUMNS = ("规格名称", "规格")


@dataclass
class SelectionConditions:
    brand_names: list[str] = field(default_factory=list)
    flavor_names: list[str] = field(default_factory=list)
    spec_names: list[str] = field(default_factory=list)
    exclusion_keywords: list[str] = field(default_factory=list)
    raw_logic: str = ""


@dataclass
class LogicRuleRow:
    plan_brand: str
    plan_spec: str
    conditions: SelectionConditions


def build_logic_index(logic_table: pd.DataFrame) -> list[LogicRuleRow]:
    if logic_table is None or logic_table.empty:
        return []

    brand_col = _find_column(logic_table, LOGIC_BRAND_COLUMNS)
    spec_col = _find_column(logic_table, LOGIC_SPEC_COLUMNS)
    rule_col = _find_column(logic_table, LOGIC_RULE_COLUMNS)
    if not brand_col or not spec_col or not rule_col:
        return []

    rules: list[LogicRuleRow] = []
    for _, row in logic_table.iterrows():
        brand = _normalize_text(row.get(brand_col))
        spec = _normalize_text(row.get(spec_col))
        logic_text = _normalize_text(row.get(rule_col))
        if not brand or not spec or not logic_text:
            continue
        conditions = parse_selection_logic(logic_text)
        rules.append(LogicRuleRow(plan_brand=brand, plan_spec=spec, conditions=conditions))
    return rules


def parse_selection_logic(text: str) -> SelectionConditions:
    exclusions = _parse_exclusion_keywords(text)
    cleaned = re.sub(r"\([^)]*不要[^)]*\)", "", text)
    cleaned = re.sub(r"（[^）]*不要[^）]*）", "", cleaned)
    conditions = SelectionConditions(raw_logic=text, exclusion_keywords=exclusions)
    segments = [segment.strip() for segment in cleaned.split("--") if segment.strip()]
    for segment in segments:
        if "：" in segment:
            name, values_text = segment.split("：", 1)
        elif ":" in segment:
            name, values_text = segment.split(":", 1)
        else:
            continue
        name = name.strip()
        values = [value.strip() for value in re.split(r"[、,，]", values_text) if value.strip()]
        if name == "品牌名称":
            conditions.brand_names = values
        elif name == "口味名称":
            conditions.flavor_names = values
        elif name == "规格名称":
            conditions.spec_names = values
    return conditions


def find_logic_rule(rules: list[LogicRuleRow], brand: str, spec: str) -> LogicRuleRow | None:
    brand_key = _normalize_text(brand)
    spec_key = _normalize_text(spec)
    for rule in rules:
        if _normalize_text(rule.plan_brand) == brand_key and _normalize_text(rule.plan_spec) == spec_key:
            return rule
    for rule in rules:
        if brand_key in _normalize_text(rule.plan_brand) and spec_key in _normalize_text(rule.plan_spec):
            return rule
    return None


def apply_flavor_hint(
    conditions: SelectionConditions,
    flavor_hint: str | None,
    aliases: dict[str, str] | None = None,
) -> SelectionConditions:
    if not flavor_hint:
        return conditions
    narrowed = SelectionConditions(
        brand_names=list(conditions.brand_names),
        flavor_names=list(conditions.flavor_names),
        spec_names=list(conditions.spec_names),
        exclusion_keywords=list(conditions.exclusion_keywords),
        raw_logic=conditions.raw_logic,
    )
    hint = flavor_hint.strip()
    if not hint:
        return narrowed

    alias_map = aliases or {}
    expanded = expand_flavor_token(hint, alias_map, narrowed.flavor_names)
    if narrowed.flavor_names:
        matched: list[str] = []
        for name in narrowed.flavor_names:
            for candidate in expanded:
                if _flavor_matches_hint(name, candidate):
                    matched.append(name)
        narrowed.flavor_names = matched or expanded
    else:
        narrowed.flavor_names = expanded
    return narrowed


def match_products(product_table: pd.DataFrame, conditions: SelectionConditions) -> pd.DataFrame:
    if product_table is None or product_table.empty:
        return pd.DataFrame()

    brand_col = _find_column(product_table, PRODUCT_BRAND_COLUMNS)
    flavor_col = _find_column(product_table, PRODUCT_FLAVOR_COLUMNS)
    spec_col = _find_column(product_table, PRODUCT_SPEC_COLUMNS)
    if not brand_col and not flavor_col and not spec_col:
        return pd.DataFrame()

    mask = pd.Series(True, index=product_table.index)
    if conditions.brand_names and brand_col:
        mask &= product_table[brand_col].apply(lambda value: _value_in_candidates(value, conditions.brand_names))
    if conditions.flavor_names and flavor_col:
        mask &= product_table[flavor_col].apply(lambda value: _value_in_candidates(value, conditions.flavor_names))
    if conditions.spec_names and spec_col:
        mask &= product_table[spec_col].apply(lambda value: _value_in_candidates(value, conditions.spec_names))

    matches = product_table[mask].copy()
    return apply_exclusion_filters(matches, conditions.exclusion_keywords)


def apply_exclusion_filters(product_table: pd.DataFrame, exclusion_keywords: list[str]) -> pd.DataFrame:
    if product_table.empty or not exclusion_keywords:
        return product_table

    filtered = product_table.copy()
    text_columns = [
        col
        for col in filtered.columns
        if _normalize_text(col) in {"标品名称", "规格名称", "口味名称", "品牌名称", "主品"}
    ]
    if not text_columns:
        text_columns = list(filtered.columns)

    for keyword in exclusion_keywords:
        if not keyword:
            continue
        row_mask = pd.Series(True, index=filtered.index)
        for column in text_columns:
            row_mask &= ~filtered[column].astype(str).str.contains(keyword, na=False, regex=False)
        filtered = filtered[row_mask]
    return filtered


def lookup_upc_rows(
    participation: ProductParticipation,
    logic_rules: list[LogicRuleRow],
    product_table: pd.DataFrame,
    issues: list[ConversionIssue],
    activity_label: str,
    aliases: dict[str, str] | None = None,
) -> list[dict[str, Any]]:
    rule = find_logic_rule(logic_rules, participation.brand, participation.spec)
    if rule is None:
        issues.append(
            ConversionIssue(
                "warning",
                "活动对应UPC表",
                participation.excel_row,
                participation.brand,
                f"未找到商品匹配逻辑：品牌={participation.brand}，规格={participation.spec}",
            )
        )
        return []

    conditions = apply_flavor_hint(rule.conditions, participation.flavor_hint, aliases)
    matches = match_products(product_table, conditions)
    if matches.empty:
        issues.append(
            ConversionIssue(
                "warning",
                "活动对应UPC表",
                participation.excel_row,
                activity_label,
                f"商品清单无匹配：{participation.brand}/{participation.spec}"
                f"{'' if not participation.flavor_hint else '/' + participation.flavor_hint}"
                f"；逻辑={conditions.raw_logic}",
            )
        )
        return []

    barcode_col = _find_column(matches, PRODUCT_BARCODE_COLUMNS)
    rows: list[dict[str, Any]] = []
    for _, product_row in matches.iterrows():
        row_data = {str(column): product_row.get(column) for column in matches.columns}
        if barcode_col:
            row_data["UPC条形码"] = product_row.get(barcode_col)
        rows.append(row_data)
    return rows


def _flavor_matches_hint(flavor_name: str, hint: str) -> bool:
    flavor_name = _normalize_text(flavor_name)
    hint = _normalize_text(hint)
    if not flavor_name or not hint:
        return False
    if hint in flavor_name or flavor_name in hint:
        return True
    return hint[:1] and hint[:1] in flavor_name


def _value_in_candidates(value: Any, candidates: list[str]) -> bool:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False
    actual = _normalize_text(value)
    if not actual:
        return False
    normalized_candidates = [_normalize_text(item) for item in candidates if _normalize_text(item)]
    if actual in normalized_candidates:
        return True
    return any(candidate in actual or actual in candidate for candidate in normalized_candidates if candidate)


def _find_column(table: pd.DataFrame, candidates: tuple[str, ...]) -> str | None:
    for column in table.columns:
        name = _normalize_text(column)
        for candidate in candidates:
            if candidate == name or candidate in name:
                return str(column)
    return None


def _parse_exclusion_keywords(text: str) -> list[str]:
    keywords: list[str] = []
    for match in re.findall(r"\(([^)]*不要[^)]*)\)", text):
        keywords.extend(_keywords_from_exclusion(match))
    for match in re.findall(r"（([^）]*不要[^）]*)）", text):
        keywords.extend(_keywords_from_exclusion(match))
    unique: list[str] = []
    for keyword in keywords:
        if keyword and keyword not in unique:
            unique.append(keyword)
    return unique


def _keywords_from_exclusion(fragment: str) -> list[str]:
    preset = ["整箱售卖", "整箱", "渠道自组装", "三桶", "箱装", "自组装"]
    found = [token for token in preset if token in fragment]
    extra = re.findall(r"[\u4e00-\u9fff]{2,}", fragment.replace("不要", ""))
    found.extend(extra)
    return found


def _normalize_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass
    return str(value).strip()
