from __future__ import annotations

import pandas as pd

from plan_to_report.flavor_aliases import expand_flavor_token, load_flavor_aliases
from plan_to_report.plan_config import PlanSheetConfig
from plan_to_report.plan_participation import collect_participations
from plan_to_report.product_matcher import apply_exclusion_filters, parse_selection_logic


def test_expand_flavor_token_uses_alias_file():
    aliases = load_flavor_aliases()
    assert expand_flavor_token("藤", aliases, []) == ["藤椒牛肉"]


def test_collect_participations_multi_flavor_matches_rows():
    grid = [[None] * 7 for _ in range(20)]
    grid[9][2] = "地方口味"
    grid[10][2] = "藤椒牛肉"
    grid[10][3] = "开心桶"
    grid[11][2] = "麻辣牛肉"
    grid[11][3] = "开心桶"
    grid[10][5] = "藤/麻"
    raw = pd.DataFrame(grid, dtype=object)

    config = PlanSheetConfig(
        end_cell="G20",
        product_row_start=11,
        brand_col=3,
        spec_col=4,
        first_activity_col=6,
    )
    parts = collect_participations(raw, config, 6)
    flavors = {part.flavor_hint for part in parts}
    assert "藤椒牛肉" in flavors
    assert "麻辣牛肉" in flavors
    assert len(parts) == 2


def test_exclusion_filter_removes_full_case_products():
    logic = "品牌名称：红烧品牌--口味名称：红烧牛肉--规格名称：开心桶(整箱售卖：渠道自组装不要)"
    conditions = parse_selection_logic(logic)
    assert "整箱" in conditions.exclusion_keywords or "渠道自组装" in conditions.exclusion_keywords

    products = pd.DataFrame(
        [
            {"标品名称": "红烧牛肉开心桶整箱", "口味名称": "红烧牛肉", "规格名称": "开心桶"},
            {"标品名称": "红烧牛肉开心桶单桶", "口味名称": "红烧牛肉", "规格名称": "开心桶"},
        ]
    )
    filtered = apply_exclusion_filters(products, conditions.exclusion_keywords)
    assert len(filtered) == 1
    assert "整箱" not in str(filtered.iloc[0]["标品名称"])
