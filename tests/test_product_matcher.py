from __future__ import annotations

import pandas as pd

from plan_to_report.product_matcher import (
    build_logic_index,
    match_products,
    parse_selection_logic,
)


def test_parse_selection_logic():
    conditions = parse_selection_logic(
        "品牌名称：红烧品牌--口味名称：红烧牛肉--规格名称：开心桶、965桶"
    )
    assert conditions.brand_names == ["红烧品牌"]
    assert conditions.flavor_names == ["红烧牛肉"]
    assert conditions.spec_names == ["开心桶", "965桶"]


def test_match_products_returns_all_rows():
    logic = pd.DataFrame(
        [
            {
                "品牌": "红烧牛肉",
                "规格": "开心桶",
                "选品逻辑": "品牌名称：红烧品牌--口味名称：红烧牛肉--规格名称：开心桶",
            }
        ]
    )
    products = pd.DataFrame(
        [
            {
                "商品条形码": "6900000000001",
                "品牌名称": "红烧品牌",
                "口味名称": "红烧牛肉",
                "规格名称": "开心桶",
            },
            {
                "商品条形码": "6900000000002",
                "品牌名称": "红烧品牌",
                "口味名称": "红烧牛肉",
                "规格名称": "开心桶",
            },
        ]
    )
    rules = build_logic_index(logic)
    matches = match_products(products, rules[0].conditions)
    assert len(matches) == 2
