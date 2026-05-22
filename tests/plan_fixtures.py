from __future__ import annotations

import pandas as pd

from plan_to_report.plan_config import (
    ActivityColumnOption,
    AiSettings,
    FieldRowMapping,
    PlanSheetConfig,
)


def build_plan_grid() -> pd.DataFrame:
    grid = [[None] * 8 for _ in range(20)]
    grid[1][3] = "活动类型"
    grid[1][4] = "单品牌机制"
    grid[3][3] = "活动机制"
    grid[3][4] = "①满16-2 ②满31-5"
    grid[4][3] = "机制力度"
    grid[4][4] = "88%/84%"
    grid[5][3] = "活动时间"
    grid[5][4] = "3/1-3/31"
    grid[6][3] = "活动区域/渠道"
    grid[6][4] = "全渠道"
    grid[7][3] = "优惠券类型"
    grid[7][4] = "品牌券"
    grid[9][3] = "规格"
    grid[9][2] = "红烧牛肉"
    grid[10][3] = "开心桶"
    grid[10][4] = "√"
    return pd.DataFrame(grid, dtype=object)


def build_logic_table() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "品牌": "红烧牛肉",
                "规格": "开心桶",
                "选品逻辑": "品牌名称：红烧品牌--口味名称：红烧牛肉--规格名称：开心桶、965桶",
            }
        ]
    )


def build_product_table() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "商品条形码": "6900000000001",
                "标品名称": "红烧牛肉开心桶",
                "品牌名称": "红烧品牌",
                "口味名称": "红烧牛肉",
                "规格名称": "开心桶",
                "建议零售价": "5.0",
                "状态": "有效",
            },
            {
                "商品条形码": "6900000000002",
                "标品名称": "红烧牛肉965桶",
                "品牌名称": "红烧品牌",
                "口味名称": "红烧牛肉",
                "规格名称": "965桶",
                "建议零售价": "5.5",
                "状态": "有效",
            },
            {
                "商品条形码": "6900000000003",
                "标品名称": "红烧牛肉开心桶整箱",
                "品牌名称": "红烧品牌",
                "口味名称": "红烧牛肉",
                "规格名称": "开心桶",
                "建议零售价": "30.0",
                "状态": "有效",
            },
        ]
    )


def default_config() -> PlanSheetConfig:
    return PlanSheetConfig(
        start_cell="D2",
        end_cell="H20",
        label_col=4,
        first_activity_col=5,
        activity_row_start=2,
        activity_row_end=9,
        product_row_start=11,
        brand_col=3,
        spec_col=4,
        field_mappings=[
            FieldRowMapping(2, "活动类型", "活动类型"),
            FieldRowMapping(4, "活动机制", "活动机制"),
            FieldRowMapping(5, "机制力度", "机制力度"),
            FieldRowMapping(6, "活动时间", "活动时间"),
            FieldRowMapping(7, "活动区域/渠道", "活动区域/渠道"),
            FieldRowMapping(8, "优惠券类型", "优惠券类型"),
        ],
        activity_columns=[
            ActivityColumnOption(5, "E", "单品牌机制 / ①满16-2 ②满31-5", selected=True)
        ],
        ai_settings=AiSettings(enabled=False),
    )
