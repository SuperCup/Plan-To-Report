from __future__ import annotations

import pandas as pd

from plan_to_report.plan_config import FieldRowMapping, PlanSheetConfig
from plan_to_report.plan_outputs import build_activity_summary
from plan_to_report.plan_parser import build_remark, parse_plan_sheet, split_mechanisms


def _sample_plan_grid() -> pd.DataFrame:
    rows = 20
    cols = 8
    grid = [[None] * cols for _ in range(rows)]
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
    grid[8][3] = "预算分配"
    grid[8][4] = "1236000"
    grid[9][3] = "规格"
    return pd.DataFrame(grid, dtype=object)


def test_split_mechanisms_by_circled_numbers():
    parts = split_mechanisms("①满16-2 ②满31-5")
    assert parts == ["满16-2", "满31-5"]


def test_build_remark_joins_fields():
    remark = build_remark(
        {
            "活动类型": "单品牌机制",
            "活动机制": "满16-2",
            "机制力度": "88%",
            "活动时间": "3/1-3/31",
            "活动区域/渠道": "全渠道",
            "优惠券类型": "品牌券",
            "预算分配": "1236000",
        }
    )
    assert remark == "单品牌机制+满16-2+88%+3/1-3/31+全渠道+品牌券+1236000"


def test_parse_plan_sheet_splits_mechanism_into_two_activities():
    raw = _sample_plan_grid()
    config = PlanSheetConfig(
        start_cell="D2",
        end_cell="H15",
        label_col=4,
        first_activity_col=5,
        activity_row_start=2,
        activity_row_end=9,
        product_row_start=11,
        field_mappings=[
            FieldRowMapping(2, "活动类型", "活动类型"),
            FieldRowMapping(4, "活动机制", "活动机制"),
            FieldRowMapping(5, "机制力度", "机制力度"),
            FieldRowMapping(6, "活动时间", "活动时间"),
            FieldRowMapping(7, "活动区域/渠道", "活动区域/渠道"),
            FieldRowMapping(8, "优惠券类型", "优惠券类型"),
            FieldRowMapping(9, "预算分配", "预算分配"),
        ],
        activity_columns=[],
    )
    from plan_to_report.plan_config import ActivityColumnOption

    config.activity_columns = [
        ActivityColumnOption(excel_col=5, column_letter="E", preview="单品牌机制 / ①满16-2 ②满31-5", selected=True)
    ]

    result = parse_plan_sheet(raw, config)
    assert len(result.activities) == 2
    assert result.activities[0].fields["活动机制"] == "满16-2"
    assert result.activities[1].fields["活动机制"] == "满31-5"

    summary = build_activity_summary(result)
    assert len(summary) == 2
    assert summary.iloc[0]["活动"] == "满16-2"
    assert summary.iloc[0]["优惠券类型"] == "同享"
    assert summary.iloc[0]["券名称"] == "康师傅品牌券"
