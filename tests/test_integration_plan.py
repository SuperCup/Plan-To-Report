from __future__ import annotations

import pandas as pd

from plan_to_report.plan_config import PlanSheetConfig
from plan_to_report.plan_participation import collect_participations
from plan_to_report.plan_pipeline import PlanConversionInput, run_plan_conversion

from .plan_fixtures import build_logic_table, build_plan_grid, build_product_table, default_config


def test_integration_pipeline_end_to_end():
    output = run_plan_conversion(
        PlanConversionInput(
            plan_raw=build_plan_grid(),
            config=default_config(),
            logic_table=build_logic_table(),
            product_table=build_product_table(),
        )
    )

    assert len(output.summary_table) == 2
    assert len(output.upc_table) >= 2
    assert "UPC条形码" in output.upc_table.columns
    assert output.upc_table["UPC条形码"].notna().any()


def test_packaging_block_participation():
    grid = [[None] * 7 for _ in range(20)]
    grid[9][2] = "地方口味"
    grid[10][2] = "藤椒牛肉"
    grid[10][3] = "开心桶"
    grid[11][2] = "麻辣牛肉"
    grid[11][3] = "开心桶"
    grid[10][5] = "藤/麻\n开心桶"
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
