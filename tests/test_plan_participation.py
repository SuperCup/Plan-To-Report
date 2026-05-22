from __future__ import annotations

import pandas as pd

from plan_to_report.plan_config import PlanSheetConfig
from plan_to_report.plan_participation import collect_participations


def test_collect_participations_checkmark():
    grid = [[None] * 6 for _ in range(15)]
    grid[9][2] = "红烧牛肉"
    grid[10][3] = "开心桶"
    grid[10][4] = "√"
    raw = pd.DataFrame(grid, dtype=object)

    config = PlanSheetConfig(
        end_cell="F15",
        product_row_start=11,
        brand_col=3,
        spec_col=4,
        first_activity_col=5,
    )
    parts = collect_participations(raw, config, 5)
    assert len(parts) == 1
    assert parts[0].brand == "红烧牛肉"
    assert parts[0].spec == "开心桶"
