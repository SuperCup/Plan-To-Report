#!/usr/bin/env python
"""规划表全链路联调脚本（合成数据，无需真实 Excel）。"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

from plan_to_report.plan_pipeline import PlanConversionInput, run_plan_conversion
from tests.plan_fixtures import build_logic_table, build_plan_grid, build_product_table, default_config


def main() -> int:
    output = run_plan_conversion(
        PlanConversionInput(
            plan_raw=build_plan_grid(),
            config=default_config(),
            logic_table=build_logic_table(),
            product_table=build_product_table(),
            project_root=ROOT,
        )
    )
    print("=== 活动汇总表 ===")
    print(output.summary_table.to_string(index=False))
    print("\n=== 活动对应 UPC 表 ===")
    print(output.upc_table.to_string(index=False))
    if not output.issues_table.empty:
        print("\n=== 异常清单 ===")
        print(output.issues_table.to_string(index=False))
    print(
        f"\n完成：活动 {len(output.summary_table)} 条，UPC {len(output.upc_table)} 条，"
        f"提示 {len(output.issues)} 条。"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
