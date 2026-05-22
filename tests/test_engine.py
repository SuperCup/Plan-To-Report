from __future__ import annotations

import pandas as pd

from plan_to_report.engine import ConversionEngine
from plan_to_report.template_loader import load_template, parse_template


def test_engine_maps_regex_and_lookup(tmp_path):
    plan_path = tmp_path / "plan.xlsx"
    products_path = tmp_path / "products.xlsx"

    pd.DataFrame(
        [
            {
                "活动名称": "三月活动",
                "活动机制": "满16-2元",
                "产品名称": "红烧牛肉面",
            }
        ]
    ).to_excel(plan_path, index=False)

    pd.DataFrame(
        [
            {
                "产品名称": "红烧牛肉面",
                "UPC条形码": "6900000000001",
            }
        ]
    ).to_excel(products_path, index=False)

    template = parse_template(
        {
            "name": "test",
            "inputs": [
                {"key": "plan", "label": "规划表"},
                {"key": "products", "label": "商品清单"},
            ],
            "outputs": [
                {
                    "key": "out",
                    "label": "输出",
                    "file_name": "out.xlsx",
                    "sheet_name": "输出",
                    "primary_input": "plan",
                    "fields": [
                        {"name": "活动", "source": {"type": "direct", "column": "活动名称"}},
                        {
                            "name": "满",
                            "source": {
                                "type": "regex_extract",
                                "column": "活动机制",
                                "pattern": "满(\\d+)[-减](\\d+)",
                                "group": 1,
                            },
                        },
                        {
                            "name": "UPC",
                            "source": {
                                "type": "lookup",
                                "input": "products",
                                "left_column": "产品名称",
                                "right_column": "产品名称",
                                "return_column": "UPC条形码",
                            },
                        },
                    ],
                }
            ],
        }
    )

    result = ConversionEngine(template).run({"plan": plan_path, "products": products_path})

    assert result.tables["out"].iloc[0]["活动"] == "三月活动"
    assert result.tables["out"].iloc[0]["满"] == "16"
    assert result.tables["out"].iloc[0]["UPC"] == "6900000000001"


def test_sample_template_maps_activity_summary_and_upc_code():
    template = load_template("templates/规划转提报示例模板.json")
    engine = ConversionEngine(template)

    plan_table = pd.DataFrame(
        [
            {
                "活动机制": "满16-2元",
                "活动时间": "2026-06-01至2026-06-30",
                "活动区域/渠道": "华东",
                "活动类型": "满减",
                "优惠券类型": "立减券",
                "产品名称": "红烧牛肉面",
            }
        ]
    )
    product_table = pd.DataFrame(
        [
            {
                "产品名称": "红烧牛肉面",
                "UPC条形码": "6900000000001",
                "指导价": "5.00",
            }
        ]
    )

    result = engine.run_tables({"规划表": plan_table, "商品清单": product_table})

    summary = result.tables["活动汇总表"].iloc[0]
    upc = result.tables["活动对应UPC表"].iloc[0]

    assert summary["方案名称"] == ""
    assert summary["活动"] == "满16-2元"
    assert summary["发券渠道"] == ""
    assert summary["券名称"] == "康师傅立减券"
    assert summary["活动时间"] == "2026-06-01至2026-06-30"
    assert summary["优惠券类型"] == "同享"
    assert summary["渠道"] == "华东"
    assert summary["机制名称"] == "满16-2元"
    assert summary["机制范围"] == "满减"
    assert summary["机制类型"] == "商品券"
    assert summary["预算"] == ""
    assert summary["自定义编码"] == upc["自定义编码"]
    assert upc["活动名称"] == "满16-2元"
    assert upc["UPC条形码"] == "6900000000001"
