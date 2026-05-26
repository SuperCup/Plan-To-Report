from __future__ import annotations

from pathlib import Path

import pandas as pd

from .flavor_aliases import load_flavor_aliases
from .models import ConversionIssue
from .plan_config import PlanSheetConfig
from .plan_participation import collect_participations
from .plan_parser import ParsedActivity, PlanParseResult
from .product_matcher import build_logic_index, lookup_upc_rows

SUMMARY_COLUMNS = [
    "方案名称",
    "自定义编码",
    "活动",
    "发券渠道",
    "券名称",
    "活动时间",
    "优惠券类型",
    "渠道",
    "机制名称",
    "机制范围",
    "机制类型",
    "预算",
    "券总库存",
    "单日库存",
    "单用户领取总限制",
    "单用户每天领取限制",
    "备注",
    "活动列",
    "机制拆分序号",
]

UPC_COLUMNS = [
    "自定义编码",
    "活动列",
    "备注",
    "活动机制",
    "活动类型",
    "子品牌",
    "规格",
    "口味提示",
    "UPC条形码",
    "标品名称",
    "品牌名称",
    "口味名称",
    "规格名称",
    "建议零售价",
    "状态",
]


def build_activity_summary(parse_result: PlanParseResult) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for activity in parse_result.activities:
        rows.append(_activity_to_summary_row(activity))
    if not rows:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)
    return pd.DataFrame(rows, columns=SUMMARY_COLUMNS)


def build_activity_upc_table(
    parse_result: PlanParseResult,
    raw: pd.DataFrame,
    config: PlanSheetConfig,
    logic_table: pd.DataFrame | None,
    product_table: pd.DataFrame | None,
    project_root: Path | None = None,
) -> tuple[pd.DataFrame, list[ConversionIssue]]:
    issues: list[ConversionIssue] = []
    if logic_table is None or logic_table.empty:
        issues.append(
            ConversionIssue("error", "活动对应UPC表", None, None, "未上传或未识别「商品匹配逻辑」Sheet。")
        )
        return pd.DataFrame(columns=UPC_COLUMNS), issues
    if product_table is None or product_table.empty:
        issues.append(
            ConversionIssue("error", "活动对应UPC表", None, None, "未上传或未识别「商品清单」Sheet。")
        )
        return pd.DataFrame(columns=UPC_COLUMNS), issues

    logic_rules = build_logic_index(logic_table)
    if not logic_rules:
        issues.append(
            ConversionIssue(
                "error",
                "活动对应UPC表",
                None,
                None,
                "商品匹配逻辑表需包含列：品牌、规格、选品逻辑。",
            )
        )
        return pd.DataFrame(columns=UPC_COLUMNS), issues

    aliases = load_flavor_aliases(project_root)
    participations_by_column: dict[int, list] = {}
    for col in config.selected_column_indices():
        participations_by_column[col] = collect_participations(raw, config, col, project_root)

    rows: list[dict[str, object]] = []
    for activity in parse_result.activities:
        parts = participations_by_column.get(activity.column_index, [])
        split_suffix = ""
        if activity.split_total > 1:
            split_suffix = f" ({activity.split_index + 1}/{activity.split_total})"
        remark = activity.remark + split_suffix
        mechanism = activity.fields.get("活动机制")

        for part in parts:
            product_rows = lookup_upc_rows(
                part,
                logic_rules,
                product_table,
                issues,
                activity.column_letter,
                aliases,
            )
            for product in product_rows:
                rows.append(
                    {
                        "自定义编码": activity.remark,
                        "活动列": activity.column_letter,
                        "备注": remark,
                        "活动机制": mechanism,
                        "活动类型": activity.fields.get("活动类型"),
                        "子品牌": part.brand,
                        "规格": part.spec,
                        "口味提示": part.flavor_hint or "",
                        "UPC条形码": _first_non_blank(product.get("UPC条形码"), product.get("商品条形码")),
                        "标品名称": _first_non_blank(product.get("标品名称")),
                        "品牌名称": _first_non_blank(product.get("品牌名称")),
                        "口味名称": _first_non_blank(product.get("口味名称")),
                        "规格名称": _first_non_blank(product.get("规格名称")),
                        "建议零售价": _first_non_blank(product.get("建议零售价")),
                        "状态": _first_non_blank(product.get("状态")),
                    }
                )

    if not rows:
        issues.append(
            ConversionIssue(
                "warning",
                "活动对应UPC表",
                None,
                None,
                "未生成 UPC 行：请检查矩阵勾选、商品匹配逻辑与商品清单。",
            )
        )
        return pd.DataFrame(columns=UPC_COLUMNS), issues

    return pd.DataFrame(rows, columns=UPC_COLUMNS), issues


def _activity_to_summary_row(activity: ParsedActivity) -> dict[str, object]:
    fields = activity.fields
    coupon_type = fields.get("优惠券类型")
    mechanism = fields.get("活动机制")
    activity_type = fields.get("活动类型")
    coupon_name = ""
    if coupon_type is not None and str(coupon_type).strip():
        coupon_name = f"康师傅{str(coupon_type).strip()}"

    split_suffix = ""
    if activity.split_total > 1:
        split_suffix = f" ({activity.split_index + 1}/{activity.split_total})"

    return {
        "方案名称": "",
        "自定义编码": activity.remark,
        "活动": mechanism,
        "发券渠道": "",
        "券名称": coupon_name,
        "活动时间": fields.get("活动时间"),
        "优惠券类型": "同享",
        "渠道": fields.get("活动区域/渠道"),
        "机制名称": mechanism,
        "机制范围": activity_type,
        "机制类型": "商品券",
        "预算": "",
        "券总库存": "",
        "单日库存": "",
        "单用户领取总限制": "",
        "单用户每天领取限制": "",
        "备注": activity.remark + split_suffix,
        "活动列": activity.column_letter,
        "机制拆分序号": activity.split_index + 1 if activity.split_total > 1 else "",
    }


def _first_non_blank(*values: object) -> object:
    for value in values:
        if isinstance(value, pd.Series):
            for item in value.tolist():
                if not _is_blank(item):
                    return item
            continue
        if not _is_blank(value):
            return value
    return None


def _is_blank(value: object) -> bool:
    if value is None:
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False
