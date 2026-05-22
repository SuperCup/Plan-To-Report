from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from .models import ConversionIssue
from .plan_config import PlanSheetConfig
from .plan_outputs import build_activity_summary, build_activity_upc_table
from .plan_parser import PlanParseResult, parse_plan_sheet


@dataclass
class PlanConversionInput:
    plan_raw: pd.DataFrame
    config: PlanSheetConfig
    logic_table: pd.DataFrame | None
    product_table: pd.DataFrame | None
    project_root: Path | None = None


@dataclass
class PlanConversionOutput:
    summary_table: pd.DataFrame
    upc_table: pd.DataFrame
    issues_table: pd.DataFrame
    parse_result: PlanParseResult
    issues: list[ConversionIssue] = field(default_factory=list)


def run_plan_conversion(plan_input: PlanConversionInput) -> PlanConversionOutput:
    parse_result = parse_plan_sheet(plan_input.plan_raw, plan_input.config)
    summary_table = build_activity_summary(parse_result)
    upc_table, upc_issues = build_activity_upc_table(
        parse_result,
        plan_input.plan_raw,
        plan_input.config,
        plan_input.logic_table,
        plan_input.product_table,
        plan_input.project_root,
    )

    issues = list(parse_result.issues) + upc_issues
    if summary_table.empty:
        issues.append(
            ConversionIssue(
                "warning",
                "活动汇总表",
                None,
                None,
                "未生成任何活动行，请检查列勾选与字段映射。",
            )
        )

    issues_table = _issues_to_dataframe(issues)
    return PlanConversionOutput(
        summary_table=summary_table,
        upc_table=upc_table,
        issues_table=issues_table,
        parse_result=parse_result,
        issues=issues,
    )


def output_tables(output: PlanConversionOutput) -> dict[str, pd.DataFrame]:
    tables = {
        "活动汇总表": output.summary_table,
        "活动对应UPC表": output.upc_table,
    }
    if not output.issues_table.empty:
        tables["异常清单"] = output.issues_table
    return tables


def _issues_to_dataframe(issues: list[ConversionIssue]) -> pd.DataFrame:
    if not issues:
        return pd.DataFrame()
    return pd.DataFrame(
        [
            {
                "级别": issue.level,
                "输出表": issue.output,
                "源数据行": issue.row_number,
                "字段": issue.field,
                "说明": issue.message,
            }
            for issue in issues
        ]
    )
