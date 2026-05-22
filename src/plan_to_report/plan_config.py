from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class SheetRange:
    start_cell: str = "A1"
    end_cell: str = ""


LOGIC_FIELD_OPTIONS = [
    "",
    "活动类型",
    "活动机制",
    "机制力度",
    "活动时间",
    "活动区域/渠道",
    "优惠券类型",
    "预算分配",
    "机制说明",
]

REMARK_FIELD_ORDER = [
    "活动类型",
    "活动机制",
    "机制力度",
    "活动时间",
    "活动区域/渠道",
    "优惠券类型",
    "预算分配",
    "机制说明",
]


@dataclass
class FieldRowMapping:
    excel_row: int
    label_text: str
    field_key: str = ""


@dataclass
class ActivityColumnOption:
    excel_col: int
    column_letter: str
    preview: str
    selected: bool = True


@dataclass
class AiSettings:
    enabled: bool = False
    api_key: str = ""
    model: str = "deepseek-chat"
    base_url: str = "https://api.deepseek.com"

    def resolved_api_key(self, fallback_key: str = "") -> str:
        import os

        return (
            self.api_key.strip()
            or fallback_key.strip()
            or os.environ.get("DEEPSEEK_API_KEY", "").strip()
        )


@dataclass
class PlanSheetConfig:
    start_cell: str = "D2"
    end_cell: str = ""
    label_col: int = 4
    first_activity_col: int = 5
    activity_row_start: int = 2
    activity_row_end: int = 9
    product_row_start: int = 11
    brand_col: int = 3
    spec_col: int = 4
    field_mappings: list[FieldRowMapping] = field(default_factory=list)
    activity_columns: list[ActivityColumnOption] = field(default_factory=list)
    ai_settings: AiSettings = field(default_factory=AiSettings)

    def selected_column_indices(self) -> list[int]:
        return [item.excel_col for item in self.activity_columns if item.selected]
