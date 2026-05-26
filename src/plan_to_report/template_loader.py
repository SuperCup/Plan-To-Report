from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .models import ConversionTemplate, FieldSpec, InputSpec, OutputSpec


class TemplateError(ValueError):
    """Raised when a conversion template is invalid."""


def load_template(path: str | Path) -> ConversionTemplate:
    template_path = Path(path)
    with template_path.open("r", encoding="utf-8") as file:
        raw = json.load(file)

    return parse_template(raw, template_path)


def parse_template(raw: dict[str, Any], path: Path | None = None) -> ConversionTemplate:
    try:
        inputs = [
            InputSpec(
                key=item["key"],
                label=item.get("label", item["key"]),
                required=item.get("required", True),
            )
            for item in raw["inputs"]
        ]
        outputs = [
            OutputSpec(
                key=item["key"],
                label=item.get("label", item["key"]),
                file_name=item["file_name"],
                sheet_name=item.get("sheet_name", item.get("label", item["key"])),
                primary_input=item["primary_input"],
                primary_sheet=item.get("primary_sheet"),
                filters=item.get("filters", []),
                fields=[
                    FieldSpec(
                        name=field["name"],
                        source=field["source"],
                        required=field.get("required", False),
                    )
                    for field in item.get("fields", [])
                ],
            )
            for item in raw["outputs"]
        ]
    except KeyError as exc:
        raise TemplateError(f"模板缺少必要字段：{exc.args[0]}") from exc

    if not outputs:
        raise TemplateError("模板至少需要定义一个输出表。")

    return ConversionTemplate(
        name=raw.get("name", path.stem if path else "未命名模板"),
        version=raw.get("version", "0.1.0"),
        inputs=inputs,
        outputs=outputs,
        path=path,
    )


def list_templates(directory: str | Path) -> list[Path]:
    template_dir = Path(directory)
    if not template_dir.exists():
        return []

    templates: list[Path] = []
    for path in sorted(template_dir.glob("*.json")):
        try:
            with path.open("r", encoding="utf-8") as file:
                raw = json.load(file)
        except (json.JSONDecodeError, OSError):
            continue
        if isinstance(raw, dict) and isinstance(raw.get("inputs"), list) and isinstance(raw.get("outputs"), list):
            templates.append(path)
    return templates
