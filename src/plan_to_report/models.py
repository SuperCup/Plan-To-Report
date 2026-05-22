from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class InputSpec:
    key: str
    label: str
    required: bool = True


@dataclass(frozen=True)
class FieldSpec:
    name: str
    source: dict[str, Any]
    required: bool = False


@dataclass(frozen=True)
class OutputSpec:
    key: str
    label: str
    file_name: str
    sheet_name: str
    primary_input: str
    primary_sheet: str | None
    fields: list[FieldSpec] = field(default_factory=list)
    filters: list[dict[str, Any]] = field(default_factory=list)


@dataclass(frozen=True)
class ConversionTemplate:
    name: str
    version: str
    inputs: list[InputSpec]
    outputs: list[OutputSpec]
    path: Path | None = None


@dataclass
class ConversionIssue:
    level: str
    output: str
    row_number: int | None
    field: str | None
    message: str


@dataclass
class ConversionResult:
    tables: dict[str, Any]
    issues: list[ConversionIssue]
