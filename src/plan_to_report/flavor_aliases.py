from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any


def load_flavor_aliases(project_root: Path | None = None) -> dict[str, str]:
    candidates: list[Path] = []
    if project_root:
        candidates.append(project_root / "templates" / "口味别名.json")
    candidates.append(Path(__file__).resolve().parents[2] / "templates" / "口味别名.json")

    for path in candidates:
        if path.exists():
            with path.open("r", encoding="utf-8") as file:
                raw = json.load(file)
            if isinstance(raw, dict):
                return {str(key).strip(): str(value).strip() for key, value in raw.items() if key and value}
    return {}


def expand_flavor_token(
    token: str,
    aliases: dict[str, str],
    known_flavors: list[str],
) -> list[str]:
    token = token.strip()
    if not token:
        return []

    if token in aliases:
        return [aliases[token]]

    for alias_key, full_name in aliases.items():
        if token == alias_key or token in full_name or full_name in token:
            return [full_name]

    matched = [name for name in known_flavors if _flavor_matches(name, token)]
    if matched:
        return matched

    return [token]


def _flavor_matches(flavor_name: str, hint: str) -> bool:
    flavor_name = flavor_name.strip()
    hint = hint.strip()
    if not flavor_name or not hint:
        return False
    if hint in flavor_name or flavor_name in hint:
        return True
    if len(hint) >= 1 and hint[0] in flavor_name:
        return True
    return False
