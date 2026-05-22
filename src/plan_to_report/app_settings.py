from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path

from .plan_config import AiSettings


@dataclass
class AppSettings:
    deepseek: AiSettings = field(default_factory=AiSettings)


def settings_path(project_root: Path) -> Path:
    return project_root / "config" / "app_settings.json"


def load_app_settings(project_root: Path) -> AppSettings:
    path = settings_path(project_root)
    if not path.exists():
        return AppSettings()

    try:
        with path.open("r", encoding="utf-8") as file:
            raw = json.load(file)
    except (json.JSONDecodeError, OSError):
        return AppSettings()

    deepseek_raw = raw.get("deepseek", {})
    if not isinstance(deepseek_raw, dict):
        return AppSettings()

    return AppSettings(
        deepseek=AiSettings(
            enabled=bool(deepseek_raw.get("enabled", False)),
            api_key=str(deepseek_raw.get("api_key", "")).strip(),
            model=str(deepseek_raw.get("model", "deepseek-chat")).strip() or "deepseek-chat",
            base_url=str(deepseek_raw.get("base_url", "https://api.deepseek.com")).strip()
            or "https://api.deepseek.com",
        )
    )


def save_app_settings(project_root: Path, settings: AppSettings) -> Path:
    path = settings_path(project_root)
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "deepseek": {
            "enabled": settings.deepseek.enabled,
            "api_key": settings.deepseek.api_key,
            "model": settings.deepseek.model,
            "base_url": settings.deepseek.base_url,
        }
    }
    with path.open("w", encoding="utf-8") as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)
    return path
