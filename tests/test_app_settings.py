from __future__ import annotations

from pathlib import Path

from plan_to_report.app_settings import load_app_settings, save_app_settings
from plan_to_report.plan_config import AiSettings


def test_save_and_load_app_settings(tmp_path: Path):
    settings = load_app_settings(tmp_path)
    settings.deepseek = AiSettings(
        enabled=True,
        api_key="test-key-123",
        model="deepseek-chat",
        base_url="https://api.deepseek.com",
    )
    path = save_app_settings(tmp_path, settings)
    assert path.exists()

    loaded = load_app_settings(tmp_path)
    assert loaded.deepseek.enabled is True
    assert loaded.deepseek.api_key == "test-key-123"
