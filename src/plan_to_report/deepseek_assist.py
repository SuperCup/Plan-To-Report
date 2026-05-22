from __future__ import annotations

import json
import re
import urllib.error
import urllib.request
from typing import Any

from .plan_config import AiSettings


class DeepseekAssistError(RuntimeError):
    pass


def assist_split_mechanism(text: str, settings: AiSettings) -> list[str] | None:
    prompt = (
        "将以下促销活动机制文本拆分为多个独立机制段落，每段一个满减规则。"
        "仅返回 JSON：{\"segments\": [\"满16-2\", \"满31-5\"]}。\n"
        f"原文：{text}"
    )
    payload = _chat_json(prompt, settings)
    if not payload:
        return None
    segments = payload.get("segments")
    if not isinstance(segments, list):
        return None
    cleaned = [str(item).strip() for item in segments if str(item).strip()]
    return cleaned or None


def assist_parse_participation_cell(
    cell_text: str,
    brand: str,
    spec: str,
    known_flavors: list[str],
    settings: AiSettings,
) -> list[str] | None:
    flavor_hint = ", ".join(known_flavors[:30])
    prompt = (
        "规划表活动列单元格可能包含口味缩写或合写。请解析需要参与活动的口味名称列表。"
        "仅返回 JSON：{\"flavors\": [\"藤椒牛肉\", \"麻辣牛肉\"]}。\n"
        f"子品牌组：{brand}\n规格：{spec}\n单元格：{cell_text}\n"
        f"可选口味行：{flavor_hint}"
    )
    payload = _chat_json(prompt, settings)
    if not payload:
        return None
    flavors = payload.get("flavors")
    if not isinstance(flavors, list):
        return None
    cleaned = [str(item).strip() for item in flavors if str(item).strip()]
    return cleaned or None


def should_try_ai_mechanism_split(text: str, rule_segments: list[str]) -> bool:
    if len(rule_segments) > 1:
        return False
    if re.search(r"[①②③④⑤]", text):
        return True
    if len(re.findall(r"满\s*\d+", text)) > 1:
        return True
    if "；" in text or ";" in text:
        return True
    return False


def _chat_json(prompt: str, settings: AiSettings) -> dict[str, Any] | None:
    api_key = settings.resolved_api_key()
    if not api_key:
        return None

    url = settings.base_url.rstrip("/") + "/chat/completions"
    body = json.dumps(
        {
            "model": settings.model,
            "messages": [
                {"role": "system", "content": "你是 Excel 规划表解析助手，只输出合法 JSON。"},
                {"role": "user", "content": prompt},
            ],
            "temperature": 0,
            "response_format": {"type": "json_object"},
        },
        ensure_ascii=False,
    ).encode("utf-8")

    request = urllib.request.Request(
        url,
        data=body,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=45) as response:
            raw = json.loads(response.read().decode("utf-8"))
    except (urllib.error.URLError, json.JSONDecodeError, TimeoutError) as exc:
        raise DeepseekAssistError(f"DeepSeek 调用失败：{exc}") from exc

    content = raw.get("choices", [{}])[0].get("message", {}).get("content", "")
    if not content:
        return None
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", content, re.DOTALL)
        if match:
            return json.loads(match.group(0))
    return None
