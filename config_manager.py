"""自訂規則 JSON 的安全讀寫。"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

from rule_models import CustomRule


def default_rules_path() -> Path:
    base = Path(os.getenv("APPDATA", Path.home())) if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
    return base / "Excel管理級數整理工具" / "custom_rules.json" if getattr(sys, "frozen", False) else base / "custom_rules.json"


def save_rules(path: str | Path, rules: list[CustomRule]) -> None:
    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_suffix(destination.suffix + ".tmp")
    temporary.write_text(json.dumps({"version": 1, "rules": [rule.to_dict() for rule in rules]}, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(destination)


def load_rules(path: str | Path) -> list[CustomRule]:
    source = Path(path)
    if not source.exists():
        return []
    try:
        payload = json.loads(source.read_text(encoding="utf-8"))
        if not isinstance(payload, dict):
            raise ValueError("規則檔最外層必須是 JSON 物件。")
        if payload.get("version") != 1 or not isinstance(payload.get("rules"), list):
            raise ValueError("不支援的規則檔格式。")
        return [CustomRule.from_dict(item) for item in payload["rules"]]
    except (OSError, json.JSONDecodeError, TypeError, ValueError) as exc:
        raise ValueError(f"規則檔無法讀取：{exc}") from exc
