"""不依賴 GUI 或 Excel COM 的自訂規則判斷引擎。"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterable

from rule_models import CustomRule


NUMERIC_OPERATORS = {"range", ">", ">=", "<", "<="}


@dataclass(frozen=True)
class EvaluationResult:
    level: str | None
    numeric_conversion_failed: bool = False


def is_blank_value(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def to_number(value: Any) -> float | None:
    if isinstance(value, bool) or is_blank_value(value):
        return None
    try:
        return float(value.strip() if isinstance(value, str) else value)
    except (TypeError, ValueError):
        return None


def evaluate_rule(value: Any, rule: CustomRule) -> bool:
    """判斷單一值是否符合規則；數值轉換失敗安全地回傳 False。"""
    if not rule.enabled or is_blank_value(value):
        return False
    if rule.operator == "exact":
        return str(value).strip() == str(rule.value).strip()
    number = to_number(value)
    if number is None:
        return False
    if rule.operator == "range":
        return float(rule.minimum) <= number <= float(rule.maximum)
    target = float(rule.value)
    return {
        ">": number > target,
        ">=": number >= target,
        "<": number < target,
        "<=": number <= target,
    }[rule.operator]


def evaluate_custom_level(value: Any, gender: Any, rules: Iterable[CustomRule]) -> EvaluationResult:
    """專用性別規則優先、各群組依列表順序採 first match wins。"""
    ordered = [rule for rule in rules if rule.enabled]
    normalized_gender = str(gender).strip() if gender is not None else ""
    groups: list[list[CustomRule]] = []
    if normalized_gender in {"男", "女"}:
        groups.append([rule for rule in ordered if rule.gender == normalized_gender])
    groups.append([rule for rule in ordered if rule.gender == "共用"])
    numeric_failed = False
    for group in groups:
        for rule in group:
            if rule.operator in NUMERIC_OPERATORS and to_number(value) is None and not is_blank_value(value):
                numeric_failed = True
            if evaluate_rule(value, rule):
                return EvaluationResult(rule.level, numeric_failed)
    return EvaluationResult(None, numeric_failed)


def find_possible_overlaps(rules: list[CustomRule]) -> set[str]:
    """保守提示具有多條啟用數值規則的欄位/性別群組。"""
    counts: dict[tuple[str, str], int] = {}
    for rule in rules:
        if rule.enabled and rule.operator in NUMERIC_OPERATORS:
            key = (rule.column, rule.gender)
            counts[key] = counts.get(key, 0) + 1
    return {column for (column, _gender), count in counts.items() if count > 1}
