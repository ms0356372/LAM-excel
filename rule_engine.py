"""不依賴 GUI 或 Excel COM 的自訂規則判斷引擎。"""

from __future__ import annotations

from dataclasses import dataclass
from math import inf
from typing import Any, Iterable

from rule_models import CustomRule


NUMERIC_OPERATORS = {"range", ">", ">=", "<", "<="}


@dataclass(frozen=True)
class EvaluationResult:
    level: str | None
    numeric_conversion_failed: bool = False


@dataclass(frozen=True)
class RuleConflict:
    """兩條規則之間的衝突資訊；索引對應傳入規則清單。"""

    first_index: int
    second_index: int
    kind: str

    @property
    def is_duplicate(self) -> bool:
        return self.kind == "duplicate"

    @property
    def is_same_condition(self) -> bool:
        return self.kind == "same_condition"


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
    if is_blank_value(value):
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
    ordered = list(rules)
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


def rules_have_same_condition(first: CustomRule, second: CustomRule) -> bool:
    """判斷運算子和條件是否完全相同（不比較分級）。"""
    if first.operator != second.operator:
        return False
    if first.operator == "range":
        return float(first.minimum) == float(second.minimum) and float(first.maximum) == float(second.maximum)
    if first.operator == "exact":
        return str(first.value).strip() == str(second.value).strip()
    return float(first.value) == float(second.value)


def rules_overlap(first: CustomRule, second: CustomRule) -> bool:
    """精確判斷兩條規則是否存在至少一個可同時符合的值。"""
    if first.operator == "exact" and second.operator == "exact":
        return str(first.value).strip() == str(second.value).strip()
    if first.operator == "exact":
        return _numeric_rule_matches(first.value, second)
    if second.operator == "exact":
        return _numeric_rule_matches(second.value, first)

    first_low, first_low_closed, first_high, first_high_closed = _numeric_bounds(first)
    second_low, second_low_closed, second_high, second_high_closed = _numeric_bounds(second)
    low = max(first_low, second_low)
    high = min(first_high, second_high)
    if low < high:
        return True
    if low > high:
        return False
    first_contains = _bound_contains(first_low, first_low_closed, first_high, first_high_closed, low)
    second_contains = _bound_contains(second_low, second_low_closed, second_high, second_high_closed, low)
    return first_contains and second_contains


def find_rule_conflicts(rules: list[CustomRule]) -> list[RuleConflict]:
    """只比較相同欄位、相同性別群組內的規則。"""
    conflicts: list[RuleConflict] = []
    for first_index, first in enumerate(rules):
        for second_index in range(first_index + 1, len(rules)):
            second = rules[second_index]
            if first.column != second.column or first.gender != second.gender:
                continue
            if not rules_overlap(first, second):
                continue
            same_condition = rules_have_same_condition(first, second)
            if same_condition and first.level == second.level:
                kind = "duplicate"
            elif same_condition:
                kind = "same_condition"
            else:
                kind = "overlap"
            conflicts.append(RuleConflict(first_index, second_index, kind))
    return conflicts


def find_override_pairs(rules: list[CustomRule]) -> list[tuple[int, int]]:
    """找出男／女專用規則與同欄共用規則可能同時命中的組合。"""
    pairs: list[tuple[int, int]] = []
    for specific_index, specific in enumerate(rules):
        if specific.gender not in {"男", "女"}:
            continue
        for common_index, common in enumerate(rules):
            if common.gender == "共用" and specific.column == common.column and rules_overlap(specific, common):
                pairs.append((specific_index, common_index))
    return pairs


def _numeric_bounds(rule: CustomRule) -> tuple[float, bool, float, bool]:
    if rule.operator == "range":
        return float(rule.minimum), True, float(rule.maximum), True
    value = float(rule.value)
    if rule.operator == ">":
        return value, False, inf, False
    if rule.operator == ">=":
        return value, True, inf, False
    if rule.operator == "<":
        return -inf, False, value, False
    return -inf, False, value, True


def _numeric_rule_matches(value: Any, rule: CustomRule) -> bool:
    number = to_number(value)
    if number is None:
        return False
    low, low_closed, high, high_closed = _numeric_bounds(rule)
    return _bound_contains(low, low_closed, high, high_closed, number)


def _bound_contains(low: float, low_closed: bool, high: float, high_closed: bool, value: float) -> bool:
    return (value > low or value == low and low_closed) and (value < high or value == high and high_closed)
