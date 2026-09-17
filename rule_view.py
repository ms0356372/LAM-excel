"""自訂規則的唯讀呈現、搜尋與狀態彙整。"""

from __future__ import annotations

from dataclasses import dataclass

from rule_engine import find_override_pairs, find_rule_conflicts
from rule_models import OPERATOR_LABELS, CustomRule, format_rule_condition


@dataclass(frozen=True)
class RuleStatus:
    severity: str
    messages: tuple[str, ...] = ()

    @property
    def text(self) -> str:
        if not self.messages:
            return "正常"
        if len(self.messages) == 1:
            return f"⚠ {self.messages[0]}"
        return f"⚠ {len(self.messages)} 個問題"


def collect_rule_statuses(
    rules: list[CustomRule], available_columns: set[str] | None = None
) -> list[RuleStatus]:
    """以既有衝突檢查結果建立每條規則共用的顯示狀態。"""
    messages: list[list[str]] = [[] for _ in rules]
    severities = ["normal" for _ in rules]
    valid_rules: list[CustomRule] = []
    original_indexes: list[int] = []
    for index, rule in enumerate(rules):
        try:
            rule.validate()
        except ValueError as exc:
            detail = str(exc)
            if "下限不可大於上限" in detail:
                label = "區間異常"
            elif "必須是數字" in detail:
                label = "數值條件錯誤"
            else:
                label = detail
            messages[index].append(label); severities[index] = "error"
        else:
            valid_rules.append(rule); original_indexes.append(index)

    for conflict in find_rule_conflicts(valid_rules):
        label = {
            "duplicate": "完全重複",
            "same_condition": "相同條件不同分級",
            "overlap": "規則重疊",
        }[conflict.kind]
        for valid_index in (conflict.first_index, conflict.second_index):
            index = original_indexes[valid_index]
            if label not in messages[index]:
                messages[index].append(label)
                severities[index] = "warning"

    for specific, common in find_override_pairs(valid_rules):
        for valid_index in (specific, common):
            index = original_indexes[valid_index]
            if "與共用規則重疊" not in messages[index]:
                messages[index].append("與共用規則重疊")
                severities[index] = "warning"

    if available_columns is not None:
        for index, rule in enumerate(rules):
            if rule.column not in available_columns:
                messages[index].append("欄位不存在")
                severities[index] = "error"

    return [RuleStatus(severities[index], tuple(items)) for index, items in enumerate(messages)]


def rule_search_text(rule: CustomRule, header: str = "") -> str:
    return " ".join((rule.column, header, rule.gender, OPERATOR_LABELS[rule.operator], format_rule_condition(rule), rule.level)).casefold()


def filter_rule_indexes(
    rules: list[CustomRule], statuses: list[RuleStatus], headers: dict[str, str],
    search: str = "", gender: str = "全部", level: str = "全部",
    status: str = "全部", enabled: str = "全部",
) -> list[int]:
    needle = search.strip().casefold()
    result: list[int] = []
    for index, rule in enumerate(rules):
        item_status = statuses[index]
        if needle and needle not in rule_search_text(rule, headers.get(rule.column, "")):
            continue
        if gender != "全部" and rule.gender != gender:
            continue
        if level != "全部" and rule.level != level:
            continue
        if status != "全部" and {"正常": "normal", "警告": "warning", "錯誤": "error"}[status] != item_status.severity:
            continue
        if enabled != "全部" and (enabled == "已啟用") != rule.enabled:
            continue
        result.append(index)
    return result
