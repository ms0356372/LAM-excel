import json

import pytest

from config_manager import load_rules, save_rules
from rule_engine import (
    evaluate_custom_level,
    evaluate_rule,
    find_override_pairs,
    find_rule_conflicts,
    rules_overlap,
)
from rule_models import CustomRule, format_rule_condition
from rule_view import collect_rule_statuses, filter_rule_indexes


def rule(operator, level="第一級", gender="共用", value=None, minimum=None, maximum=None, column="A"):
    item = CustomRule(column, gender, operator, level, value=value, minimum=minimum, maximum=maximum)
    item.validate()
    return item


def test_inclusive_range_and_numeric_text():
    item = rule("range", minimum=1, maximum=10)
    assert evaluate_rule(5, item)
    assert evaluate_rule(10, item)
    assert evaluate_rule(" 1 ", item)
    assert not evaluate_rule(10.1, item)


def test_strict_comparisons_and_invalid_numeric_value():
    greater = rule(">", "第三級", value=10.6)
    lower = rule("<", "第二級", value=25.8)
    assert evaluate_rule(10.7, greater)
    assert not evaluate_rule(10.6, greater)
    assert evaluate_rule(25.7, lower)
    assert not evaluate_rule("ABC", greater)
    assert evaluate_custom_level("ABC", "男", [greater]).numeric_conversion_failed


@pytest.mark.parametrize("expected,actual,matched", [("++", "++", True), ("++", "+++", False), ("代謝症候群", " 代謝症候群 ", True), ("代謝症候群", "疑似代謝症候群", False)])
def test_exact_match(expected, actual, matched):
    assert evaluate_rule(actual, rule("exact", value=expected)) is matched


def test_gender_override_then_common_fallback():
    common = rule(">=", "第二級", value=80)
    male = rule(">=", "第三級", gender="男", value=90)
    female = rule(">=", "第三級", gender="女", value=80)
    assert evaluate_custom_level(95, "男", [common, male]).level == "第三級"
    assert evaluate_custom_level(85, "男", [common, male]).level == "第二級"
    assert evaluate_custom_level(92, "女", [male]).level is None
    assert evaluate_custom_level(82, "女", [female]).level == "第三級"
    assert evaluate_custom_level(82, "未知", [common, female]).level == "第二級"


def test_first_match_wins_and_disabled_rule_is_ignored():
    first = rule(">", "第二級", value=10)
    second = rule(">", "第三級", value=20)
    assert evaluate_custom_level(25, "", [first, second]).level == "第二級"
    first.enabled = False
    assert evaluate_custom_level(25, "", [first, second]).level == "第三級"


def test_blank_skips():
    assert evaluate_custom_level(None, "男", [rule(">", value=10)]).level is None


def test_validation_and_json_roundtrip(tmp_path):
    with pytest.raises(ValueError):
        rule("range", minimum=10, maximum=1)
    path = tmp_path / "rules.json"
    rules = [rule(">=", "第三級", "男", value=90), rule("exact", "第二級", value="++")]
    save_rules(path, rules)
    assert [item.to_dict() for item in load_rules(path)] == [item.to_dict() for item in rules]
    path.write_text("{broken", encoding="utf-8")
    with pytest.raises(ValueError, match="規則檔無法讀取"):
        load_rules(path)
    path.write_text(json.dumps([]), encoding="utf-8")
    with pytest.raises(ValueError, match="最外層"):
        load_rules(path)


@pytest.mark.parametrize(
    "first,second",
    [
        (rule(">", value=10), rule(">", value=20)),
        (rule("<", value=30), rule("<", value=20)),
        (rule("range", minimum=1, maximum=10), rule("range", minimum=5, maximum=15)),
        (rule("range", minimum=1, maximum=10), rule(">", value=8)),
        (rule(">=", value=10), rule("<=", value=20)),
        (rule("exact", value="++"), rule("exact", value=" ++ ")),
    ],
)
def test_required_overlap_examples(first, second):
    assert rules_overlap(first, second)


@pytest.mark.parametrize(
    "first,second",
    [
        (rule("range", minimum=1, maximum=10), rule(">", value=10)),
        (rule("range", minimum=1, maximum=10), rule("<", value=1)),
        (rule("exact", value="++"), rule("exact", value="+++")),
        (rule("<", value=10), rule(">=", value=10)),
    ],
)
def test_non_overlapping_boundaries(first, second):
    assert not rules_overlap(first, second)


def test_conflict_scope_duplicate_and_same_condition_classification():
    duplicate = rule(">=", "第二級", value=27)
    different_level = rule(">=", "第三級", value=27)
    other_gender = rule(">=", "第四級", "男", value=27)
    other_column = rule(">=", "第四級", value=27, column="B")
    conflicts = find_rule_conflicts([duplicate, rule(">=", "第二級", value=27), different_level, other_gender, other_column])
    assert [conflict.kind for conflict in conflicts] == ["duplicate", "same_condition", "same_condition"]
    assert all(conflict.second_index not in {3, 4} for conflict in conflicts)


def test_disabled_rules_are_not_conflicts_and_override_is_separate():
    common = rule(">=", "第二級", value=80)
    male = rule(">=", "第三級", "男", value=90)
    disabled = rule(">=", "第四級", value=90)
    disabled.enabled = False
    assert find_rule_conflicts([common, male, disabled]) == []
    assert find_override_pairs([common, male, disabled]) == [(1, 0)]


def test_shared_condition_formatting():
    assert format_rule_condition(rule("range", minimum=1, maximum=10)) == "1 ～ 10"
    assert format_rule_condition(rule(">=", value=90)) == ">= 90"
    assert format_rule_condition(rule("exact", value="++")) == "完全相符：++"


def test_rule_view_search_filters_and_statuses():
    rules = [
        rule(">", "第二級", value=10, column="M"),
        rule(">", "第三級", value=20, column="M"),
        rule(">=", "第三級", "男", value=90, column="N"),
    ]
    headers = {"M": "BMI", "N": "腰圍"}
    statuses = collect_rule_statuses(rules, {"M", "N"})
    assert statuses[0].severity == statuses[1].severity == "warning"
    assert "規則重疊" in statuses[0].messages
    assert filter_rule_indexes(rules, statuses, headers, search="BMI") == [0, 1]
    assert filter_rule_indexes(rules, statuses, headers, gender="男") == [2]
    assert filter_rule_indexes(rules, statuses, headers, level="第三級") == [1, 2]
    assert filter_rule_indexes(rules, statuses, headers, status="警告") == [0, 1]


def test_rule_view_reports_duplicate_override_and_missing_column():
    common = rule(">=", "第二級", value=80, column="N")
    duplicate = rule(">=", "第二級", value=80, column="N")
    male = rule(">=", "第三級", "男", value=90, column="N")
    missing = rule("exact", value="++", column="O")
    statuses = collect_rule_statuses([common, duplicate, male, missing], {"N"})
    assert "完全重複" in statuses[0].messages
    assert "與共用規則重疊" in statuses[2].messages
    assert statuses[3].severity == "error"
    assert statuses[3].text == "⚠ 欄位不存在"


def test_rule_view_reuses_model_validation_for_invalid_conditions():
    bad_range = CustomRule("A", "共用", "range", "第一級", minimum=10, maximum=1)
    bad_number = CustomRule("B", "共用", ">", "第一級", value="not-a-number")
    statuses = collect_rule_statuses([bad_range, bad_number])
    assert statuses[0].messages == ("區間異常",)
    assert statuses[1].messages == ("數值條件錯誤",)
