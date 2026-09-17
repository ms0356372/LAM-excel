import json

import pytest

from config_manager import load_rules, save_rules
from rule_engine import evaluate_custom_level, evaluate_rule
from rule_models import CustomRule


def rule(operator, level="第一級", gender="共用", value=None, minimum=None, maximum=None):
    item = CustomRule("A", gender, operator, level, value=value, minimum=minimum, maximum=maximum)
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
