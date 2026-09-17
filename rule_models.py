"""自訂分級規則的資料模型與驗證。"""

from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any


VALID_GENDERS = ("共用", "男", "女")
VALID_OPERATORS = ("range", ">", ">=", "<", "<=", "exact")
VALID_LEVELS = ("第一級", "第二級", "第三級", "第四級")
OPERATOR_LABELS = {
    "range": "區間",
    ">": ">",
    ">=": ">=",
    "<": "<",
    "<=": "<=",
    "exact": "完全相符",
}
LABEL_TO_OPERATOR = {label: operator for operator, label in OPERATOR_LABELS.items()}


@dataclass
class CustomRule:
    """一條可序列化、可排序的自訂分級規則。"""

    column: str
    gender: str
    operator: str
    level: str
    enabled: bool = True
    value: Any = None
    minimum: float | None = None
    maximum: float | None = None

    def __post_init__(self) -> None:
        self.column = str(self.column).strip().upper()
        self.gender = str(self.gender).strip()
        self.operator = str(self.operator).strip()
        self.level = str(self.level).strip()

    def validate(self) -> None:
        if not self.column.isalpha() or not 1 <= len(self.column) <= 3:
            raise ValueError("規則來源欄位格式錯誤。")
        column_number = 0
        for char in self.column:
            column_number = column_number * 26 + ord(char) - 64
        if column_number > 16_384:
            raise ValueError("規則來源欄位超出 Excel 可用範圍。")
        if self.gender not in VALID_GENDERS:
            raise ValueError("規則性別必須為共用、男或女。")
        if self.operator not in VALID_OPERATORS:
            raise ValueError("規則判斷方式無效。")
        if self.level not in VALID_LEVELS:
            raise ValueError("規則分級必須為第一級至第四級。")
        if not isinstance(self.enabled, bool):
            raise ValueError("規則啟用狀態必須為布林值。")
        if self.operator == "range":
            self.minimum = _number(self.minimum, "區間下限")
            self.maximum = _number(self.maximum, "區間上限")
            if self.minimum > self.maximum:
                raise ValueError("區間下限不可大於上限。")
        elif self.operator in {">", ">=", "<", "<="}:
            self.value = _number(self.value, "比較數值")
        elif self.value is None or str(self.value).strip() == "":
            raise ValueError("完全相符的內容不可空白。")

    @property
    def condition_text(self) -> str:
        return format_rule_condition(self)

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        return {key: value for key, value in data.items() if value is not None}

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "CustomRule":
        if not isinstance(data, dict):
            raise ValueError("規則內容必須是 JSON 物件。")
        rule = cls(
            column=data.get("column", ""), gender=data.get("gender", "共用"),
            operator=data.get("operator", ""), level=data.get("level", ""),
            enabled=data.get("enabled", True), value=data.get("value"),
            minimum=data.get("minimum", data.get("min")),
            maximum=data.get("maximum", data.get("max")),
        )
        rule.validate()
        return rule


def _number(value: Any, name: str) -> float:
    try:
        if isinstance(value, str):
            value = value.strip()
        if value == "" or value is None or isinstance(value, bool):
            raise ValueError
        return float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{name}必須是數字。") from exc


def _display_number(value: Any) -> str:
    number = float(value)
    return str(int(number)) if number.is_integer() else str(number)


def format_rule_condition(rule: CustomRule) -> str:
    """將規則條件轉為列表、總覽及警告訊息共用的使用者文字。"""
    if rule.operator == "range":
        return f"{_display_number(rule.minimum)} ～ {_display_number(rule.maximum)}"
    if rule.operator == "exact":
        return f"完全相符：{str(rule.value).strip()}"
    return f"{rule.operator} {_display_number(rule.value)}"
