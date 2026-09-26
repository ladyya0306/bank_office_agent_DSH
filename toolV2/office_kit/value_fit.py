"""Fit a source value into a template's already printed wrapper or unit.

Only adapt values to literal text the template already supplies.  Same-currency
renminbi units are converted exactly; no cross-currency conversion is attempted.
"""
from __future__ import annotations

from decimal import Decimal, InvalidOperation, localcontext
import re

from .common import OfficeKitError


_PRINTED_ENDINGS = ("有限公司", "万元", "人民币", "公司", "元", "年", "号")
_YUAN_AMOUNT = re.compile(r"(?:人民币\s*)?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?\s*(?:万|亿)?")
_AMOUNT_VALUE = re.compile(
    r"^\s*(?:人民币\s*)?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?\s*(?P<unit>亿元|万元|元|亿|万)?\s*$"
)
_AMOUNT_UNITS = ("万元", "亿元", "元", "万", "亿")
_AMOUNT_FACTORS = {"元": Decimal(1), "万": Decimal(10_000), "万元": Decimal(10_000),
                   "亿": Decimal(100_000_000), "亿元": Decimal(100_000_000)}
_VALUE_UNITS = ("万元", "亿元", "元", "年", "个月", "月", "日", "人", "户", "份", "笔", "个", "次", "家", "%")
_UNIT_BLANK_RE = re.compile(
    r"(?:[ \u3000]+|_{2,}|＿{2,}|[□☐]+)(?=(?:万元|亿元|元|年|个月|月|日|人|户|份|笔|个|次|家|%))"
)
_PLACEHOLDER_MARK_RE = re.compile(r"[ \u3000]+|_{2,}|＿{2,}|[□☐]+")
_DATE_PART = r"\d{4}(?:\s*年(?:\s*\d{1,2}\s*月(?:\s*\d{1,2}\s*日?)?)?|[-./]\d{1,2}(?:[-./]\d{1,2})?)?"
_DATE_RANGE = re.compile(_DATE_PART + r"\s*起?\s*(?:至|到|[-—–~～]+)\s*" + _DATE_PART)


def _printed_amount_unit(text: str, start: int, end: int) -> str | None:
    right = text[end:].lstrip()
    suffix = next((unit for unit in _AMOUNT_UNITS if right.startswith(unit)), None)
    if suffix:
        return suffix
    # A label may already supply the unit before its blank, e.g.
    # “金额（万元）：____”. Reuse exactly the same Decimal conversion as
    # “金额：____万元”; do not append a second unit inside the value.
    label_unit = re.search(r'[（(]\s*(亿元|万元|元)\s*[）)]\s*[：:]?\s*$', text[:start])
    return label_unit.group(1) if label_unit else None


def value_fit_issue(text: str, start: int, end: int, source_value: str) -> str | None:
    """Return a local, explainable incompatibility without changing a template.

    A date range cannot be one component before a printed year/month unit.
    It may be a duration count or calendar component; neither accepts a range.
    """
    right = text[end:].lstrip()
    value = str(source_value).strip()
    if right.startswith(("年", "个月", "月")) and _DATE_RANGE.search(value):
        unit = "年" if right.startswith("年") else ("个月" if right.startswith("个月") else "月")
        return "模板此处固定印有“%s”，需要期限数量；来源值是日期区间，不能直接填写" % unit
    amount_unit = _printed_amount_unit(text, start, end)
    if amount_unit and not _AMOUNT_VALUE.fullmatch(value):
        return "模板此处固定印有人民币金额单位“%s”，来源值不是可识别的人民币数字金额；不能进行跨币种换算" % amount_unit
    return None


def unit_wrapper_span(text: str) -> tuple[int, int] | None:
    """Find one blank amount/unit slot in a compact Excel value cell.

    This is intentionally stricter than general blank discovery: a neighbouring
    label is needed by callers, and the blank must sit directly before a known
    printed unit. Multiple blanks are ambiguous and are left for review.
    """
    if not isinstance(text, str):
        return None
    candidates = []
    for match in _UNIT_BLANK_RE.finditer(text):
        right = text[match.end():].lstrip()
        if any(right.startswith(unit) for unit in _VALUE_UNITS):
            candidates.append(match.span())
    if len(candidates) != 1 or len(list(_PLACEHOLDER_MARK_RE.finditer(text))) != 1:
        return None
    return candidates[0]


def _amount_unit(value: str) -> str | None:
    match = _AMOUNT_VALUE.fullmatch(value)
    return match.group('unit') if match else None


def _converted_amount(value: str, target_unit: str) -> str:
    """Return a source amount in the template's printed unit, without float math."""
    match = _AMOUNT_VALUE.fullmatch(value.strip())
    if not match:
        raise OfficeKitError(
            f"模板此处固定印有人民币金额单位“{target_unit}”，来源值不是可识别的人民币数字金额；不能进行跨币种换算")
    source_unit = match.group("unit")
    number = re.sub(r"[ ,，\s]", "", value.strip())
    number = re.sub(r"^人民币", "", number).strip()
    number = re.sub(r"(?:亿元|万元|元|亿|万)$", "", number).strip()
    try:
        with localcontext() as context:
            context.prec = max(50, len(number) + 20)
            amount = Decimal(number)
            if source_unit:
                amount = amount * _AMOUNT_FACTORS[source_unit] / _AMOUNT_FACTORS[target_unit]
    except InvalidOperation as exc:
        raise OfficeKitError("来源金额无法精确解析，不能按模板单位换算") from exc
    rendered = format(amount.normalize(), "f")
    return "0" if rendered in ("-0", "") else rendered


def fit_value(text: str, start: int, end: int, source_value: str) -> str:
    issue = value_fit_issue(text, start, end, source_value)
    if issue:
        raise OfficeKitError(issue)
    value = str(source_value)
    left = text[:start].rstrip()
    right = text[end:].lstrip()

    printed_amount_unit = _printed_amount_unit(text, start, end)
    if printed_amount_unit:
        value = _converted_amount(value, printed_amount_unit)
    if left.endswith("人民币") and value.strip().startswith("人民币"):
        value = value.strip()[len("人民币"):].lstrip()

    # Keep a contract-number prefix already printed by the template. Match
    # literal text, not a bank/year-specific naming convention.
    number_prefix = re.search(r'(?:合同编号(?:为)?|编号为)\s*[：:]?\s*([^\n，。；;（）()【】]+)$', left)
    if number_prefix and right.startswith('号'):
        prefix = number_prefix.group(1).strip()
        if prefix and value.strip().startswith(prefix):
            value = value.strip()[len(prefix):].lstrip()

    # A printed 元 can complete 800万 + 元. It cannot turn 800美元 into a
    # renminbi amount. Require a plain numeric amount before this suffix;
    # complex or foreign-currency expressions need a human mapping.
    if right.startswith("元"):
        amount_before_printed_unit = value[:-1] if value.endswith("元") else value
        if not _YUAN_AMOUNT.fullmatch(amount_before_printed_unit.strip()):
            raise OfficeKitError("模板已印“元”，源值不是可直接拼接的人民币数字金额")

    if left.endswith("【") and right.startswith("】"):
        if value.startswith("【"):
            value = value[1:]
        if value.endswith("】号"):
            value = value[:-2]
        elif value.endswith("】"):
            value = value[:-1]

    # The rule may capture only the blank before a fixed suffix.  Prefer the
    # longest literal overlap; e.g. keep the template's "有限公司" rather than
    # writing a second copy of it.
    for ending in _PRINTED_ENDINGS:
        if right.startswith(ending) and value.endswith(ending):
            value = value[:-len(ending)]
            break

    # A rule capturing the inside of fixed quotation marks can also receive a
    # complete value.  Strip the marks only when the template supplies both.
    for opening, closing in (("（", "）"), ("(", ")"), ("《", "》")):
        if left.endswith(opening) and right.startswith(closing) \
                and value.startswith(opening) and value.endswith(closing):
            value = value[1:-1]
            break

    if not value and str(source_value).strip():
        raise OfficeKitError("值与模板固定文字完全重叠，无法判断是否应留空")
    return value
