"""Fit a source value into a template's already printed wrapper or unit.

Only remove literal text the template itself already supplies.  Never convert
units or invent missing parts of a contract number.
"""
from __future__ import annotations

import re

from .common import OfficeKitError


_PRINTED_ENDINGS = ("有限公司", "万元", "人民币", "公司", "元", "年", "号")
_YUAN_AMOUNT = re.compile(r"(?:人民币\s*)?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?\s*(?:万|亿)?")


def fit_value(text: str, start: int, end: int, source_value: str) -> str:
    value = str(source_value)
    left = text[:start].rstrip()
    right = text[end:].lstrip()

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
