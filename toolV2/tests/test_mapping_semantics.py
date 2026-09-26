"""Position meaning must veto stale shared-rule labels."""
from __future__ import annotations

import sys
import unittest
from pathlib import Path


TOOL = Path(__file__).resolve().parents[1]
if str(TOOL) not in sys.path:
    sys.path.insert(0, str(TOOL))

from workflow.mapping import semantic_compatible  # noqa: E402


def slot(text: str) -> dict:
    return {"target": {"expected_text": text}, "context": {}}


class MappingSemanticTests(unittest.TestCase):
    def test_credit_term_cannot_fill_board_meeting_date(self) -> None:
        self.assertFalse(semantic_compatible("授信期限", slot("二、会议时间：____年____月____日")))

    def test_contract_number_cannot_fill_product_or_count(self) -> None:
        self.assertFalse(semantic_compatible("最高额保证合同编号", slot("品种：____ 份数：____")))

    def test_amount_in_credit_sentence_remains_eligible(self) -> None:
        self.assertTrue(semantic_compatible("借款金额", slot("申请授信金额：____万元")))


if __name__ == "__main__":
    unittest.main()
