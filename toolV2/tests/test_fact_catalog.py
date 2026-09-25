"""Synthetic tests for subject-qualified fact aliases."""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
from office_kit.fact_catalog import qualified_facts
from office_kit.store_v2 import StoreV2


class FactCatalogTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.store = StoreV2(Path(self.tmp.name) / "workflow.db")
        self.batch = "20260925-41"
        self.store.batch_scope = self.batch

    def tearDown(self) -> None:
        self.store.close()
        self.tmp.cleanup()

    def entity(self, name: str, *, personal: bool = False) -> int:
        return self.store.ensure_entity(name, personal=personal)[0]

    def fact(self, key: str, value: str, eid: int, provenance: str = "材料第1行") -> None:
        self.store.put_fact(self.batch, key, value, entity_id=eid, provenance=provenance)

    def test_explicit_roles_and_person_identity_are_qualified(self) -> None:
        borrower = self.entity("借款主体")
        guarantor_one = self.entity("保证主体一")
        guarantor_two = self.entity("保证主体二")
        # Historical migrated rows may classify these entities as legal_person;
        # the explicit representative name fact must still resolve them.
        borrower_rep = self.entity("借款代表")
        guarantor_rep = self.entity("保证代表")
        for eid, role, evidence in ((borrower, "借款人", "材料明确借款人"),
                                     (guarantor_one, "保证人", "材料明确保证人1"),
                                     (guarantor_two, "保证人", "材料明确保证人2")):
            self.store.set_role(eid, role, evidence=evidence, batch_no=self.batch)
        self.fact("地址", "借款地址", borrower)
        self.fact("联系电话", "保证电话一", guarantor_one)
        self.fact("地址", "保证地址二", guarantor_two)
        self.fact("借款人法定代表人", "借款代表", borrower)
        self.fact("保证人法定代表人", "保证代表", guarantor_one)
        self.fact("证件类型", "护照", borrower_rep)
        self.fact("证件号码", "REP-PASS", borrower_rep)
        self.fact("证件类型", "其他证件", guarantor_rep)
        self.fact("证件号码", "REP-OTHER", guarantor_rep)

        facts = qualified_facts(self.store)
        self.assertEqual("借款地址", facts["借款人地址"]["value"])
        self.assertEqual(borrower, facts["借款人地址"]["entity_id"])
        self.assertEqual("保证电话一", facts["保证人1联系电话"]["value"])
        self.assertEqual("保证地址二", facts["保证人2地址"]["value"])
        self.assertEqual(borrower_rep, facts["借款人法定代表人"]["entity_id"])
        self.assertEqual(borrower, facts["借款人法定代表人"]["_relation_owner_eid"])
        self.assertEqual("REP-PASS", facts["借款人法定代表人证件号码"]["value"])
        self.assertEqual("护照", facts["借款人法定代表人证件类型"]["value"])
        self.assertEqual("REP-OTHER", facts["保证人1法定代表人证件号码"]["value"])

    def test_unqualified_multiple_guarantors_are_ambiguous_and_not_numbered(self) -> None:
        first = self.entity("无编号保证一")
        second = self.entity("无编号保证二")
        for eid in (first, second):
            self.store.set_role(eid, "保证人", evidence="材料只写保证人", batch_no=self.batch)
            self.fact("联系电话", "同字段值" + str(eid), eid)
        facts = qualified_facts(self.store)
        self.assertNotIn("保证人1联系电话", facts)
        self.assertNotIn("保证人2联系电话", facts)
        self.assertTrue(facts["保证人联系电话"]["_ambiguous"])
        self.assertEqual({first, second},
                         {c["entity_id"] for c in facts["保证人联系电话"]["_candidates"]})

    def test_batch_scope_excludes_other_batch_and_keeps_raw_fact(self) -> None:
        current = self.entity("当前主体")
        other = self.entity("其他批次主体")
        self.store.set_role(current, "借款人", evidence="借款人", batch_no=self.batch)
        self.store.set_role(other, "借款人", evidence="借款人", batch_no="20260925-42")
        self.fact("地址", "当前地址", current)
        self.store.put_fact("20260925-42", "地址", "其他地址", entity_id=other)
        facts = qualified_facts(self.store)
        self.assertEqual("当前地址", facts["借款人地址"]["value"])
        self.assertNotIn("其他地址", {facts["借款人地址"]["value"]})
        self.assertEqual("当前地址", facts["地址"]["value"])


if __name__ == "__main__":
    unittest.main()
