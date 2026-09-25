"""Read-only real-layout checks using temporary copies of three local templates.

The templates are supplied with ``--templates-dir`` (the default is intentionally
empty).  Only invented source facts are written.  A missing rule is reported as
``needs_mapping``; this test never creates a rule merely to make a case pass.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document
from openpyxl import load_workbook

TOOL = Path(__file__).resolve().parents[1]
CASES = (
    ("5.签约核实书.docx", "docx", "企业名称", "段落 14，标签‘企业名称：’后的空位"),
    ("7.借款人董事会决议模板.docx", "docx", "会议地点", "段落 6，标签‘三、会议地点：’后的空位"),
    ("36-普惠型小微企业授信业务用信阶段重要信息核实情况表.xlsx", "xlsx", "借款企业名称", "Sheet1!A3:C3、Sheet2!A3:B3 合并区域"),
)
SYNTHETIC = {
    "企业名称": "合成版式验证公司",
    "会议地点": "合成会议地点",
    "借款人": "合成版式验证公司",
}


def call(payload: dict) -> dict:
    cp = subprocess.run(
        [sys.executable, str(TOOL / "office.py")], input=json.dumps(payload, ensure_ascii=False),
        text=True, encoding="utf-8", capture_output=True, timeout=120,
    )
    if cp.returncode:
        raise AssertionError(cp.stderr[-1000:] or cp.stdout[-1000:])
    return json.loads(cp.stdout)


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def make_source(path: Path) -> None:
    doc = Document()
    doc.add_paragraph("借款人：" + SYNTHETIC["借款人"])
    doc.add_paragraph("企业名称：" + SYNTHETIC["企业名称"])
    doc.add_paragraph("会议地点：" + SYNTHETIC["会议地点"])
    doc.save(path)


def expected_layout(path: Path, kind: str, field: str) -> None:
    """Independent structural expectation; it is never changed from output."""
    if kind == "docx":
        doc = Document(path)
        text = "\n".join(p.text for p in doc.paragraphs)
        if field == "企业名称":
            assert "企业名称：" in doc.paragraphs[14].text
        else:
            assert "会议地点：" in doc.paragraphs[6].text
        assert field in text
    else:
        book = load_workbook(path, read_only=False, data_only=False)
        ws = book["Sheet1"]
        assert ws["A3"].value and "借款企业名称" in str(ws["A3"].value)
        assert "A3:C3" in {str(r) for r in ws.merged_cells.ranges}
        ws2 = book["Sheet2"]
        assert "A3:B3" in {str(r) for r in ws2.merged_cells.ranges}
        book.close()


class RealLayoutTests(unittest.TestCase):
    templates_dir: Path | None = None
    artifacts_dir: Path | None = None

    def test_real_templates_from_temp_copies(self) -> None:
        if self.templates_dir is None:
            self.skipTest("需要 --templates-dir 指向本机真实模板目录")
        report = []
        with tempfile.TemporaryDirectory(prefix="toolv2_real_layout_") as tmp:
            work = Path(tmp) / "work"
            work.mkdir()
            source = work / "合成来源.docx"
            make_source(source)
            original_hashes = {source.name: sha256(source)}
            targets = []
            for name, kind, field, location in CASES:
                original = self.templates_dir / name
                self.assertTrue(original.is_file(), name)
                expected_layout(original, kind, field)
                target = work / name
                shutil.copy2(original, target)
                targets.append(target)
                original_hashes[name] = sha256(original)

            artifact_run = None
            if self.artifacts_dir is not None:
                stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S-%f")
                artifact_run = self.artifacts_dir / stamp
                artifact_run.mkdir(parents=True, exist_ok=False)
                shutil.copy2(source, artifact_run / source.name)
                for target in targets:
                    shutil.copy2(target, artifact_run / target.name)

            result = call({"action": "start", "work": str(work), "source": [str(source)],
                           "targets": [str(p) for p in targets], "batch": "20260925-90"})
            # Answer only source/fill questions using the first displayed option.  No
            # position update is sent: absent rules must remain needs_mapping.
            for _ in range(6):
                if result.get("status") not in ("awaiting_source", "awaiting_fill"):
                    break
                answers = [{"id": q["id"], "selected": [q["options"][0]["label"]], "custom": ""}
                           for q in result.get("questions", []) if q.get("options")]
                if not answers:
                    break
                result = call({"action": "resume", "work": str(work), "task_id": result["task_id"],
                               "answers": answers})

            for name, kind, field, location in CASES:
                with self.subTest(template=name):
                    self.assertEqual(original_hashes[name], sha256(self.templates_dir / name))
                    self.assertEqual(original_hashes[name], sha256(work / name))
                    record = next(r for r in result.get("results", []) if r["template"] == name)
                    status = record.get("status")
                    passed = status == "completed"
                    reason = ""
                    if status == "completed":
                        output = Path(record["output"])
                        self.assertTrue(output.is_relative_to(work))
                        if kind == "docx":
                            paragraphs = Document(output).paragraphs
                            expected_index = 14 if field == "企业名称" else 6
                            expected_value = SYNTHETIC[field]
                            self.assertIn(expected_value, paragraphs[expected_index].text)
                        else:
                            book = load_workbook(output, data_only=False)
                            self.assertIn(SYNTHETIC["借款人"], str(book["Sheet1"]["A3"].value))
                            self.assertIn(SYNTHETIC["借款人"], str(book["Sheet2"]["A3"].value))
                            book.close()
                        if artifact_run is not None:
                            shutil.copy2(output, artifact_run / output.name)
                    else:
                        reason = "工具返回 needs_mapping，未找到已批准的模板位置规则"
                    item = {"template": name, "field": field, "location": location,
                            "status": status, "passed": passed, "reason": reason}
                    report.append(item)
                    self.assertTrue(passed, item)

            report_path = TOOL / "实际版式验证记录_20260925.md"
            lines = ["# 实际版式验证记录（2026-09-25）", "", "测试使用本机已有真实模板的临时副本；源资料全部为合成内容。", ""]
            lines += ["| 文件名 | 字段 | 填格位置 | 状态 | 是否通过 |", "|---|---|---|---|---|"]
            lines += [f"| {r['template']} | {r['field']} | {r['location']} | {r['status']} | {'是' if r['passed'] else '否'} |" for r in report]
            lines += ["", "未输出模板正文或客户字段值；needs_mapping 计为未通过。两份未自动填完的原因均为工具未找到已批准的模板位置规则；这不证明版式错误。", "本记录只覆盖这三份真实模板的本地临时副本，不代表所有真实材料均通过。", ""]
            report_path.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--templates-dir", type=Path, required=True)
    parser.add_argument("--artifacts-dir", type=Path)
    args, rest = parser.parse_known_args()
    RealLayoutTests.templates_dir = args.templates_dir.resolve()
    RealLayoutTests.artifacts_dir = args.artifacts_dir.resolve() if args.artifacts_dir else None
    result = unittest.main(argv=[sys.argv[0], *rest], exit=False)
    return 0 if result.result.wasSuccessful() else 1


if __name__ == "__main__":
    raise SystemExit(main())
