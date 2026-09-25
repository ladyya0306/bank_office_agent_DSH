"""Narrow synthetic checks for target label extraction only."""
from pathlib import Path
import sys
import unittest
from tempfile import TemporaryDirectory
from zipfile import ZipFile

from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Border, Side

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from office_kit.harness import extract_labels, current_docx_target_errors, current_xlsx_gaps


def test_xlsx_scans_sheets_and_requires_formatted_adjacent_cell():
    with TemporaryDirectory() as tmp:
        path = Path(tmp) / "target.xlsx"
        book = Workbook()
        book.active["A1"] = "联系电话"
        book.active["B1"] = ""
        sheet = book.create_sheet("填写页")
        sheet["A1"] = "联系电话"
        sheet["B1"].border = Border(bottom=Side(style="thin"))
        book.save(path)
        labels = extract_labels(path, {"联系电话"})
        assert [(x["target"]["sheet"], x["target"]["cell"]) for x in labels] == [("填写页", "B1")]


def test_xlsx_merged_label_starts_after_merge():
    with TemporaryDirectory() as tmp:
        path = Path(tmp) / "target.xlsx"
        book = Workbook()
        sheet = book.active
        sheet.merge_cells("A1:C1")
        sheet["A1"] = "联系电话"
        sheet["D1"].border = Border(bottom=Side(style="thin"))
        book.save(path)
        labels = extract_labels(path, {"联系电话"})
        assert labels[0]["target"]["cell"] == "D1"


def test_word_table_merged_label_uses_one_tc():
    with TemporaryDirectory() as tmp:
        path = Path(tmp) / "target.docx"
        doc = Document()
        table = doc.add_table(rows=1, cols=3)
        table.cell(0, 0).merge(table.cell(0, 1))
        table.cell(0, 0).text = "联系电话"
        doc.save(path)
        labels = extract_labels(path, {"联系电话"})
        assert len(labels) == 1
        assert labels[0]["target"]["col"] == 2


def test_word_header_reports_existing_part_without_materializing_more_parts():
    with TemporaryDirectory() as tmp:
        path = Path(tmp) / "target.docx"
        doc = Document()
        doc.sections[0].header.paragraphs[0].text = "联系电话："
        doc.save(path)
        before = sorted(n for n in ZipFile(path).namelist() if "header" in n)
        labels = extract_labels(path, {"联系电话"})
        after = sorted(n for n in ZipFile(path).namelist() if "header" in n)
        assert before == after
        assert labels[0]["target"]["part"] == "word/header1.xml"


def test_numbered_label_keeps_anchor_but_cleans_field_name():
    with TemporaryDirectory() as tmp:
        path = Path(tmp) / "target.xlsx"
        book = Workbook()
        book.active["A1"] = "三、联系电话："
        book.save(path)
        labels = extract_labels(path, {"联系电话"})
        assert labels[0]["text"] == "联系电话"
        assert labels[0]["target"]["anchor"] == "三、联系电话："


def test_xlsx_formula_is_not_an_adjacent_target():
    with TemporaryDirectory() as tmp:
        path = Path(tmp) / "target.xlsx"
        book = Workbook()
        sheet = book.active
        sheet["A1"] = "联系电话"
        sheet["B1"] = "=1+2"
        sheet["B1"].border = Border(bottom=Side(style="thin"))
        book.save(path)
        assert extract_labels(path, {"联系电话"}) == []


def test_word_second_cell_is_not_covered_by_first_rule():
    with TemporaryDirectory() as tmp:
        path = Path(tmp) / 'target.docx'
        doc = Document()
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = '联系电话'
        table.cell(1, 0).text = '联系电话'
        doc.save(path)
        plan = {'rows': [{'kind': 'slot', 'template': path.name, 'field': '联系电话',
                          'target': {'kind': 'cell', 'table': 0, 'row': 0, 'col': 1}}]}
        errors = current_docx_target_errors([path], plan)
        assert any(e.get('target', {}).get('row') == 1 for e in errors), errors


def test_excel_second_cell_is_not_covered_by_first_rule():
    with TemporaryDirectory() as tmp:
        path = Path(tmp) / 'target.xlsx'
        book = Workbook()
        sheet = book.active
        for row in (1, 2):
            sheet.cell(row, 1).value = '联系电话'
            sheet.cell(row, 2).border = Border(bottom=Side(style='thin'))
        book.save(path)
        class Facts:
            def facts_by_key(self):
                return {'联系电话': {'value': '13800001234'}}
        plan = {'rows': [{'kind': 'slot', 'template': path.name, 'field': '联系电话',
                          'target': {'kind': 'xlsx_cell', 'sheet': sheet.title, 'cell': 'B1'}}]}
        gaps = current_xlsx_gaps(Facts(), [path], plan)
        assert len(gaps) == 1 and gaps[0]['target']['cell'] == 'B2', gaps


if __name__ == '__main__':
    suite = unittest.TestSuite(unittest.FunctionTestCase(func) for name, func in list(globals().items())
                               if name.startswith('test_') and callable(func))
    sys.exit(0 if unittest.TextTestRunner(verbosity=2).run(suite).wasSuccessful() else 1)
