"""Integration review regressions; all inputs are synthetic."""
from pathlib import Path
from zipfile import ZipFile

from openpyxl import Workbook
import pytest

from app.contracts import ContractError, ErrorCode
from app.excel_parser import WorkbookParser
from app.excel_parser.ooxml import scan_formulas


def test_text_only_data_has_no_guessed_header(tmp_path):
    w = Workbook()
    w.active.append(["PN-FIRST", "synthetic supplier"])
    w.active.append(["PN-SECOND", "synthetic note"])
    p = tmp_path / "text.xlsx"
    w.save(p)
    sheet = WorkbookParser().parse_workbook(str(p)).sheets[0]
    assert sheet.tables[0].header_rows == ()
    assert [r.row_number for r in sheet.tables[0].rows] == [1, 2]
    assert any(x.code == "HEADER_NOT_FOUND" for x in sheet.warnings)


def test_repeated_header_without_empty_rows_starts_new_region(tmp_path):
    w = Workbook()
    for row in (["型号", "供应商"], ["X-1", "supplier"],
                ["型号", "数量"], ["Y-2", 3]):
        w.active.append(row)
    p = tmp_path / "regions.xlsx"
    w.save(p)
    tables = WorkbookParser().parse_workbook(str(p)).sheets[0].tables
    assert [t.header_rows for t in tables] == [(1,), (3,)]
    assert [r.row_number for t in tables for r in t.rows] == [2, 4]


def test_shared_formula_preserves_relative_and_absolute_references(tmp_path):
    p = tmp_path / "shared.xlsx"
    xml = '''<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData><row r="2"><c r="C2"><f t="shared" si="0">A2+$B$1</f><v/></c></row>
    <row r="3"><c r="C3"><f t="shared" si="0"/><v/></c></row></sheetData></worksheet>'''
    with ZipFile(p, "w") as z:
        z.writestr("xl/worksheets/sheet1.xml", xml)
    with ZipFile(p) as z:
        formulas = scan_formulas(z, "xl/worksheets/sheet1.xml")
    assert formulas["C3"].text == "A3+$B$1"
    assert not formulas["C3"].has_cached_value


def test_excessive_merge_fails_explicitly_instead_of_truncating(tmp_path):
    w = Workbook()
    w.active.append(["型号", "备注"])
    w.active["A2"] = "X"
    w.active.merge_cells("A2:A20003")
    p = tmp_path / "large.xlsx"
    w.save(p)
    with pytest.raises(ContractError) as exc:
        WorkbookParser().parse_workbook(str(p))
    assert exc.value.code == ErrorCode.VALIDATION_FAILED
