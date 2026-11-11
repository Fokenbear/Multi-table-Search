"""Synthetic boundary tests for the Excel parser.

Real company workbooks are never committed; every fixture here is generated at
test time inside pytest's temporary directory.
"""

from __future__ import annotations

import hashlib
import re
import zipfile
from datetime import datetime
from pathlib import Path

import openpyxl
import pytest
from openpyxl import Workbook
from openpyxl.styles import PatternFill

from app.contracts import ContractError, ErrorCode, ExcelParser
from app.excel_parser import WorkbookParser
from app.excel_parser.ooxml import count_empty_fills, repair_empty_fills
from app.excel_parser.values import classify_value, format_number


def test_workbook_parser_implements_contract_protocol() -> None:
    assert isinstance(WorkbookParser(), ExcelParser)


def _save(wb: Workbook, path: Path) -> Path:
    wb.save(path)
    return path


def _parse(path: Path):
    return WorkbookParser().parse_workbook(str(path))


def _add_cached_formula_value(path: Path, address: str, cached: str) -> None:
    """Inject a cached ``<v>`` value into a formula written by openpyxl."""

    pattern = re.compile(
        rf'(<c r="{re.escape(address)}"[^>]*>)(<f>[^<]*</f>)(<v\s*/>)?(</c>)'
    )
    with zipfile.ZipFile(path) as source:
        entries = {info.filename: source.read(info.filename) for info in source.infolist()}
        infos = {info.filename: info for info in source.infolist()}
    sheet_name = "xl/worksheets/sheet1.xml"
    text = entries[sheet_name].decode("utf-8")

    def repl(match: re.Match[str]) -> str:
        return match.group(1) + match.group(2) + f"<v>{cached}</v>" + match.group(4)

    patched, count = pattern.subn(repl, text)
    assert count == 1, f"formula cell {address} not found"
    entries[sheet_name] = patched.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as target:
        for name, data in entries.items():
            target.writestr(infos[name], data)


def _inject_empty_fills(path: Path) -> None:
    with zipfile.ZipFile(path) as source:
        entries = {info.filename: source.read(info.filename) for info in source.infolist()}
        infos = {info.filename: info for info in source.infolist()}
    styles = entries["xl/styles.xml"].decode("utf-8")
    assert "</fills>" in styles
    styles = styles.replace("</fills>", "<fill/><fill></fill></fills>", 1)
    entries["xl/styles.xml"] = styles.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as target:
        for name, data in entries.items():
            target.writestr(infos[name], data)


def test_header_first_row_and_complete_row(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "供应商A"
    ws.append(["日期", "型号", "包装量", "供应商"])
    ws.append(["2026-04-13", "A-0001", 10000, "0.13未 25+ 供应商甲"])
    ws.append(["2026-04-14", "A-0002", None, "0.14未 25+ 供应商乙"])
    path = _save(wb, tmp_path / "header_first.xlsx")

    workbook = _parse(path)
    sheet = workbook.sheets[0]
    assert sheet.warnings == ()
    table = sheet.tables[0]
    assert table.header_rows == (1,)
    assert table.table_name == "供应商A"
    assert table.min_row == 1 and table.max_row == 3
    assert len(table.rows) == 2
    second = table.rows[1]
    assert [cell.header for cell in second.cells] == ["日期", "型号", "包装量", "供应商"]
    assert second.cells[2].raw_value is None
    assert second.cells[2].display_value == ""
    assert second.cells[2].data_type == "blank"
    assert second.matched_cell_addresses == ()


def test_header_on_second_row(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws["A2"] = "料号"
    ws["B2"] = "供应商"
    ws["A3"] = "ABC-1"
    ws["B3"] = "0.5未"
    path = _save(wb, tmp_path / "header_second.xlsx")

    table = _parse(path).sheets[0].tables[0]
    assert table.header_rows == (2,)
    assert table.min_row == 2
    assert table.rows[0].row_number == 3
    assert [cell.header for cell in table.rows[0].cells] == ["料号", "供应商"]


def test_merged_title_and_table_name(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.merge_cells("A1:D1")
    ws["A1"] = "供应商报价汇总表"
    ws.append(["日期", "型号", "供应商", "备注"])
    ws.append(["4.13", "A-0001", "0.13未", "优先"])
    path = _save(wb, tmp_path / "title.xlsx")

    table = _parse(path).sheets[0].tables[0]
    assert table.table_name == "供应商报价汇总表"
    assert table.header_rows == (2,)
    assert table.min_row == 1
    assert table.max_row == 3
    assert "A1:D1" in table.merged_ranges
    assert table.rows[0].row_number == 3


def test_duplicate_headers_are_not_merged(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(["日期", "型号", "供应商", "供应商", "供应商"])
    ws.append(["4.13", "A-1", "0.13未 供应商甲", "0.14未 供应商乙", "0.15未 供应商丙"])
    path = _save(wb, tmp_path / "duplicate_headers.xlsx")

    row = _parse(path).sheets[0].tables[0].rows[0]
    supplier_cells = [cell for cell in row.cells if cell.header == "供应商"]
    assert len(supplier_cells) == 3
    assert [cell.column_letter for cell in supplier_cells] == ["C", "D", "E"]
    assert [cell.raw_value for cell in supplier_cells] == [
        "0.13未 供应商甲",
        "0.14未 供应商乙",
        "0.15未 供应商丙",
    ]


def test_blank_header_keeps_data_column(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(["日期", None, "型号"])
    ws.append(["4.13", "未命名值", "A-1"])
    path = _save(wb, tmp_path / "blank_header.xlsx")

    table = _parse(path).sheets[0].tables[0]
    row = table.rows[0]
    assert [cell.header for cell in row.cells] == ["日期", None, "型号"]
    assert row.cells[1].raw_value == "未命名值"
    assert row.cells[1].is_supplementary is False


def test_no_header_degrades_with_warning(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append([datetime(2026, 4, 13), "A-0001", "0.13未 供应商甲"])
    ws.append([datetime(2026, 4, 14), "A-0002", "0.14未 供应商乙"])
    path = _save(wb, tmp_path / "no_header.xlsx")

    sheet = _parse(path).sheets[0]
    assert [warning.code for warning in sheet.warnings] == ["HEADER_NOT_FOUND"]
    table = sheet.tables[0]
    assert table.header_rows == ()
    assert table.min_row == 1
    assert len(table.rows) == 2
    assert all(cell.header is None for cell in table.rows[0].cells)


def test_vertical_merge_propagates_value_and_origin(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(["日期", "询货人", "型号"])
    ws["A2"] = "2026-04-13"
    ws["B2"] = "示例用户"
    ws["C2"] = "P-1"
    ws["C3"] = "P-2"
    ws["C4"] = "P-3"
    ws.merge_cells("A2:A4")
    ws.merge_cells("B2:B3")
    path = _save(wb, tmp_path / "vertical_merge.xlsx")

    table = _parse(path).sheets[0].tables[0]
    rows = {row.row_number: row for row in table.rows}
    assert rows[2].cells[0].raw_value == "2026-04-13"
    assert rows[3].cells[0].raw_value == "2026-04-13"
    assert rows[4].cells[0].raw_value == "2026-04-13"
    assert rows[3].cells[0].origin_address == "A2"
    assert rows[4].cells[0].origin_address == "A2"
    assert rows[2].cells[1].origin_address is None
    assert rows[3].cells[1].origin_address == "B2"
    assert rows[4].cells[1].raw_value is None
    assert set(table.merged_ranges) == {"A2:A4", "B2:B3"}


def test_horizontal_merge_is_not_duplicated(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(["日期", "供应商", "供应商", "备注"])
    ws["A2"] = "4.13"
    ws["B2"] = "0.13未 供应商甲"
    ws.merge_cells("B2:C2")
    ws["D2"] = "价格有优势"
    path = _save(wb, tmp_path / "horizontal_merge.xlsx")

    row = _parse(path).sheets[0].tables[0].rows[0]
    cells = {cell.column_letter: cell for cell in row.cells}
    assert cells["B"].raw_value == "0.13未 供应商甲"
    assert cells["C"].raw_value is None
    assert cells["C"].display_value == ""


def test_right_side_note_is_supplementary(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(["日期", "型号", "供应商"])
    ws.append(["4.13", "A-1", "0.13未"])
    ws["D2"] = "价格有优势"
    ws["E2"] = "可议价"
    path = _save(wb, tmp_path / "right_note.xlsx")

    row = _parse(path).sheets[0].tables[0].rows[0]
    cells = {cell.column_letter: cell for cell in row.cells}
    assert cells["D"].header is None
    assert cells["D"].is_supplementary is True
    assert cells["D"].raw_value == "价格有优势"
    assert cells["E"].is_supplementary is True
    assert cells["A"].is_supplementary is False


def test_vertical_merged_note_propagates(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(["日期", "型号", "供应商"])
    ws["A2"] = "4.13"
    ws["B2"] = "A-1"
    ws["C2"] = "0.13未"
    ws["D2"] = "整批可议价"
    ws["A3"] = "4.14"
    ws["B3"] = "A-2"
    ws["C3"] = "0.14未"
    ws.merge_cells("D2:D3")
    path = _save(wb, tmp_path / "merged_note.xlsx")

    table = _parse(path).sheets[0].tables[0]
    rows = {row.row_number: row for row in table.rows}
    for row_number in (2, 3):
        note = next(cell for cell in rows[row_number].cells if cell.column_letter == "D")
        assert note.raw_value == "整批可议价"
        assert note.is_supplementary is True
        assert note.origin_address == ("D2" if row_number == 3 else None)
    assert rows[3].cells[-1].address == "D3"


def test_hidden_rows_and_columns_are_parsed(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(["日期", "型号", "供应商"])
    ws.append(["4.13", "A-1", "0.13未"])
    ws.append(["4.14", "A-2", "0.14未"])
    ws.row_dimensions[3].hidden = True
    ws.column_dimensions["B"].hidden = True
    path = _save(wb, tmp_path / "hidden.xlsx")

    table = _parse(path).sheets[0].tables[0]
    assert [row.row_number for row in table.rows] == [2, 3]
    hidden_row = table.rows[1]
    assert hidden_row.cells[1].raw_value == "A-2"
    assert hidden_row.cells[1].header == "型号"


def test_formula_cached_and_missing_cache(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(["A", "B", "C", "合计", "无缓存"])
    ws.append([1, 2, 3, "=SUM(A2:C2)", "=A2+B2"])
    path = _save(wb, tmp_path / "formula.xlsx")
    _add_cached_formula_value(path, "D2", "6")

    sheet = _parse(path).sheets[0]
    table = sheet.tables[0]
    row = table.rows[0]
    cached = row.cells[3]
    assert cached.raw_value == "=SUM(A2:C2)"
    assert cached.display_value == "6"
    assert cached.data_type == "formula"
    assert cached.formula_has_cached_value is True
    missing = row.cells[4]
    assert missing.raw_value == "=A2+B2"
    assert missing.display_value == "=A2+B2"
    assert missing.formula_has_cached_value is False
    assert "FORMULA_WITHOUT_CACHED_VALUE" in [warning.code for warning in sheet.warnings]


def test_dates_numbers_and_leading_zeros(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(["日期", "零填充", "文本料号", "零值", "布尔"])
    ws["A2"] = datetime(2026, 5, 14)
    ws["A2"].number_format = 'm"月"d"日";@'
    ws["B2"] = 7
    ws["B2"].number_format = "000000"
    ws["C2"] = "007"
    ws["C2"].number_format = "@"
    ws["D2"] = 0
    ws["D2"].number_format = "0.00"
    ws["E2"] = False
    path = _save(wb, tmp_path / "formats.xlsx")

    row = _parse(path).sheets[0].tables[0].rows[0]
    cells = {cell.column_letter: cell for cell in row.cells}
    assert cells["A"].display_value == "5月14日"
    assert cells["A"].data_type in {"date", "datetime"}
    assert cells["B"].display_value == "000007"
    assert cells["C"].display_value == "007"
    assert cells["D"].display_value == "0.00"
    assert cells["D"].data_type == "number"
    assert cells["E"].display_value == "FALSE"
    assert cells["E"].data_type == "boolean"


def test_multiple_independent_regions(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(["日期", "型号", "供应商"])
    ws.append(["4.13", "A-1", "0.13未"])
    ws.append(["4.14", "A-2", "0.14未"])
    ws.append([None, None, None])
    ws.append([None, None, None])
    ws.append(["日期", "型号", "供应商"])
    ws.append(["5.1", "B-1", "0.2未"])
    path = _save(wb, tmp_path / "regions.xlsx")

    sheet = _parse(path).sheets[0]
    assert len(sheet.tables) == 2
    first, second = sheet.tables
    assert first.table_index == 1 and second.table_index == 2
    assert first.header_rows == (1,) and second.header_rows == (6,)
    assert [row.row_number for row in first.rows] == [2, 3]
    assert [row.row_number for row in second.rows] == [7]
    assert all(row.table_index == 2 for row in second.rows)


def test_side_by_side_regions(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    for col, value in zip("ABCEFG", ["日期", "型号", "供应商", "料号", "数量", "价格"]):
        ws[f"{col}1"] = value
    for col, value in zip("ABCEFG", ["4.13", "A-1", "0.13未", "B-1", 5, "0.2未"]):
        ws[f"{col}2"] = value
    path = _save(wb, tmp_path / "side_by_side.xlsx")

    sheet = _parse(path).sheets[0]
    assert len(sheet.tables) == 2
    left, right = sheet.tables
    assert left.min_column == 1 and left.max_column == 3
    assert right.min_column == 5 and right.max_column == 7
    assert [cell.header for cell in left.rows[0].cells] == ["日期", "型号", "供应商"]
    assert [cell.header for cell in right.rows[0].cells] == ["料号", "数量", "价格"]
    assert left.rows[0].cells[0].raw_value == "4.13"
    assert right.rows[0].cells[0].raw_value == "B-1"


def test_header_only_sheet(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(["日期", "型号", "供应商"])
    path = _save(wb, tmp_path / "header_only.xlsx")

    sheet = _parse(path).sheets[0]
    assert len(sheet.tables) == 1
    assert sheet.tables[0].rows == ()
    assert sheet.tables[0].header_rows == (1,)
    assert sheet.tables[0].max_row == 1


def test_empty_sheet_and_empty_workbook(tmp_path: Path) -> None:
    wb = Workbook()
    wb.active.title = "空表"
    path = _save(wb, tmp_path / "empty.xlsx")

    workbook = _parse(path)
    assert workbook.sheets[0].tables == ()
    assert "EMPTY_SHEET" in [warning.code for warning in workbook.sheets[0].warnings]
    assert "EMPTY_WORKBOOK" in [warning.code for warning in workbook.warnings]


def test_empty_fill_repair_keeps_original_unchanged(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(["日期", "型号", "供应商"])
    ws.append(["4.13", "A-1", "0.13未"])
    ws["A2"].fill = PatternFill(fill_type="solid", fgColor="FFFF00")
    path = _save(wb, tmp_path / "empty_fill.xlsx")
    _inject_empty_fills(path)
    before = hashlib.sha256(path.read_bytes()).hexdigest()

    workbook = _parse(path)
    after = hashlib.sha256(path.read_bytes()).hexdigest()

    assert before == after
    assert "STYLE_REPAIRED" in [warning.code for warning in workbook.warnings]
    assert workbook.sheets[0].tables[0].rows[0].cells[1].raw_value == "A-1"


def test_style_repair_function_preserves_other_xml() -> None:
    styles = (
        b'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b'<fonts count="1"><font><sz val="11"/></font></fonts>'
        b'<fills count="3"><fill><patternFill patternType="none"/></fill>'
        b'<fill/><fill></fill></fills>'
        b'<numFmts count="1"><numFmt numFmtId="164" formatCode="0000"/></numFmts>'
        b"</styleSheet>"
    )
    repaired, count = repair_empty_fills(styles)
    assert count == 2
    assert b"<fill/>" not in repaired
    assert b'<patternFill patternType="none"/>' in repaired
    assert b'formatCode="0000"' in repaired
    assert count_empty_fills(repaired) == 0


def test_corrupt_styles_raise_parse_failed(tmp_path: Path) -> None:
    wb = Workbook()
    wb.active["A1"] = "x"
    path = _save(wb, tmp_path / "bad_styles.xlsx")
    with zipfile.ZipFile(path) as source:
        entries = {info.filename: source.read(info.filename) for info in source.infolist()}
        infos = {info.filename: info for info in source.infolist()}
    entries["xl/styles.xml"] = b"<styleSheet><broken"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as target:
        for name, data in entries.items():
            target.writestr(infos[name], data)

    with pytest.raises(ContractError) as error:
        _parse(path)
    assert error.value.code == ErrorCode.PARSE_FAILED


@pytest.mark.parametrize("slot", [0, 4, True, "1"])
def test_invalid_workbook_argument_rejected(tmp_path: Path, slot) -> None:
    path = tmp_path / "x.xlsx"
    path.write_bytes(b"not a workbook")
    with pytest.raises(ContractError) as error:
        WorkbookParser().parse_workbook(slot)  # type: ignore[arg-type]
    assert error.value.code == ErrorCode.INVALID_ARGUMENT


def test_invalid_inputs(tmp_path: Path) -> None:
    parser = WorkbookParser()
    with pytest.raises(ContractError) as error:
        parser.parse_workbook("relative.xlsx")
    assert error.value.code == ErrorCode.INVALID_ARGUMENT

    with pytest.raises(ContractError) as error:
        parser.parse_workbook(str(tmp_path / "missing.xlsx"))
    assert error.value.code == ErrorCode.FILE_NOT_FOUND

    old = tmp_path / "old.xls"
    old.write_bytes(b"x")
    with pytest.raises(ContractError) as error:
        parser.parse_workbook(str(old))
    assert error.value.code == ErrorCode.UNSUPPORTED_FORMAT

    broken = tmp_path / "broken.xlsx"
    broken.write_bytes(b"not a zip")
    with pytest.raises(ContractError) as error:
        parser.parse_workbook(str(broken))
    assert error.value.code == ErrorCode.PARSE_FAILED


def test_multiple_hits_in_one_row_stay_one_complete_row(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(["日期", "型号", "供应商", "备注"])
    ws.append(["4.13", "A-0001", "A-0001 报价", "A-0001 有货"])
    path = _save(wb, tmp_path / "multi_hit.xlsx")

    table = _parse(path).sheets[0].tables[0]
    assert len(table.rows) == 1
    row = table.rows[0]
    assert row.matched_cell_addresses == ()
    assert row.cells[1].raw_value == "A-0001"
    assert row.cells[2].raw_value == "A-0001 报价"
    assert row.cells[3].raw_value == "A-0001 有货"


def test_multi_row_header_with_merged_label(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "供应商信息"
    ws.merge_cells("A1:B1")
    ws["C1"] = "型号"
    ws["A2"] = "名称"
    ws["B2"] = "价格"
    ws["A3"] = "供应商甲"
    ws["B3"] = "0.13未"
    ws["C3"] = "A-1"
    path = _save(wb, tmp_path / "multi_header.xlsx")

    table = _parse(path).sheets[0].tables[0]
    assert table.header_rows == (1, 2)
    row = table.rows[0]
    assert row.cells[0].header == "供应商信息\n名称"
    assert row.cells[1].header == "供应商信息\n价格"
    assert row.cells[2].header == "型号"


def test_format_helpers() -> None:
    assert format_number(7, "000000") == ("000007", False)
    assert format_number(1234567, "#,##0") == ("1,234,567", False)
    assert format_number(-1.234, '"￥"#,##0.00;[Red]"￥"-#,##0.00')[0] == "￥-1.23"
    assert classify_value(False, "General").display_value == "FALSE"
    assert classify_value(None, "General").data_type == "blank"
