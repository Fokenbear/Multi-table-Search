from openpyxl import Workbook
from app.excel_parser import WorkbookParser


def test_inventory_header_after_title_is_not_indexed_as_a_quote(tmp_path):
    book = Workbook()
    book.active.append(["synthetic inventory title"])
    book.active.append(["型号", "品牌", "未税价", "含税价"])
    book.active.append(["PN-01", "synthetic", 1, 2])
    path = tmp_path / "inventory.xlsx"
    book.save(path)
    table = WorkbookParser().parse_workbook(str(path)).sheets[0].tables[0]
    assert table.header_rows == (2,)
    assert [row.row_number for row in table.rows] == [3]


def test_numeric_cell_with_text_format_retains_its_value(tmp_path):
    book = Workbook()
    book.active.append(["型号", "数量"])
    book.active.append([12345, 7])
    book.active["A2"].number_format = "@"
    path = tmp_path / "numeric_text_format.xlsx"
    book.save(path)
    cell = WorkbookParser().parse_workbook(str(path)).sheets[0].tables[0].rows[0].cells[0]
    assert cell.raw_value == 12345
    assert cell.display_value == "12345"
