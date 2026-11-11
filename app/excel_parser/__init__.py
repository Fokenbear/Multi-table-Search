"""Excel 解析模块：把 .xlsx 工作簿还原成可检索的行与单元格。

Public entry point::

    from app.excel_parser.parser import WorkbookParser

    parsed = WorkbookParser().parse_workbook(path)

The module is independent of Qt, SQLite, WPS and the search service.  It
imports the frozen DTOs from :mod:`app.contracts`.
"""

from app.excel_parser.parser import WorkbookParser

__all__ = ["WorkbookParser"]
