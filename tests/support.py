"""Helpers to build small .xlsx files inside tests."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

DEFAULT_HEADERS = ("日期", "型号", "供应商", "备注")


def build_workbook(path: Path, rows, headers=DEFAULT_HEADERS, sheet_name="示例供应商A") -> Path:
    book = Workbook()
    sheet = book.active
    sheet.title = sheet_name
    if headers:
        sheet.append(list(headers))
    for row in rows:
        sheet.append(list(row))
    path.parent.mkdir(parents=True, exist_ok=True)
    book.save(path)
    return path


def sample_rows() -> list[list[object]]:
    return [
        ["2026-04-13", "A-0001", "供应商甲 0.13未", "现货"],
        ["2026-04-14", "A-0002", "供应商乙 0.14未", "订货"],
    ]
