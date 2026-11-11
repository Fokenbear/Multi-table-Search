"""Public interface checks: immutability, protocol conformance, versioning."""

from __future__ import annotations

import dataclasses
from datetime import datetime, timezone

import pytest

from app import contracts as c
from app.contracts import WorkbookInfo, WorkbookStatus


def test_interface_version_is_declared():
    assert c.INTERFACE_VERSION == "2.0"


def test_dtos_are_immutable():
    cell = c.ParsedCell(1, "A", "型号", "A-0001", "A-0001", "A2", False)
    with pytest.raises(dataclasses.FrozenInstanceError):
        cell.header = "x"  # type: ignore[misc]


def test_workbook_info_defaults():
    info = WorkbookInfo(workbook_id=7, status=WorkbookStatus.READY, original_filename="a.xlsx")
    assert info.sheet_count == 0
    assert info.row_count == 0
    assert info.imported_at is None


def test_search_result_carries_workbook_identity():
    result = c.SearchResult(
        workbook_id=3, workbook_name="a.xlsx", workbook_path="C:/x/a.xlsx",
        sheet_name="S", table_name="S", row_number=2, cells=(), matched_cell_addresses=("A2",),
        table_index=1, table_refs=(c.TableReference(1, "S"),),
    )
    assert result.workbook_id == 3
    assert result.table_refs[0].table_name == "S"


def test_services_implement_the_protocols(services):
    assert isinstance(services.data_manager, c.DataManager)
    assert isinstance(services.search_service, c.SearchService)
    assert isinstance(services.wps_opener, c.WpsOpener)
    assert isinstance(services.result_formatter, c.ResultFormatter)


def test_data_manager_exposes_no_fixed_slot_api(services):
    for name in ("list_slots", "replace_slot"):
        assert not hasattr(services.data_manager, name)
    for name in ("add_workbook", "replace_workbook", "remove_workbook", "list_workbooks"):
        assert callable(getattr(services.data_manager, name))


def test_import_summary_reports_failure_code():
    summary = c.ImportSummary(None, 0, 0, 0, 0.1, False,
                              error_code=c.ErrorCode.UNSUPPORTED_FORMAT, message="仅支持 .xlsx")
    assert summary.success is False
    assert summary.error_code is c.ErrorCode.UNSUPPORTED_FORMAT


def test_workbook_info_accepts_import_time():
    info = WorkbookInfo(
        workbook_id=1, status=WorkbookStatus.READY, original_filename="a.xlsx",
        imported_at=datetime(2026, 4, 13, tzinfo=timezone.utc), row_count=10,
    )
    assert info.imported_at.year == 2026
    assert info.row_count == 10
