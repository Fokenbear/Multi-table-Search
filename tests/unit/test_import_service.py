"""Import pipeline: unbounded add, replace, remove and failure containment."""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from app.contracts import ContractError, ErrorCode, WorkbookStatus
from tests.support import build_workbook, sample_rows


def test_add_workbook_indexes_and_keeps_a_managed_copy(manager, tmp_path):
    source = build_workbook(tmp_path / "a.xlsx", sample_rows())
    summary = manager.add_workbook(str(source))

    assert summary.success, summary.message
    assert summary.workbook_id == 1
    assert summary.imported_sheets == 1
    assert summary.imported_rows == 2

    info = manager.list_workbooks()[0]
    assert info.original_filename == "a.xlsx"
    assert info.sheet_count == 1
    assert info.row_count == 2
    assert Path(info.managed_path).is_file()
    assert source.is_file(), "原件不能被移动或修改"


def test_workbook_count_has_no_upper_bound(manager, tmp_path):
    ids = []
    for index in range(6):
        source = build_workbook(
            tmp_path / f"w{index}.xlsx",
            [["2026-04-13", f"A-000{index}", "供应商甲", ""]],
        )
        summary = manager.add_workbook(str(source))
        assert summary.success, summary.message
        ids.append(summary.workbook_id)

    assert ids == sorted(ids)
    assert len(manager.list_workbooks()) == 6


def test_duplicate_content_is_rejected(manager, tmp_path):
    source = build_workbook(tmp_path / "a.xlsx", sample_rows())
    assert manager.add_workbook(str(source)).success

    duplicate = tmp_path / "again.xlsx"
    shutil.copyfile(source, duplicate)
    summary = manager.add_workbook(str(duplicate))

    assert summary.success is False
    assert summary.error_code == ErrorCode.DUPLICATE_WORKBOOK
    assert len(manager.list_workbooks()) == 1


def test_replace_keeps_identity_and_drops_the_previous_copy(manager, tmp_path):
    first = build_workbook(tmp_path / "a.xlsx", sample_rows())
    workbook_id = manager.add_workbook(str(first)).workbook_id
    old_managed = Path(manager.list_workbooks()[0].managed_path)

    second = build_workbook(tmp_path / "b.xlsx", [["2026-04-13", "B-0001", "供应商乙", ""]])
    replaced = manager.replace_workbook(workbook_id, str(second))

    assert replaced.success, replaced.message
    assert replaced.workbook_id == workbook_id
    infos = manager.list_workbooks()
    assert len(infos) == 1
    assert infos[0].workbook_id == workbook_id
    assert infos[0].original_filename == "b.xlsx"
    assert infos[0].row_count == 1
    assert not old_managed.parent.exists()


def test_remove_deletes_index_and_copy(manager, services, tmp_path):
    source = build_workbook(tmp_path / "a.xlsx", sample_rows())
    workbook_id = manager.add_workbook(str(source)).workbook_id
    managed = Path(manager.list_workbooks()[0].managed_path)

    manager.remove_workbook(workbook_id)

    assert manager.list_workbooks() == ()
    assert not managed.parent.exists()
    assert services.search_service.search("A-0001").total_count == 0


def test_failed_import_keeps_the_previous_index(manager, services, tmp_path):
    source = build_workbook(tmp_path / "a.xlsx", sample_rows())
    workbook_id = manager.add_workbook(str(source)).workbook_id

    broken = tmp_path / "broken.xlsx"
    broken.write_bytes(b"definitely not a zip archive")
    summary = manager.replace_workbook(workbook_id, str(broken))

    assert summary.success is False
    infos = manager.list_workbooks()
    assert len(infos) == 1
    assert infos[0].status is WorkbookStatus.FAILED
    assert services.search_service.search("A-0001").total_count == 1, "旧索引必须仍然可查"


def test_relative_path_is_rejected(manager):
    with pytest.raises(ContractError) as error:
        manager.add_workbook("relative.xlsx")
    assert error.value.code == ErrorCode.INVALID_ARGUMENT


def test_unsupported_extension_is_rejected(manager, tmp_path):
    legacy = tmp_path / "old.xls"
    legacy.write_bytes(b"x")
    summary = manager.add_workbook(str(legacy))
    assert summary.error_code == ErrorCode.UNSUPPORTED_FORMAT


def test_cancellation_leaves_no_partial_index(manager, tmp_path):
    class ImmediateCancel:
        def is_cancelled(self) -> bool:
            return True

    source = build_workbook(tmp_path / "a.xlsx", sample_rows())
    summary = manager.add_workbook(str(source), cancellation=ImmediateCancel())

    assert summary.success is False
    assert summary.error_code == ErrorCode.CANCELLED
    assert manager.list_workbooks() == ()


def test_progress_reports_stages(manager, tmp_path):
    seen = []
    source = build_workbook(tmp_path / "a.xlsx", sample_rows())
    summary = manager.add_workbook(str(source), on_progress=seen.append)

    assert summary.success
    stages = [item.stage.value for item in seen]
    assert "copying" in stages
    assert "completed" in stages
    assert all(item.workbook_id is None for item in seen)
