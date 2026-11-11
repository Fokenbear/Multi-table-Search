"""Test fixtures: a real database, real parser and a recording WPS opener."""

from __future__ import annotations

import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest

from app.contracts import ApplicationServices, OpenResult, OpenStatus
from app.core.database import Database
from app.core.import_service import LocalDataManager
from app.core.result_formatter import PlainTextResultFormatter
from app.core.search_service import SqliteSearchService
from app.excel_parser import WorkbookParser
from app.windows.windows_paths import AppPaths


class RecordingOpener:
    """Stands in for WPS: records the call and reports a canned outcome."""

    def __init__(self) -> None:
        self.calls: list[tuple[str, str, str]] = []
        self.result = OpenResult(OpenStatus.LOCATED, "", "", "", "located")

    def open_location(self, path: str, sheet_name: str, cell_address: str) -> OpenResult:
        self.calls.append((path, sheet_name, cell_address))
        return self.result


@pytest.fixture
def data_paths(tmp_path):
    return AppPaths.create(tmp_path / "data")


@pytest.fixture
def database(data_paths):
    return Database(data_paths)


@pytest.fixture
def manager(database):
    return LocalDataManager(database, WorkbookParser())


@pytest.fixture
def opener():
    return RecordingOpener()


@pytest.fixture
def services(database, manager, opener):
    return ApplicationServices(
        manager, SqliteSearchService(database), opener, PlainTextResultFormatter())
