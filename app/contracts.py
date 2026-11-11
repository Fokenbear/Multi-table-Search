"""Public data model and service interfaces.

Standard library only: no Qt, no openpyxl, no sqlite3 imports.  Everything that
crosses a module boundary is declared here so the parser, the storage layer and
the UI can be developed and tested independently.

DTOs are immutable tuples/dataclasses.  Type annotations are documentation, not
runtime validation: implementations validate external input themselves.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from enum import StrEnum
from typing import Callable, Literal, Protocol, TypeAlias, runtime_checkable

INTERFACE_VERSION = "2.0"

WorkbookId: TypeAlias = int
RawValue: TypeAlias = str | int | float | bool | date | datetime | time | timedelta | None
CellDataType: TypeAlias = Literal[
    "blank", "text", "number", "boolean", "date", "datetime", "time",
    "duration", "error", "formula",
]


class ErrorCode(StrEnum):
    INVALID_ARGUMENT = "invalid_argument"
    UNSUPPORTED_FORMAT = "unsupported_format"
    FILE_NOT_FOUND = "file_not_found"
    FILE_ACCESS = "file_access"
    PARSE_FAILED = "parse_failed"
    VALIDATION_FAILED = "validation_failed"
    STORAGE_FAILED = "storage_failed"
    WORKBOOK_NOT_FOUND = "workbook_not_found"
    DUPLICATE_WORKBOOK = "duplicate_workbook"
    BUSY = "busy"
    CANCELLED = "cancelled"
    WPS_UNAVAILABLE = "wps_unavailable"
    OPEN_FAILED = "open_failed"
    INTERNAL_ERROR = "internal_error"


class ContractError(Exception):
    """Expected service failure. ``message`` is safe to show in the desktop UI."""

    def __init__(self, code: ErrorCode, message: str, *, details: str = "") -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.details = details


@dataclass(frozen=True, slots=True)
class Diagnostic:
    code: str
    message: str
    sheet_name: str | None = None
    cell_address: str | None = None


@dataclass(frozen=True, slots=True)
class ParsedCell:
    column_index: int
    column_letter: str
    header: str | None
    raw_value: RawValue
    display_value: str
    address: str
    is_supplementary: bool
    data_type: CellDataType = "text"
    number_format: str = "General"
    origin_address: str | None = None
    formula_has_cached_value: bool | None = None


@dataclass(frozen=True, slots=True)
class ParsedRow:
    workbook_name: str
    sheet_name: str
    table_name: str
    row_number: int
    cells: tuple[ParsedCell, ...]
    table_index: int
    matched_cell_addresses: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class ParsedTable:
    table_index: int
    table_name: str
    min_row: int
    max_row: int
    min_column: int
    max_column: int
    header_rows: tuple[int, ...]
    rows: tuple[ParsedRow, ...]
    merged_ranges: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class ParsedSheet:
    sheet_name: str
    tables: tuple[ParsedTable, ...]
    warnings: tuple[Diagnostic, ...] = ()


@dataclass(frozen=True, slots=True)
class ParsedWorkbook:
    original_path: str
    original_filename: str
    sheets: tuple[ParsedSheet, ...]
    warnings: tuple[Diagnostic, ...] = ()


@dataclass(frozen=True, slots=True)
class TableReference:
    table_index: int
    table_name: str


@dataclass(frozen=True, slots=True)
class SearchResult:
    workbook_id: WorkbookId
    workbook_name: str
    workbook_path: str
    sheet_name: str
    table_name: str
    row_number: int
    cells: tuple[ParsedCell, ...]
    matched_cell_addresses: tuple[str, ...]
    table_index: int
    table_refs: tuple[TableReference, ...]


@dataclass(frozen=True, slots=True)
class ImportSummary:
    workbook_id: WorkbookId | None
    imported_sheets: int
    imported_rows: int
    warning_count: int
    elapsed_seconds: float
    success: bool
    warnings: tuple[Diagnostic, ...] = ()
    error_code: ErrorCode | None = None
    message: str = ""


@dataclass(frozen=True, slots=True)
class SearchPage:
    keyword: str
    total_count: int
    results: tuple[SearchResult, ...]
    page: int
    page_size: int
    offset: int


class ImportStage(StrEnum):
    COPYING = "copying"
    PARSING = "parsing"
    VALIDATING = "validating"
    COMMITTING = "committing"
    COMPLETED = "completed"
    FAILED = "failed"
    CANCELLED = "cancelled"


@dataclass(frozen=True, slots=True)
class ImportProgress:
    workbook_id: WorkbookId | None
    stage: ImportStage
    percent: int | None
    message: str


ProgressCallback: TypeAlias = Callable[[ImportProgress], None]


@runtime_checkable
class CancellationToken(Protocol):
    def is_cancelled(self) -> bool: ...


class WorkbookStatus(StrEnum):
    READY = "ready"
    IMPORTING = "importing"
    FAILED = "failed"


@dataclass(frozen=True, slots=True)
class WorkbookInfo:
    """One indexed workbook.  There is no upper bound on how many exist."""

    workbook_id: WorkbookId
    status: WorkbookStatus
    original_filename: str
    managed_path: str | None = None
    source_path: str | None = None
    file_size_bytes: int | None = None
    imported_at: datetime | None = None
    sha256: str | None = None
    sheet_count: int = 0
    row_count: int = 0
    last_message: str = ""


class OpenStatus(StrEnum):
    LOCATED = "located"
    OPENED_WITHOUT_LOCATION = "opened_without_location"
    FAILED = "failed"


@dataclass(frozen=True, slots=True)
class OpenResult:
    status: OpenStatus
    path: str
    sheet_name: str
    cell_address: str
    message: str
    error_code: ErrorCode | None = None


@runtime_checkable
class ExcelParser(Protocol):
    def parse_workbook(self, path: str) -> ParsedWorkbook: ...


@runtime_checkable
class DataManager(Protocol):
    def list_workbooks(self) -> tuple[WorkbookInfo, ...]: ...

    def add_workbook(
        self, file_path: str, *,
        on_progress: ProgressCallback | None = None,
        cancellation: CancellationToken | None = None,
    ) -> ImportSummary: ...

    def replace_workbook(
        self, workbook_id: WorkbookId, file_path: str, *,
        on_progress: ProgressCallback | None = None,
        cancellation: CancellationToken | None = None,
    ) -> ImportSummary: ...

    def remove_workbook(self, workbook_id: WorkbookId) -> None: ...


@runtime_checkable
class SearchService(Protocol):
    def search(
        self, keyword: str, offset: int = 0, limit: int = 100,
    ) -> SearchPage: ...


@runtime_checkable
class WpsOpener(Protocol):
    def open_location(
        self, path: str, sheet_name: str, cell_address: str,
    ) -> OpenResult: ...


@runtime_checkable
class ResultFormatter(Protocol):
    def format_results(self, results: tuple[SearchResult, ...]) -> str: ...


@dataclass(frozen=True, slots=True)
class ApplicationServices:
    data_manager: DataManager
    search_service: SearchService
    wps_opener: WpsOpener
    result_formatter: ResultFormatter
