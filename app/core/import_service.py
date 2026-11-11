"""Import pipeline: immutable copy first, atomic index commit last.

Each import stage is cancellable except the commit itself, which is a single
SQLite transaction.  A failed or cancelled import leaves the previous index and
the previous managed copy in place.
"""

from dataclasses import asdict
from datetime import datetime, timezone
import hashlib
import logging
import os
from pathlib import Path
import shutil
import sqlite3
import tempfile
import threading
import time
from uuid import uuid4

from openpyxl.utils import get_column_letter

from app.contracts import (
    CancellationToken,
    ContractError,
    ErrorCode,
    ExcelParser,
    ImportProgress,
    ImportStage,
    ImportSummary,
    ParsedWorkbook,
    ProgressCallback,
    WorkbookId,
    WorkbookInfo,
    WorkbookStatus,
)
from app.core import codec
from app.core.database import Database

log = logging.getLogger(__name__)


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def purge_directory(path: Path) -> None:
    """Delete a managed-copy directory, clearing the read-only bit first.

    Managed copies are chmod 0o444 so nothing can silently rewrite them; on
    Windows that attribute also blocks deletion.
    """
    if not path.exists():
        return
    for item in [path, *path.rglob("*")]:
        try:
            item.chmod(0o666)
        except OSError:
            pass
    shutil.rmtree(path, ignore_errors=True)


class LocalDataManager:
    """Owns the set of indexed workbooks.  The set has no fixed size."""

    def __init__(self, database: Database, parser: ExcelParser):
        self.database = database
        self.parser = parser
        self._states: dict[WorkbookId, tuple[WorkbookStatus, str]] = {}
        self._state_lock = threading.Lock()
        self.last_import_timings: dict[str, float] = {}

    # -- queries ---------------------------------------------------------

    def list_workbooks(self) -> tuple[WorkbookInfo, ...]:
        try:
            with self.database.connect() as conn:
                saved = list(conn.execute("SELECT * FROM workbooks ORDER BY workbook_id"))
            with self._state_lock:
                states = self._states.copy()
            items = []
            for row in saved:
                workbook_id = row["workbook_id"]
                status, message = states.get(workbook_id, (WorkbookStatus.READY, ""))
                items.append(WorkbookInfo(
                    workbook_id=workbook_id,
                    status=status,
                    original_filename=row["original_filename"],
                    managed_path=row["managed_path"],
                    source_path=row["source_path"],
                    file_size_bytes=row["file_size_bytes"],
                    imported_at=datetime.fromisoformat(row["imported_at"]),
                    sha256=row["sha256"],
                    sheet_count=row["sheet_count"],
                    row_count=row["row_count"],
                    last_message=message,
                ))
            return tuple(items)
        except sqlite3.Error as exc:
            raise ContractError(ErrorCode.STORAGE_FAILED, "无法读取已索引的工作簿", details=str(exc)) from exc

    def _state(self, workbook_id, status: WorkbookStatus, message: str) -> None:
        with self._state_lock:
            self._states[workbook_id] = (status, message)

    # -- mutations -------------------------------------------------------

    def add_workbook(
        self, file_path: str, *,
        on_progress: ProgressCallback | None = None,
        cancellation: CancellationToken | None = None,
    ) -> ImportSummary:
        return self._ingest(file_path, None, on_progress, cancellation)

    def replace_workbook(
        self, workbook_id: WorkbookId, file_path: str, *,
        on_progress: ProgressCallback | None = None,
        cancellation: CancellationToken | None = None,
    ) -> ImportSummary:
        if type(workbook_id) is not int or workbook_id <= 0:
            raise ContractError(ErrorCode.INVALID_ARGUMENT, "工作簿编号无效")
        self._require_workbook(workbook_id)
        return self._ingest(file_path, workbook_id, on_progress, cancellation)

    def remove_workbook(self, workbook_id: WorkbookId) -> None:
        if type(workbook_id) is not int or workbook_id <= 0:
            raise ContractError(ErrorCode.INVALID_ARGUMENT, "工作簿编号无效")
        if not self.database.writer_lock.acquire(blocking=False):
            raise ContractError(ErrorCode.BUSY, "已有数据写入任务，请稍后重试")
        try:
            with self.database.connect() as conn:
                row = conn.execute(
                    "SELECT managed_path FROM workbooks WHERE workbook_id=?", (workbook_id,)
                ).fetchone()
                if row is None:
                    raise ContractError(ErrorCode.WORKBOOK_NOT_FOUND, "该工作簿已不在索引中")
                managed = Path(row["managed_path"])
                backup = self.database.backup()
                conn.execute("BEGIN IMMEDIATE")
                conn.execute("DELETE FROM workbooks WHERE workbook_id=?", (workbook_id,))
                conn.execute(
                    "INSERT INTO import_history(workbook_id,action,created_at,success,message,backup_path)"
                    " VALUES (?,?,?,1,?,?)",
                    (workbook_id, "remove", datetime.now(timezone.utc).isoformat(),
                     "removed", str(backup)),
                )
                conn.commit()
            # Files are only deleted after the transaction is durable.
            purge_directory(managed.parent)
            with self._state_lock:
                self._states.pop(workbook_id, None)
        except (sqlite3.Error, OSError) as exc:
            raise ContractError(ErrorCode.STORAGE_FAILED, "移除工作簿失败，索引未改变", details=str(exc)) from exc
        finally:
            self.database.writer_lock.release()

    def _require_workbook(self, workbook_id: WorkbookId) -> None:
        with self.database.connect() as conn:
            row = conn.execute(
                "SELECT workbook_id FROM workbooks WHERE workbook_id=?", (workbook_id,)
            ).fetchone()
        if row is None:
            raise ContractError(ErrorCode.WORKBOOK_NOT_FOUND, "该工作簿已不在索引中")

    # -- the pipeline ----------------------------------------------------

    def _ingest(self, file_path, target_id, on_progress, cancellation) -> ImportSummary:
        if not isinstance(file_path, str) or not Path(file_path).is_absolute():
            raise ContractError(ErrorCode.INVALID_ARGUMENT, "文件路径必须是绝对路径")
        started = time.perf_counter()
        warnings: tuple = ()

        def emit(stage: ImportStage, percent: int | None, message: str) -> None:
            if on_progress is not None:
                try:
                    on_progress(ImportProgress(target_id, stage, percent, message))
                except Exception:
                    log.exception("导入进度回调失败")

        def check_cancel() -> None:
            if cancellation is not None and cancellation.is_cancelled():
                raise ContractError(ErrorCode.CANCELLED, "导入已取消，原有数据保持不变")

        if not self.database.writer_lock.acquire(blocking=False):
            return ImportSummary(target_id, 0, 0, 0, 0.0, False,
                                 error_code=ErrorCode.BUSY, message="已有数据写入任务，请稍后重试")
        state_key = target_id if target_id is not None else -1
        try:
            try:
                self._state(state_key, WorkbookStatus.IMPORTING, "正在导入；旧数据仍可查询")
                source = Path(file_path)
                if source.suffix.lower() != ".xlsx":
                    raise ContractError(ErrorCode.UNSUPPORTED_FORMAT, "仅支持 .xlsx，不会静默转换文件")
                if not source.is_file():
                    raise ContractError(ErrorCode.FILE_NOT_FOUND, "所选工作簿不存在或不是文件")
                check_cancel()
                timings: dict[str, float] = {}
                stage_started = time.perf_counter()
                emit(ImportStage.COPYING, 0, "复制并校验工作簿")
                with tempfile.TemporaryDirectory(
                    prefix="import-", dir=self.database.paths.root / "staging",
                    ignore_cleanup_errors=True,
                ) as temp:
                    staged = Path(temp) / source.name
                    before = file_sha256(source)
                    with source.open("rb") as src, staged.open("wb") as dest:
                        shutil.copyfileobj(src, dest, 1024 * 1024)
                        dest.flush()
                        os.fsync(dest.fileno())
                    digest = file_sha256(staged)
                    if digest != before or file_sha256(source) != before:
                        raise ContractError(ErrorCode.VALIDATION_FAILED, "原文件在复制时发生变化，请关闭编辑后重试")
                    timings["copy_seconds"] = time.perf_counter() - stage_started
                    self._reject_duplicate(digest, target_id)
                    check_cancel()
                    emit(ImportStage.PARSING, None, "解析工作簿；此阶段取消可能延迟")
                    stage_started = time.perf_counter()
                    workbook = self.parser.parse_workbook(str(staged))
                    timings["parse_seconds"] = time.perf_counter() - stage_started
                    warnings = tuple(workbook.warnings) + tuple(
                        w for sheet in workbook.sheets for w in sheet.warnings)
                    check_cancel()
                    emit(ImportStage.VALIDATING, 65, "校验完整行、坐标与区域")
                    stage_started = time.perf_counter()
                    physical = self._validate(workbook, source.name)
                    timings["validation_seconds"] = time.perf_counter() - stage_started
                    check_cancel()
                    emit(ImportStage.COMMITTING, 80, "备份并原子更新索引；提交阶段不再取消")
                    version = uuid4().hex
                    version_dir = self.database.paths.root / "workbooks" / version
                    version_dir.mkdir(parents=True)
                    # Office cannot reliably open equal basenames from different
                    # directories at once, so each managed copy gets a unique name.
                    managed = version_dir / f"{source.stem[:40]}__{version}.xlsx"
                    os.replace(staged, managed)
                    managed.chmod(0o444)
                    backup = self.database.backup()
                    stage_started = time.perf_counter()
                    workbook_id, index_seconds = self._write(
                        workbook, physical, source, managed, digest, version, backup, target_id)
                    timings["database_seconds"] = time.perf_counter() - stage_started
                    timings["cell_index_write_seconds"] = index_seconds
                    timings["total_seconds"] = time.perf_counter() - started
                    self.last_import_timings = timings
                with self._state_lock:
                    self._states.pop(state_key, None)
                emit(ImportStage.COMPLETED, 100, "索引更新完成")
                return ImportSummary(workbook_id, len(workbook.sheets), sum(len(r) for r in physical),
                                     len(warnings), time.perf_counter() - started, True, warnings,
                                     message=source.name)
            except ContractError as exc:
                error = exc
            except (PermissionError, FileNotFoundError) as exc:
                error = ContractError(ErrorCode.FILE_ACCESS, "文件无法访问，请检查权限及文件是否仍存在", details=str(exc))
            except (sqlite3.Error, OSError) as exc:
                error = ContractError(ErrorCode.STORAGE_FAILED, "更新未提交，旧数据保留；请检查磁盘空间和数据目录权限", details=str(exc))
            except Exception as exc:
                log.exception("导入出现非预期错误")
                error = ContractError(ErrorCode.INTERNAL_ERROR, "导入发生内部错误，更新未提交", details=str(exc))
            self._state(state_key, WorkbookStatus.FAILED, error.message)
            emit(ImportStage.CANCELLED if error.code == ErrorCode.CANCELLED else ImportStage.FAILED,
                 None, error.message)
            try:
                with self.database.connect() as conn:
                    conn.execute(
                        "INSERT INTO import_history(workbook_id,action,created_at,success,message)"
                        " VALUES (?,?,?,0,?)",
                        (target_id, "import", datetime.now(timezone.utc).isoformat(), error.code.value),
                    )
                    conn.commit()
            except sqlite3.Error:
                log.exception("无法记录失败导入历史")
            log.warning("导入失败，code=%s", error.code.value)
            return ImportSummary(target_id, 0, 0, len(warnings), time.perf_counter() - started,
                                 False, warnings, error.code, error.message)
        finally:
            self.database.writer_lock.release()

    def _reject_duplicate(self, digest: str, target_id) -> None:
        with self.database.connect() as conn:
            row = conn.execute(
                "SELECT workbook_id,original_filename FROM workbooks WHERE sha256=?", (digest,)
            ).fetchone()
        if row is not None and row["workbook_id"] != target_id:
            raise ContractError(
                ErrorCode.DUPLICATE_WORKBOOK,
                f"《{row['original_filename']}》已经索引，同一份文件不重复导入",
            )

    def _validate(self, workbook, filename):
        def require(condition, message):
            if not condition:
                raise ContractError(ErrorCode.VALIDATION_FAILED, message)

        require(isinstance(workbook, ParsedWorkbook), "解析结果类型不正确")
        require(workbook.original_filename == filename, "解析返回的来源信息不一致")
        require(type(workbook.sheets) is tuple, "解析工作表集合必须为 tuple")
        names = set()
        physical = []
        for sheet in workbook.sheets:
            require(sheet.sheet_name not in names, "解析包含重复工作表")
            names.add(sheet.sheet_name)
            merged = {}
            for expected, table in enumerate(sheet.tables, 1):
                require(table.table_index == expected, "表格索引必须从 1 连续编号")
                for row in table.rows:
                    require(row.workbook_name == filename and row.sheet_name == sheet.sheet_name
                            and row.table_index == table.table_index and row.table_name == table.table_name
                            and 1 <= row.row_number <= 1048576 and not row.matched_cell_addresses,
                            "解析行来源、表格或行号不合法")
                    item = merged.setdefault(row.row_number, {"cells": {}, "refs": [], "owners": {}})
                    ref = {"table_index": table.table_index, "table_name": table.table_name}
                    require(ref not in item["refs"], "同一区域重复返回同一行")
                    item["refs"].append(ref)
                    previous = 0
                    for cell in row.cells:
                        col = cell.column_index
                        require(type(col) is int and previous < col <= 16384, "单元格列序不合法")
                        previous = col
                        require(cell.column_letter == get_column_letter(col)
                                and cell.address == f"{cell.column_letter}{row.row_number}",
                                "单元格坐标不一致")
                        require(col not in item["cells"] or item["cells"][col] == cell,
                                "不同区域的单元格发生冲突")
                        codec.cell_json(cell)  # Validate the closed scalar representation.
                        item["cells"][col] = cell
                        item["owners"][col] = table.table_index
            physical.append(merged)
        return physical

    def _write(self, workbook, physical, source, managed, digest, version, backup, target_id):
        now = datetime.now(timezone.utc).isoformat()
        index_seconds = 0.0
        sheet_count = len(workbook.sheets)
        row_count = sum(len(rows) for rows in physical)
        previous_managed: Path | None = None
        with self.database.connect() as conn:
            try:
                conn.execute("BEGIN IMMEDIATE")
                if target_id is None:
                    workbook_id = conn.execute(
                        "INSERT INTO workbooks(original_filename,source_path,managed_path,sha256,"
                        "file_size_bytes,imported_at,sheet_count,row_count,import_version)"
                        " VALUES (?,?,?,?,?,?,?,?,?)",
                        (source.name, str(source), str(managed), digest,
                         managed.stat().st_size, now, sheet_count, row_count, version)).lastrowid
                else:
                    workbook_id = target_id
                    previous = conn.execute(
                        "SELECT managed_path FROM workbooks WHERE workbook_id=?",
                        (workbook_id,)).fetchone()
                    previous_managed = Path(previous["managed_path"]) if previous else None
                    conn.execute(
                        "UPDATE workbooks SET original_filename=?,source_path=?,managed_path=?,sha256=?,"
                        "file_size_bytes=?,imported_at=?,sheet_count=?,row_count=?,import_version=?"
                        " WHERE workbook_id=?",
                        (source.name, str(source), str(managed), digest, managed.stat().st_size,
                         now, sheet_count, row_count, version, workbook_id))
                    conn.execute("DELETE FROM worksheets WHERE workbook_id=?", (workbook_id,))
                for ordinal, (sheet, rows) in enumerate(zip(workbook.sheets, physical)):
                    sheet_id = conn.execute(
                        "INSERT INTO worksheets(workbook_id,sheet_name,sheet_order) VALUES (?,?,?)",
                        (workbook_id, sheet.sheet_name, ordinal)).lastrowid
                    table_ids = {}
                    for table in sheet.tables:
                        table_ids[table.table_index] = conn.execute(
                            "INSERT INTO tables(sheet_id,table_index,table_name,min_row,max_row,"
                            "min_column,max_column,header_rows,merged_ranges)"
                            " VALUES (?,?,?,?,?,?,?,?,?)",
                            (sheet_id, table.table_index, table.table_name, table.min_row, table.max_row,
                             table.min_column, table.max_column, codec.dumps(table.header_rows),
                             codec.dumps(table.merged_ranges))).lastrowid
                    for number, item in sorted(rows.items()):
                        cells = tuple(item["cells"][col] for col in sorted(item["cells"]))
                        refs = sorted(item["refs"], key=lambda x: x["table_index"])
                        row_id = conn.execute(
                            "INSERT INTO rows(sheet_id,row_number,cells_json,table_refs) VALUES (?,?,?,?)",
                            (sheet_id, number, codec.dumps([asdict(c) for c in cells]),
                             codec.dumps(refs))).lastrowid
                        conn.executemany("INSERT INTO row_table_memberships VALUES (?,?)",
                                         [(row_id, table_ids[ref["table_index"]]) for ref in refs])
                        index_started = time.perf_counter()
                        conn.executemany(
                            "INSERT INTO cells VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", [
                                (row_id, table_ids[item["owners"][c.column_index]], c.column_index,
                                 c.column_letter, c.header, codec.dumps(c.raw_value), c.display_value,
                                 c.address, c.is_supplementary, c.data_type, c.number_format,
                                 c.origin_address, c.formula_has_cached_value,
                                 c.display_value.strip().casefold()) for c in cells])
                        index_seconds += time.perf_counter() - index_started
                conn.execute(
                    "INSERT INTO import_history(workbook_id,action,version,sha256,created_at,success,"
                    "message,backup_path) VALUES (?,?,?,?,?,1,?,?)",
                    (workbook_id, "replace" if target_id is not None else "add", version, digest,
                     now, "committed", str(backup)))
                conn.commit()
                if previous_managed is not None:
                    # The new index is durable; the superseded copy can go.
                    purge_directory(previous_managed.parent)
                return workbook_id, index_seconds
            except BaseException:
                conn.rollback()
                raise
