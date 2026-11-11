"""SQLite storage for an unbounded set of indexed workbooks.

The database only ever references immutable managed copies under
``<data>/workbooks/<version>/``.  Writing a new index for one workbook is a
single transaction; a failure leaves the previous index untouched.
"""

from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
import sqlite3
import threading
from uuid import uuid4

from app.windows.windows_paths import AppPaths

SCHEMA_VERSION = 2

SCHEMA = """
CREATE TABLE IF NOT EXISTS schema_version(version INTEGER NOT NULL);
CREATE TABLE IF NOT EXISTS workbooks(
 workbook_id INTEGER PRIMARY KEY AUTOINCREMENT,
 original_filename TEXT NOT NULL,
 source_path TEXT NOT NULL,
 managed_path TEXT NOT NULL,
 sha256 TEXT NOT NULL,
 file_size_bytes INTEGER NOT NULL,
 imported_at TEXT NOT NULL,
 sheet_count INTEGER NOT NULL DEFAULT 0,
 row_count INTEGER NOT NULL DEFAULT 0,
 import_version TEXT NOT NULL);
CREATE UNIQUE INDEX IF NOT EXISTS idx_workbooks_sha256 ON workbooks(sha256);
CREATE TABLE IF NOT EXISTS worksheets(
 id INTEGER PRIMARY KEY,
 workbook_id INTEGER NOT NULL REFERENCES workbooks(workbook_id) ON DELETE CASCADE,
 sheet_name TEXT NOT NULL,
 sheet_order INTEGER NOT NULL,
 UNIQUE(workbook_id,sheet_name));
CREATE TABLE IF NOT EXISTS tables(
 id INTEGER PRIMARY KEY,
 sheet_id INTEGER NOT NULL REFERENCES worksheets(id) ON DELETE CASCADE,
 table_index INTEGER NOT NULL,
 table_name TEXT NOT NULL,
 min_row INTEGER, max_row INTEGER, min_column INTEGER, max_column INTEGER,
 header_rows TEXT NOT NULL, merged_ranges TEXT NOT NULL,
 UNIQUE(sheet_id,table_index));
CREATE TABLE IF NOT EXISTS rows(
 id INTEGER PRIMARY KEY,
 sheet_id INTEGER NOT NULL REFERENCES worksheets(id) ON DELETE CASCADE,
 row_number INTEGER NOT NULL,
 cells_json TEXT NOT NULL,
 table_refs TEXT NOT NULL,
 UNIQUE(sheet_id,row_number));
CREATE TABLE IF NOT EXISTS row_table_memberships(
 row_id INTEGER NOT NULL REFERENCES rows(id) ON DELETE CASCADE,
 table_id INTEGER NOT NULL REFERENCES tables(id) ON DELETE CASCADE,
 PRIMARY KEY(row_id,table_id));
CREATE TABLE IF NOT EXISTS cells(
 row_id INTEGER NOT NULL REFERENCES rows(id) ON DELETE CASCADE,
 table_id INTEGER NOT NULL REFERENCES tables(id) ON DELETE CASCADE,
 column_index INTEGER NOT NULL,
 column_letter TEXT NOT NULL,
 header TEXT,
 raw_json TEXT NOT NULL,
 display_value TEXT NOT NULL,
 address TEXT NOT NULL,
 is_supplementary INTEGER NOT NULL,
 data_type TEXT NOT NULL,
 number_format TEXT NOT NULL,
 origin_address TEXT,
 formula_has_cached_value INTEGER,
 search_text TEXT NOT NULL,
 PRIMARY KEY(row_id,column_index));
CREATE TABLE IF NOT EXISTS import_history(
 id INTEGER PRIMARY KEY,
 workbook_id INTEGER,
 action TEXT NOT NULL,
 version TEXT,
 sha256 TEXT,
 created_at TEXT NOT NULL,
 success INTEGER NOT NULL,
 message TEXT NOT NULL,
 backup_path TEXT);
CREATE INDEX IF NOT EXISTS idx_worksheets_workbook ON worksheets(workbook_id,sheet_order);
CREATE INDEX IF NOT EXISTS idx_rows_sheet ON rows(sheet_id,row_number);
CREATE INDEX IF NOT EXISTS idx_cells_table ON cells(table_id);
CREATE INDEX IF NOT EXISTS idx_cells_row ON cells(row_id);
CREATE INDEX IF NOT EXISTS idx_memberships_table ON row_table_memberships(table_id);
"""


class Database:
    def __init__(self, paths: AppPaths):
        self.paths = paths
        self.writer_lock = threading.Lock()
        with self.connect() as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            existing = conn.execute(
                "SELECT name FROM sqlite_master WHERE name='schema_version'"
            ).fetchone()
            if existing:
                row = conn.execute("SELECT version FROM schema_version").fetchone()
                if row is None or row[0] != SCHEMA_VERSION:
                    raise RuntimeError("数据库版本不兼容，未修改数据库")
            conn.executescript(SCHEMA)
            if not existing:
                conn.execute("INSERT INTO schema_version VALUES (?)", (SCHEMA_VERSION,))
            conn.commit()

    @contextmanager
    def connect(self):
        conn = sqlite3.connect(self.paths.database, timeout=30)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys=ON")
        conn.execute("PRAGMA synchronous=FULL")
        try:
            yield conn
        finally:
            conn.close()

    def backup(self) -> Path:
        """Transactionally consistent copy of the live database."""
        name = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S") + "-" + uuid4().hex + ".db"
        target = self.paths.root / "backups" / name
        try:
            with self.connect() as source, sqlite3.connect(target) as dest:
                source.backup(dest)
            return target
        except Exception:
            target.unlink(missing_ok=True)
            raise

    def restore(self, backup: str | Path) -> Path:
        """Offline restore.  The caller must hold the single-instance lock."""
        source_path = Path(backup).resolve()
        if source_path == self.paths.database:
            raise ValueError("不能用数据库自身作为恢复来源")
        with self.writer_lock:
            with sqlite3.connect(source_path.as_uri() + "?mode=ro", uri=True) as source:
                if source.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
                    raise ValueError("备份完整性校验失败")
                version = source.execute("SELECT version FROM schema_version").fetchone()
                if version is None or version[0] != SCHEMA_VERSION:
                    raise ValueError("备份数据库版本不兼容")
                for (path,) in source.execute("SELECT managed_path FROM workbooks"):
                    if not Path(path).is_file():
                        raise ValueError("备份引用的托管工作簿缺失，未执行恢复")
                previous = self.backup()
                with self.connect() as target:
                    source.backup(target)
            return previous
