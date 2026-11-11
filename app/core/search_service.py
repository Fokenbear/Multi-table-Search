"""Exact Unicode substring matching across every indexed cell of every workbook.

A row matches when any single cell contains the keyword; results are de-duplicated
per physical row and paged in a stable order (workbook, sheet, row number).
"""

import sqlite3
import threading

from app.contracts import ContractError, ErrorCode, SearchPage, SearchResult, TableReference
from app.core.codec import decode_cells, loads
from app.core.database import Database


class SqliteSearchService:
    def __init__(self, database: Database):
        self.database = database
        self._cache_lock = threading.Lock()
        self._cache_key = None
        self._cache_ids: tuple[int, ...] = ()

    def search(self, keyword: str, offset: int = 0, limit: int = 100) -> SearchPage:
        if (not isinstance(keyword, str) or not keyword.strip()
                or type(offset) is not int or offset < 0
                or type(limit) is not int or not 1 <= limit <= 1000):
            raise ContractError(ErrorCode.INVALID_ARGUMENT,
                                "关键词不能为空，offset 必须非负，limit 必须为 1..1000")
        original = keyword.strip()
        normalized = original.casefold()
        try:
            with self.database.connect() as conn:
                conn.execute("BEGIN")
                revision = tuple(
                    tuple(row) for row in conn.execute(
                        "SELECT workbook_id,import_version FROM workbooks ORDER BY workbook_id"))
                cache_key = (normalized, revision)
                with self._cache_lock:
                    cached = self._cache_ids if self._cache_key == cache_key else None
                # No LIKE wildcards, concatenation or tokenization: exactly Python
                # containment, evaluated per cell so neighbouring cells never merge.
                ids = cached if cached is not None else tuple(row[0] for row in conn.execute("""
                    SELECT r.id FROM rows r JOIN worksheets w ON w.id=r.sheet_id
                    WHERE r.id IN (SELECT DISTINCT row_id FROM cells WHERE instr(search_text, ?) > 0)
                    ORDER BY w.workbook_id, w.sheet_order, r.row_number
                    """, (normalized,)))
                if cached is None:
                    with self._cache_lock:
                        self._cache_key, self._cache_ids = cache_key, ids
                results = []
                for row_id in ids[offset:offset + limit]:
                    row = conn.execute("""
                        SELECT r.*, w.sheet_name, b.workbook_id, b.original_filename, b.managed_path
                        FROM rows r
                        JOIN worksheets w ON w.id = r.sheet_id
                        JOIN workbooks b ON b.workbook_id = w.workbook_id
                        WHERE r.id = ?""", (row_id,)).fetchone()
                    matched = tuple(x[0] for x in conn.execute(
                        "SELECT address FROM cells WHERE row_id=? AND instr(search_text,?)>0"
                        " ORDER BY column_index", (row_id, normalized)))
                    refs = tuple(TableReference(**x) for x in loads(row["table_refs"]))
                    results.append(SearchResult(
                        row["workbook_id"], row["original_filename"], row["managed_path"],
                        row["sheet_name"], refs[0].table_name, row["row_number"],
                        decode_cells(row["cells_json"]), matched, refs[0].table_index, refs,
                    ))
                return SearchPage(original, len(ids), tuple(results), offset // limit + 1,
                                  limit, offset)
        except (sqlite3.Error, ValueError, KeyError, TypeError) as exc:
            raise ContractError(ErrorCode.STORAGE_FAILED,
                                "查询数据库失败，请检查数据目录或从备份恢复",
                                details=str(exc)) from exc
