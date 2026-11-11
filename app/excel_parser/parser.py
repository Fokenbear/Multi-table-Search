"""Excel parser implementation for the Multi-Table Search application.

Public entry point: :class:`WorkbookParser.parse_workbook`.  The parser is
synchronous, read-only with respect to the input workbook, and returns the
frozen DTOs from :mod:`app.contracts`.  It does not import Qt, SQLite or COM.
"""

from __future__ import annotations

import gc
import re
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.contracts import (
    INTERFACE_VERSION,
    ContractError,
    Diagnostic,
    ErrorCode,
    ParsedCell,
    ParsedRow,
    ParsedSheet,
    ParsedTable,
    ParsedWorkbook,
)
from app.excel_parser import ooxml
from app.excel_parser.ooxml import FormulaInfo
from app.excel_parser.values import classify_value, formula_display

MIN_REGION_BLANK_ROWS = 2
MAX_PROPAGATED_MERGE_ROWS = 20000
MAX_HEADER_TEXT_LENGTH = 40
MAX_TITLE_TEXT_LENGTH = 120
_HEADER_LABELS = {
    "日期", "询货人", "型号", "料号", "物料编码", "物料号", "产品型号", "包装量",
    "数量", "供应商", "供应商信息", "备注", "名称", "价格", "已成交", "合计",
    "无缓存", "零填充", "文本料号", "零值", "布尔", "品名", "规格", "单价",
    "part", "part number", "part no", "part no.", "qty", "quantity", "supplier",
    "date", "name", "price", "description", "notes", "value", "品牌", "库存",
    "库存数量", "库存数", "库存量", "单位", "产品名称", "货期", "工厂", "封装",
    "manufacturer", "brand", "stock", "inventory", "package", "mfr", "mfg",
}


@dataclass(frozen=True, slots=True)
class _CellRange:
    min_row: int
    max_row: int
    min_col: int
    max_col: int
    coordinate: str

    @property
    def height(self) -> int:
        return self.max_row - self.min_row + 1

    @property
    def width(self) -> int:
        return self.max_col - self.min_col + 1


@dataclass(frozen=True, slots=True)
class _CellData:
    raw_value: object
    display_value: str
    data_type: str
    number_format: str
    formula_cached: bool | None = None
    format_fallback: bool = False


@dataclass(slots=True)
class _SheetGrid:
    worksheet: Worksheet
    own: dict[tuple[int, int], _CellData]
    merges: tuple[_CellRange, ...]
    missing_formula_caches: list[str]
    fallback_cells: list[str]

    def merge_at(self, row: int, col: int) -> _CellRange | None:
        for merge in self.merges:
            if merge.min_row <= row <= merge.max_row and merge.min_col <= col <= merge.max_col:
                return merge
        return None

    def logical(self, row: int, col: int) -> tuple[_CellData | None, str | None]:
        """Return the logical cell data and optional merged origin address."""

        merge = self.merge_at(row, col)
        if merge is not None and merge.coordinate:
            if col != merge.min_col:
                return None, None
            anchor = self.own.get((merge.min_row, merge.min_col))
            if anchor is None:
                return None, None
            if row == merge.min_row:
                return anchor, None
            origin = f"{get_column_letter(merge.min_col)}{merge.min_row}"
            return anchor, origin
        return self.own.get((row, col)), None


class WorkbookParser:
    """Concrete :class:`app.contracts.ExcelParser` implementation.

    The constructor takes no arguments and performs no I/O.  A new parser can
    be reused for any number of workbooks; per-call state lives on the stack.
    """

    contract_version = INTERFACE_VERSION

    def parse_workbook(self, path: str) -> ParsedWorkbook:
        source_path = self._validate_input(path)
        warnings: list[Diagnostic] = []
        formula_map: dict[str, dict[str, FormulaInfo]] = {}
        try:
            with ooxml.open_archive(source_path) as archive:
                index = ooxml.read_workbook_index(archive)
                for name in index.sheet_order:
                    sheet_path = index.sheet_paths.get(name)
                    if sheet_path:
                        formula_map[name] = ooxml.scan_formulas(archive, sheet_path)
                needs_repair = ooxml.count_empty_fills(index.styles_xml) > 0
        except ooxml.OoxmlError as exc:
            raise self._wrap_ooxml_error(source_path, exc) from exc

        workbook = self._load_workbook(source_path, needs_repair=needs_repair, warnings=warnings)
        try:
            sheets: list[ParsedSheet] = []
            total_tables = 0
            for worksheet in workbook.worksheets:
                try:
                    parsed_sheet = self._parse_sheet(
                        worksheet,
                        filename=source_path.name,
                        formulas=formula_map.get(worksheet.title, {}),
                    )
                except ContractError:
                    raise
                except Exception as exc:
                    raise ContractError(
                        ErrorCode.PARSE_FAILED,
                        f"工作表“{worksheet.title}”解析失败",
                        details=f"{type(exc).__name__}: {exc}",
                    ) from exc
                total_tables += len(parsed_sheet.tables)
                sheets.append(parsed_sheet)
            if not sheets or total_tables == 0:
                warnings.append(
                    Diagnostic(
                        code="EMPTY_WORKBOOK",
                        message="工作簿没有包含可解析的数据区域",
                    )
                )
            return ParsedWorkbook(
                original_path=str(source_path),
                original_filename=source_path.name,
                sheets=tuple(sheets),
                warnings=tuple(warnings),
            )
        finally:
            try:
                workbook.close()
            finally:
                del workbook
                gc.collect()

    # ------------------------------------------------------------------
    # input validation / loading

    def _validate_input(self, path: str) -> Path:
        if not isinstance(path, str) or not path.strip():
            raise ContractError(ErrorCode.INVALID_ARGUMENT, "工作簿路径不能为空")
        source_path = Path(path)
        if not source_path.is_absolute():
            raise ContractError(ErrorCode.INVALID_ARGUMENT, "工作簿路径必须是绝对路径")
        if source_path.suffix.lower() != ".xlsx":
            raise ContractError(
                ErrorCode.UNSUPPORTED_FORMAT,
                "仅支持 .xlsx 格式；其他格式不会被静默转换",
            )
        if not source_path.exists():
            raise ContractError(ErrorCode.FILE_NOT_FOUND, "工作簿文件不存在")
        if not source_path.is_file():
            raise ContractError(ErrorCode.FILE_ACCESS, "工作簿路径不是文件")
        return source_path

    def _wrap_ooxml_error(self, source_path: Path, exc: ooxml.OoxmlError) -> ContractError:
        message = str(exc)
        if "不存在" in message:
            return ContractError(ErrorCode.FILE_NOT_FOUND, "工作簿文件不存在")
        if "权限" in message:
            return ContractError(ErrorCode.FILE_ACCESS, "工作簿文件无法读取（权限不足）")
        return ContractError(
            ErrorCode.PARSE_FAILED,
            "工作簿文件无法解析，可能已损坏或不是有效的 .xlsx 文件",
            details=f"{source_path.name}: {message}",
        )

    def _load_workbook(self, source_path: Path, *, needs_repair: bool, warnings: list[Diagnostic]):
        attempts = [needs_repair]
        if not needs_repair:
            attempts.append(True)
        last_error: Exception | None = None
        for repaired in attempts:
            try:
                with ooxml.RepairedWorkbook(source_path, needed=repaired) as load_path:
                    workbook = load_workbook(
                        filename=str(load_path),
                        read_only=False,
                        data_only=True,
                        keep_links=False,
                    )
                if repaired:
                    count = 0
                    try:
                        with ooxml.open_archive(source_path) as archive:
                            count = ooxml.count_empty_fills(archive.read("xl/styles.xml"))
                    except Exception:
                        count = 0
                    if count:
                        warnings.append(
                            Diagnostic(
                                code="STYLE_REPAIRED",
                                message=(
                                    f"检测到 {count} 个空 fill 节点，已在临时副本上修复；"
                                    "原始文件未被修改"
                                ),
                            )
                        )
                return workbook
            except Exception as exc:  # openpyxl raises several exception types
                last_error = exc
                if repaired:
                    break
        details = ""
        if last_error is not None:
            details = f"{type(last_error).__name__}: {last_error}"
        raise ContractError(
            ErrorCode.PARSE_FAILED,
            "工作簿无法解析；请确认文件未损坏且格式为 .xlsx",
            details=details,
        )

    # ------------------------------------------------------------------
    # sheet parsing

    def _parse_sheet(
        self,
        worksheet: Worksheet,
        *,
        filename: str,
        formulas: dict[str, FormulaInfo],
    ) -> ParsedSheet:
        grid = self._build_grid(worksheet, formulas)
        warnings: list[Diagnostic] = []
        if not grid.own:
            warnings.append(
                Diagnostic(
                    code="EMPTY_SHEET",
                    message="工作表没有数据",
                    sheet_name=worksheet.title,
                )
            )
            return ParsedSheet(sheet_name=worksheet.title, tables=(), warnings=tuple(warnings))

        if grid.missing_formula_caches:
            warnings.append(
                Diagnostic(
                    code="FORMULA_WITHOUT_CACHED_VALUE",
                    message=(
                        f"{len(grid.missing_formula_caches)} 个公式单元格没有缓存值，"
                        "已保留公式原文"
                    ),
                    sheet_name=worksheet.title,
                    cell_address=grid.missing_formula_caches[0],
                )
            )
        if grid.fallback_cells:
            warnings.append(
                Diagnostic(
                    code="DISPLAY_FORMAT_FALLBACK",
                    message=(
                        f"{len(grid.fallback_cells)} 个单元格的数字格式无法完整复现，"
                        "已使用常规显示"
                    ),
                    sheet_name=worksheet.title,
                    cell_address=grid.fallback_cells[0],
                )
            )

        tables: list[ParsedTable] = []
        for start_row, end_row in self._split_regions(grid):
            region_tables = self._build_tables(
                grid, workbook_name=filename, start_row=start_row, end_row=end_row
            )
            for table in region_tables:
                table_index = len(tables) + 1
                table = replace(
                    table,
                    table_index=table_index,
                    rows=tuple(replace(row, table_index=table_index) for row in table.rows),
                )
                tables.append(table)
                if not table.header_rows:
                    warnings.append(
                        Diagnostic(
                            code="HEADER_NOT_FOUND",
                            message="未能可靠识别表头，已使用列字母降级解析并保留原始数据",
                            sheet_name=worksheet.title,
                        )
                    )
        if not tables:
            warnings.append(
                Diagnostic(
                    code="NO_DATA_REGION",
                    message="未能从该工作表识别出数据区域",
                    sheet_name=worksheet.title,
                )
            )
        # Every nonempty source cell must be represented as data, a header or
        # a retained title. Never publish a partial workbook as a success.
        data_addresses = {cell.address for table in tables for row in table.rows for cell in row.cells}
        for (row, col), value in grid.own.items():
            address = f"{get_column_letter(col)}{row}"
            if address in data_addresses:
                continue
            retained = any(
                (table.min_column <= col <= table.max_column and row in table.header_rows) or (
                    table.header_rows and row == table.min_row < min(table.header_rows)
                    and self._text_of(value) in table.table_name)
                for table in tables
            )
            if not retained:
                raise ContractError(ErrorCode.VALIDATION_FAILED,
                                    "存在未关联到数据、表头或标题的原始单元格，已拒绝不完整导入",
                                    details=f"{worksheet.title}!{address}")
        return ParsedSheet(sheet_name=worksheet.title, tables=tuple(tables), warnings=tuple(warnings))

    # ------------------------------------------------------------------
    # grid construction

    def _build_grid(
        self,
        worksheet: Worksheet,
        formulas: dict[str, FormulaInfo],
    ) -> _SheetGrid:
        own: dict[tuple[int, int], _CellData] = {}
        missing_formula: list[str] = []
        fallback_cells: list[str] = []
        formula_by_address = {addr.upper(): info for addr, info in formulas.items()}
        raw_cells = getattr(worksheet, "_cells", {})
        cells: Iterable = raw_cells.values() if hasattr(raw_cells, "values") else ()
        for cell in cells:
            address = str(cell.coordinate).upper()
            formula = formula_by_address.get(address)
            if cell.value is None and formula is None:
                continue
            number_format = str(cell.number_format or "General")
            if formula is not None:
                raw_formula = "=" + formula.text
                if formula.has_cached_value:
                    display, format_fallback = formula_display(cell.value, number_format)
                    data_type = "formula"
                    cached = True
                else:
                    display = raw_formula
                    data_type = "formula"
                    cached = False
                    format_fallback = False
                    missing_formula.append(address)
                own[(cell.row, cell.column)] = _CellData(
                    raw_value=raw_formula,
                    display_value=display,
                    data_type=data_type,
                    number_format=number_format,
                    formula_cached=cached,
                    format_fallback=format_fallback,
                )
                if format_fallback:
                    fallback_cells.append(address)
                continue
            parsed = classify_value(cell.value, number_format)
            own[(cell.row, cell.column)] = _CellData(
                raw_value=parsed.raw_value,
                display_value=parsed.display_value,
                data_type=parsed.data_type,
                number_format=parsed.number_format,
                format_fallback=parsed.used_format_fallback,
            )
            if parsed.used_format_fallback:
                fallback_cells.append(address)

        # Formula addresses seen in XML but absent from openpyxl's cell store
        # (defensive; normally impossible) are still retained.
        for address, formula in formula_by_address.items():
            try:
                row, col = self._split_address(address)
            except ValueError:
                continue
            if (row, col) in own:
                continue
            raw_formula = "=" + formula.text
            own[(row, col)] = _CellData(
                raw_value=raw_formula,
                display_value=raw_formula if not formula.has_cached_value else "",
                data_type="formula",
                number_format="General",
                formula_cached=formula.has_cached_value,
            )
            if not formula.has_cached_value:
                missing_formula.append(address)

        merges = tuple(
            _CellRange(
                min_row=merged.min_row,
                max_row=merged.max_row,
                min_col=merged.min_col,
                max_col=merged.max_col,
                coordinate=str(merged),
            )
            for merged in worksheet.merged_cells.ranges
        )
        return _SheetGrid(
            worksheet=worksheet,
            own=own,
            merges=merges,
            missing_formula_caches=missing_formula,
            fallback_cells=fallback_cells,
        )

    # ------------------------------------------------------------------
    # region / table detection

    def _split_regions(self, grid: _SheetGrid) -> list[tuple[int, int]]:
        meaningful: set[int] = set()
        by_row: dict[int, dict[int, _CellData]] = {}
        for (row, col), value in grid.own.items():
            meaningful.add(row)
            by_row.setdefault(row, {})[col] = value
        for merge in grid.merges:
            anchor = grid.own.get((merge.min_row, merge.min_col))
            if anchor is None:
                continue
            if merge.height > MAX_PROPAGATED_MERGE_ROWS:
                raise ContractError(
                    ErrorCode.VALIDATION_FAILED,
                    "纵向合并区域过大，无法安全完整解析；旧数据将保留",
                    details=merge.coordinate,
                )
            meaningful.update(range(merge.min_row, merge.max_row + 1))
        rows = sorted(meaningful)
        if not rows:
            return []
        regions: list[tuple[int, int]] = []
        start = rows[0]
        previous = rows[0]
        for row in rows[1:]:
            # A new explicit header can begin a table without blank separators.
            current = by_row.get(row, {})
            if (row - previous - 1 >= MIN_REGION_BLANK_ROWS
                    or (len(current) >= 2 and self._is_header_like(current)
                        and row > start + 1)):
                regions.append((start, previous))
                start = row
            previous = row
        regions.append((start, previous))
        return regions

    def _build_tables(
        self,
        grid: _SheetGrid,
        *,
        workbook_name: str,
        start_row: int,
        end_row: int,
    ) -> list[ParsedTable]:
        row_values: dict[int, dict[int, _CellData]] = {}
        for (row, col), data in grid.own.items():
            if start_row <= row <= end_row:
                row_values.setdefault(row, {})[col] = data
        present_rows = [row for row in range(start_row, end_row + 1) if row_values.get(row)]
        if not present_rows:
            return []

        title_text, title_row, header_rows = self._detect_header(
            grid, row_values, present_rows
        )
        header_values = self._build_header_map(grid, header_rows)
        data_start = (max(header_rows) + 1) if header_rows else start_row
        data_start = self._first_content_row(row_values, data_start, end_row)
        if data_start is None:
            data_start = end_row + 1

        active = set(header_values)
        for row in range(data_start, end_row + 1):
            active.update(row_values.get(row, {}))
        for merge in grid.merges:
            if merge.min_row >= data_start and merge.min_row <= end_row:
                if (merge.min_row, merge.min_col) in grid.own:
                    active.add(merge.min_col)
        active = {col for col in active if col >= 1}
        if not active:
            return []

        header_columns = {col for col, header in header_values.items() if header is not None}
        last_header_col = max(header_columns) if header_columns else None
        table_name = title_text if title_text else grid.worksheet.title

        tables: list[ParsedTable] = []
        for block in self._column_blocks(active, header_values):
            table = self._build_block(
                grid,
                workbook_name=workbook_name,
                table_name=table_name,
                title_row=title_row,
                header_rows=header_rows,
                header_values=header_values,
                data_start=data_start,
                end_row=end_row,
                start_row=start_row,
                block=block,
                last_header_col=last_header_col,
            )
            if table is not None:
                tables.append(table)
        return tables

    @staticmethod
    def _column_blocks(
        active: set[int],
        header_values: dict[int, str | None],
    ) -> list[set[int]]:
        """Split a row band into side-by-side tables when headers are disjoint.

        A gap is only accepted when every column in it is entirely empty in the
        whole region.  Right-side no-header notes therefore stay attached to
        the last header block instead of becoming a separate table.
        """

        header_columns = sorted(col for col, header in header_values.items() if header is not None)
        if len(header_columns) < 2:
            return [set(active)]
        clusters: list[list[int]] = [[header_columns[0]]]
        for col in header_columns[1:]:
            previous = clusters[-1][-1]
            gap = range(previous + 1, col)
            if col == previous + 1 or any(gap_col in active for gap_col in gap):
                clusters[-1].append(col)
            else:
                clusters.append([col])
        if len(clusters) == 1:
            return [set(active)]
        blocks: list[set[int]] = []
        for index, cluster in enumerate(clusters):
            low = min(active) if index == 0 else min(cluster)
            high = max(active) if index == len(clusters) - 1 else max(cluster)
            block = {col for col in active if low <= col <= high}
            if block:
                blocks.append(block)
        return blocks or [set(active)]

    def _build_block(
        self,
        grid: _SheetGrid,
        *,
        workbook_name: str,
        table_name: str,
        title_row: int | None,
        header_rows: list[int],
        header_values: dict[int, str | None],
        data_start: int,
        end_row: int,
        start_row: int,
        block: set[int],
        last_header_col: int | None,
    ) -> ParsedTable | None:
        min_col, max_col = min(block), max(block)
        rows: list[ParsedRow] = []
        last_data_row = max(header_rows) if header_rows else start_row
        for row in range(data_start, end_row + 1):
            cells: list[ParsedCell] = []
            has_value = False
            for col in sorted(block):
                data, origin = grid.logical(row, col)
                header = header_values.get(col)
                supplementary = header is None and last_header_col is not None and col > last_header_col
                address = f"{get_column_letter(col)}{row}"
                if data is None:
                    cells.append(
                        ParsedCell(
                            column_index=col,
                            column_letter=get_column_letter(col),
                            header=header,
                            raw_value=None,
                            display_value="",
                            address=address,
                            is_supplementary=supplementary,
                            data_type="blank",
                            number_format="General",
                            origin_address=origin,
                        )
                    )
                    continue
                has_value = True
                cells.append(
                    ParsedCell(
                        column_index=col,
                        column_letter=get_column_letter(col),
                        header=header,
                        raw_value=data.raw_value,  # type: ignore[arg-type]
                        display_value=data.display_value,
                        address=address,
                        is_supplementary=supplementary,
                        data_type=data.data_type,  # type: ignore[arg-type]
                        number_format=data.number_format,
                        origin_address=origin,
                        formula_has_cached_value=data.formula_cached,
                    )
                )
            if not has_value:
                continue
            last_data_row = row
            rows.append(
                ParsedRow(
                    workbook_name=workbook_name,
                    sheet_name=grid.worksheet.title,
                    table_name=table_name,
                    row_number=row,
                    cells=tuple(cells),
                    table_index=0,
                )
            )

        if not rows and not header_rows:
            return None
        min_row = min(title_row, min(header_rows)) if (title_row and header_rows) else (
            title_row or (min(header_rows) if header_rows else start_row)
        )
        table_bounds = _CellRange(
            min_row=min_row,
            max_row=max(last_data_row, min_row),
            min_col=min_col,
            max_col=max_col,
            coordinate="",
        )
        merged_ranges = tuple(
            merge.coordinate
            for merge in grid.merges
            if not (
                merge.max_row < table_bounds.min_row
                or merge.min_row > table_bounds.max_row
                or merge.max_col < table_bounds.min_col
                or merge.min_col > table_bounds.max_col
            )
        )
        return ParsedTable(
            table_index=0,
            table_name=table_name,
            min_row=min_row,
            max_row=table_bounds.max_row,
            min_column=min_col,
            max_column=max_col,
            header_rows=tuple(header_rows),
            rows=tuple(rows),
            merged_ranges=merged_ranges,
        )

    def _first_content_row(
        self,
        row_values: dict[int, dict[int, _CellData]],
        start: int,
        end: int,
    ) -> int | None:
        for row in range(start, end + 1):
            if row_values.get(row):
                return row
        return None

    def _detect_header(
        self,
        grid: _SheetGrid,
        row_values: dict[int, dict[int, _CellData]],
        present_rows: list[int],
    ) -> tuple[str | None, int | None, list[int]]:
        first = present_rows[0]
        headers: list[int] = []
        title: str | None = None
        title_row: int | None = None

        def header_like(row: int) -> bool:
            return self._is_header_like(row_values.get(row, {}))

        if (
            len(present_rows) >= 2
            and self._is_title_like(row_values.get(first, {}))
            and header_like(present_rows[1])
        ):
            title = self._row_text(row_values.get(first, {}))
            title_row = first
            headers = [present_rows[1]]
        elif header_like(first):
            headers = [first]

        if headers:
            # Conservative multi-row header support: only combine when the
            # upper header row contains a horizontally merged label, so an
            # ordinary short data row is never mistaken for a header.
            index = present_rows.index(headers[-1])
            if index + 1 < len(present_rows):
                candidate = present_rows[index + 1]
                if header_like(candidate) and self._has_horizontal_merge(grid, headers[-1]):
                    headers.append(candidate)
        return title, title_row, headers

    def _build_header_map(
        self,
        grid: _SheetGrid,
        header_rows: list[int],
    ) -> dict[int, str | None]:
        if not header_rows:
            return {}
        values: dict[int, list[str]] = {}
        for row in header_rows:
            for col in self._header_columns(grid, row):
                text = self._header_text(grid, row, col)
                if text is None:
                    continue
                values.setdefault(col, []).append(text)
            # Propagate a horizontally merged header label to covered columns.
            for merge in grid.merges:
                if merge.min_row <= row <= merge.max_row and merge.width > 1:
                    anchor = grid.own.get((merge.min_row, merge.min_col))
                    if anchor is None or anchor.data_type == "blank":
                        continue
                    anchor_text = self._text_of(anchor)
                    for col in range(merge.min_col + 1, merge.max_col + 1):
                        bucket = values.setdefault(col, [])
                        if anchor_text and anchor_text not in bucket:
                            bucket.append(anchor_text)
        return {
            col: "\n".join(parts) if parts else None
            for col, parts in values.items()
        }

    def _header_columns(self, grid: _SheetGrid, row: int) -> set[int]:
        columns = {col for (r, col) in grid.own if r == row}
        for merge in grid.merges:
            if merge.min_row <= row <= merge.max_row:
                columns.update(range(merge.min_col, merge.max_col + 1))
        return columns

    def _header_text(self, grid: _SheetGrid, row: int, col: int) -> str | None:
        data = grid.own.get((row, col))
        if data is not None and data.data_type != "blank":
            return self._text_of(data)
        return None

    @staticmethod
    def _text_of(data: _CellData) -> str:
        if isinstance(data.raw_value, str):
            return data.raw_value
        return data.display_value

    def _row_text(self, values: dict[int, _CellData]) -> str | None:
        parts = []
        for col in sorted(values):
            data = values[col]
            if data.data_type == "blank":
                continue
            text = self._text_of(data)
            if text:
                parts.append(text)
        return " ".join(parts) if parts else None

    @staticmethod
    def _is_title_like(values: dict[int, _CellData]) -> bool:
        # A title row cannot be a normal multi-column header: a header row with
        # two or more labels is always treated as a header.  Titles are a single
        # label (usually merged across the table width) followed by a header row.
        if len(values) != 1:
            return False
        for data in values.values():
            if not isinstance(data.raw_value, str):
                return False
            if "\n" in data.raw_value or len(data.raw_value) > MAX_TITLE_TEXT_LENGTH:
                return False
        return True

    @staticmethod
    def _is_header_like(values: dict[int, _CellData]) -> bool:
        if not values:
            return False
        for data in values.values():
            if not isinstance(data.raw_value, str):
                return False
            if "\n" in data.raw_value:
                return False
            if len(data.raw_value) > MAX_HEADER_TEXT_LENGTH:
                return False
        recognized = sum(
            data.raw_value.strip().casefold() in _HEADER_LABELS
            for data in values.values()
        )
        # All-text business rows must not disappear as guessed headers.
        return recognized >= min(2, len(values))

    @staticmethod
    def _has_horizontal_merge(grid: _SheetGrid, row: int) -> bool:
        return any(
            merge.min_row <= row <= merge.max_row and merge.width > 1
            for merge in grid.merges
        )

    @staticmethod
    def _split_address(address: str) -> tuple[int, int]:
        match = re.fullmatch(r"([A-Z]+)(\d+)", address)
        if not match:
            raise ValueError(address)
        from openpyxl.utils import column_index_from_string

        return int(match.group(2)), column_index_from_string(match.group(1))
