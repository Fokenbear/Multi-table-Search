"""Display formatting for UI labels and values.

Rules come from docs/ARCHITECTURE.md「格式器」section and the product spec:
duplicate headers get a `（列字母列）` suffix, unnamed columns get stable labels,
blanks show 「无」, 0/False keep their real text, missing formula cache is stated
explicitly.  Nothing here mutates data; display_value/raw_value stay untouched.
"""

from __future__ import annotations

from datetime import datetime

from app.contracts import (
    ImportStage,
    ParsedCell,
    SearchResult,
    WorkbookInfo,
    WorkbookStatus,
)

BLANK_TEXT = "无"
NO_CACHE_SUFFIX = "（公式无缓存值，显示原文）"
_MISSING = "—"


def build_field_labels(cells: tuple[ParsedCell, ...]) -> tuple[tuple[ParsedCell, str], ...]:
    """Return (cell, display label) pairs, in column order.

    Duplicate non-empty headers get `（X列）`; unnamed columns use
    「未命名列（X列）」 or 「补充信息（X列）」.
    """
    counts: dict[str, int] = {}
    for cell in cells:
        if cell.header:
            counts[cell.header] = counts.get(cell.header, 0) + 1

    labeled: list[tuple[ParsedCell, str]] = []
    for cell in cells:
        if cell.header:
            if counts[cell.header] > 1:
                label = f"{cell.header}（{cell.column_letter}列）"
            else:
                label = cell.header
        elif cell.is_supplementary:
            label = f"补充信息（{cell.column_letter}列）"
        else:
            label = f"未命名列（{cell.column_letter}列）"
        labeled.append((cell, label))
    return tuple(labeled)


def field_value_text(cell: ParsedCell) -> str:
    """User-facing value text for one cell; blanks become 「无」."""
    if cell.data_type == "blank" or (cell.raw_value is None and not cell.display_value):
        return BLANK_TEXT
    text = cell.display_value
    if text == "":
        return BLANK_TEXT
    if cell.data_type == "formula" and cell.formula_has_cached_value is False:
        return f"{text}{NO_CACHE_SUFFIX}"
    return text


def field_line(cell: ParsedCell, label: str) -> str:
    return f"{label}：{field_value_text(cell)}"


def result_summary_text(result: SearchResult, number: int) -> str:
    addresses = "、".join(result.matched_cell_addresses) or _MISSING
    return (
        f"结果 {number}　{result.workbook_name}　{result.sheet_name}　"
        f"{result.table_name}　第 {result.row_number} 行　命中 {addresses}"
    )


def result_location_text(result: SearchResult) -> str:
    """Copyable location info used when WPS cannot auto-locate."""
    addresses = "、".join(result.matched_cell_addresses) or _MISSING
    return (
        f"文件：{result.workbook_name}\n"
        f"路径：{result.workbook_path}\n"
        f"工作表：{result.sheet_name}\n"
        f"表格：{result.table_name}\n"
        f"Excel 行号：{result.row_number}\n"
        f"命中单元格：{addresses}"
    )


def format_file_size(size_bytes: int | None) -> str:
    if size_bytes is None:
        return _MISSING
    if size_bytes < 1024:
        return f"{size_bytes} B"
    if size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    return f"{size_bytes / (1024 * 1024):.1f} MB"


def format_upload_time(value: datetime | None) -> str:
    """uploaded_at is tz-aware UTC; show local wall clock."""
    if value is None:
        return _MISSING
    if value.tzinfo is not None:
        value = value.astimezone()
    return value.strftime("%Y-%m-%d %H:%M")


def workbook_status_text(status: WorkbookStatus) -> str:
    return {
        WorkbookStatus.READY: "已就绪",
        WorkbookStatus.IMPORTING: "导入中",
        WorkbookStatus.FAILED: "导入失败",
    }.get(status, str(status))


def workbook_headline(info: WorkbookInfo) -> str:
    return info.original_filename or _MISSING


def import_stage_text(stage: ImportStage) -> str:
    return {
        ImportStage.COPYING: "正在复制文件…",
        ImportStage.PARSING: "正在解析工作簿…",
        ImportStage.VALIDATING: "正在校验数据…",
        ImportStage.COMMITTING: "正在提交数据库…",
        ImportStage.COMPLETED: "导入完成",
        ImportStage.FAILED: "导入失败",
        ImportStage.CANCELLED: "导入已取消",
    }.get(stage, str(stage))
