"""Plain-text export of search results, independent of Qt."""

from collections import Counter

from app.contracts import SearchResult


class PlainTextResultFormatter:
    def format_results(self, results: tuple[SearchResult, ...]) -> str:
        records = []
        for result in results:
            lines = [
                f"文件名称：{result.workbook_name}",
                f"工作表名称：{result.sheet_name}",
                "表格名称：" + "、".join(t.table_name for t in result.table_refs),
                f"Excel 原始行号：{result.row_number}",
                "命中单元格：" + "、".join(result.matched_cell_addresses),
            ]
            counts = Counter(c.header for c in result.cells if c.header is not None)
            for cell in result.cells:
                if cell.is_supplementary:
                    label = f"补充信息（{cell.column_letter}列）"
                elif cell.header is None:
                    label = f"未命名列（{cell.column_letter}列）"
                elif counts[cell.header] > 1:
                    label = f"{cell.header}（{cell.column_letter}列）"
                else:
                    label = cell.header
                value = "无" if cell.data_type == "blank" else cell.display_value
                if cell.data_type == "formula" and cell.formula_has_cached_value is False:
                    value += "（公式无缓存值）"
                lines.append(f"{label}：{value}")
            records.append("\n".join(lines))
        return "\n\n".join(records)
