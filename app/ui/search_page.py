"""Search page: one keyword box, a paged list of full original rows."""

from __future__ import annotations

from html import escape

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from app.contracts import ApplicationServices, ContractError, SearchPage as SearchPageDTO
from app.contracts import SearchResult
from app.ui import display
from app.ui.components import EmptyState, PageHeader
from app.ui.dialogs import OpenResultDialog, TextContentDialog
from app.ui.style import plain_label, pill
from app.ui.workers import TaskRunner

PAGE_SIZE = 50
COLLAPSED_FIELDS = 6


def _fields_html(result: SearchResult, limit: int | None) -> str:
    pairs = display.build_field_labels(result.cells)
    if limit is not None:
        pairs = pairs[:limit]
    lines = []
    for cell, label in pairs:
        value = escape(display.field_value_text(cell))
        lines.append(
            f'<span style="color:#5B6270">{escape(label)}：</span>'
            f'<span style="color:#14161A">{value}</span>'
        )
    return "<br>".join(lines)


class ResultCard(QFrame):
    """One hit: where it came from, the full row, and the two row actions."""

    def __init__(self, result: SearchResult, services: ApplicationServices, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("Card")
        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        self._result = result
        self._services = services
        self._expanded = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 10)
        layout.setSpacing(8)

        top = QHBoxLayout()
        top.setSpacing(8)
        origin = plain_label(
            f"{result.workbook_name}　·　{result.sheet_name}　·　第 {result.row_number} 行",
            self, "SectionTitle")
        origin.setWordWrap(True)
        top.addWidget(origin, 1)
        top.addWidget(pill("命中 " + "、".join(result.matched_cell_addresses), "ok", self))
        layout.addLayout(top)

        self.body = QLabel(self)
        self.body.setTextFormat(Qt.TextFormat.RichText)
        self.body.setWordWrap(True)
        self.body.setText(_fields_html(result, COLLAPSED_FIELDS))
        self.body.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        layout.addWidget(self.body)

        footer = QHBoxLayout()
        footer.setSpacing(4)
        self._more_button = QPushButton("", self)
        self._more_button.setObjectName("Ghost")
        self._more_button.clicked.connect(self._toggle_fields)
        total = len(display.build_field_labels(result.cells))
        if total > COLLAPSED_FIELDS:
            self._more_button.setText(f"展开全部 {total} 项")
            footer.addWidget(self._more_button)
        footer.addStretch(1)
        copy_button = QPushButton("复制本行", self)
        copy_button.setObjectName("Ghost")
        copy_button.clicked.connect(self._copy_row)
        open_button = QPushButton("打开原始数据", self)
        open_button.clicked.connect(self._open_location)
        footer.addWidget(copy_button)
        footer.addWidget(open_button)
        layout.addLayout(footer)

    def _toggle_fields(self) -> None:
        self._expanded = not self._expanded
        total = len(display.build_field_labels(self._result.cells))
        self.body.setText(_fields_html(self._result, None if self._expanded else COLLAPSED_FIELDS))
        self._more_button.setText("收起" if self._expanded else f"展开全部 {total} 项")

    def _copy_row(self) -> None:
        text = self._services.result_formatter.format_results((self._result,))
        QApplication.clipboard().setText(text)

    def _open_location(self) -> None:
        address = (self._result.matched_cell_addresses or ("",))[0]
        opened = self._services.wps_opener.open_location(
            self._result.workbook_path, self._result.sheet_name, address)
        location = display.result_location_text(self._result)
        if opened.status.value == "located":
            QApplication.clipboard().setText(location)
            return
        OpenResultDialog(opened, location, self).exec()

    def show_full_text(self, label: str, value: str) -> None:
        TextContentDialog(label, value, self).exec()


class SearchPage(QWidget):
    """Keyword entry, paged results and the empty states around them."""

    message = Signal(str)
    request_library = Signal()

    def __init__(self, services: ApplicationServices, parent=None) -> None:
        super().__init__(parent)
        self._services = services
        self._runner = TaskRunner(self)
        self._request_id = 0
        self._keyword = ""
        self._offset = 0
        self._total = 0
        self._page = 0
        self._last_page: SearchPageDTO | None = None
        self._workbook_count = 0

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        header = PageHeader("检索", "跨全部已索引工作簿逐单元格匹配，命中即返回完整原始行")
        self._copy_page_button = QPushButton("复制本页", self)
        self._copy_page_button.setEnabled(False)
        self._copy_page_button.clicked.connect(self._copy_page)
        header.add_action(self._copy_page_button)
        outer.addWidget(header)

        body = QWidget(self)
        body.setObjectName("PageBody")
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(24, 16, 24, 16)
        body_layout.setSpacing(12)

        search_row = QHBoxLayout()
        search_row.setSpacing(8)
        self.input = QLineEdit(body)
        self.input.setObjectName("SearchInput")
        self.input.setPlaceholderText("输入料号或关键词，按 Enter 检索")
        self.input.setClearButtonEnabled(True)
        self.input.returnPressed.connect(self.submit)
        self.search_button = QPushButton("检索", body)
        self.search_button.setObjectName("Primary")
        self.search_button.clicked.connect(self.submit)
        search_row.addWidget(self.input, 1)
        search_row.addWidget(self.search_button)
        body_layout.addLayout(search_row)

        meta_row = QHBoxLayout()
        meta_row.setSpacing(8)
        self._status_label = plain_label("尚未检索", body, "Meta")
        meta_row.addWidget(self._status_label)
        meta_row.addStretch(1)
        self._prev_button = QPushButton("上一页", body)
        self._prev_button.setObjectName("Ghost")
        self._prev_button.clicked.connect(lambda: self._step(-1))
        self._next_button = QPushButton("下一页", body)
        self._next_button.setObjectName("Ghost")
        self._next_button.clicked.connect(lambda: self._step(1))
        self._prev_button.setEnabled(False)
        self._next_button.setEnabled(False)
        meta_row.addWidget(self._prev_button)
        meta_row.addWidget(self._next_button)
        body_layout.addLayout(meta_row)

        self._scroll = QScrollArea(body)
        self._scroll.setWidgetResizable(True)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._results_host = QWidget(self._scroll)
        self._results_layout = QVBoxLayout(self._results_host)
        self._results_layout.setContentsMargins(0, 0, 8, 0)
        self._results_layout.setSpacing(8)
        self._results_layout.addStretch(1)
        self._scroll.setWidget(self._results_host)
        body_layout.addWidget(self._scroll, 1)

        self._empty = EmptyState(
            "还没有索引任何工作簿",
            "先在「数据源」添加一个或多个 .xlsx 工作簿，之后就能在这里检索。",
            body,
        )
        open_library = QPushButton("去添加工作簿", self._empty)
        open_library.setObjectName("Primary")
        open_library.clicked.connect(self.request_library.emit)
        self._empty.add_button(open_library)
        self._empty.setVisible(False)
        body_layout.addWidget(self._empty, 1)

        outer.addWidget(body, 1)
        self._scroll.setVisible(False)

    # -- public API ------------------------------------------------------

    def focus_input(self) -> None:
        self.input.setFocus()

    def set_workbook_count(self, count: int) -> None:
        self._workbook_count = count
        if count == 0:
            self._clear_cards()
            self._empty.setVisible(True)
            self._scroll.setVisible(False)
            self._status_label.setText("尚未索引任何工作簿")
        elif self._last_page is None:
            self._empty.setVisible(False)
            self._scroll.setVisible(True)
            self._status_label.setText(f"已索引 {count} 个工作簿，输入关键词开始检索")

    # -- search ----------------------------------------------------------

    def submit(self) -> None:
        keyword = self.input.text().strip()
        if not keyword:
            self.message.emit("请输入料号或关键词")
            self.input.setFocus()
            return
        if self._workbook_count == 0:
            self.message.emit("还没有索引任何工作簿")
            self.request_library.emit()
            return
        self._keyword = keyword
        self._offset = 0
        self._run_search()

    def _run_search(self) -> None:
        self._request_id += 1
        request_id = self._request_id
        self._status_label.setText("检索中…")
        self.search_button.setEnabled(False)
        self._prev_button.setEnabled(False)
        self._next_button.setEnabled(False)
        self._runner.submit(
            self._services.search_service.search, self._keyword, self._offset, PAGE_SIZE,
            finished=lambda page: self._on_results(request_id, page),
            failed=lambda error: self._on_failed(request_id, error),
        )

    def _on_results(self, request_id: int, page: SearchPageDTO) -> None:
        if request_id != self._request_id:
            return  # a newer keyword already replaced this request
        self._last_page = page
        self._total = page.total_count
        self._page = page.page
        self.search_button.setEnabled(True)
        self._render(page)
        self._update_pager()

    def _on_failed(self, request_id: int, error: ContractError) -> None:
        if request_id != self._request_id:
            return
        self.search_button.setEnabled(True)
        self._status_label.setText("检索失败")
        self.message.emit(error.message)
        self._clear_cards()
        self._show_empty("检索失败", error.message)

    def _render(self, page: SearchPageDTO) -> None:
        self._clear_cards()
        if page.total_count == 0:
            self._scroll.setVisible(False)
            self._show_empty(
                f"没有匹配「{page.keyword}」的记录",
                "匹配规则是逐单元格包含匹配，仅去掉首尾空格；不会模糊纠错或删除连字符。",
            )
            self._status_label.setText("0 条结果")
            return
        self._empty.setVisible(False)
        self._scroll.setVisible(True)
        first = page.offset + 1
        last = page.offset + len(page.results)
        pages = max(1, (page.total_count + PAGE_SIZE - 1) // PAGE_SIZE)
        self._status_label.setText(
            f"「{page.keyword}」命中 {page.total_count} 条　第 {first}–{last} 条　"
            f"第 {page.page}/{pages} 页"
        )
        cards = [ResultCard(r, self._services, self._results_host) for r in page.results]
        layout = self._results_layout
        for card in cards:
            layout.insertWidget(layout.count() - 1, card)
        self._scroll.verticalScrollBar().setValue(0)

    def _clear_cards(self) -> None:
        layout = self._results_layout
        while layout.count() > 1:
            item = layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

    def _show_empty(self, title: str, hint: str) -> None:
        self._empty.title_label.setText(title)
        self._empty.hint_label.setText(hint)
        self._empty.setVisible(True)

    def _update_pager(self) -> None:
        pages = max(1, (self._total + PAGE_SIZE - 1) // PAGE_SIZE) if self._total else 1
        self._prev_button.setEnabled(self._page > 1)
        self._next_button.setEnabled(self._page < pages and self._total > 0)
        self._copy_page_button.setEnabled(bool(self._last_page and self._last_page.results))

    def _step(self, direction: int) -> None:
        if not self._keyword:
            return
        target = self._page + direction
        if target < 1:
            return
        self._offset = (target - 1) * PAGE_SIZE
        self._run_search()

    def _copy_page(self) -> None:
        if not self._last_page or not self._last_page.results:
            return
        text = self._services.result_formatter.format_results(self._last_page.results)
        QApplication.clipboard().setText(text)
        self.message.emit(f"已复制 {len(self._last_page.results)} 条结果")
