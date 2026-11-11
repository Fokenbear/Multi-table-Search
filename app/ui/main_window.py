"""Application shell: rail navigation, two pages, status bar."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QFrame,
    QHBoxLayout,
    QMainWindow,
    QPushButton,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from app.contracts import ApplicationServices, ContractError
from app.ui.library_page import LibraryPage
from app.ui.search_page import SearchPage
from app.ui.style import apply_app_style, plain_label
from app.version import VERSION

APP_TITLE = "多表格统一检索系统"

RAIL_WIDTH = 216


class MainWindow(QMainWindow):
    def __init__(self, services: ApplicationServices) -> None:
        super().__init__(None)
        self._services = services
        self.setWindowTitle(APP_TITLE)
        self.resize(1180, 780)
        self.setMinimumSize(960, 620)

        self._search_page = SearchPage(services, self)
        self._library_page = LibraryPage(services, self)

        root = QWidget(self)
        root.setObjectName("Root")
        self.setCentralWidget(root)
        outer = QHBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        outer.addWidget(self._build_rail(root))

        self._stack = QStackedWidget(root)
        self._stack.addWidget(self._search_page)
        self._stack.addWidget(self._library_page)
        outer.addWidget(self._stack, 1)

        self.statusBar().showMessage("就绪")
        self._wire()
        self._install_shortcuts()
        self._refresh_counts()
        self._search_page.focus_input()

    # -- construction ----------------------------------------------------

    def _build_rail(self, parent: QWidget) -> QFrame:
        rail = QFrame(parent)
        rail.setObjectName("Rail")
        rail.setFixedWidth(RAIL_WIDTH)
        layout = QVBoxLayout(rail)
        layout.setContentsMargins(12, 16, 12, 12)
        layout.setSpacing(4)

        brand = plain_label(APP_TITLE, rail, "Brand")
        brand.setWordWrap(True)
        layout.addWidget(brand)
        hint = plain_label("本地离线检索 · 不上传文件", rail, "BrandHint")
        hint.setWordWrap(True)
        layout.addWidget(hint)
        layout.addSpacing(16)

        self._nav_group = QButtonGroup(self)
        self._nav_group.setExclusive(True)
        self._nav_search = QPushButton("检索", rail)
        self._nav_library = QPushButton("数据源", rail)
        for index, button in enumerate((self._nav_search, self._nav_library)):
            button.setObjectName("NavButton")
            button.setCheckable(True)
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            self._nav_group.addButton(button, index)
            layout.addWidget(button)
        self._nav_search.setChecked(True)

        layout.addStretch(1)
        self._footprint = plain_label("", rail, "Caption")
        self._footprint.setWordWrap(True)
        layout.addWidget(self._footprint)
        version = plain_label(f"v{VERSION}", rail, "Caption")
        layout.addWidget(version)
        return rail

    def _wire(self) -> None:
        self._nav_group.idClicked.connect(self._switch_page)
        self._search_page.message.connect(self._show_message)
        self._search_page.request_library.connect(lambda: self._switch_page(1))
        self._library_page.message.connect(self._show_message)
        self._library_page.changed.connect(self._refresh_counts)

    def _install_shortcuts(self) -> None:
        QShortcut(QKeySequence("Ctrl+F"), self, activated=self._focus_search)
        QShortcut(QKeySequence("Ctrl+1"), self, activated=lambda: self._switch_page(0))
        QShortcut(QKeySequence("Ctrl+2"), self, activated=lambda: self._switch_page(1))

    # -- behaviour -------------------------------------------------------

    def _switch_page(self, index: int) -> None:
        self._stack.setCurrentIndex(index)
        button = self._nav_search if index == 0 else self._nav_library
        button.setChecked(True)
        if index == 0:
            self._search_page.focus_input()
        else:
            self._library_page.refresh()

    def _focus_search(self) -> None:
        self._switch_page(0)
        self._search_page.focus_input()

    def _refresh_counts(self) -> None:
        try:
            workbooks = self._services.data_manager.list_workbooks()
        except ContractError:
            workbooks = ()
        rows = sum(info.row_count for info in workbooks)
        if workbooks:
            self._footprint.setText(f"{len(workbooks)} 个工作簿 · {rows:,} 行")
        else:
            self._footprint.setText("尚未索引工作簿")
        self._search_page.set_workbook_count(len(workbooks))

    def _show_message(self, text: str) -> None:
        self.statusBar().showMessage(text, 6000)


def create_main_window(services: ApplicationServices) -> QMainWindow:
    """Factory used by the composition root and by tests; the window is not shown."""
    if QApplication.instance() is not None:
        apply_app_style()
    return MainWindow(services)
