"""Small reusable widgets shared by the two pages."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from app.ui.style import plain_label, pill


class PageHeader(QWidget):
    """Page title on the left, actions on the right, separated by a 1px rule."""

    def __init__(self, title: str, subtitle: str = "", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        bar = QFrame(self)
        bar.setObjectName("HeaderBar")
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        row = QHBoxLayout(bar)
        row.setContentsMargins(24, 16, 24, 16)
        row.setSpacing(12)

        text_column = QVBoxLayout()
        text_column.setSpacing(2)
        self.title_label = plain_label(title, bar, "PageTitle")
        text_column.addWidget(self.title_label)
        self.subtitle_label = plain_label(subtitle, bar, "Hint")
        self.subtitle_label.setVisible(bool(subtitle))
        text_column.addWidget(self.subtitle_label)
        row.addLayout(text_column)
        row.addStretch(1)
        self._actions = QHBoxLayout()
        self._actions.setSpacing(8)
        row.addLayout(self._actions)
        outer.addWidget(bar)

    def add_action(self, widget: QWidget) -> None:
        self._actions.addWidget(widget)


class EmptyState(QWidget):
    """Centered placeholder with an optional call to action."""

    def __init__(self, title: str, hint: str = "", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 48, 24, 48)
        layout.setSpacing(6)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title_label = plain_label(title, self, "EmptyTitle")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.title_label)
        self.hint_label = plain_label(hint, self, "EmptyHint")
        self.hint_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.hint_label.setWordWrap(True)
        self.hint_label.setVisible(bool(hint))
        layout.addWidget(self.hint_label)
        layout.addSpacing(8)
        self._button_row = QHBoxLayout()
        self._button_row.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addLayout(self._button_row)

    def add_button(self, button: QPushButton) -> None:
        self._button_row.addWidget(button)


class WorkbookCard(QFrame):
    """One indexed workbook: name, metrics, state and per-row actions."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("Card")
        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(8)

        top = QHBoxLayout()
        top.setSpacing(8)
        self.name_label = plain_label("", self, "SectionTitle")
        self.name_label.setWordWrap(True)
        top.addWidget(self.name_label, 1)
        self.status_pill = pill("", "muted", self)
        top.addWidget(self.status_pill)
        layout.addLayout(top)

        self.meta_label = plain_label("", self, "Meta")
        self.meta_label.setWordWrap(True)
        layout.addWidget(self.meta_label)

        self.path_label = plain_label("", self, "Mono")
        self.path_label.setWordWrap(True)
        self.path_label.setVisible(False)
        layout.addWidget(self.path_label)

        bottom = QHBoxLayout()
        bottom.setSpacing(4)
        bottom.addStretch(1)
        self.replace_button = QPushButton("替换…", self)
        self.replace_button.setObjectName("Ghost")
        self.remove_button = QPushButton("移除", self)
        self.remove_button.setObjectName("Danger")
        bottom.addWidget(self.replace_button)
        bottom.addWidget(self.remove_button)
        layout.addLayout(bottom)
