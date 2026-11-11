"""Small modal dialogs: open-location outcome and full-value viewer."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QDialogButtonBox,
    QLabel,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
)

from app.contracts import OpenResult, OpenStatus
from app.ui.style import force_plain_text, plain_label


class TextContentDialog(QDialog):
    """Read-only viewer for long cell text (「展开查看」)."""

    def __init__(self, title: str, text: str, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(520, 320)
        layout = QVBoxLayout(self)
        viewer = QPlainTextEdit(self)
        viewer.setReadOnly(True)
        viewer.setPlainText(text)
        viewer.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
        layout.addWidget(viewer)
        buttons = QDialogButtonBox(self)
        copy_button = QPushButton("复制内容", self)
        close_button = QPushButton("关闭", self)
        buttons.addButton(copy_button, QDialogButtonBox.ButtonRole.ActionRole)
        buttons.addButton(close_button, QDialogButtonBox.ButtonRole.RejectRole)
        copy_button.clicked.connect(lambda: QApplication.clipboard().setText(text))
        close_button.clicked.connect(self.accept)
        layout.addWidget(buttons)
        force_plain_text(self)


class OpenResultDialog(QDialog):
    """Shows the real outcome of 打开原始数据; never pretends success.

    OPENED_WITHOUT_LOCATION and FAILED both keep the exact sheet/cell info and
    offer 复制定位信息 so the user can navigate by hand.
    """

    def __init__(
        self,
        open_result: OpenResult,
        location_text: str,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._location_text = location_text
        self.setWindowTitle("打开原始数据")
        self.resize(520, 340)
        layout = QVBoxLayout(self)

        if open_result.status == OpenStatus.OPENED_WITHOUT_LOCATION:
            headline = "已打开文件，但未能自动定位到目标单元格。"
        elif open_result.status == OpenStatus.FAILED:
            headline = "未能打开原始文件。"
        else:
            headline = "已定位到目标位置。"

        title = plain_label(headline, self)
        title.setWordWrap(True)
        layout.addWidget(title)

        detail_lines = [open_result.message or ""]
        if open_result.error_code is not None:
            detail_lines.append(f"错误代码：{open_result.error_code.value}")
        detail = plain_label("\n".join(line for line in detail_lines if line), self)
        detail.setWordWrap(True)
        layout.addWidget(detail)

        location = QPlainTextEdit(self)
        location.setReadOnly(True)
        location.setPlainText(location_text)
        location.setFixedHeight(140)
        layout.addWidget(location)

        buttons = QDialogButtonBox(self)
        copy_button = QPushButton("复制定位信息", self)
        close_button = QPushButton("关闭", self)
        buttons.addButton(copy_button, QDialogButtonBox.ButtonRole.ActionRole)
        buttons.addButton(close_button, QDialogButtonBox.ButtonRole.RejectRole)
        copy_button.clicked.connect(lambda: QApplication.clipboard().setText(location_text))
        close_button.clicked.connect(self.accept)
        layout.addWidget(buttons)

        self.setWindowModality(Qt.WindowModality.WindowModal)
        force_plain_text(self)
