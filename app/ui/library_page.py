"""Data source page: an unbounded list of indexed workbooks."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QObject, Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from app.contracts import ApplicationServices, ContractError, ImportProgress, ImportSummary
from app.contracts import WorkbookInfo, WorkbookStatus
from app.ui import display
from app.ui.components import EmptyState, PageHeader, WorkbookCard
from app.ui.style import plain_label
from app.ui.workers import CancellableToken, TaskRunner


class _ProgressBridge(QObject):
    """Re-emits worker-thread progress callbacks on the UI thread."""

    progress = Signal(object)


class LibraryPage(QWidget):
    message = Signal(str)
    changed = Signal()

    def __init__(self, services: ApplicationServices, parent=None) -> None:
        super().__init__(parent)
        self._services = services
        self._runner = TaskRunner(self)
        self._cards: list[WorkbookCard] = []
        self._importing = False
        self._cancel_token: CancellableToken | None = None

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self._header = PageHeader("数据源", "")
        self._add_button = QPushButton("添加工作簿", self)
        self._add_button.setObjectName("Primary")
        self._add_button.clicked.connect(self.choose_files)
        self._header.add_action(self._add_button)
        self._refresh_button = QPushButton("刷新", self)
        self._refresh_button.setObjectName("Ghost")
        self._refresh_button.clicked.connect(self.refresh)
        self._header.add_action(self._refresh_button)
        outer.addWidget(self._header)

        body = QWidget(self)
        body.setObjectName("PageBody")
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(24, 16, 24, 16)
        body_layout.setSpacing(12)

        self._progress_frame = QFrame(body)
        self._progress_frame.setObjectName("Card")
        progress_layout = QVBoxLayout(self._progress_frame)
        progress_layout.setContentsMargins(16, 10, 16, 10)
        progress_layout.setSpacing(6)
        progress_top = QHBoxLayout()
        self._progress_label = plain_label("准备导入…", self._progress_frame, "Meta")
        progress_top.addWidget(self._progress_label, 1)
        self._cancel_button = QPushButton("取消", self._progress_frame)
        self._cancel_button.setObjectName("Ghost")
        self._cancel_button.clicked.connect(self._cancel_import)
        progress_top.addWidget(self._cancel_button)
        progress_layout.addLayout(progress_top)
        self._progress_bar = QProgressBar(self._progress_frame)
        self._progress_bar.setRange(0, 0)
        self._progress_bar.setTextVisible(False)
        progress_layout.addWidget(self._progress_bar)
        self._progress_frame.setVisible(False)
        body_layout.addWidget(self._progress_frame)

        self._scroll = QScrollArea(body)
        self._scroll.setWidgetResizable(True)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._host = QWidget(self._scroll)
        self._list_layout = QVBoxLayout(self._host)
        self._list_layout.setContentsMargins(0, 0, 8, 0)
        self._list_layout.setSpacing(8)
        self._list_layout.addStretch(1)
        self._scroll.setWidget(self._host)
        body_layout.addWidget(self._scroll, 1)

        self._empty = EmptyState(
            "还没有索引任何工作簿",
            "支持 .xlsx，数量没有上限。导入后原件不会被修改，"
            "索引保存在本机数据目录，可随时替换或移除。",
            body,
        )
        add_here = QPushButton("添加工作簿", self._empty)
        add_here.setObjectName("Primary")
        add_here.clicked.connect(self.choose_files)
        self._empty.add_button(add_here)
        body_layout.addWidget(self._empty, 1)

        outer.addWidget(body, 1)
        self.refresh()

    # -- listing ---------------------------------------------------------

    def refresh(self) -> None:
        try:
            workbooks = self._services.data_manager.list_workbooks()
        except ContractError as error:
            self.message.emit(error.message)
            return
        self._render(workbooks)

    def _render(self, workbooks: tuple[WorkbookInfo, ...]) -> None:
        for card in self._cards:
            card.setParent(None)
            card.deleteLater()
        self._cards = []
        total_rows = sum(info.row_count for info in workbooks)
        if not workbooks:
            self._scroll.setVisible(False)
            self._empty.setVisible(True)
            self._header.subtitle_label.setText("还没有导入工作簿；数量不设上限")
            self._header.subtitle_label.setVisible(True)
            self.changed.emit()
            return
        self._empty.setVisible(False)
        self._scroll.setVisible(True)
        self._header.subtitle_label.setText(
            f"{len(workbooks)} 个工作簿　{total_rows:,} 行　导入数量不限")
        self._header.subtitle_label.setVisible(True)
        for info in workbooks:
            card = self._make_card(info)
            self._list_layout.insertWidget(self._list_layout.count() - 1, card)
            self._cards.append(card)
        self.changed.emit()

    def _make_card(self, info: WorkbookInfo) -> WorkbookCard:
        card = WorkbookCard(self._host)
        card.name_label.setText(info.original_filename)
        parts = [
            display.format_file_size(info.file_size_bytes),
            f"{info.sheet_count} 个工作表",
            f"{info.row_count:,} 行",
            display.format_upload_time(info.imported_at),
        ]
        if info.status == WorkbookStatus.FAILED and info.last_message:
            parts.append(info.last_message)
        card.meta_label.setText("　·　".join(parts))
        if info.managed_path:
            card.path_label.setText(info.managed_path)
        tone = {
            WorkbookStatus.READY: "ok",
            WorkbookStatus.IMPORTING: "warn",
            WorkbookStatus.FAILED: "danger",
        }.get(info.status, "muted")
        card.status_pill.setText(display.workbook_status_text(info.status))
        card.status_pill.setObjectName(
            {"ok": "PillOk", "warn": "PillWarn", "danger": "PillDanger"}.get(tone, "PillMuted"))
        card.status_pill.style().unpolish(card.status_pill)
        card.status_pill.style().polish(card.status_pill)
        card.replace_button.clicked.connect(lambda _=False, i=info: self._choose_replace(i))
        card.remove_button.clicked.connect(lambda _=False, i=info: self._remove(i))
        enabled = not self._importing and info.status != WorkbookStatus.IMPORTING
        card.replace_button.setEnabled(enabled)
        card.remove_button.setEnabled(enabled)
        return card

    # -- importing -------------------------------------------------------

    def choose_files(self) -> None:
        if self._importing:
            return
        paths, _ = QFileDialog.getOpenFileNames(
            self, "选择要索引的 Excel 工作簿", "", "Excel 工作簿 (*.xlsx)")
        if not paths:
            return
        self._import_many([Path(p) for p in paths], replace_id=None)

    def _choose_replace(self, info: WorkbookInfo) -> None:
        if self._importing:
            return
        path, _ = QFileDialog.getOpenFileName(
            self, f"选择用来替换《{info.original_filename}》的文件", "",
            "Excel 工作簿 (*.xlsx)")
        if not path:
            return
        answer = QMessageBox.question(
            self, "替换工作簿",
            f"导入成功后将替换《{info.original_filename}》的索引。\n"
            "原件不会被修改，旧的托管副本会在提交成功后清理。\n\n继续吗？",
            QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Cancel,
        )
        if answer != QMessageBox.StandardButton.Ok:
            return
        self._import_many([Path(path)], replace_id=info.workbook_id)

    def _import_many(self, paths: list[Path], replace_id: int | None) -> None:
        """Import one file, then continue with the rest of the batch."""
        if not paths:
            self._finish_import()
            return
        path = paths[0]
        rest = paths[1:]
        self._start_import(path, replace_id, rest)

    def _start_import(self, path: Path, replace_id: int | None, rest: list[Path]) -> None:
        self._importing = True
        self._add_button.setEnabled(False)
        self._progress_frame.setVisible(True)
        self._progress_bar.setRange(0, 0)
        self._progress_label.setText(f"正在导入 {path.name}…")
        self._set_cards_enabled(False)
        bridge = _ProgressBridge(self)
        bridge.progress.connect(self._on_progress)
        token = CancellableToken()
        self._cancel_token = token
        manager = self._services.data_manager
        if replace_id is None:
            args = (str(path),)
            fn = manager.add_workbook
        else:
            args = (replace_id, str(path))
            fn = manager.replace_workbook
        self._runner.submit(
            fn, *args,
            on_progress=bridge.progress.emit,
            cancellation=token,
            finished=lambda summary: self._on_imported(summary, rest),
            failed=lambda error: self._on_import_failed(error, rest),
        )

    def _on_progress(self, progress: ImportProgress) -> None:
        self._progress_label.setText(display.import_stage_text(progress.stage) + (
            f"　{progress.message}" if progress.message else ""))
        if progress.percent is None:
            self._progress_bar.setRange(0, 0)
        else:
            self._progress_bar.setRange(0, 100)
            self._progress_bar.setValue(progress.percent)

    def _cancel_import(self) -> None:
        if self._cancel_token is not None:
            self._cancel_token.cancel()
            self._progress_label.setText("正在取消…")

    def _on_imported(self, summary: ImportSummary, rest: list[Path]) -> None:
        if summary.success:
            self.message.emit(
                f"《{summary.message}》导入完成：{summary.imported_sheets} 个工作表，"
                f"{summary.imported_rows:,} 行，用时 {summary.elapsed_seconds:.1f}s")
        else:
            self.message.emit(f"导入失败：{summary.message}")
        self.refresh()
        if rest:
            self._import_many(rest, replace_id=None)
        else:
            self._finish_import()

    def _on_import_failed(self, error: ContractError, rest: list[Path]) -> None:
        self.message.emit(f"导入失败：{error.message}")
        self.refresh()
        if rest:
            self._import_many(rest, replace_id=None)
        else:
            self._finish_import()

    def _finish_import(self) -> None:
        self._importing = False
        self._cancel_token = None
        self._add_button.setEnabled(True)
        self._progress_frame.setVisible(False)
        self._set_cards_enabled(True)

    def _set_cards_enabled(self, enabled: bool) -> None:
        for card in self._cards:
            card.replace_button.setEnabled(enabled)
            card.remove_button.setEnabled(enabled)

    # -- removal ---------------------------------------------------------

    def _remove(self, info: WorkbookInfo) -> None:
        if self._importing:
            return
        answer = QMessageBox.question(
            self, "移除工作簿",
            f"从索引中移除《{info.original_filename}》？\n"
            "原件不受影响，只会删除本机的托管副本和索引。",
            QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Cancel,
        )
        if answer != QMessageBox.StandardButton.Ok:
            return
        self._importing = True
        self._set_cards_enabled(False)
        self._runner.submit(
            self._services.data_manager.remove_workbook, info.workbook_id,
            finished=lambda _result: self._on_removed(info),
            failed=self._on_remove_failed,
        )

    def _on_removed(self, info: WorkbookInfo) -> None:
        self._importing = False
        self.message.emit(f"已移除《{info.original_filename}》")
        self.refresh()

    def _on_remove_failed(self, error: ContractError) -> None:
        self._importing = False
        self.message.emit(f"移除失败：{error.message}")
        self.refresh()
