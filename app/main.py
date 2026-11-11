"""Composition root, offline backup restore and the installed self-test."""

from __future__ import annotations

import argparse
import json
import logging
from logging.handlers import RotatingFileHandler
import os
from pathlib import Path
import sys
import tempfile

from app.version import VERSION


def create_services(paths):
    from app.contracts import ApplicationServices
    from app.core.database import Database
    from app.core.import_service import LocalDataManager
    from app.core.result_formatter import PlainTextResultFormatter
    from app.core.search_service import SqliteSearchService
    from app.excel_parser import WorkbookParser
    from app.windows.wps_opener import WindowsWpsOpener

    database = Database(paths)
    return ApplicationServices(
        LocalDataManager(database, WorkbookParser()),
        SqliteSearchService(database),
        WindowsWpsOpener(),
        PlainTextResultFormatter(),
    )


def _self_test() -> int:
    """Build a synthetic workbook, index it, query it and render the window offscreen."""
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from openpyxl import Workbook
    from PySide6.QtCore import QThreadPool
    from PySide6.QtWidgets import QApplication

    from app.ui.main_window import create_main_window
    from app.windows.windows_paths import AppPaths

    with tempfile.TemporaryDirectory(prefix="mts-selftest-") as temp:
        paths = AppPaths.create(Path(temp) / "data")
        source = Path(temp) / "synthetic.xlsx"
        book = Workbook()
        sheet = book.active
        sheet.title = "示例表"
        sheet.append(["型号", "供应商", "供应商"])
        sheet.append(["SELFTEST-001", "示例供应商甲", None])
        book.save(source)

        services = create_services(paths)
        summary = services.data_manager.add_workbook(str(source))
        assert summary.success, summary.message
        second = services.data_manager.add_workbook(str(source))
        assert not second.success, "重复导入同一份文件必须被拒绝"

        result = services.search_service.search("selftest-001")
        assert result.total_count == 1
        text = services.result_formatter.format_results(result.results)
        assert "供应商（C列）：无" in text, text

        app = QApplication.instance() or QApplication([])
        window = create_main_window(services)
        window.show()
        QThreadPool.globalInstance().waitForDone()
        app.processEvents()
        assert window._stack.currentIndex() == 0
        window.close()
        print(json.dumps({
            "version": VERSION,
            "self_test": "passed",
            "parser": "real",
            "database": "real",
            "ui": "offscreen",
            "wps": "not_tested",
        }, ensure_ascii=False))
    return 0


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="多表格统一检索系统")
    parser.add_argument("--version", action="version", version=VERSION)
    parser.add_argument("--data-dir", help="指定数据目录（默认 LocalAppData）")
    parser.add_argument("--self-test", action="store_true", help="用合成数据在临时目录做安装自检")
    parser.add_argument("--restore-backup", help="离线恢复 SQLite 备份，恢复前自动备份当前库")
    args = parser.parse_args(argv)
    if args.self_test:
        return _self_test()

    from PySide6.QtCore import QLockFile, QThreadPool, Qt
    from PySide6.QtGui import QGuiApplication
    from PySide6.QtWidgets import QApplication, QMessageBox

    from app.windows.windows_paths import AppPaths

    paths = AppPaths.create(args.data_dir)
    handler = RotatingFileHandler(
        paths.root / "logs" / "app.log", maxBytes=2_000_000, backupCount=3, encoding="utf-8")
    logging.basicConfig(level=logging.INFO, handlers=[handler],
                        format="%(asctime)s %(levelname)s %(name)s %(message)s")
    QGuiApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv[:1])
    app.setApplicationName("MultiTableSearch")
    app.setApplicationVersion(VERSION)

    lock = QLockFile(str(paths.root / "settings" / "instance.lock"))
    lock.setStaleLockTime(0)
    if not lock.tryLock(0):
        QMessageBox.information(None, "程序已运行", "该数据目录已有程序实例运行，请切换到原窗口。")
        return 1
    try:
        services = create_services(paths)
        if args.restore_backup:
            previous = services.data_manager.database.restore(args.restore_backup)
            print("恢复完成，恢复前备份：" + str(previous))
            return 0
        from app.ui.main_window import create_main_window
        window = create_main_window(services)
        window.show()
        result = app.exec()
        QThreadPool.globalInstance().waitForDone()
        return result
    except Exception as exc:  # noqa: BLE001 - top level guard shows a readable message
        logging.exception("应用启动或维护失败")
        QMessageBox.critical(
            None, "无法启动", f"{exc}\n请检查用户数据目录，详细原因见 logs/app.log。")
        return 1
    finally:
        lock.unlock()


if __name__ == "__main__":
    raise SystemExit(main())
