"""Widget-level tests: the two pages, their empty states and the shell."""

from __future__ import annotations

from PySide6.QtCore import Qt

from app.ui.library_page import LibraryPage
from app.ui.main_window import create_main_window
from app.ui.search_page import SearchPage
from tests.support import build_workbook


def _index(manager, tmp_path, name="a.xlsx", keyword="A-0001"):
    source = build_workbook(tmp_path / name, [["2026-04-13", keyword, "供应商甲", "现货"]])
    summary = manager.add_workbook(str(source))
    assert summary.success, summary.message
    return summary.workbook_id


def test_library_page_lists_every_workbook(qtbot, services, tmp_path):
    for index in range(3):
        _index(services.data_manager, tmp_path, f"w{index}.xlsx", f"A-000{index}")

    page = LibraryPage(services)
    qtbot.addWidget(page)
    page.refresh()

    assert len(page._cards) == 3
    assert page._cards[0].name_label.text() == "w0.xlsx"
    assert page._empty.isHidden()
    assert "3 个工作簿" in page._header.subtitle_label.text()


def test_library_page_empty_state(qtbot, services):
    page = LibraryPage(services)
    qtbot.addWidget(page)

    assert page._cards == []
    assert not page._empty.isHidden()
    assert page._scroll.isHidden()


def test_search_page_renders_one_card_per_hit(qtbot, services, tmp_path):
    _index(services.data_manager, tmp_path)

    page = SearchPage(services)
    qtbot.addWidget(page)
    page.set_workbook_count(1)
    page.input.setText("A-0001")
    page.submit()

    qtbot.waitUntil(lambda: page._last_page is not None, timeout=10_000)
    assert page._last_page.total_count == 1
    assert page._results_layout.count() == 2  # one card + trailing stretch
    assert page._empty.isHidden()
    assert "命中 1 条" in page._status_label.text()


def test_search_page_reports_zero_results(qtbot, services, tmp_path):
    _index(services.data_manager, tmp_path)

    page = SearchPage(services)
    qtbot.addWidget(page)
    page.set_workbook_count(1)
    page.input.setText("nothing-matches-this")
    page.submit()

    qtbot.waitUntil(lambda: page._last_page is not None, timeout=10_000)
    assert page._last_page.total_count == 0
    assert not page._empty.isHidden()
    assert "0 条结果" in page._status_label.text()


def test_search_page_without_workbooks_prompts_the_library(qtbot, services):
    page = SearchPage(services)
    qtbot.addWidget(page)
    page.set_workbook_count(0)

    asked = []
    page.request_library.connect(lambda: asked.append(True))
    page.input.setText("A-0001")
    page.submit()

    assert asked == [True]
    assert page._last_page is None


def test_main_window_switches_between_pages(qtbot, services, tmp_path):
    _index(services.data_manager, tmp_path)

    window = create_main_window(services)
    qtbot.addWidget(window)

    assert window._stack.currentIndex() == 0
    window._switch_page(1)
    assert window._stack.currentIndex() == 1
    assert window._nav_library.isChecked()
    window._switch_page(0)
    assert window._stack.currentIndex() == 0
    assert window._nav_search.isChecked()


def test_main_window_footer_counts_workbooks(qtbot, services, tmp_path):
    _index(services.data_manager, tmp_path, "a.xlsx", "A-0001")
    _index(services.data_manager, tmp_path, "b.xlsx", "A-0002")

    window = create_main_window(services)
    qtbot.addWidget(window)

    assert "2 个工作簿" in window._footprint.text()


def test_nav_buttons_stay_exclusive(qtbot, services):
    window = create_main_window(services)
    qtbot.addWidget(window)
    window._nav_library.click()
    assert window._nav_library.isChecked()
    assert not window._nav_search.isChecked()
    assert window._stack.currentIndex() == 1
