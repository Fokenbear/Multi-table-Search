"""Search semantics: per-cell containment, de-duplication, order and paging."""

from __future__ import annotations

import pytest

from app.contracts import ContractError, ErrorCode
from tests.support import build_workbook


def _index(manager, tmp_path, rows, name="a.xlsx", sheet="示例供应商A"):
    source = build_workbook(tmp_path / name, rows, sheet_name=sheet)
    summary = manager.add_workbook(str(source))
    assert summary.success, summary.message
    return summary.workbook_id


def test_keyword_matches_any_cell_of_the_row(manager, services, tmp_path):
    _index(manager, tmp_path, [["2026-04-13", "A-0001", "供应商甲 0.13未", "现货"]])
    search = services.search_service

    assert search.search("A-0001").total_count == 1
    assert search.search("供应商甲").total_count == 1
    assert search.search("现货").total_count == 1
    assert search.search("不存在").total_count == 0


def test_matching_is_case_insensitive_and_trims_surrounding_spaces(manager, services, tmp_path):
    _index(manager, tmp_path, [["2026-04-13", "A-0001", "供应商甲", ""]])
    search = services.search_service

    assert search.search("a-0001").total_count == 1
    assert search.search("  A-0001  ").total_count == 1
    assert search.search("A-0001").keyword == "A-0001"


def test_adjacent_cells_are_never_concatenated(manager, services, tmp_path):
    _index(manager, tmp_path, [["2026-04-13", "AB", "CD", ""]])
    assert services.search_service.search("BCD").total_count == 0


def test_two_hits_in_one_row_are_reported_once(manager, services, tmp_path):
    _index(manager, tmp_path, [["2026-04-13", "A-0001", "供应商甲", "A-0001 有货"]])
    page = services.search_service.search("A-0001")

    assert page.total_count == 1
    assert len(page.results[0].matched_cell_addresses) == 2


def test_results_cover_every_workbook_in_storage_order(manager, services, tmp_path):
    first = _index(manager, tmp_path, [["2026-04-13", "A-0001", "供应商甲", ""]], "one.xlsx")
    second = _index(manager, tmp_path, [["2026-04-13", "A-0001", "供应商乙", ""]], "two.xlsx")

    page = services.search_service.search("A-0001")

    assert page.total_count == 2
    assert [row.workbook_id for row in page.results] == [first, second]
    assert [row.workbook_name for row in page.results] == ["one.xlsx", "two.xlsx"]


def test_paging_reports_stable_offsets(manager, services, tmp_path):
    rows = [["2026-04-13", f"MATCH-{index}", "供应商甲", ""] for index in range(3)]
    _index(manager, tmp_path, rows)
    search = services.search_service

    first = search.search("MATCH", offset=0, limit=2)
    assert first.total_count == 3
    assert len(first.results) == 2
    assert first.page == 1

    second = search.search("MATCH", offset=2, limit=2)
    assert len(second.results) == 1
    assert second.page == 2
    assert second.results[0].row_number == first.results[1].row_number + 1


def test_like_wildcards_are_literal(manager, services, tmp_path):
    _index(manager, tmp_path, [["2026-04-13", "A-0001", "供应商甲", ""]])
    search = services.search_service

    assert search.search("%").total_count == 0
    assert search.search("_").total_count == 0


@pytest.mark.parametrize("keyword", ["", "   "])
def test_blank_keyword_is_rejected(services, keyword):
    with pytest.raises(ContractError) as error:
        services.search_service.search(keyword)
    assert error.value.code == ErrorCode.INVALID_ARGUMENT


def test_invalid_paging_arguments_are_rejected(services):
    with pytest.raises(ContractError):
        services.search_service.search("A", offset=-1)
    with pytest.raises(ContractError):
        services.search_service.search("A", limit=0)
    with pytest.raises(ContractError):
        services.search_service.search("A", limit=10_000)


def test_formatter_emits_the_complete_row(manager, services, tmp_path):
    _index(manager, tmp_path, [["2026-04-13", "A-0001", "供应商甲 0.13未", "现货"]])
    page = services.search_service.search("A-0001")

    text = services.result_formatter.format_results(page.results)

    assert "文件名称：a.xlsx" in text
    assert "工作表名称：示例供应商A" in text
    assert "供应商甲 0.13未" in text
    assert "现货" in text
