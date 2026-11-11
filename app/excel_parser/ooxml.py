"""OOXML helpers: style compatibility repair and formula extraction.

The parser must never modify the original workbook.  This module reads the
ZIP container directly, detects the known openpyxl 3.1.x incompatibility
(empty ``<fill/>`` nodes in ``xl/styles.xml``), writes a repaired *temporary*
copy, and scans worksheet XML for formula cells and cached values.
"""

from __future__ import annotations

import os
import posixpath
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from openpyxl.formula.translate import Translator
from dataclasses import dataclass
from pathlib import Path

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
C_TAG = f"{{{MAIN_NS}}}c"
F_TAG = f"{{{MAIN_NS}}}f"
V_TAG = f"{{{MAIN_NS}}}v"
SHEET_TAG = f"{{{MAIN_NS}}}sheet"


class OoxmlError(Exception):
    """Raised when the OOXML container cannot be read or repaired."""


@dataclass(frozen=True, slots=True)
class FormulaInfo:
    text: str
    has_cached_value: bool


@dataclass(frozen=True, slots=True)
class WorkbookIndex:
    """Minimal workbook metadata needed by the parser."""

    sheet_paths: dict[str, str]
    sheet_order: tuple[str, ...]
    styles_xml: bytes


# ---------------------------------------------------------------------------
# style repair


_FILLS_BLOCK_RE = re.compile(
    r"(<(?P<p>[A-Za-z_][\w.-]*:)?fills\b[^>]*>)(?P<body>.*?)(</(?P=p)?fills>)",
    re.S,
)
_EMPTY_FILL_SINGLE_RE = re.compile(r"<(?P<p>[A-Za-z_][\w.-]*:)?fill\b[^>]*/>")
_EMPTY_FILL_PAIR_RE = re.compile(
    r"<(?P<p>[A-Za-z_][\w.-]*:)?fill\b[^>]*>\s*</(?:(?P=p))?fill>"
)


def _replacement(prefix: str | None) -> str:
    p = prefix or ""
    return f'<{p}fill><{p}patternFill patternType="none"/></{p}fill>'


def repair_empty_fills(styles_xml: bytes) -> tuple[bytes, int]:
    """Return styles XML with empty fill nodes replaced by a default fill.

    ``openpyxl`` 3.1.5 raises ``TypeError: expected Fill`` for a bare
    ``<fill/>``.  An empty fill is semantically the default "none" pattern, so
    adding ``<patternFill patternType="none"/>`` preserves the intended style
    while keeping every other byte (fonts, borders, number formats) intact.
    """

    try:
        text = styles_xml.decode("utf-8")
    except UnicodeDecodeError as exc:  # pragma: no cover - defensive
        raise OoxmlError("styles.xml 不是 UTF-8 编码") from exc

    counter = {"count": 0}

    def fix_block(match: re.Match[str]) -> str:
        body = match.group("body")
        body, count_single = _EMPTY_FILL_SINGLE_RE.subn(
            lambda m: _replacement(m.group("p")), body
        )
        body, count_pair = _EMPTY_FILL_PAIR_RE.subn(
            lambda m: _replacement(m.group("p")), body
        )
        counter["count"] += count_single + count_pair
        return match.group(1) + body + match.group(4)

    repaired = _FILLS_BLOCK_RE.sub(fix_block, text)
    if counter["count"] == 0:
        return styles_xml, 0
    return repaired.encode("utf-8"), counter["count"]


def count_empty_fills(styles_xml: bytes) -> int:
    """Cheap probe used to decide whether a repaired copy is required."""

    try:
        text = styles_xml.decode("utf-8")
    except UnicodeDecodeError:
        return 0
    match = _FILLS_BLOCK_RE.search(text)
    if not match:
        return 0
    body = match.group("body")
    return len(_EMPTY_FILL_SINGLE_RE.findall(body)) + len(
        _EMPTY_FILL_PAIR_RE.findall(body)
    )


# ---------------------------------------------------------------------------
# workbook metadata


def _resolve_target(target: str) -> str:
    """Resolve a relationship target to a normalised ZIP member name."""

    target = target.replace("\\", "/")
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join("xl", target))


def read_workbook_index(archive: zipfile.ZipFile) -> WorkbookIndex:
    """Read workbook sheet order and per-sheet worksheet XML paths."""

    try:
        workbook_xml = archive.read("xl/workbook.xml")
        rels_xml = archive.read("xl/_rels/workbook.xml.rels")
    except KeyError as exc:
        raise OoxmlError(f"缺少必要的 OOXML 部件: {exc.args[0]}") from exc
    try:
        styles_xml = archive.read("xl/styles.xml")
    except KeyError:
        styles_xml = b""  # Styles are optional in a valid workbook.

    try:
        workbook_root = ET.fromstring(workbook_xml)
        rels_root = ET.fromstring(rels_xml)
    except ET.ParseError as exc:
        raise OoxmlError("workbook.xml 或关系文件损坏") from exc

    rel_map: dict[str, str] = {}
    for rel in rels_root.iter(f"{{{PKG_REL_NS}}}Relationship"):
        rel_id = rel.get("Id")
        target = rel.get("Target")
        if rel_id and target:
            rel_map[rel_id] = _resolve_target(target)

    paths: dict[str, str] = {}
    order: list[str] = []
    for sheet in workbook_root.iter(SHEET_TAG):
        name = sheet.get("name")
        rel_id = sheet.get(f"{{{REL_NS}}}id")
        if not name or not rel_id:
            continue
        target = rel_map.get(rel_id)
        if not target:
            continue
        order.append(name)
        paths[name] = target
    if not order:
        raise OoxmlError("工作簿中没有可读取的工作表")
    return WorkbookIndex(sheet_paths=paths, sheet_order=tuple(order), styles_xml=styles_xml)


# ---------------------------------------------------------------------------
# formula scan


def scan_formulas(archive: zipfile.ZipFile, sheet_path: str) -> dict[str, FormulaInfo]:
    """Return ``address -> FormulaInfo`` for all formula cells in one sheet.

    Shared-formula slave cells are translated from the master's origin using
    OOXML relative-reference semantics. Formula results are never calculated.
    """

    raw: dict[str, tuple[str, bool, str | None]] = {}
    shared_master: dict[str, tuple[str, str]] = {}
    try:
        with archive.open(sheet_path) as stream:
            for _event, elem in ET.iterparse(stream, events=("end",)):
                if elem.tag != C_TAG:
                    continue
                formula = elem.find(F_TAG)
                if formula is not None:
                    address = elem.get("r")
                    if address:
                        value_node = elem.find(V_TAG)
                        # openpyxl writes ``<v />`` for a formula without a
                        # cached result.  Only a value with text (or an
                        # explicitly typed string/error result) is a cache.
                        has_cached = value_node is not None and (
                            value_node.text is not None
                            or elem.get("t") in {"str", "e"}
                        )
                        raw[address] = (
                            formula.text or "",
                            has_cached,
                            formula.get("si"),
                        )
                        if (
                            formula.get("t") == "shared"
                            and formula.text
                            and formula.get("si") is not None
                        ):
                            shared_master[formula.get("si", "")] = (address, formula.text)
                elem.clear()
    except (KeyError, ET.ParseError) as exc:
        raise OoxmlError(f"工作表 XML 无法读取: {sheet_path}") from exc

    result: dict[str, FormulaInfo] = {}
    for address, (text, has_cached, si) in raw.items():
        if not text and si is not None:
            master = shared_master.get(si)
            if master is None:
                raise OoxmlError("共享公式缺少主公式，无法忠实保留原始公式")
            try:
                origin, formula = master
                text = Translator("=" + formula, origin=origin).translate_formula(address)[1:]
            except Exception as exc:
                raise OoxmlError("共享公式引用无法完整展开") from exc
        result[address] = FormulaInfo(text=text, has_cached_value=has_cached)
    return result


# ---------------------------------------------------------------------------
# repaired copy


class RepairedWorkbook:
    """Context manager that owns the temporary repaired copy.

    When repair is unnecessary the original path is returned and nothing is
    written.  The original file is opened read-only and its bytes are copied
    verbatim except for ``xl/styles.xml``.
    """

    def __init__(self, source: Path, *, needed: bool) -> None:
        self._source = source
        self._needed = needed
        self._temp_path: Path | None = None
        self.repaired_fill_count = 0

    def __enter__(self) -> Path:
        if not self._needed:
            return self._source
        fd, temp_name = tempfile.mkstemp(prefix="pns_parser_", suffix=".xlsx")
        os.close(fd)
        self._temp_path = Path(temp_name)
        try:
            with zipfile.ZipFile(self._source, "r") as zin, zipfile.ZipFile(
                self._temp_path, "w"
            ) as zout:
                for info in zin.infolist():
                    data = zin.read(info.filename)
                    if info.filename == "xl/styles.xml":
                        data, self.repaired_fill_count = repair_empty_fills(data)
                    zout.writestr(info, data)
        except Exception:
            self._cleanup()
            raise
        return self._temp_path

    def __exit__(self, exc_type, exc, tb) -> None:
        self._cleanup()

    def _cleanup(self) -> None:
        if self._temp_path is not None:
            try:
                self._temp_path.unlink(missing_ok=True)
            except OSError:  # pragma: no cover - best effort cleanup
                pass
            self._temp_path = None


def open_archive(path: Path) -> zipfile.ZipFile:
    """Open a workbook ZIP with a stable error type."""

    try:
        return zipfile.ZipFile(path, "r")
    except FileNotFoundError as exc:
        raise OoxmlError("文件不存在") from exc
    except PermissionError as exc:
        raise OoxmlError("文件无法读取（权限不足）") from exc
    except (zipfile.BadZipFile, OSError) as exc:
        raise OoxmlError("不是有效的 .xlsx 文件或文件已损坏") from exc
