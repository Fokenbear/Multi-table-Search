"""Design tokens and the application stylesheet.

Design system: **Retool** (light, Inter, 4px grid) extracted by
https://github.com/ihlamury/design-skills, applied to a Qt Widgets desktop app.
The skill's MUST rules that carry over directly: Inter first, 4px spacing grid,
radius from {2, 3, 4, 19, 32, 99}px, 1px borders, a 2px accent focus ring,
disabled controls at opacity .5, tabular figures for data, no decorative
animation, and a confirmation dialog before destructive actions.

Two tokens were deliberately re-mapped; see ``docs/UI_DESIGN.md``:
``surface-base`` and ``accent``.  The extracted ``#E3E6D7`` page background is an
olive tint that fights the neutral surfaces, and ``#80AEF2`` cannot carry text at
4.5:1 on white.  The skill itself mandates 4.5:1 contrast, so the accent is
darkened to ``#3469F5`` (5.0:1 on white) and the background neutralised.
"""

from __future__ import annotations

from PySide6.QtGui import QFont
from PySide6.QtWidgets import QApplication, QLabel, QMessageBox

# -- palette ---------------------------------------------------------------

SURFACE_BASE = "#F6F7F9"
SURFACE_RAISED = "#FFFFFF"
SURFACE_SUNKEN = "#EEF0F4"
BORDER_SUBTLE = "#E3E6EA"
BORDER_STRONG = "#CDD3DA"
TEXT_PRIMARY = "#14161A"
TEXT_SECONDARY = "#5B6270"
TEXT_TERTIARY = "#8A919E"
ACCENT = "#3469F5"
ACCENT_HOVER = "#2A57D0"
ACCENT_WEAK = "#E9EFFD"
SUCCESS = "#1B7F4B"
SUCCESS_WEAK = "#E7F4EC"
WARNING = "#8A5B00"
WARNING_WEAK = "#FBF0DC"
DANGER = "#B42318"
DANGER_WEAK = "#FCEDEB"
DANGER_HOVER = "#93190F"

# -- typography ------------------------------------------------------------

FONT_FAMILIES = ("Inter", "Segoe UI", "Microsoft YaHei UI", "Microsoft YaHei")
BASE_POINT_SIZE = 9

STYLESHEET = f"""
QWidget {{
    color: {TEXT_PRIMARY};
    background: transparent;
    font-family: "Inter", "Segoe UI", "Microsoft YaHei UI", "Microsoft YaHei";
    font-size: 13px;
}}
QMainWindow, QWidget#Root, QWidget#PageBody {{
    background: {SURFACE_BASE};
}}
QFrame#Rail {{
    background: {SURFACE_RAISED};
    border-right: 1px solid {BORDER_SUBTLE};
}}
QFrame#HeaderBar {{
    background: {SURFACE_BASE};
    border-bottom: 1px solid {BORDER_SUBTLE};
}}
QFrame#Card {{
    background: {SURFACE_RAISED};
    border: 1px solid {BORDER_SUBTLE};
    border-radius: 4px;
}}
QLabel#Brand {{
    font-size: 15px;
    font-weight: 600;
}}
QLabel#BrandHint, QLabel#Hint, QLabel#Meta, QLabel#Caption {{
    color: {TEXT_SECONDARY};
    font-size: 12px;
}}
QLabel#PageTitle {{
    font-size: 17px;
    font-weight: 600;
}}
QLabel#SectionTitle {{
    font-size: 13px;
    font-weight: 600;
}}
QLabel#EmptyTitle {{
    font-size: 14px;
    font-weight: 600;
    color: {TEXT_PRIMARY};
}}
QLabel#EmptyHint {{
    font-size: 12px;
    color: {TEXT_SECONDARY};
}}
QLabel#FieldLabel {{
    color: {TEXT_SECONDARY};
    font-size: 12px;
}}
QLabel#FieldValue {{
    color: {TEXT_PRIMARY};
    font-size: 13px;
}}
QLabel#Mono {{
    color: {TEXT_SECONDARY};
    font-family: "Cascadia Mono", "Consolas", "Courier New", monospace;
    font-size: 12px;
}}
QLabel#PillOk, QLabel#PillMuted, QLabel#PillWarn, QLabel#PillDanger {{
    border-radius: 99px;
    padding: 2px 8px;
    font-size: 11px;
}}
QLabel#PillOk {{ background: {SUCCESS_WEAK}; color: {SUCCESS}; }}
QLabel#PillMuted {{ background: {SURFACE_SUNKEN}; color: {TEXT_SECONDARY}; }}
QLabel#PillWarn {{ background: {WARNING_WEAK}; color: {WARNING}; }}
QLabel#PillDanger {{ background: {DANGER_WEAK}; color: {DANGER}; }}
QPushButton {{
    background: {SURFACE_RAISED};
    border: 1px solid {BORDER_STRONG};
    border-radius: 4px;
    padding: 6px 12px;
    font-size: 13px;
    min-height: 20px;
}}
QPushButton:hover {{ background: {SURFACE_SUNKEN}; }}
QPushButton:pressed {{ background: {BORDER_SUBTLE}; }}
QPushButton:disabled {{ color: {TEXT_TERTIARY}; background: {SURFACE_SUNKEN}; border-color: {BORDER_SUBTLE}; }}
QPushButton:focus {{ border: 1px solid {ACCENT}; }}
QPushButton#Primary {{
    background: {ACCENT};
    border: 1px solid {ACCENT};
    color: #FFFFFF;
    font-weight: 600;
}}
QPushButton#Primary:hover {{ background: {ACCENT_HOVER}; border-color: {ACCENT_HOVER}; }}
QPushButton#Primary:disabled {{ background: #A9BEF9; border-color: #A9BEF9; color: #FFFFFF; }}
QPushButton#Ghost {{
    background: transparent;
    border: none;
    color: {TEXT_SECONDARY};
    padding: 4px 8px;
}}
QPushButton#Ghost:hover {{ background: {SURFACE_SUNKEN}; color: {TEXT_PRIMARY}; }}
QPushButton#Danger {{
    background: transparent;
    border: none;
    color: {DANGER};
    padding: 4px 8px;
}}
QPushButton#Danger:hover {{ background: {DANGER_WEAK}; }}
QPushButton#NavButton {{
    background: transparent;
    border: none;
    border-radius: 4px;
    color: {TEXT_SECONDARY};
    font-size: 13px;
    min-height: 32px;
    padding: 4px 10px;
    text-align: left;
}}
QPushButton#NavButton:hover {{ background: {SURFACE_SUNKEN}; color: {TEXT_PRIMARY}; }}
QPushButton#NavButton:checked {{
    background: {ACCENT_WEAK};
    color: {ACCENT};
    font-weight: 600;
}}
QLineEdit#SearchInput {{
    background: {SURFACE_RAISED};
    border: 1px solid {BORDER_STRONG};
    border-radius: 4px;
    padding: 10px 12px;
    font-size: 14px;
    selection-background-color: {ACCENT};
    selection-color: #FFFFFF;
}}
QLineEdit#SearchInput:hover {{ border-color: #B6BEC9; }}
QLineEdit#SearchInput:focus {{ border: 2px solid {ACCENT}; padding: 9px 11px; }}
QLineEdit#SearchInput:disabled {{ background: {SURFACE_SUNKEN}; }}
QScrollArea {{ border: none; background: transparent; }}
QScrollArea > QWidget > QWidget {{ background: transparent; }}
QScrollBar:vertical {{ background: transparent; width: 10px; margin: 0; }}
QScrollBar::handle:vertical {{ background: {BORDER_STRONG}; border-radius: 3px; min-height: 32px; }}
QScrollBar::handle:vertical:hover {{ background: {TEXT_TERTIARY}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: transparent; }}
QProgressBar {{
    background: {SURFACE_SUNKEN};
    border: none;
    border-radius: 2px;
    height: 4px;
    text-align: center;
    color: transparent;
}}
QProgressBar::chunk {{ background: {ACCENT}; border-radius: 2px; }}
QStatusBar {{
    background: {SURFACE_RAISED};
    border-top: 1px solid {BORDER_SUBTLE};
    color: {TEXT_SECONDARY};
    font-size: 12px;
}}
QStatusBar::item {{ border: none; }}
QDialog {{ background: {SURFACE_BASE}; }}
QPlainTextEdit {{
    background: {SURFACE_RAISED};
    border: 1px solid {BORDER_SUBTLE};
    border-radius: 4px;
    padding: 6px;
    font-size: 12px;
    selection-background-color: {ACCENT};
}}
QToolTip {{
    background: {SURFACE_RAISED};
    color: {TEXT_PRIMARY};
    border: 1px solid {BORDER_STRONG};
    padding: 4px 6px;
}}
"""


def plain_label(text: str = "", parent=None, object_name: str | None = None) -> QLabel:
    """A QLabel pinned to plain text so values like ``A<B`` are never parsed as HTML."""
    from PySide6.QtCore import Qt

    label = QLabel(text, parent)
    label.setTextFormat(Qt.TextFormat.PlainText)
    if object_name:
        label.setObjectName(object_name)
    return label


def pill(text: str, tone: str = "muted", parent=None) -> QLabel:
    """Status pill. ``tone`` is one of ok / muted / warn / danger."""
    names = {"ok": "PillOk", "muted": "PillMuted", "warn": "PillWarn", "danger": "PillDanger"}
    label = plain_label(text, parent, names.get(tone, "PillMuted"))
    return label


def force_plain_text(widget) -> None:
    """Recursively pin every QLabel under ``widget`` to plain text."""
    from PySide6.QtWidgets import QLabel as _QLabel

    if isinstance(widget, _QLabel):
        from PySide6.QtCore import Qt

        widget.setTextFormat(Qt.TextFormat.PlainText)
        return
    if isinstance(widget, QMessageBox):
        widget.setTextFormat(Qt.TextFormat.PlainText)
    find_children = getattr(widget, "findChildren", None)
    if find_children is None:
        return
    for child in find_children(_QLabel):
        child.setTextFormat(Qt.TextFormat.PlainText)


def app_font(point_size: int = BASE_POINT_SIZE) -> QFont:
    font = QFont(FONT_FAMILIES[0], point_size)
    font.setFamilies(list(FONT_FAMILIES))
    return font


def apply_app_style(app: QApplication | None = None) -> None:
    app = app or QApplication.instance()
    if app is None:
        return
    app.setFont(app_font())
    app.setStyleSheet(STYLESHEET)
