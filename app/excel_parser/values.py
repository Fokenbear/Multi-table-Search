"""Cell value classification and Excel-like display formatting.

The parser keeps the original value separate from the display text.  Search
uses ``display_value``; raw values preserve Python scalar types for storage.
No value is trimmed, split or rewritten here.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import TypeAlias

from app.contracts import CellDataType, RawValue

_ERROR_VALUES = {
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#GETTING_DATA",
    "#SPILL!",
    "#CALC!",
    "#FIELD!",
    "#BLOCKED!",
    "#UNKNOWN!",
    "#CONNECT!",
}

_AM_PM_RE = re.compile(r"(AM/PM|A/P)", re.IGNORECASE)
_QUOTED_RE = re.compile(r'"[^"]*"')

FormatFallback: TypeAlias = bool


@dataclass(frozen=True, slots=True)
class ParsedValue:
    raw_value: RawValue
    display_value: str
    data_type: CellDataType
    number_format: str
    used_format_fallback: bool = False


def classify_value(value: RawValue, number_format: str) -> ParsedValue:
    """Classify a non-formula value and build its display text."""

    if value is None:
        return ParsedValue(None, "", "blank", number_format)
    if isinstance(value, bool):
        return ParsedValue(value, "TRUE" if value else "FALSE", "boolean", number_format)
    if isinstance(value, datetime):
        text, fallback = _format_temporal(value, number_format)
        return ParsedValue(value, text, "datetime", number_format, fallback)
    if isinstance(value, date):
        text, fallback = _format_temporal(value, number_format)
        return ParsedValue(value, text, "date", number_format, fallback)
    if isinstance(value, time):
        text, fallback = _format_temporal(value, number_format)
        return ParsedValue(value, text, "time", number_format, fallback)
    if isinstance(value, timedelta):
        text, fallback = _format_temporal(value, number_format)
        return ParsedValue(value, text, "duration", number_format, fallback)
    if isinstance(value, (int, float)):
        text, fallback = format_number(value, number_format)
        return ParsedValue(value, text, "number", number_format, fallback)
    if isinstance(value, str):
        if value in _ERROR_VALUES:
            return ParsedValue(value, value, "error", number_format)
        return ParsedValue(value, value, "text", number_format)
    # RawValue is a closed union; keep an explicit, lossless fallback.
    text = str(value)
    return ParsedValue(text, text, "text", number_format, True)


def formula_display(value: RawValue, number_format: str) -> tuple[str, bool]:
    """Display text for a formula's cached typed value."""

    if value is None:
        return "", False
    parsed = classify_value(value, number_format)
    return parsed.display_value, parsed.used_format_fallback


# ---------------------------------------------------------------------------
# general formatting


def _format_general(value: int | float) -> str:
    if isinstance(value, int):
        return str(value)
    if value != value:  # NaN cannot come from xlsx but keep it explicit
        return "#NUM!"
    if value in (float("inf"), float("-inf")):
        return "#NUM!"
    if value.is_integer() and abs(value) < 1e16:
        return str(int(value))
    text = format(value, ".15g")
    if "e" in text or "E" in text:
        text = text.upper()
    return text


# ---------------------------------------------------------------------------
# date / time formats


def _has_date_tokens(fmt: str) -> bool:
    if not fmt:
        return False
    sans_literals = _QUOTED_RE.sub("", fmt)
    sans_literals = re.sub(r"\\\\.", "", sans_literals)
    sans_literals = re.sub(r"\[[^\]]*\]", "", sans_literals)
    letters = re.findall(r"[A-Za-z]", sans_literals)
    lowered = "".join(letters).lower()
    if "am/pm" in sans_literals.lower() or "a/p" in sans_literals.lower():
        return True
    if any(ch in lowered for ch in "ydhs"):
        return True
    # A bare month/minute token is treated as a date only next to another
    # temporal token, avoiding false positives for ordinary number formats.
    return False


def _format_temporal(value: date | datetime | time | timedelta, fmt: str) -> tuple[str, bool]:
    if isinstance(value, timedelta):
        return _format_duration(value, fmt)
    if fmt and ";" in fmt:
        fmt = _split_sections(fmt)[0]
    if not fmt or fmt == "General" or not _has_date_tokens(fmt):
        return _format_iso(value), False
    if not isinstance(value, datetime):
        if isinstance(value, date):
            base = datetime(value.year, value.month, value.day)
        else:  # time
            base = datetime(1899, 12, 30, value.hour, value.minute, value.second, value.microsecond)
    else:
        base = value
    try:
        return _render_datetime_format(base, fmt), False
    except ValueError:
        return _format_iso(value), True


def _format_duration(value: timedelta, fmt: str) -> tuple[str, bool]:
    total_seconds = value.total_seconds()
    section = _split_sections(fmt)[0] if fmt else ""
    lowered = section.lower()
    hours = int(total_seconds // 3600)
    minutes = int((total_seconds % 3600) // 60)
    seconds = int(total_seconds % 60)
    if "[h]" in lowered or "[hh]" in lowered:
        if "ss" in lowered:
            return f"{hours}:{minutes:02d}:{seconds:02d}", False
        if "mm" in lowered:
            return f"{hours}:{minutes:02d}", False
        return str(hours), False
    if "[m]" in lowered or "[mm]" in lowered:
        total_minutes = int(total_seconds // 60)
        if "ss" in lowered:
            return f"{total_minutes}:{seconds:02d}", False
        return str(total_minutes), False
    if "[s]" in lowered or "[ss]" in lowered:
        return str(int(total_seconds)), False
    return f"{hours}:{minutes:02d}:{seconds:02d}", False


def _format_iso(value: date | datetime | time) -> str:
    if isinstance(value, datetime):
        if value.microsecond:
            return value.isoformat(sep=" ", timespec="seconds")
        return value.isoformat(sep=" ")
    return value.isoformat()


def _render_datetime_format(dt: datetime, fmt: str) -> str:
    out: list[str] = []
    i = 0
    length = len(fmt)
    am_pm_format = bool(_AM_PM_RE.search(fmt))
    while i < length:
        ch = fmt[i]
        if ch == '"':
            end = fmt.find('"', i + 1)
            if end == -1:
                out.append(fmt[i + 1 :])
                break
            out.append(fmt[i + 1 : end])
            i = end + 1
            continue
        if ch == "\\":
            if i + 1 < length:
                out.append(fmt[i + 1])
            i += 2
            continue
        if ch == "_":
            out.append(" ")
            i += 2
            continue
        if ch == "*":
            i += 2
            continue
        if ch == "[":
            end = fmt.find("]", i + 1)
            if end == -1:
                out.append(fmt[i:])
                break
            inner = fmt[i + 1 : end].strip().lower()
            if inner in {"h", "hh"}:
                out.append(str(int(dt.hour)))
            elif inner in {"m", "mm"}:
                out.append(str(int(dt.minute)))
            elif inner in {"s", "ss"}:
                out.append(str(int(dt.second)))
            i = end + 1
            continue
        upper = ch.upper()
        if upper == "A" and fmt[i : i + 5].upper() in {"AM/PM", "A/P"}:
            token = fmt[i : i + 5] if fmt[i : i + 5].upper() == "AM/PM" else fmt[i : i + 3]
            if len(token) == 5:
                out.append("AM" if dt.hour < 12 else "PM")
            else:
                out.append("A" if dt.hour < 12 else "P")
            i += len(token)
            continue
        if upper in {"Y", "M", "D", "H", "S"}:
            start = i
            while i < length and fmt[i].upper() == upper:
                i += 1
            count = i - start
            if upper == "Y":
                out.append(f"{dt.year:04d}" if count >= 3 else f"{dt.year % 100:02d}")
            elif upper == "D":
                if count == 1:
                    out.append(str(dt.day))
                elif count == 2:
                    out.append(f"{dt.day:02d}")
                elif count == 3:
                    out.append(dt.strftime("%a"))
                else:
                    out.append(dt.strftime("%A"))
            elif upper == "H":
                hour = dt.hour
                if am_pm_format:
                    hour = hour % 12 or 12
                out.append(f"{hour:02d}" if count >= 2 else str(hour))
            elif upper == "S":
                out.append(f"{dt.second:02d}" if count >= 2 else str(dt.second))
            else:  # M
                if _minute_context(fmt, start):
                    out.append(f"{dt.minute:02d}" if count >= 2 else str(dt.minute))
                else:
                    if count == 1:
                        out.append(str(dt.month))
                    elif count == 2:
                        out.append(f"{dt.month:02d}")
                    elif count == 3:
                        out.append(dt.strftime("%b"))
                    else:
                        out.append(dt.strftime("%B"))
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def _minute_context(fmt: str, index: int) -> bool:
    """True when the m token is a minute, not a month."""

    before = _QUOTED_RE.sub("", fmt[:index])
    after = _QUOTED_RE.sub("", fmt[index + 1 :])
    before = re.sub(r"\\\\.", "", before)
    after = re.sub(r"\\\\.", "", after)
    if re.search(r"[hH]", before):
        return True
    if re.search(r"[sS]", after) and not re.search(r"[dDyY]", after):
        return True
    return False


# ---------------------------------------------------------------------------
# number formats


def _split_sections(fmt: str) -> list[str]:
    sections: list[str] = []
    buf: list[str] = []
    in_quotes = False
    in_brackets = False
    i = 0
    while i < len(fmt):
        ch = fmt[i]
        if ch == "\\" and i + 1 < len(fmt):
            buf.append(fmt[i : i + 2])
            i += 2
            continue
        if ch == '"':
            in_quotes = not in_quotes
        elif ch == "[" and not in_quotes:
            in_brackets = True
        elif ch == "]" and not in_quotes:
            in_brackets = False
        elif ch == ";" and not in_quotes and not in_brackets:
            sections.append("".join(buf))
            buf = []
            i += 1
            continue
        buf.append(ch)
        i += 1
    sections.append("".join(buf))
    return sections


def _strip_bracket_tokens(fmt: str) -> str:
    return re.sub(r"\[[^\]]*\]", "", fmt)


def _literal_text(fmt: str) -> str:
    """Convert a non-numeric format fragment to its literal output."""

    out: list[str] = []
    i = 0
    while i < len(fmt):
        ch = fmt[i]
        if ch == '"':
            end = fmt.find('"', i + 1)
            if end == -1:
                out.append(fmt[i + 1 :])
                break
            out.append(fmt[i + 1 : end])
            i = end + 1
        elif ch == "\\":
            if i + 1 < len(fmt):
                out.append(fmt[i + 1])
            i += 2
        elif ch == "_":
            out.append(" ")
            i += 2
        elif ch == "*":
            i += 2
        elif ch == "[":
            end = fmt.find("]", i + 1)
            i = len(fmt) if end == -1 else end + 1
        else:
            out.append(ch)
            i += 1
    return "".join(out)


def format_number(value: int | float, fmt: str) -> tuple[str, FormatFallback]:
    """Format a number, returning ``(display, used_fallback)``."""

    if not fmt or fmt.casefold() == "general" or fmt == "@":
        return _format_general(value), False
    try:
        return _format_number_strict(value, fmt), False
    except (ValueError, InvalidOperation, OverflowError):
        return _format_general(value), True


def _format_number_strict(value: int | float, fmt: str) -> str:
    if re.search(r"\[(?:[<>=])", fmt):
        raise ValueError("conditional format requires fallback")
    sections = _split_sections(fmt)
    if len(sections) > 4:
        raise ValueError("too many sections")
    text_section = sections[3] if len(sections) == 4 else None
    if text_section is not None and isinstance(value, str):
        return _literal_text(text_section).replace("@", str(value))

    if not isinstance(value, (int, float)):
        raise ValueError("not numeric")

    section_index = 0
    if value < 0:
        if len(sections) >= 2:
            section_index = 1
        elif len(sections) == 1:
            section_index = 0
    elif value == 0 and len(sections) >= 3:
        section_index = 2
    section = sections[section_index]
    if not re.search(r"[0#?]", section):
        # A literal-only section (rare) still has a defined output.
        return _literal_text(section)
    if "E" in _strip_bracket_tokens(section).upper() or "?" in section:
        raise ValueError("scientific/optional format not supported")

    percent = "%" in _strip_bracket_tokens(section)
    number = float(value)
    if percent:
        number *= 100

    clean_section = _strip_bracket_tokens(section)
    core_match = re.search(r"[0#.,]+", clean_section)
    if not core_match:
        raise ValueError("no numeric core")
    core = core_match.group(0)
    if core.count(".") > 1:
        raise ValueError("multiple decimal points")
    integer_part, _, fractional_part = core.partition(".")
    decimals = sum(1 for ch in fractional_part if ch in "0#")
    mandatory_decimals = sum(1 for ch in fractional_part if ch == "0")
    thousands = "," in integer_part
    integer_zeros = integer_part.count("0")
    scaling_commas = 0
    stripped_int = integer_part.rstrip(",")
    if stripped_int != integer_part:
        scaling_commas = integer_part[len(stripped_int) :].count(",")
    if scaling_commas:
        number /= 1000**scaling_commas

    negative = number < 0
    magnitude = abs(number)
    if decimals:
        quant = Decimal(1).scaleb(-decimals)
        magnitude = float(Decimal(str(magnitude)).quantize(quant, rounding=ROUND_HALF_UP))
    rounded = f"{magnitude:,.{decimals}f}" if thousands else f"{magnitude:.{decimals}f}"
    if decimals == 0:
        rounded = rounded.split(".")[0]

    integer_text, _, frac_text = rounded.partition(".")
    integer_text = integer_text.replace(",", "")
    if integer_zeros > len(integer_text):
        integer_text = integer_text.zfill(integer_zeros)
    if thousands:
        sign = ""
        if integer_text.startswith("-"):
            sign, integer_text = "-", integer_text[1:]
        groups: list[str] = []
        while len(integer_text) > 3:
            groups.insert(0, integer_text[-3:])
            integer_text = integer_text[:-3]
        groups.insert(0, integer_text)
        integer_text = sign + ",".join(groups)
    if decimals:
        while len(frac_text) > mandatory_decimals and frac_text.endswith("0"):
            frac_text = frac_text[:-1]
        if mandatory_decimals:
            frac_text = (frac_text or "").ljust(mandatory_decimals, "0")
        rounded = f"{integer_text}.{frac_text}" if frac_text else integer_text
    else:
        rounded = integer_text

    # Build prefix/suffix around the numeric core, replacing the placeholders.
    start, end = core_match.span()
    prefix = _literal_text(clean_section[:start])
    suffix = _literal_text(clean_section[end:])
    if negative and section_index == 0 and not prefix.strip().startswith("-"):
        rounded = "-" + rounded
    result = prefix + rounded + suffix
    return result
