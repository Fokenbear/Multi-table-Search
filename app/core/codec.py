"""Typed JSON storage. Never deserialize executable objects from a database."""
from dataclasses import asdict
from datetime import date, datetime, time, timedelta
import json

from app.contracts import ParsedCell


def _default(value):
    if isinstance(value, (datetime, date, time)):
        return {"$type": type(value).__name__, "value": value.isoformat()}
    if isinstance(value, timedelta):
        return {"$type": "timedelta", "days": value.days,
                "seconds": value.seconds, "microseconds": value.microseconds}
    raise TypeError(type(value).__name__)


def _object(value):
    kind = value.get("$type")
    if kind in {"datetime", "date", "time"}:
        return {"datetime": datetime, "date": date, "time": time}[kind].fromisoformat(value["value"])
    if kind == "timedelta":
        return timedelta(days=value["days"], seconds=value["seconds"], microseconds=value["microseconds"])
    return value


def dumps(value) -> str:
    return json.dumps(value, default=_default, ensure_ascii=False, allow_nan=False, separators=(",", ":"))


def loads(value):
    return json.loads(value, object_hook=_object)


def cell_json(cell: ParsedCell) -> str:
    return dumps(asdict(cell))


def decode_cells(value: str) -> tuple[ParsedCell, ...]:
    return tuple(ParsedCell(**cell) for cell in loads(value))
