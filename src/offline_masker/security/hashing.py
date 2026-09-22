from __future__ import annotations

import hashlib
from datetime import date, datetime, time
from decimal import Decimal
import json


def value_kind(value: object) -> str:
    if value is None:
        return "blank"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, datetime):
        return "datetime"
    if isinstance(value, date):
        return "date"
    if isinstance(value, time):
        return "time"
    if isinstance(value, Decimal):
        return "decimal"
    if isinstance(value, int):
        return "integer"
    if isinstance(value, float):
        return "float"
    if isinstance(value, str):
        return "text"
    return type(value).__name__


def _canonical_value(value: object) -> str:
    if isinstance(value, Decimal):
        return format(value, "f")
    if isinstance(value, (date, datetime, time)):
        return value.isoformat()
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)


def value_hash(value: object, kind: str | None = None) -> str:
    payload = f"{kind or value_kind(value)}\0{_canonical_value(value)}"
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def short_fingerprint(value: object, kind: str | None = None) -> str:
    return value_hash(value, kind)[:12]
