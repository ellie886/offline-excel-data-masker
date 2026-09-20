from __future__ import annotations

import hashlib


def value_hash(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def short_fingerprint(value: str) -> str:
    return value_hash(value)[:12]

