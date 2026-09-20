from __future__ import annotations

from collections import defaultdict

from offline_masker.domain.enums import MAPPING_PREFIXES, SensitiveType


class MappingManager:
    """In-memory mapping; no persistence or logging by design."""

    def __init__(self, prefixes: dict[SensitiveType, str] | None = None) -> None:
        self._prefixes = dict(MAPPING_PREFIXES)
        if prefixes:
            self._prefixes.update(prefixes)
        self._counters: dict[SensitiveType, int] = defaultdict(int)
        self._mappings: dict[tuple[SensitiveType, str], str] = {}

    def replacement_for(self, sensitive_type: SensitiveType, original: str) -> str:
        key = (sensitive_type, original)
        if key not in self._mappings:
            self._counters[sensitive_type] += 1
            prefix = self._prefixes.get(sensitive_type, "对象")
            self._mappings[key] = f"{prefix}{self._counters[sensitive_type]:03d}"
        return self._mappings[key]

    def clear(self) -> None:
        self._mappings.clear()
        self._counters.clear()

