from __future__ import annotations

import re

from offline_masker.domain.models import CustomTerm


def find_custom_term(text: str, term: CustomTerm) -> list[tuple[int, int]]:
    if not term.original:
        return []
    flags = 0 if term.case_sensitive else re.IGNORECASE
    return [(match.start(), match.end()) for match in re.finditer(re.escape(term.original), text, flags)]

