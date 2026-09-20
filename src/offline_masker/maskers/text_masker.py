from __future__ import annotations

from offline_masker.domain.enums import SensitiveType


def mask_value(value: str, sensitive_type: SensitiveType) -> str:
    if sensitive_type == SensitiveType.PHONE:
        return value[:3] + "****" + value[-4:] if len(value) >= 7 else "*" * len(value)
    if sensitive_type == SensitiveType.ID_CARD:
        return value[:4] + "*" * max(len(value) - 8, 1) + value[-4:]
    if sensitive_type == SensitiveType.BANK_ACCOUNT:
        digits = "".join(char for char in value if char.isdigit())
        return "****" + digits[-4:] if len(digits) >= 4 else "****"
    if sensitive_type == SensitiveType.EMAIL:
        local, separator, domain = value.partition("@")
        if not separator:
            return _generic_mask(value)
        if len(local) <= 1:
            masked_local = "*"
        elif len(local) == 2:
            masked_local = local[0] + "*"
        else:
            masked_local = local[0] + "*" * (len(local) - 2) + local[-1]
        return f"{masked_local}@{domain}"
    return _generic_mask(value)


def _generic_mask(value: str) -> str:
    if len(value) <= 2:
        return value[:1] + "*" * max(len(value) - 1, 1)
    return value[0] + "*" * (len(value) - 2) + value[-1]

