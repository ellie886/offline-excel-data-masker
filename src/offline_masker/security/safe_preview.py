from __future__ import annotations

from offline_masker.domain.enums import SensitiveType


def _mask_middle(value: str, left: int, right: int) -> str:
    if len(value) <= left + right:
        return "*" * max(len(value), 1)
    return value[:left] + "*" * min(max(len(value) - left - right, 1), 8) + value[-right:]


def safe_preview(value: object, sensitive_type: SensitiveType) -> str:
    if not value:
        return "[空值]"
    if sensitive_type == SensitiveType.FINANCIAL_AMOUNT:
        return "[财务数值已隐藏]"
    value = str(value)
    if sensitive_type == SensitiveType.PHONE:
        return _mask_middle(value, 3, 4)
    if sensitive_type == SensitiveType.ID_CARD:
        return _mask_middle(value, 4, 4)
    if sensitive_type == SensitiveType.BANK_ACCOUNT:
        return "****" + value[-4:] if len(value) >= 4 else "****"
    if sensitive_type == SensitiveType.EMAIL:
        local, separator, domain = value.partition("@")
        if not separator:
            return _mask_middle(value, 1, 1)
        shown = local[:1] + "***" if local else "***"
        return f"{shown}@{domain}"
    if len(value) == 1:
        return "*"
    if len(value) == 2:
        return value[0] + "*"
    return value[0] + "*" * min(len(value) - 2, 6) + value[-1]
