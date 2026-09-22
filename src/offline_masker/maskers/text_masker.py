from __future__ import annotations

from offline_masker.domain.enums import SensitiveType


def mask_value(value: str, sensitive_type: SensitiveType) -> str:
    if sensitive_type == SensitiveType.PHONE:
        return "【已脱敏-手机号】"
    if sensitive_type == SensitiveType.ID_CARD:
        return "【已脱敏-身份证号】"
    if sensitive_type == SensitiveType.BANK_ACCOUNT:
        return "【已脱敏-银行账号】"
    if sensitive_type == SensitiveType.EMAIL:
        return "【已脱敏-邮箱】"
    return _generic_mask(value)


def _generic_mask(value: str) -> str:
    if len(value) <= 2:
        return value[:1] + "*" * max(len(value) - 1, 1)
    return value[0] + "*" * (len(value) - 2) + value[-1]
