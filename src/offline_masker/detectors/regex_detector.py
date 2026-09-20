from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date

from offline_masker.domain.enums import SensitiveType


@dataclass(frozen=True, slots=True)
class RegexMatch:
    sensitive_type: SensitiveType
    value: str
    start: int
    end: int
    confidence: float
    basis: str


PHONE_PATTERN = re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")
ID_PATTERN = re.compile(r"(?<![0-9A-Za-z])\d{17}[0-9Xx](?![0-9A-Za-z])")
EMAIL_PATTERN = re.compile(r"(?<![\w.+-])[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+(?![\w.-])")
BANK_PATTERN = re.compile(r"(?<!\d)(?:\d[ -]?){12,19}(?!\d)")

ID_WEIGHTS = (7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2)
ID_CHECK_CODES = "10X98765432"


def valid_chinese_id(value: str) -> bool:
    normalized = value.upper()
    if not re.fullmatch(r"\d{17}[0-9X]", normalized):
        return False
    try:
        date(int(normalized[6:10]), int(normalized[10:12]), int(normalized[12:14]))
    except ValueError:
        return False
    total = sum(int(char) * weight for char, weight in zip(normalized[:17], ID_WEIGHTS, strict=True))
    return ID_CHECK_CODES[total % 11] == normalized[-1]


def detect_regex(text: str, enabled_types: set[SensitiveType], allow_bank_account: bool = False) -> list[RegexMatch]:
    matches: list[RegexMatch] = []
    if SensitiveType.ID_CARD in enabled_types:
        for match in ID_PATTERN.finditer(text):
            if valid_chinese_id(match.group()):
                matches.append(RegexMatch(SensitiveType.ID_CARD, match.group(), match.start(), match.end(), 0.99, "身份证格式、出生日期及校验位有效"))
    if SensitiveType.EMAIL in enabled_types:
        matches.extend(
            RegexMatch(SensitiveType.EMAIL, match.group(), match.start(), match.end(), 0.98, "电子邮箱格式匹配")
            for match in EMAIL_PATTERN.finditer(text)
        )
    if SensitiveType.PHONE in enabled_types:
        matches.extend(
            RegexMatch(SensitiveType.PHONE, match.group(), match.start(), match.end(), 0.97, "中国大陆手机号格式匹配")
            for match in PHONE_PATTERN.finditer(text)
        )
    if allow_bank_account and SensitiveType.BANK_ACCOUNT in enabled_types:
        for match in BANK_PATTERN.finditer(text):
            value = match.group().strip()
            digits = re.sub(r"[ -]", "", value)
            if len(digits) in {11, 18} or not 12 <= len(digits) <= 19:
                continue
            matches.append(RegexMatch(SensitiveType.BANK_ACCOUNT, value, match.start(), match.end(), 0.72, "银行账号列中的 12–19 位数字标识符"))
    return _remove_overlaps(matches)


def _remove_overlaps(matches: list[RegexMatch]) -> list[RegexMatch]:
    priority = {
        SensitiveType.ID_CARD: 4,
        SensitiveType.EMAIL: 3,
        SensitiveType.PHONE: 2,
        SensitiveType.BANK_ACCOUNT: 1,
    }
    accepted: list[RegexMatch] = []
    for item in sorted(matches, key=lambda value: (-priority[value.sensitive_type], value.start, -(value.end - value.start))):
        if any(item.start < other.end and other.start < item.end for other in accepted):
            continue
        accepted.append(item)
    return sorted(accepted, key=lambda value: value.start)

