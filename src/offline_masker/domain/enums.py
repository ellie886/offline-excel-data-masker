from __future__ import annotations

from enum import StrEnum


class SensitiveType(StrEnum):
    PHONE = "手机号码"
    ID_CARD = "身份证号码"
    EMAIL = "电子邮箱"
    BANK_ACCOUNT = "银行账号"
    PERSON_NAME = "自然人姓名"
    CUSTOMER = "客户名称"
    SUPPLIER = "供应商名称"
    COMPANY = "企业名称"
    ADDRESS = "地址"
    CONTRACT = "合同编号"
    CUSTOM = "自定义敏感词"


class MaskMethod(StrEnum):
    MASK = "掩码"
    CONSISTENT = "一致性编号"
    CUSTOM = "自定义替换"


DEFAULT_METHODS: dict[SensitiveType, MaskMethod] = {
    SensitiveType.PHONE: MaskMethod.MASK,
    SensitiveType.ID_CARD: MaskMethod.MASK,
    SensitiveType.EMAIL: MaskMethod.MASK,
    SensitiveType.BANK_ACCOUNT: MaskMethod.MASK,
    SensitiveType.PERSON_NAME: MaskMethod.CONSISTENT,
    SensitiveType.CUSTOMER: MaskMethod.CONSISTENT,
    SensitiveType.SUPPLIER: MaskMethod.CONSISTENT,
    SensitiveType.COMPANY: MaskMethod.CONSISTENT,
    SensitiveType.ADDRESS: MaskMethod.CONSISTENT,
    SensitiveType.CONTRACT: MaskMethod.CONSISTENT,
    SensitiveType.CUSTOM: MaskMethod.CUSTOM,
}


MAPPING_PREFIXES: dict[SensitiveType, str] = {
    SensitiveType.PERSON_NAME: "人员",
    SensitiveType.CUSTOMER: "客户",
    SensitiveType.SUPPLIER: "供应商",
    SensitiveType.COMPANY: "企业",
    SensitiveType.ADDRESS: "地址",
    SensitiveType.CONTRACT: "合同",
    SensitiveType.CUSTOM: "敏感词",
}

