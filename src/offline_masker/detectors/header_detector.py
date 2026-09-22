from __future__ import annotations

import re

from offline_masker.domain.enums import SensitiveType


HEADER_RULES: list[tuple[SensitiveType, re.Pattern[str]]] = [
    (
        SensitiveType.FINANCIAL_AMOUNT,
        re.compile(r"销售收入|营业收入|收入|销售额|合同金额|采购金额|含税金额|金额|成本|单价|毛利|利润|税额|回款|应收|应付|工资|奖金|薪酬"),
    ),
    (SensitiveType.PHONE, re.compile(r"手机|手机号|联系电话|移动电话")),
    (SensitiveType.ID_CARD, re.compile(r"身份证|证件号码|身份证号")),
    (SensitiveType.EMAIL, re.compile(r"邮箱|电子邮件|e-?mail", re.I)),
    (SensitiveType.BANK_ACCOUNT, re.compile(r"银行账号|银行卡|账户号码|收款账号|付款账号")),
    (SensitiveType.CUSTOMER, re.compile(r"客户|购货方|委托方")),
    (SensitiveType.SUPPLIER, re.compile(r"供应商|供货方|受托方")),
    (SensitiveType.PERSON_NAME, re.compile(r"姓名|联系人|经办人|负责人|员工")),
    (SensitiveType.ADDRESS, re.compile(r"地址|住址|通讯地址|注册地址")),
    (SensitiveType.CONTRACT, re.compile(r"合同编号|合同号|协议编号|协议号")),
    (SensitiveType.COMPANY, re.compile(r"企业名称|公司名称|单位名称|主体名称")),
]


def infer_header_type(value: object) -> tuple[SensitiveType, str] | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    for sensitive_type, pattern in HEADER_RULES:
        if pattern.search(text):
            return sensitive_type, f"表头“{text}”命中关键词规则"
    return None
