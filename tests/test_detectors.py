from __future__ import annotations

import unittest

from offline_masker.detectors.regex_detector import detect_regex, valid_chinese_id
from offline_masker.domain.enums import SensitiveType
from offline_masker.maskers.mapping_manager import MappingManager
from offline_masker.maskers.text_masker import mask_value


class DetectorTests(unittest.TestCase):
    def test_chinese_id_requires_valid_date_and_checksum(self) -> None:
        self.assertTrue(valid_chinese_id("11010519491231002X"))
        self.assertFalse(valid_chinese_id("11010519491331002X"))
        self.assertFalse(valid_chinese_id("110105194912310021"))

    def test_regex_detection_finds_phone_email_and_id(self) -> None:
        text = "联系人 13812345678，邮箱 test.user@example.com，证件 11010519491231002X"
        matches = detect_regex(
            text,
            {SensitiveType.PHONE, SensitiveType.EMAIL, SensitiveType.ID_CARD},
        )
        self.assertEqual(
            [match.sensitive_type for match in matches],
            [SensitiveType.PHONE, SensitiveType.EMAIL, SensitiveType.ID_CARD],
        )

    def test_masking_examples(self) -> None:
        self.assertEqual(mask_value("13812345678", SensitiveType.PHONE), "138****5678")
        self.assertEqual(mask_value("test@example.com", SensitiveType.EMAIL), "t**t@example.com")
        self.assertEqual(mask_value("6222021234567890", SensitiveType.BANK_ACCOUNT), "****7890")

    def test_mapping_is_consistent_and_type_scoped(self) -> None:
        manager = MappingManager()
        first = manager.replacement_for(SensitiveType.CUSTOMER, "北京甲公司")
        second = manager.replacement_for(SensitiveType.CUSTOMER, "北京甲公司")
        supplier = manager.replacement_for(SensitiveType.SUPPLIER, "北京甲公司")
        self.assertEqual(first, second)
        self.assertEqual(first, "客户001")
        self.assertEqual(supplier, "供应商001")


if __name__ == "__main__":
    unittest.main()

