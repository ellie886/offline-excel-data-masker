from __future__ import annotations

import unittest
from io import BytesIO
from zipfile import ZipFile

from offline_masker.domain.exceptions import UnsafeWorkbookError
from offline_masker.security.file_guard import validate_xlsx_package


class FileGuardTests(unittest.TestCase):
    def test_rejects_wrong_extension(self) -> None:
        with self.assertRaises(UnsafeWorkbookError):
            validate_xlsx_package(b"PK", "data.xls")

    def test_rejects_zip_traversal_entry(self) -> None:
        buffer = BytesIO()
        with ZipFile(buffer, "w") as archive:
            archive.writestr("[Content_Types].xml", "<Types/>")
            archive.writestr("xl/workbook.xml", "<workbook/>")
            archive.writestr("_rels/.rels", "<Relationships/>")
            archive.writestr("../outside.txt", "unsafe")
        with self.assertRaises(UnsafeWorkbookError):
            validate_xlsx_package(buffer.getvalue(), "data.xlsx")


if __name__ == "__main__":
    unittest.main()

