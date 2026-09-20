from __future__ import annotations

import posixpath
from dataclasses import dataclass
from io import BytesIO
from zipfile import BadZipFile, ZipFile

from offline_masker.domain.exceptions import UnsafeWorkbookError


@dataclass(frozen=True, slots=True)
class FileLimits:
    max_file_bytes: int = 100 * 1024 * 1024
    max_uncompressed_bytes: int = 500 * 1024 * 1024
    max_compression_ratio: int = 200
    max_entries: int = 20_000


REQUIRED_PARTS = {"[Content_Types].xml", "xl/workbook.xml", "_rels/.rels"}


def validate_xlsx_package(data: bytes, filename: str, limits: FileLimits | None = None) -> list[str]:
    limits = limits or FileLimits()
    if not filename.lower().endswith(".xlsx"):
        raise UnsafeWorkbookError("第一版只支持 .xlsx 文件。")
    if len(data) > limits.max_file_bytes:
        raise UnsafeWorkbookError("文件超过 100 MB 的安全限制。")
    if not data.startswith(b"PK"):
        raise UnsafeWorkbookError("文件不是有效的未加密 XLSX 压缩包，可能已加密或扩展名不正确。")

    try:
        with ZipFile(BytesIO(data)) as archive:
            infos = archive.infolist()
            if len(infos) > limits.max_entries:
                raise UnsafeWorkbookError("工作簿包含过多内部文件，已拒绝处理。")
            names = {item.filename for item in infos}
            missing = REQUIRED_PARTS - names
            if missing:
                raise UnsafeWorkbookError("工作簿缺少必要的 XLSX 组件。")

            total_uncompressed = 0
            for item in infos:
                normalized = posixpath.normpath(item.filename)
                if item.filename.startswith(("/", "\\")) or normalized == ".." or normalized.startswith("../"):
                    raise UnsafeWorkbookError("工作簿包含不安全的内部路径。")
                total_uncompressed += item.file_size
                if total_uncompressed > limits.max_uncompressed_bytes:
                    raise UnsafeWorkbookError("工作簿解压后超过 500 MB 的安全限制。")
                if item.file_size and item.compress_size == 0:
                    raise UnsafeWorkbookError("工作簿包含异常压缩条目。")
                if item.compress_size and item.file_size / item.compress_size > limits.max_compression_ratio:
                    raise UnsafeWorkbookError("工作簿包含压缩比异常的条目。")

            lowered = {name.lower() for name in names}
            if any("_xmlsignatures/" in name or name.endswith("origin.sigs") for name in lowered):
                raise UnsafeWorkbookError("工作簿带有数字签名，修改会使签名失效，第一版拒绝处理。")
            if any(name.endswith("vbaproject.bin") for name in lowered):
                raise UnsafeWorkbookError("检测到 VBA 宏内容；第一版仅支持不含宏的 .xlsx 文件。")

            warnings: list[str] = []
            if any(name.startswith("xl/externalLinks/") for name in names):
                warnings.append("存在外部工作簿链接：将原样保留，但不会验证链接目标。")
            if any(name.startswith("xl/embeddings/") for name in names):
                warnings.append("存在嵌入对象：将原样保留，但不会扫描其中内容。")
            return warnings
    except BadZipFile as exc:
        raise UnsafeWorkbookError("文件不是有效的 XLSX 压缩包。") from exc

