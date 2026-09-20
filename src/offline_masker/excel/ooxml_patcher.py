from __future__ import annotations

import posixpath
import xml.etree.ElementTree as ET
from collections import defaultdict
from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

from offline_masker.domain.exceptions import ProcessingConflictError, UnsafeWorkbookError
from offline_masker.domain.models import PatchOperation, PatchResult
from offline_masker.security.hashing import value_hash


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"m": MAIN_NS, "r": DOC_REL_NS, "p": PKG_REL_NS}

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", DOC_REL_NS)


class OoxmlPatcher:
    """Replace confirmed plain-text cells while leaving unrelated package parts untouched."""

    def patch(self, source: bytes, operations: list[PatchOperation]) -> PatchResult:
        if not operations:
            raise ProcessingConflictError("没有已确认的替换项目。")

        with ZipFile(BytesIO(source), "r") as archive:
            names = set(archive.namelist())
            sheet_paths = resolve_sheet_paths(archive)
            shared_strings, rich_shared_indexes = self._shared_strings(archive)
            grouped: dict[str, list[PatchOperation]] = defaultdict(list)
            for operation in operations:
                if operation.sheet not in sheet_paths:
                    raise ProcessingConflictError(f"找不到工作表：{operation.sheet}。")
                grouped[sheet_paths[operation.sheet]].append(operation)

            replacements: dict[str, bytes] = {}
            modified_parts: set[str] = set()
            replacement_count = 0
            modified_cells = 0
            for path, sheet_operations in grouped.items():
                if path not in names:
                    raise UnsafeWorkbookError(f"工作表组件缺失：{path}。")
                root = parse_xml_preserving_namespaces(archive.read(path))
                seen_coordinates: set[str] = set()
                for operation in sheet_operations:
                    if operation.coordinate in seen_coordinates:
                        raise ProcessingConflictError(f"处理计划中存在重复单元格：{operation.sheet}!{operation.coordinate}。")
                    seen_coordinates.add(operation.coordinate)
                    cell = root.find(f".//{{{MAIN_NS}}}c[@r='{operation.coordinate}']")
                    if cell is None:
                        raise ProcessingConflictError(f"找不到待替换单元格：{operation.sheet}!{operation.coordinate}。")
                    if cell.find(f"{{{MAIN_NS}}}f") is not None:
                        raise ProcessingConflictError(f"拒绝修改公式单元格：{operation.sheet}!{operation.coordinate}。")

                    current_value, rich_text = self._cell_text(cell, shared_strings, rich_shared_indexes)
                    if value_hash(current_value) != operation.expected_value_hash:
                        raise ProcessingConflictError(f"单元格内容在确认后发生变化：{operation.sheet}!{operation.coordinate}。")
                    if rich_text:
                        raise ProcessingConflictError(f"目标单元格包含富文本，第一版拒绝修改以避免破坏字符级格式：{operation.sheet}!{operation.coordinate}。")
                    self._set_inline_string(cell, operation.replacement_value)
                    replacement_count += operation.candidate_count
                    modified_cells += 1

                replacements[path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                modified_parts.add(path)

            output = BytesIO()
            with ZipFile(output, "w", compression=ZIP_DEFLATED, allowZip64=True) as target:
                target.comment = archive.comment
                for info in archive.infolist():
                    payload = replacements.get(info.filename, archive.read(info.filename))
                    target.writestr(info, payload)

        return PatchResult(
            output_bytes=output.getvalue(),
            modified_parts=modified_parts,
            modified_cells=modified_cells,
            replacement_count=replacement_count,
        )

    @staticmethod
    def _shared_strings(archive: ZipFile) -> tuple[list[str], set[int]]:
        if "xl/sharedStrings.xml" not in archive.namelist():
            return [], set()
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        values: list[str] = []
        rich_indexes: set[int] = set()
        for index, item in enumerate(root.findall(f"{{{MAIN_NS}}}si")):
            if item.find(f"{{{MAIN_NS}}}r") is not None:
                rich_indexes.add(index)
            values.append("".join(node.text or "" for node in item.iter(f"{{{MAIN_NS}}}t")))
        return values, rich_indexes

    @staticmethod
    def _cell_text(cell: ET.Element, shared_strings: list[str], rich_shared_indexes: set[int]) -> tuple[str, bool]:
        cell_type = cell.attrib.get("t")
        if cell_type == "s":
            value_node = cell.find(f"{{{MAIN_NS}}}v")
            if value_node is None or value_node.text is None:
                return "", False
            try:
                index = int(value_node.text)
                return shared_strings[index], index in rich_shared_indexes
            except (ValueError, IndexError) as exc:
                raise UnsafeWorkbookError("共享字符串索引无效。") from exc
        if cell_type == "inlineStr":
            inline = cell.find(f"{{{MAIN_NS}}}is")
            if inline is None:
                return "", False
            rich = inline.find(f"{{{MAIN_NS}}}r") is not None
            return "".join(node.text or "" for node in inline.iter(f"{{{MAIN_NS}}}t")), rich
        value_node = cell.find(f"{{{MAIN_NS}}}v")
        return (value_node.text or "") if value_node is not None else "", False

    @staticmethod
    def _set_inline_string(cell: ET.Element, replacement: str) -> None:
        cell.attrib["t"] = "inlineStr"
        for tag in (f"{{{MAIN_NS}}}v", f"{{{MAIN_NS}}}is"):
            existing = cell.find(tag)
            if existing is not None:
                cell.remove(existing)
        inline = ET.Element(f"{{{MAIN_NS}}}is")
        extension = cell.find(f"{{{MAIN_NS}}}extLst")
        if extension is None:
            cell.append(inline)
        else:
            cell.insert(list(cell).index(extension), inline)
        text = ET.SubElement(inline, f"{{{MAIN_NS}}}t")
        if replacement[:1].isspace() or replacement[-1:].isspace():
            text.set(f"{{{XML_NS}}}space", "preserve")
        text.text = replacement


def resolve_sheet_paths(archive: ZipFile) -> dict[str, str]:
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    relationships = {
        item.attrib["Id"]: item.attrib["Target"]
        for item in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    result: dict[str, str] = {}
    sheets = workbook_root.find(f"{{{MAIN_NS}}}sheets")
    if sheets is None:
        raise UnsafeWorkbookError("工作簿不包含工作表清单。")
    for sheet in sheets:
        relationship_id = sheet.attrib.get(f"{{{DOC_REL_NS}}}id")
        target = relationships.get(relationship_id or "")
        if not target:
            continue
        path = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
        result[sheet.attrib["name"]] = path
    return result


def parse_xml_preserving_namespaces(payload: bytes) -> ET.Element:
    for _, namespace in ET.iterparse(BytesIO(payload), events=("start-ns",)):
        prefix, uri = namespace
        try:
            ET.register_namespace(prefix or "", uri)
        except ValueError:
            # ElementTree reserves generated ns<number> prefixes. The URI is
            # still preserved even when the exact prefix cannot be registered.
            pass
    parser = ET.XMLParser(target=ET.TreeBuilder(insert_comments=True))
    return ET.fromstring(payload, parser=parser)
