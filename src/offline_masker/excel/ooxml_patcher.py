from __future__ import annotations

import posixpath
import xml.etree.ElementTree as ET
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

from offline_masker.domain.exceptions import ProcessingConflictError, UnsafeWorkbookError
from offline_masker.domain.models import PatchOperation, PatchResult
from offline_masker.security.hashing import value_hash

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"
ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", DOC_REL_NS)


class OoxmlPatcher:
    """Perform narrow, verified OOXML edits without rebuilding the workbook."""

    def patch(self, source: bytes, operations: list[PatchOperation]) -> PatchResult:
        if not operations:
            raise ProcessingConflictError("没有已确认的替换项目。")
        with ZipFile(BytesIO(source), "r") as archive:
            names = set(archive.namelist())
            sheet_paths = resolve_sheet_paths(archive)
            path_to_sheet = {path: sheet for sheet, path in sheet_paths.items()}
            shared_root, shared_values, rich_indexes = self._shared_strings(archive)
            sheet_roots = {
                path: parse_xml_preserving_namespaces(archive.read(path))
                for path in sheet_paths.values() if path in names
            }
            shared_refs = self._shared_references(sheet_roots, path_to_sheet)
            target_keys = {(item.sheet, item.coordinate) for item in operations}
            grouped: dict[str, list[PatchOperation]] = defaultdict(list)
            for operation in operations:
                path = sheet_paths.get(operation.sheet)
                if not path:
                    raise ProcessingConflictError(f"找不到工作表：{operation.sheet}。")
                grouped[path].append(operation)

            removable_shared: set[int] = set()
            changed_sheets: set[str] = set()
            replacement_count = modified_cells = 0
            for path, items in grouped.items():
                root = sheet_roots.get(path)
                if root is None:
                    raise UnsafeWorkbookError(f"工作表组件缺失：{path}。")
                seen: set[str] = set()
                for operation in items:
                    if operation.coordinate in seen:
                        raise ProcessingConflictError(f"处理计划中存在重复单元格：{operation.sheet}!{operation.coordinate}。")
                    seen.add(operation.coordinate)
                    cell = root.find(f".//{{{MAIN_NS}}}c[@r='{operation.coordinate}']")
                    if cell is None:
                        raise ProcessingConflictError(f"找不到待替换单元格：{operation.sheet}!{operation.coordinate}。")
                    if cell.find(f"{{{MAIN_NS}}}f") is not None:
                        raise ProcessingConflictError(f"拒绝修改公式单元格：{operation.sheet}!{operation.coordinate}。")
                    current, kind, rich, shared_index = self._cell_value(
                        cell, shared_values, rich_indexes, operation.original_kind
                    )
                    if kind != operation.original_kind or value_hash(current, kind) != operation.expected_value_hash:
                        raise ProcessingConflictError(f"单元格内容或类型在确认后发生变化：{operation.sheet}!{operation.coordinate}。")
                    if rich:
                        raise ProcessingConflictError(f"目标单元格包含富文本，拒绝修改：{operation.sheet}!{operation.coordinate}。")
                    if shared_index is not None:
                        other_refs = [
                            (sheet, coordinate) for sheet, coordinate, _ in shared_refs.get(shared_index, [])
                            if (sheet, coordinate) not in target_keys
                        ]
                        if other_refs:
                            locations = "、".join(f"{sheet}!{coordinate}" for sheet, coordinate in other_refs[:8])
                            raise UnsafeWorkbookError(
                                "敏感共享字符串仍被未选中单元格引用，已阻止导出："
                                f"{locations}。请将这些单元格一并确认脱敏。"
                            )
                        removable_shared.add(shared_index)
                    if operation.replacement_kind == "text":
                        self._set_inline_string(cell, str(operation.replacement_value))
                    elif operation.replacement_kind in {"integer", "float", "decimal", "number"}:
                        self._set_numeric_value(cell, operation.replacement_value)
                    else:
                        raise ProcessingConflictError(f"不支持的替换值类型：{operation.replacement_kind}。")
                    replacement_count += operation.candidate_count
                    modified_cells += 1
                    changed_sheets.add(path)

            formula_caches_cleared = 0
            formula_changed_paths: set[str] = set()
            for path, root in sheet_roots.items():
                cleared = self._clear_formula_caches(root)
                if cleared:
                    formula_caches_cleared += cleared
                    formula_changed_paths.add(path)
                    changed_sheets.add(path)

            replacements: dict[str, bytes] = {}
            modified_parts: set[str] = set()
            reasons: dict[str, str] = {}
            shared_removed = 0
            shared_changed_paths: set[str] = set()
            if shared_root is not None and removable_shared:
                referenced_after = self._referenced_shared_indexes(sheet_roots)
                safe_remove = {index for index in removable_shared if index not in referenced_after}
                if safe_remove:
                    shared_changed_paths = self._rebuild_shared_strings(shared_root, safe_remove, sheet_roots)
                    shared_removed = len(safe_remove)
                    replacements["xl/sharedStrings.xml"] = ET.tostring(shared_root, encoding="utf-8", xml_declaration=True)
                    modified_parts.add("xl/sharedStrings.xml")
                    reasons["xl/sharedStrings.xml"] = "删除无引用敏感共享字符串并重建索引"
                    changed_sheets.update(shared_changed_paths)

            calc_changed = False
            if formula_caches_cleared:
                workbook_root = parse_xml_preserving_namespaces(archive.read("xl/workbook.xml"))
                calc = workbook_root.find(f"{{{MAIN_NS}}}calcPr")
                if calc is None:
                    calc = ET.SubElement(workbook_root, f"{{{MAIN_NS}}}calcPr")
                calc.set("calcMode", "auto")
                calc.set("fullCalcOnLoad", "1")
                calc.set("forceFullCalc", "1")
                replacements["xl/workbook.xml"] = ET.tostring(workbook_root, encoding="utf-8", xml_declaration=True)
                modified_parts.add("xl/workbook.xml")
                reasons["xl/workbook.xml"] = "强制打开后完整重新计算"
                calc_changed = True

            for path in changed_sheets:
                replacements[path] = ET.tostring(sheet_roots[path], encoding="utf-8", xml_declaration=True)
                modified_parts.add(path)
                changes: list[str] = []
                if path in grouped:
                    changes.append("用户确认目标单元格")
                if path in formula_changed_paths:
                    changes.append("公式缓存清理")
                if path in shared_changed_paths:
                    changes.append("共享字符串索引同步")
                reasons[path] = "、".join(changes)

            output = BytesIO()
            with ZipFile(output, "w", compression=ZIP_DEFLATED, allowZip64=True) as target:
                target.comment = archive.comment
                for info in archive.infolist():
                    target.writestr(info, replacements.get(info.filename, archive.read(info.filename)))

        return PatchResult(
            output_bytes=output.getvalue(), modified_parts=modified_parts,
            modified_cells=modified_cells, replacement_count=replacement_count,
            shared_strings_removed=shared_removed,
            formula_caches_cleared=formula_caches_cleared,
            calculation_properties_changed=calc_changed,
            allowed_part_changes=reasons,
        )

    @staticmethod
    def _shared_strings(archive: ZipFile) -> tuple[ET.Element | None, list[str], set[int]]:
        if "xl/sharedStrings.xml" not in archive.namelist():
            return None, [], set()
        root = parse_xml_preserving_namespaces(archive.read("xl/sharedStrings.xml"))
        values: list[str] = []
        rich: set[int] = set()
        for index, item in enumerate(root.findall(f"{{{MAIN_NS}}}si")):
            if item.find(f"{{{MAIN_NS}}}r") is not None:
                rich.add(index)
            values.append("".join(node.text or "" for node in item.iter(f"{{{MAIN_NS}}}t")))
        return root, values, rich

    @staticmethod
    def _shared_references(roots: dict[str, ET.Element], names: dict[str, str]) -> dict[int, list[tuple[str, str, ET.Element]]]:
        refs: dict[int, list[tuple[str, str, ET.Element]]] = defaultdict(list)
        for path, root in roots.items():
            for cell in root.findall(f".//{{{MAIN_NS}}}c[@t='s']"):
                node = cell.find(f"{{{MAIN_NS}}}v")
                if node is None or node.text is None:
                    continue
                try:
                    refs[int(node.text)].append((names.get(path, path), cell.attrib.get("r", "?"), cell))
                except ValueError as exc:
                    raise UnsafeWorkbookError("共享字符串索引无效。") from exc
        return refs

    @staticmethod
    def _referenced_shared_indexes(roots: dict[str, ET.Element]) -> set[int]:
        indexes: set[int] = set()
        for root in roots.values():
            for cell in root.findall(f".//{{{MAIN_NS}}}c[@t='s']"):
                node = cell.find(f"{{{MAIN_NS}}}v")
                if node is not None and node.text is not None:
                    indexes.add(int(node.text))
        return indexes

    @staticmethod
    def _rebuild_shared_strings(root: ET.Element, removed: set[int], roots: dict[str, ET.Element]) -> set[str]:
        items = list(root.findall(f"{{{MAIN_NS}}}si"))
        mapping: dict[int, int] = {}
        kept: list[ET.Element] = []
        for old, item in enumerate(items):
            if old not in removed:
                mapping[old] = len(kept)
                kept.append(item)
        for item in items:
            root.remove(item)
        for item in kept:
            root.append(item)
        changed_paths: set[str] = set()
        count = 0
        for path, sheet_root in roots.items():
            for cell in sheet_root.findall(f".//{{{MAIN_NS}}}c[@t='s']"):
                node = cell.find(f"{{{MAIN_NS}}}v")
                if node is None or node.text is None:
                    continue
                old = int(node.text)
                if old not in mapping:
                    raise UnsafeWorkbookError("已删除的共享字符串仍有引用。")
                if mapping[old] != old:
                    node.text = str(mapping[old])
                    changed_paths.add(path)
                count += 1
        root.set("count", str(count))
        root.set("uniqueCount", str(len(kept)))
        return changed_paths

    @staticmethod
    def _cell_value(cell: ET.Element, strings: list[str], rich: set[int], expected_kind: str) -> tuple[object, str, bool, int | None]:
        cell_type = cell.attrib.get("t")
        if cell_type == "s":
            node = cell.find(f"{{{MAIN_NS}}}v")
            try:
                index = int(node.text) if node is not None and node.text is not None else -1
                return strings[index], "text", index in rich, index
            except (ValueError, IndexError) as exc:
                raise UnsafeWorkbookError("共享字符串索引无效。") from exc
        if cell_type == "inlineStr":
            inline = cell.find(f"{{{MAIN_NS}}}is")
            if inline is None:
                return "", "text", False, None
            return "".join(n.text or "" for n in inline.iter(f"{{{MAIN_NS}}}t")), "text", inline.find(f"{{{MAIN_NS}}}r") is not None, None
        if cell_type in {"str", "e"}:
            node = cell.find(f"{{{MAIN_NS}}}v")
            return (node.text or "") if node is not None else "", "text", False, None
        if cell_type == "b":
            node = cell.find(f"{{{MAIN_NS}}}v")
            return bool(node is not None and node.text == "1"), "boolean", False, None
        node = cell.find(f"{{{MAIN_NS}}}v")
        lexical = (node.text or "") if node is not None else ""
        try:
            number = Decimal(lexical)
        except InvalidOperation as exc:
            raise UnsafeWorkbookError("数值单元格内容无法解析。") from exc
        if expected_kind == "integer":
            return int(number), "integer", False, None
        if expected_kind == "float":
            return float(number), "float", False, None
        if expected_kind == "decimal":
            return number, "decimal", False, None
        return number, "number", False, None

    @staticmethod
    def _set_inline_string(cell: ET.Element, replacement: str) -> None:
        cell.attrib["t"] = "inlineStr"
        for tag in (f"{{{MAIN_NS}}}v", f"{{{MAIN_NS}}}is"):
            node = cell.find(tag)
            if node is not None:
                cell.remove(node)
        inline = ET.Element(f"{{{MAIN_NS}}}is")
        ext = cell.find(f"{{{MAIN_NS}}}extLst")
        cell.append(inline) if ext is None else cell.insert(list(cell).index(ext), inline)
        text = ET.SubElement(inline, f"{{{MAIN_NS}}}t")
        if replacement[:1].isspace() or replacement[-1:].isspace():
            text.set(f"{{{XML_NS}}}space", "preserve")
        text.text = replacement

    @staticmethod
    def _set_numeric_value(cell: ET.Element, replacement: object) -> None:
        cell.attrib.pop("t", None)
        inline = cell.find(f"{{{MAIN_NS}}}is")
        if inline is not None:
            cell.remove(inline)
        node = cell.find(f"{{{MAIN_NS}}}v")
        if node is None:
            node = ET.Element(f"{{{MAIN_NS}}}v")
            ext = cell.find(f"{{{MAIN_NS}}}extLst")
            cell.append(node) if ext is None else cell.insert(list(cell).index(ext), node)
        node.text = str(replacement)

    @staticmethod
    def _clear_formula_caches(root: ET.Element) -> int:
        count = 0
        for cell in root.findall(f".//{{{MAIN_NS}}}c"):
            if cell.find(f"{{{MAIN_NS}}}f") is None:
                continue
            node = cell.find(f"{{{MAIN_NS}}}v")
            if node is not None:
                cell.remove(node)
                count += 1
        return count


def resolve_sheet_paths(archive: ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    relationships = {item.attrib["Id"]: item.attrib["Target"] for item in rels.findall(f"{{{PKG_REL_NS}}}Relationship")}
    sheets = workbook.find(f"{{{MAIN_NS}}}sheets")
    if sheets is None:
        raise UnsafeWorkbookError("工作簿不包含工作表清单。")
    result: dict[str, str] = {}
    for sheet in sheets:
        target = relationships.get(sheet.attrib.get(f"{{{DOC_REL_NS}}}id", ""))
        if target:
            result[sheet.attrib["name"]] = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
    return result


def parse_xml_preserving_namespaces(payload: bytes) -> ET.Element:
    for _, (prefix, uri) in ET.iterparse(BytesIO(payload), events=("start-ns",)):
        try:
            ET.register_namespace(prefix or "", uri)
        except ValueError:
            pass
    return ET.fromstring(payload, parser=ET.XMLParser(target=ET.TreeBuilder(insert_comments=True)))
