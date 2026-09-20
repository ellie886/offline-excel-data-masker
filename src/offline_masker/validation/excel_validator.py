from __future__ import annotations

from io import BytesIO
from numbers import Number
import xml.etree.ElementTree as ET
from collections import defaultdict
from zipfile import ZipFile

from openpyxl import load_workbook

from offline_masker.domain.models import (
    PatchOperation,
    ValidationCheck,
    ValidationResult,
    WorkbookSnapshot,
)
from offline_masker.excel.ooxml_patcher import MAIN_NS, parse_xml_preserving_namespaces, resolve_sheet_paths


class ExcelValidator:
    def snapshot(self, data: bytes) -> WorkbookSnapshot:
        workbook = load_workbook(BytesIO(data), read_only=False, data_only=False, keep_links=True, rich_text=True)
        formulas: dict[str, str] = {}
        numeric_constants: dict[str, tuple[str, object]] = {}
        cell_styles: dict[str, int] = {}
        merged_ranges: dict[str, tuple[str, ...]] = {}
        sheet_features: dict[str, tuple[object, ...]] = {}
        try:
            sheet_names = tuple(workbook.sheetnames)
            sheet_states = {worksheet.title: worksheet.sheet_state for worksheet in workbook.worksheets}
            dimensions = {
                worksheet.title: (worksheet.max_row, worksheet.max_column) for worksheet in workbook.worksheets
            }
            for worksheet in workbook.worksheets:
                for row in worksheet.iter_rows():
                    for cell in row:
                        reference = f"{worksheet.title}!{cell.coordinate}"
                        if cell.data_type == "f":
                            formulas[reference] = str(cell.value)
                        elif isinstance(cell.value, Number) and not isinstance(cell.value, bool):
                            numeric_constants[reference] = (type(cell.value).__name__, cell.value)
                        if cell.has_style:
                            cell_styles[reference] = cell.style_id
                merged_ranges[worksheet.title] = tuple(sorted(str(item) for item in worksheet.merged_cells.ranges))
                row_dimensions = tuple(
                    sorted(
                        (index, dimension.hidden, dimension.height, dimension.outlineLevel)
                        for index, dimension in worksheet.row_dimensions.items()
                    )
                )
                column_dimensions = tuple(
                    sorted(
                        (
                            key,
                            dimension.hidden,
                            dimension.width,
                            dimension.min,
                            dimension.max,
                            dimension.outlineLevel,
                        )
                        for key, dimension in worksheet.column_dimensions.items()
                    )
                )
                sheet_features[worksheet.title] = (
                    str(worksheet.freeze_panes or ""),
                    worksheet.auto_filter.ref or "",
                    str(worksheet.print_area or ""),
                    str(worksheet.print_title_rows or ""),
                    str(worksheet.print_title_cols or ""),
                    worksheet.sheet_view.showGridLines,
                    row_dimensions,
                    column_dimensions,
                )
        finally:
            workbook.close()

        with ZipFile(BytesIO(data)) as archive:
            package_parts = tuple(sorted(archive.namelist()))
        return WorkbookSnapshot(
            sheet_names=sheet_names,
            sheet_states=sheet_states,
            dimensions=dimensions,
            formulas=formulas,
            numeric_constants=numeric_constants,
            cell_styles=cell_styles,
            merged_ranges=merged_ranges,
            sheet_features=sheet_features,
            package_parts=package_parts,
        )

    def validate(
        self,
        source: bytes,
        output: bytes,
        operations: list[PatchOperation],
        modified_parts: set[str],
        actual_replacement_count: int,
    ) -> ValidationResult:
        checks: list[ValidationCheck] = []
        try:
            before = self.snapshot(source)
            after = self.snapshot(output)
        except Exception as exc:
            return ValidationResult(
                [ValidationCheck("输出文件可读取", False, f"无法重新读取输出文件：{type(exc).__name__}。")]
            )

        checks.extend(
            [
                self._equal_check("工作表名称和顺序", before.sheet_names, after.sheet_names),
                self._equal_check("工作表隐藏状态", before.sheet_states, after.sheet_states),
                self._equal_check("工作表行列范围", before.dimensions, after.dimensions),
                self._equal_check("公式数量和内容", before.formulas, after.formulas),
                self._equal_check("数值常量及金额基础", before.numeric_constants, after.numeric_constants),
                self._equal_check("单元格样式引用", before.cell_styles, after.cell_styles),
                self._equal_check("合并单元格", before.merged_ranges, after.merged_ranges),
                self._equal_check("冻结窗格、筛选、打印及行列设置", before.sheet_features, after.sheet_features),
                self._equal_check("工作簿内部组件清单", before.package_parts, after.package_parts),
            ]
        )

        expected_count = sum(operation.candidate_count for operation in operations)
        checks.append(
            ValidationCheck(
                "实际替换数量",
                actual_replacement_count == expected_count,
                f"确认 {expected_count} 项，实际执行 {actual_replacement_count} 项。",
            )
        )
        checks.append(self._validate_target_values(output, operations))
        checks.append(self._validate_modified_sheet_structure(source, output, operations))
        checks.append(self._validate_untouched_parts(source, output, modified_parts))
        return ValidationResult(checks)

    @staticmethod
    def _equal_check(name: str, before: object, after: object) -> ValidationCheck:
        passed = before == after
        if passed:
            count = len(before) if hasattr(before, "__len__") else None
            detail = f"一致，共 {count} 项。" if count is not None else "一致。"
            return ValidationCheck(name, True, detail)
        if isinstance(before, dict) and isinstance(after, dict):
            before_keys = set(before)
            after_keys = set(after)
            missing = sorted(before_keys - after_keys)
            added = sorted(after_keys - before_keys)
            changed = sorted(key for key in before_keys & after_keys if before[key] != after[key])
            samples = [*(f"缺失:{key}" for key in missing[:3]), *(f"新增:{key}" for key in added[:3]), *(f"变化:{key}" for key in changed[:5])]
            return ValidationCheck(name, False, "；".join(samples) or "处理前后字典内容不同。")
        return ValidationCheck(name, False, f"处理前为 {before!r}，处理后为 {after!r}。")

    @staticmethod
    def _validate_target_values(output: bytes, operations: list[PatchOperation]) -> ValidationCheck:
        workbook = load_workbook(BytesIO(output), read_only=False, data_only=False, keep_links=True, rich_text=True)
        differences: list[str] = []
        try:
            for operation in operations:
                actual = workbook[operation.sheet][operation.coordinate].value
                if str(actual) != operation.replacement_value:
                    differences.append(f"{operation.sheet}!{operation.coordinate}")
        finally:
            workbook.close()
        if differences:
            sample = "、".join(differences[:5])
            return ValidationCheck("目标单元格替换结果", False, f"以下位置结果不符：{sample}。")
        return ValidationCheck("目标单元格替换结果", True, f"{len(operations)} 个目标单元格均与确认计划一致。")

    @staticmethod
    def _validate_untouched_parts(source: bytes, output: bytes, modified_parts: set[str]) -> ValidationCheck:
        with ZipFile(BytesIO(source)) as before_zip, ZipFile(BytesIO(output)) as after_zip:
            before_names = set(before_zip.namelist())
            after_names = set(after_zip.namelist())
            if before_names != after_names:
                return ValidationCheck("非目标 OOXML 组件", False, "内部组件清单发生变化。")
            changed = [
                name
                for name in sorted(before_names - modified_parts)
                if before_zip.read(name) != after_zip.read(name)
            ]
        if changed:
            sample = "、".join(changed[:5])
            return ValidationCheck("非目标 OOXML 组件", False, f"发现 {len(changed)} 个非目标组件变化：{sample}。")
        return ValidationCheck("非目标 OOXML 组件", True, "除目标工作表 XML 外，其余组件字节保持不变。")

    @staticmethod
    def _validate_modified_sheet_structure(
        source: bytes,
        output: bytes,
        operations: list[PatchOperation],
    ) -> ValidationCheck:
        coordinates_by_sheet: dict[str, set[str]] = defaultdict(set)
        for operation in operations:
            coordinates_by_sheet[operation.sheet].add(operation.coordinate)

        differences: list[str] = []
        with ZipFile(BytesIO(source)) as before_zip, ZipFile(BytesIO(output)) as after_zip:
            before_paths = resolve_sheet_paths(before_zip)
            after_paths = resolve_sheet_paths(after_zip)
            for sheet, coordinates in coordinates_by_sheet.items():
                before_path = before_paths.get(sheet)
                after_path = after_paths.get(sheet)
                if not before_path or not after_path:
                    differences.append(sheet)
                    continue
                before_root = parse_xml_preserving_namespaces(before_zip.read(before_path))
                after_root = parse_xml_preserving_namespaces(after_zip.read(after_path))
                for root in (before_root, after_root):
                    for coordinate in coordinates:
                        cell = root.find(f".//{{{MAIN_NS}}}c[@r='{coordinate}']")
                        if cell is None:
                            continue
                        cell.attrib["t"] = "__masked_target__"
                        for tag in (f"{{{MAIN_NS}}}v", f"{{{MAIN_NS}}}is"):
                            node = cell.find(tag)
                            if node is not None:
                                cell.remove(node)
                if _element_signature(before_root) != _element_signature(after_root):
                    differences.append(sheet)

        if differences:
            return ValidationCheck(
                "目标工作表非内容结构",
                False,
                f"目标值以外的 XML 结构发生变化：{'、'.join(differences[:5])}。",
            )
        return ValidationCheck("目标工作表非内容结构", True, "目标单元格值以外的工作表 XML 结构一致。")


def _element_signature(element: ET.Element) -> tuple[object, ...]:
    text = element.text if element.text and element.text.strip() else ""
    if not isinstance(element.tag, str):
        return ("#comment", element.text or "")
    return (
        element.tag,
        tuple(sorted(element.attrib.items())),
        text,
        tuple(_element_signature(child) for child in list(element)),
    )
