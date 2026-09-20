from __future__ import annotations

from io import BytesIO
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from offline_masker.domain.exceptions import UnsafeWorkbookError
from offline_masker.domain.models import SheetInfo, WorkbookInspection
from offline_masker.security.file_guard import FileLimits, validate_xlsx_package


RISK_PARTS: tuple[tuple[str, str], ...] = (
    ("xl/charts/", "图表中的标题、标签和缓存数据不会扫描"),
    ("xl/drawings/", "绘图、文本框和形状中的文字不会扫描"),
    ("xl/pivot", "数据透视表及其缓存不会进行功能验证"),
    ("xl/threadedComments/", "线程批注不会扫描"),
    ("xl/embeddings/", "嵌入对象不会扫描"),
    ("customXml/", "自定义 XML 内容不会扫描"),
    ("xl/connections.xml", "数据连接不会刷新或验证"),
)


class WorkbookInspector:
    def __init__(self, *, max_scannable_cells: int = 1_000_000, limits: FileLimits | None = None) -> None:
        self.max_scannable_cells = max_scannable_cells
        self.limits = limits or FileLimits()

    def inspect(self, data: bytes, filename: str) -> WorkbookInspection:
        warnings = validate_xlsx_package(data, filename, self.limits)
        with ZipFile(BytesIO(data)) as archive:
            package_parts = sorted(archive.namelist())
        unsupported = [message for prefix, message in RISK_PARTS if any(name.startswith(prefix) for name in package_parts)]

        try:
            workbook = load_workbook(BytesIO(data), read_only=False, data_only=False, keep_links=True, rich_text=True)
        except Exception as exc:
            raise UnsafeWorkbookError(f"无法安全读取工作簿：{type(exc).__name__}。") from exc

        sheets: list[SheetInfo] = []
        total_cells = 0
        try:
            for worksheet in workbook.worksheets:
                estimated_cells = worksheet.max_row * worksheet.max_column
                total_cells += estimated_cells
                if total_cells > self.max_scannable_cells:
                    raise UnsafeWorkbookError(
                        f"工作簿使用区域约含 {total_cells:,} 个单元格，超过第一版 {self.max_scannable_cells:,} 个的扫描限制。"
                    )

                formula_count = 0
                comment_count = 0
                for row in worksheet.iter_rows():
                    for cell in row:
                        if cell.data_type == "f":
                            formula_count += 1
                        if cell.comment is not None:
                            comment_count += 1

                hidden_rows = sum(1 for dimension in worksheet.row_dimensions.values() if dimension.hidden)
                hidden_columns = 0
                for key, dimension in worksheet.column_dimensions.items():
                    if dimension.hidden:
                        start = dimension.min or worksheet[key][0].column
                        end = dimension.max or start
                        hidden_columns += end - start + 1

                if worksheet.sheet_state != "visible":
                    warnings.append(f"工作表“{worksheet.title}”处于隐藏状态（{worksheet.sheet_state}）。")
                if hidden_rows:
                    warnings.append(f"工作表“{worksheet.title}”包含 {hidden_rows} 个隐藏行定义。")
                if hidden_columns:
                    warnings.append(f"工作表“{worksheet.title}”包含 {hidden_columns} 个隐藏列定义。")
                if comment_count:
                    warnings.append(f"工作表“{worksheet.title}”包含 {comment_count} 条普通批注，第一版只提示、不扫描批注正文。")

                sheets.append(
                    SheetInfo(
                        name=worksheet.title,
                        state=worksheet.sheet_state,
                        max_row=worksheet.max_row,
                        max_column=worksheet.max_column,
                        formula_count=formula_count,
                        merged_range_count=len(worksheet.merged_cells.ranges),
                        hidden_row_count=hidden_rows,
                        hidden_column_count=hidden_columns,
                        comment_count=comment_count,
                    )
                )
        finally:
            workbook.close()

        if unsupported:
            warnings.extend(f"未处理内容：{item}。" for item in unsupported)
        return WorkbookInspection(
            filename=filename,
            size_bytes=len(data),
            sheets=sheets,
            warnings=_deduplicate(warnings),
            unsupported_items=unsupported,
            package_parts=package_parts,
        )

    @staticmethod
    def list_headers(data: bytes, sheet_name: str, header_row: int = 1) -> list[tuple[str, str]]:
        workbook = load_workbook(BytesIO(data), read_only=False, data_only=False, keep_links=True, rich_text=True)
        try:
            worksheet = workbook[sheet_name]
            headers: list[tuple[str, str]] = []
            for column_index in range(1, worksheet.max_column + 1):
                value = worksheet.cell(header_row, column_index).value
                headers.append((get_column_letter(column_index), "" if value is None else str(value)))
            return headers
        finally:
            workbook.close()


def _deduplicate(items: list[str]) -> list[str]:
    return list(dict.fromkeys(items))

