from __future__ import annotations

from datetime import datetime
import hashlib
from io import BytesIO
from pathlib import Path
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from offline_masker.domain.models import Candidate, ProcessingResult
from offline_masker.domain.exceptions import UnsafeWorkbookError
from offline_masker.excel.inspector import WorkbookInspector
from offline_masker.excel.ooxml_patcher import OoxmlPatcher
from offline_masker.maskers.replacement_planner import ReplacementPlanner
from offline_masker.reports.audit_report import build_audit_report
from offline_masker.validation.excel_validator import ExcelValidator, report_contains_sensitive_value
from offline_masker.domain.models import ValidationCheck


class ProcessingService:
    def __init__(self) -> None:
        self.planner = ReplacementPlanner()
        self.patcher = OoxmlPatcher()
        self.validator = ExcelValidator()

    def process(
        self,
        *,
        source: bytes,
        source_filename: str,
        scanned_sheets: list[str],
        candidates: list[Candidate],
        warnings: list[str] | None = None,
    ) -> ProcessingResult:
        base_name = Path(Path(source_filename).name).stem
        output_filename = f"{base_name}_脱敏版.xlsx"
        report_filename = f"{base_name}_脱敏报告.xlsx"
        try:
            source_sha256 = hashlib.sha256(source).hexdigest()
            inspection = WorkbookInspector().inspect(source, source_filename)
            blocking = list(inspection.blocking_risks)
            selected = set(scanned_sheets)
            for sheet in inspection.sheets:
                if sheet.state == "veryHidden":
                    blocking.append(f"工作表“{sheet.name}”为 veryHidden")
                elif sheet.state != "visible" and sheet.name not in selected:
                    blocking.append(f"隐藏工作表“{sheet.name}”未扫描")
            if blocking:
                raise UnsafeWorkbookError("发现无法安全清理的高风险内容，已阻止导出：" + "；".join(dict.fromkeys(blocking)) + "。")
            operations = self.planner.build(candidates)
            self._reject_sensitive_defined_names(source, operations)
            patch_result = self.patcher.patch(source, operations)
            validation = self.validator.validate(
                source,
                patch_result.output_bytes,
                operations,
                patch_result.modified_parts,
                patch_result.replacement_count,
                source_sha256_before=source_sha256,
            )
            processed_at = datetime.now().astimezone()
            report_bytes = build_audit_report(
                processed_at=processed_at,
                source_filename=Path(source_filename).name,
                output_filename=output_filename,
                scanned_sheets=scanned_sheets,
                candidates=candidates,
                validation=validation,
                warnings=warnings or [],
                source_sha256=source_sha256,
                output_sha256=hashlib.sha256(patch_result.output_bytes).hexdigest(),
                patch_result=patch_result,
                high_risk_items=[],
            )
            report_findings = report_contains_sensitive_value(report_bytes, operations)
            validation.checks.append(ValidationCheck(
                "审计报告敏感原值扫描", not report_findings,
                "报告中未发现完整敏感原值。" if not report_findings else
                f"报告内部组件仍包含敏感原值：{'、'.join(report_findings[:5])}。",
            ))
            if report_findings:
                report_bytes = build_audit_report(
                    processed_at=processed_at, source_filename=Path(source_filename).name,
                    output_filename=output_filename, scanned_sheets=scanned_sheets,
                    candidates=candidates, validation=validation,
                    warnings=[*(warnings or []), "审计报告敏感残留校验失败。"],
                    source_sha256=source_sha256,
                    output_sha256=hashlib.sha256(patch_result.output_bytes).hexdigest(),
                    patch_result=patch_result, high_risk_items=[],
                )
        finally:
            self.planner.mapping_manager.clear()
        return ProcessingResult(
            output_filename=output_filename,
            report_filename=report_filename,
            output_bytes=patch_result.output_bytes,
            report_bytes=report_bytes,
            validation=validation,
            processed_at=processed_at,
            replacement_count=patch_result.replacement_count,
            risk_warnings=warnings or [],
            shared_strings_removed=patch_result.shared_strings_removed,
            formula_caches_cleared=patch_result.formula_caches_cleared,
        )

    @staticmethod
    def _reject_sensitive_defined_names(source: bytes, operations) -> None:
        secrets = {str(value) for item in operations for value in item.sensitive_originals if value not in (None, "")}
        if not secrets:
            return
        with ZipFile(BytesIO(source)) as archive:
            root = ET.fromstring(archive.read("xl/workbook.xml"))
        findings: list[str] = []
        for node in root.iter():
            if isinstance(node.tag, str) and node.tag.rsplit("}", 1)[-1] == "definedName":
                content = node.text or ""
                if any(secret in content for secret in secrets):
                    findings.append(node.attrib.get("name", "(未命名)"))
        if findings:
            raise UnsafeWorkbookError("定义名称中包含已确认敏感常量，已阻止导出：" + "、".join(findings[:8]) + "。")
