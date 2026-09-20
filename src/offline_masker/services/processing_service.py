from __future__ import annotations

from datetime import datetime
from pathlib import Path

from offline_masker.domain.models import Candidate, ProcessingResult
from offline_masker.excel.ooxml_patcher import OoxmlPatcher
from offline_masker.maskers.replacement_planner import ReplacementPlanner
from offline_masker.reports.audit_report import build_audit_report
from offline_masker.validation.excel_validator import ExcelValidator


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
            operations = self.planner.build(candidates)
            patch_result = self.patcher.patch(source, operations)
            validation = self.validator.validate(
                source,
                patch_result.output_bytes,
                operations,
                patch_result.modified_parts,
                patch_result.replacement_count,
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
        )
