from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

from .enums import MaskMethod, SensitiveType


@dataclass(slots=True)
class SheetInfo:
    name: str
    state: str
    max_row: int
    max_column: int
    formula_count: int
    merged_range_count: int
    hidden_row_count: int
    hidden_column_count: int
    comment_count: int


@dataclass(slots=True)
class WorkbookInspection:
    filename: str
    size_bytes: int
    sheets: list[SheetInfo]
    warnings: list[str] = field(default_factory=list)
    unsupported_items: list[str] = field(default_factory=list)
    package_parts: list[str] = field(default_factory=list)


@dataclass(slots=True)
class ColumnRule:
    sheet: str
    column: str
    sensitive_type: SensitiveType
    header_row: int = 1
    source: str = "用户指定列"


@dataclass(slots=True)
class CustomTerm:
    original: str
    replacement: str = ""
    case_sensitive: bool = True
    whole_workbook: bool = True


@dataclass(slots=True)
class Candidate:
    candidate_id: str
    sheet: str
    coordinate: str
    original_value: str
    safe_preview: str
    sensitive_type: SensitiveType
    suggested_method: MaskMethod
    confidence: float
    basis: str
    start: int
    end: int
    enabled: bool = True
    replacement_override: str = ""
    method: MaskMethod | None = None

    @property
    def effective_method(self) -> MaskMethod:
        return self.method or self.suggested_method


@dataclass(slots=True)
class ScanResult:
    candidates: list[Candidate]
    warnings: list[str] = field(default_factory=list)


@dataclass(slots=True)
class PatchOperation:
    sheet: str
    coordinate: str
    expected_value_hash: str
    replacement_value: str
    candidate_count: int


@dataclass(slots=True)
class PatchResult:
    output_bytes: bytes
    modified_parts: set[str]
    modified_cells: int
    replacement_count: int


@dataclass(slots=True)
class ValidationCheck:
    name: str
    passed: bool
    detail: str
    severity: str = "critical"


@dataclass(slots=True)
class ValidationResult:
    checks: list[ValidationCheck]

    @property
    def passed(self) -> bool:
        return all(check.passed for check in self.checks if check.severity == "critical")


@dataclass(slots=True)
class ProcessingResult:
    output_filename: str
    report_filename: str
    output_bytes: bytes
    report_bytes: bytes
    validation: ValidationResult
    processed_at: datetime
    replacement_count: int
    risk_warnings: list[str]


@dataclass(slots=True)
class WorkbookSnapshot:
    sheet_names: tuple[str, ...]
    sheet_states: dict[str, str]
    dimensions: dict[str, tuple[int, int]]
    formulas: dict[str, str]
    numeric_constants: dict[str, tuple[str, Any]]
    cell_styles: dict[str, int]
    merged_ranges: dict[str, tuple[str, ...]]
    sheet_features: dict[str, tuple[Any, ...]]
    package_parts: tuple[str, ...]
