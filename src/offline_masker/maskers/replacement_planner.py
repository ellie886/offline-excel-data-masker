from __future__ import annotations

from collections import defaultdict

from offline_masker.domain.enums import MaskMethod, SensitiveType
from offline_masker.domain.exceptions import ProcessingConflictError
from offline_masker.domain.models import Candidate, PatchOperation
from offline_masker.security.hashing import value_hash

from .mapping_manager import MappingManager
from .text_masker import mask_value


class ReplacementPlanner:
    def __init__(self, mapping_manager: MappingManager | None = None) -> None:
        self.mapping_manager = mapping_manager or MappingManager()

    def build(self, candidates: list[Candidate]) -> list[PatchOperation]:
        selected = [candidate for candidate in candidates if candidate.enabled]
        grouped: dict[tuple[str, str], list[Candidate]] = defaultdict(list)
        for candidate in selected:
            grouped[(candidate.sheet, candidate.coordinate)].append(candidate)

        operations: list[PatchOperation] = []
        for (sheet, coordinate), cell_candidates in grouped.items():
            originals = {(candidate.original_kind, repr(candidate.original_value)) for candidate in cell_candidates}
            if len(originals) != 1:
                raise ProcessingConflictError(f"同一单元格存在不一致的原始值：{sheet}!{coordinate}。")
            original = cell_candidates[0].original_value
            original_kind = cell_candidates[0].original_kind
            if original_kind != "text":
                if len(cell_candidates) != 1 or cell_candidates[0].effective_method != MaskMethod.CONSTANT_888:
                    raise ProcessingConflictError(f"数值单元格只能使用固定数值888：{sheet}!{coordinate}。")
                operations.append(
                    PatchOperation(
                        sheet=sheet,
                        coordinate=coordinate,
                        expected_value_hash=value_hash(original, original_kind),
                        original_value=original,
                        original_kind=original_kind,
                        replacement_value=888,
                        replacement_kind="integer",
                        candidate_count=1,
                        sensitive_types=(SensitiveType.FINANCIAL_AMOUNT,),
                        sensitive_originals=(original,),
                    )
                )
                continue
            ordered = sorted(cell_candidates, key=lambda item: (item.start, item.end))
            for left, right in zip(ordered, ordered[1:]):
                if right.start < left.end:
                    raise ProcessingConflictError(f"同一单元格存在重叠的替换范围：{sheet}!{coordinate}。")

            result = original
            for candidate in reversed(ordered):
                fragment = original[candidate.start : candidate.end]
                replacement = self._replacement(candidate, fragment)
                candidate.replacement_value = replacement
                candidate.replacement_kind = "text"
                result = result[: candidate.start] + replacement + result[candidate.end :]

            operations.append(
                PatchOperation(
                    sheet=sheet,
                    coordinate=coordinate,
                    expected_value_hash=value_hash(original),
                    original_value=original,
                    original_kind="text",
                    replacement_value=result,
                    replacement_kind="text",
                    candidate_count=len(cell_candidates),
                    sensitive_types=tuple(dict.fromkeys(item.sensitive_type for item in cell_candidates)),
                    sensitive_originals=tuple(original[item.start:item.end] for item in cell_candidates),
                )
            )
        return operations

    def _replacement(self, candidate: Candidate, fragment: str) -> str:
        if candidate.replacement_override:
            return candidate.replacement_override
        if candidate.effective_method == MaskMethod.MASK:
            return mask_value(fragment, candidate.sensitive_type)
        if candidate.effective_method == MaskMethod.CONSISTENT:
            return self.mapping_manager.replacement_for(candidate.sensitive_type, fragment)
        if candidate.effective_method == MaskMethod.CUSTOM:
            return "【已脱敏】"
        if candidate.effective_method == MaskMethod.CONSTANT_888:
            return "888"
        raise ProcessingConflictError(f"不支持的脱敏方式：{candidate.effective_method}。")
