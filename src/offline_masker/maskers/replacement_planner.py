from __future__ import annotations

from collections import defaultdict

from offline_masker.domain.enums import MaskMethod
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
            original_values = {candidate.original_value for candidate in cell_candidates}
            if len(original_values) != 1:
                raise ProcessingConflictError(f"同一单元格存在不一致的原始值：{sheet}!{coordinate}。")
            original = original_values.pop()
            ordered = sorted(cell_candidates, key=lambda item: (item.start, item.end))
            for left, right in zip(ordered, ordered[1:]):
                if right.start < left.end:
                    raise ProcessingConflictError(f"同一单元格存在重叠的替换范围：{sheet}!{coordinate}。")

            result = original
            for candidate in reversed(ordered):
                fragment = original[candidate.start : candidate.end]
                replacement = self._replacement(candidate, fragment)
                result = result[: candidate.start] + replacement + result[candidate.end :]

            operations.append(
                PatchOperation(
                    sheet=sheet,
                    coordinate=coordinate,
                    expected_value_hash=value_hash(original),
                    replacement_value=result,
                    candidate_count=len(cell_candidates),
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
            raise ProcessingConflictError(f"自定义替换缺少替换内容：{candidate.sheet}!{candidate.coordinate}。")
        raise ProcessingConflictError(f"不支持的脱敏方式：{candidate.effective_method}。")

