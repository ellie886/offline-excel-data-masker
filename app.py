from __future__ import annotations

import hashlib
import json
import re
import sys
from pathlib import Path

SOURCE_ROOT = Path(__file__).resolve().parent / "src"
if str(SOURCE_ROOT) not in sys.path:
    sys.path.insert(0, str(SOURCE_ROOT))

import streamlit as st

from offline_masker.detectors.coordinator import (
    DetectionCoordinator,
    infer_column_rules,
    make_manual_candidate,
)
from offline_masker.domain.enums import MaskMethod, SensitiveType
from offline_masker.domain.exceptions import MaskerError
from offline_masker.domain.models import ColumnRule, CustomTerm
from offline_masker.excel.inspector import WorkbookInspector
from offline_masker.security.safe_preview import safe_preview
from offline_masker.services.processing_service import ProcessingService


st.set_page_config(page_title="离线 Excel 数据脱敏工具", page_icon="🔒", layout="wide")

TYPE_VALUES = [item.value for item in SensitiveType]
METHOD_VALUES = [item.value for item in MaskMethod]


def reset_processing_state() -> None:
    for key in (
        "inspection",
        "column_rows",
        "column_seed",
        "candidates",
        "scan_warnings",
        "scan_configuration_fingerprint",
        "processing_result",
        "processed_plan_signature",
    ):
        st.session_state.pop(key, None)


def parse_custom_terms(raw: str, case_sensitive: bool) -> list[CustomTerm]:
    terms: list[CustomTerm] = []
    for line in raw.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if "=>" in stripped:
            original, replacement = (part.strip() for part in stripped.split("=>", 1))
        else:
            original, replacement = stripped, "[已脱敏]"
        if original:
            terms.append(CustomTerm(original, replacement, case_sensitive, True))
    return terms


def rows_from_rules(data: bytes, rules: list[ColumnRule]) -> list[dict[str, object]]:
    header_cache: dict[tuple[str, int], dict[str, str]] = {}
    rows: list[dict[str, object]] = []
    for rule in rules:
        key = (rule.sheet, rule.header_row)
        if key not in header_cache:
            header_cache[key] = dict(WorkbookInspector.list_headers(data, rule.sheet, rule.header_row))
        rows.append(
            {
                "处理": True,
                "工作表": rule.sheet,
                "列": rule.column,
                "表头": header_cache[key].get(rule.column, ""),
                "类型": rule.sensitive_type.value,
                "依据": rule.source,
            }
        )
    return rows


def dataframe_records(value) -> list[dict[str, object]]:
    if hasattr(value, "to_dict"):
        return value.to_dict("records")
    return list(value)


def stable_digest(value: object) -> str:
    payload = json.dumps(value, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def candidate_plan_digest(candidates) -> str:
    return stable_digest(
        [
            (
                candidate.candidate_id,
                candidate.enabled,
                candidate.sensitive_type.value,
                candidate.effective_method.value,
                candidate.replacement_override,
            )
            for candidate in candidates
        ]
    )


st.title("离线 Excel 数据脱敏工具")
st.caption("文件只在本机的浏览器与本地 Python 进程之间流转，不调用云端 API，不覆盖原文件。")
st.info("第一版仅支持未加密、无宏、无数字签名的 .xlsx 文件。请先预览并确认，再执行脱敏。")

st.header("第一步：导入文件")
uploaded = st.file_uploader("选择一个 Excel 工作簿", type=["xlsx"], accept_multiple_files=False)
if uploaded is None:
    st.stop()

source_bytes = uploaded.getvalue()
source_digest = hashlib.sha256(source_bytes).hexdigest()
if st.session_state.get("source_digest") != source_digest:
    reset_processing_state()
    st.session_state.source_digest = source_digest
    st.session_state.source_bytes = source_bytes
    try:
        st.session_state.inspection = WorkbookInspector().inspect(source_bytes, uploaded.name)
    except MaskerError as exc:
        st.error(str(exc))
        st.stop()

inspection = st.session_state.inspection
col1, col2, col3 = st.columns(3)
col1.metric("文件名称", inspection.filename)
col2.metric("文件大小", f"{inspection.size_bytes / 1024 / 1024:.2f} MB")
col3.metric("工作表数量", len(inspection.sheets))
st.dataframe(
    [
        {
            "工作表": sheet.name,
            "状态": sheet.state,
            "行数": sheet.max_row,
            "列数": sheet.max_column,
            "公式": sheet.formula_count,
            "合并区域": sheet.merged_range_count,
            "隐藏行": sheet.hidden_row_count,
            "隐藏列": sheet.hidden_column_count,
            "批注": sheet.comment_count,
        }
        for sheet in inspection.sheets
    ],
    width="stretch",
    hide_index=True,
)
for warning in inspection.warnings:
    st.warning(warning)

st.header("第二步：设置规则")
all_sheets = [sheet.name for sheet in inspection.sheets]
visible_sheets = [sheet.name for sheet in inspection.sheets if sheet.state == "visible"]
selected_sheets = st.multiselect("选择需要扫描的工作表", all_sheets, default=visible_sheets)
selected_type_values = st.multiselect("选择需要识别的信息类型", TYPE_VALUES, default=TYPE_VALUES)
enabled_types = {SensitiveType(value) for value in selected_type_values}
header_row = int(st.number_input("表头所在行", min_value=1, max_value=100, value=1, step=1))

seed = (tuple(selected_sheets), tuple(selected_type_values), header_row, source_digest)
if st.session_state.get("column_seed") != seed:
    inferred = infer_column_rules(source_bytes, selected_sheets, enabled_types, header_row) if selected_sheets else []
    st.session_state.column_rows = rows_from_rules(source_bytes, inferred)
    st.session_state.column_seed = seed
    for stale_key in ("candidates", "scan_warnings", "scan_configuration_fingerprint", "processing_result", "processed_plan_signature"):
        st.session_state.pop(stale_key, None)

st.markdown("程序根据表头生成了列规则。可以取消、修改，或新增一行指定列。")
column_editor = st.data_editor(
    st.session_state.column_rows,
    num_rows="dynamic",
    width="stretch",
    hide_index=True,
    column_config={
        "处理": st.column_config.CheckboxColumn(default=True),
        "工作表": st.column_config.SelectboxColumn(options=selected_sheets),
        "列": st.column_config.TextColumn(help="Excel 列字母，例如 A、BC"),
        "类型": st.column_config.SelectboxColumn(options=TYPE_VALUES),
        "表头": st.column_config.TextColumn(disabled=True),
        "依据": st.column_config.TextColumn(disabled=True),
    },
    key=f"column_rule_editor_{stable_digest(seed)[:12]}",
)

custom_raw = st.text_area(
    "自定义敏感词",
    placeholder="每行一个，格式：原始敏感词 => 替换内容\n未填写替换内容时使用 [已脱敏]",
    height=110,
)
custom_case_sensitive = st.checkbox("自定义敏感词区分大小写", value=True)
column_records = dataframe_records(column_editor)
scan_configuration_fingerprint = stable_digest(
    {
        "sheets": selected_sheets,
        "types": selected_type_values,
        "header_row": header_row,
        "columns": column_records,
        "custom_terms": custom_raw,
        "case_sensitive": custom_case_sensitive,
    }
)

if st.button("扫描敏感信息", type="primary", disabled=not selected_sheets or not enabled_types):
    try:
        rules: list[ColumnRule] = []
        for row in column_records:
            if not row.get("处理", True):
                continue
            sheet = str(row.get("工作表") or "").strip()
            column = str(row.get("列") or "").strip().upper()
            type_value = str(row.get("类型") or "").strip()
            if sheet not in selected_sheets or not re.fullmatch(r"[A-Z]{1,3}", column) or type_value not in TYPE_VALUES:
                raise ValueError("列规则包含无效的工作表、列字母或类型。")
            rules.append(
                ColumnRule(
                    sheet=sheet,
                    column=column,
                    sensitive_type=SensitiveType(type_value),
                    header_row=header_row,
                    source=str(row.get("依据") or "用户指定列"),
                )
            )
        custom_terms = parse_custom_terms(custom_raw, custom_case_sensitive)
        scan_result = DetectionCoordinator().scan(
            source_bytes,
            selected_sheets,
            enabled_types,
            rules,
            custom_terms,
        )
        st.session_state.candidates = scan_result.candidates
        st.session_state.scan_warnings = scan_result.warnings
        st.session_state.scan_configuration_fingerprint = scan_configuration_fingerprint
        st.session_state.pop("processing_result", None)
        st.session_state.pop("processed_plan_signature", None)
    except (MaskerError, ValueError) as exc:
        st.error(str(exc))

candidates = st.session_state.get("candidates")
if not candidates:
    if candidates == []:
        st.info("未发现候选敏感信息。可以调整列规则或补充自定义敏感词后重新扫描。")
    st.stop()

scan_is_current = st.session_state.get("scan_configuration_fingerprint") == scan_configuration_fingerprint
if not scan_is_current:
    st.warning("扫描规则已经变化。请重新点击“扫描敏感信息”后再确认处理。")

st.header("第三步：预览确认")
for warning in st.session_state.get("scan_warnings", []):
    st.warning(warning)

preview_rows = [
    {
        "候选ID": candidate.candidate_id,
        "执行": candidate.enabled,
        "工作表": candidate.sheet,
        "单元格": candidate.coordinate,
        "安全预览": candidate.safe_preview,
        "识别类型": candidate.sensitive_type.value,
        "脱敏方式": candidate.effective_method.value,
        "置信度(%)": candidate.confidence * 100,
        "识别依据": candidate.basis,
        "替换内容": candidate.replacement_override,
    }
    for candidate in candidates
]
preview_editor = st.data_editor(
    preview_rows,
    width="stretch",
    hide_index=True,
    disabled=["工作表", "单元格", "安全预览", "置信度(%)", "识别依据"],
    column_config={
        "候选ID": None,
        "执行": st.column_config.CheckboxColumn(),
        "识别类型": st.column_config.SelectboxColumn(options=TYPE_VALUES),
        "脱敏方式": st.column_config.SelectboxColumn(options=METHOD_VALUES),
        "置信度(%)": st.column_config.NumberColumn(format="%.0f"),
        "替换内容": st.column_config.TextColumn(help="填写后优先使用该内容"),
    },
    key=f"candidate_editor_{stable_digest([item.candidate_id for item in candidates])[:12]}",
)

candidate_by_id = {candidate.candidate_id: candidate for candidate in candidates}
for row in dataframe_records(preview_editor):
    candidate = candidate_by_id[str(row["候选ID"])]
    candidate.enabled = bool(row["执行"])
    candidate.sensitive_type = SensitiveType(str(row["识别类型"]))
    candidate.method = MaskMethod(str(row["脱敏方式"]))
    candidate.replacement_override = str(row.get("替换内容") or "")
    fragment = candidate.original_value[candidate.start : candidate.end]
    candidate.safe_preview = safe_preview(fragment, candidate.sensitive_type)
st.session_state.candidates = candidates
current_plan_signature = candidate_plan_digest(candidates)
if st.session_state.get("processed_plan_signature") not in (None, current_plan_signature):
    st.session_state.pop("processing_result", None)
    st.session_state.pop("processed_plan_signature", None)

with st.expander("补充一个未识别的单元格"):
    manual_sheet = st.selectbox("工作表", selected_sheets, key="manual_sheet")
    manual_coordinate = st.text_input("单元格位置", placeholder="例如 D18")
    manual_type = st.selectbox("敏感信息类型", TYPE_VALUES, key="manual_type")
    if st.button("加入预览列表"):
        try:
            new_candidate = make_manual_candidate(
                source_bytes,
                manual_sheet,
                manual_coordinate,
                SensitiveType(manual_type),
            )
            if new_candidate.candidate_id not in candidate_by_id:
                candidates.append(new_candidate)
                st.session_state.candidates = candidates
                st.rerun()
            else:
                st.info("该单元格已经在预览列表中。")
        except (ValueError, KeyError) as exc:
            st.error(str(exc))

selected_count = sum(candidate.enabled for candidate in candidates)
st.metric("预计替换数量", selected_count)
confirmed = st.checkbox("我已检查预览结果并确认执行；程序将生成新文件，不覆盖原文件。")

if st.button("执行脱敏并校验", type="primary", disabled=not confirmed or selected_count == 0 or not scan_is_current):
    try:
        warnings = inspection.warnings + st.session_state.get("scan_warnings", [])
        st.session_state.processing_result = ProcessingService().process(
            source=source_bytes,
            source_filename=uploaded.name,
            scanned_sheets=selected_sheets,
            candidates=candidates,
            warnings=list(dict.fromkeys(warnings)),
        )
        st.session_state.processed_plan_signature = current_plan_signature
    except MaskerError as exc:
        st.error(str(exc))

result = st.session_state.get("processing_result")
if result is None:
    st.stop()

st.header("第四步：处理与导出")
st.dataframe(
    [
        {"校验项目": check.name, "结果": "通过" if check.passed else "失败", "说明": check.detail}
        for check in result.validation.checks
    ],
    width="stretch",
    hide_index=True,
)

if result.validation.passed:
    st.success(f"处理和关键校验均已通过，共替换 {result.replacement_count} 项。")
    left, right = st.columns(2)
    left.download_button(
        "下载脱敏文件",
        result.output_bytes,
        file_name=result.output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    right.download_button(
        "下载脱敏报告",
        result.report_bytes,
        file_name=result.report_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.error("处理后校验未全部通过。脱敏文件不提供下载，请查看报告中的具体差异。")
    st.download_button(
        "下载失败校验报告",
        result.report_bytes,
        file_name=result.report_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

if st.button("清除本次会话数据"):
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()
