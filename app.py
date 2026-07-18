# -*- coding: utf-8 -*-
"""
소포수령증 자동화 웹앱 — v48 큐텐 월평균환율 적용
실행: streamlit run app.py
"""

from __future__ import annotations

import contextlib
import io
import os
import hashlib
import re
import sys
import tempfile
import time
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from modules.pdf_parser import parse_pdf, detect_pdf_type
from modules.excel_writer import generate_excel, period_labels
from modules.exchange_rate import (
    fetch_all_currencies_for_period,
    fetch_monthly_avg_currencies_for_period,
    merge_monthly_rates,
    RATE_LOOKBACK_DAYS,
)
from modules.extra_docs import (
    build_declaration_rows,
    company_name_from_results,
    create_export_performance,
    create_zero_rate_attachments,
    safe_filename,
)

CURRENCIES = ["MYR", "PHP", "SGD", "THB", "TWD", "VND", "IDR", "JPY", "BRL", "MXN", "USD", "EUR", "GBP", "CAD", "AUD"]

# ── 선택적 로그인: Streamlit secrets [auth]가 있을 때만 사용 ────────────
ALLOWED_EMAILS = [
    "guwjd2298@gmail.com",
    "help@taxexpert.kr",
    "m0120@taxexpert.kr",
    "m0227@taxexpert.kr",
    "m0125@taxexpert.kr",
    "ayoung9976@gmail.com",
    "m0429@taxexpert.kr",
    "m0607@taxexpert.kr",
    "m1007@taxexpert.kr",
    "m1211@taxexpert.kr",
    "m1225@taxexpert.kr",
    "m1018@taxexpert.kr",
]

st.set_page_config(page_title="소포수령증 자동화", page_icon="📦", layout="centered")

_AUTH_ENABLED = False
try:
    _AUTH_ENABLED = "auth" in st.secrets
except Exception:
    _AUTH_ENABLED = False

if _AUTH_ENABLED:
    try:
        _logged_in = st.user.is_logged_in
    except Exception as e:
        st.error(f"로그인 상태 확인 오류: {type(e).__name__}: {e}")
        st.stop()
    if not _logged_in:
        st.markdown(
            """
            <div style="text-align:center; padding:3rem 1rem;">
                <p style="font-size:2rem; font-weight:700; color:#1f4e79;">📦 소포수령증 자동화</p>
                <p style="color:#555; margin-bottom:1.5rem;">사용하려면 Google 계정으로 로그인하세요.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        c1, c2, c3 = st.columns([1, 1.4, 1])
        with c2:
            st.button("🔓 Google 계정으로 로그인", type="primary", on_click=st.login, use_container_width=True)
        st.stop()
    _user_email = st.user.get("email", "")
    if ALLOWED_EMAILS and _user_email not in ALLOWED_EMAILS:
        st.error(f"❌ 접근 권한이 없습니다. ({_user_email})")
        if st.button("로그아웃"):
            st.logout()
        st.stop()
    with st.sidebar:
        st.markdown(f"**👤 {st.user.get('name','') or _user_email}**")
        st.caption(_user_email)
        if st.button("로그아웃"):
            st.logout()

st.markdown(
    """
<style>
.main-title { font-size:2rem; font-weight:700; color:#1f4e79; margin-bottom:0.2rem; }
.sub-title  { font-size:1rem; color:#555; margin-bottom:1.5rem; }
.warn-box   { background:#fff8e1; border-radius:8px; padding:0.8rem 1.2rem; border-left:4px solid #f9a825; margin-bottom:0.5rem; }
.info-box   { background:#e3f2fd; border-radius:8px; padding:0.8rem 1.2rem; border-left:4px solid #1565c0; margin-bottom:0.5rem; }
.log-box    { background:#0b1020; color:#e7edf8; border-radius:10px; padding:1rem; font-family:Consolas, monospace; font-size:0.85rem; max-height:320px; overflow:auto; white-space:pre-wrap; }
div[data-testid="stColumn"] .stButton > button { white-space: nowrap !important; }
</style>
""",
    unsafe_allow_html=True,
)

st.markdown('<p class="main-title">📦 소포수령증 자동화</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">기존 GitHub 방식의 PDF 파싱·엑셀 생성을 유지하고, 환율은 서울외국환중개에서 자동 수집합니다.</p>', unsafe_allow_html=True)
st.divider()

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "qoo10_entries" not in st.session_state:
    st.session_state.qoo10_entries = []
if "result_files" not in st.session_state:
    st.session_state.result_files = []
if "qoo10_auto_imported_keys" not in st.session_state:
    st.session_state.qoo10_auto_imported_keys = set()

# 파일명으로 구분되지 않는 PDF에만 직접 선택 항목을 표시합니다.
# 일반 화면의 문구와 배치는 v38과 동일하게 유지합니다.
UNKNOWN_PDF_TYPE_OPTIONS = ["쇼피", "라자다", "큐텐재팬", "이베이"]
UNKNOWN_PDF_TYPE_TO_CODE = {
    "쇼피": "shopee",
    "라자다": "lazada",
    "큐텐재팬": "qoo10",
    "이베이": "ebay",
}

# ══════════════════════════════════════════════════════════════════
# STEP 1 — PDF 업로드
# ══════════════════════════════════════════════════════════════════
st.markdown("### 📄 STEP 1 — 소포수령증 PDF 업로드")
c_desc, c_reset = st.columns([6, 1])
c_desc.caption("쇼피, 라자다, 큐텐재팬 PDF를 한꺼번에 올려주세요.")
if c_reset.button("🔄 초기화"):
    st.session_state.uploader_key += 1
    st.session_state.qoo10_entries = []
    st.session_state.result_files = []
    st.session_state.qoo10_auto_imported_keys = set()
    st.rerun()

uploaded_files = st.file_uploader(
    "PDF 파일 선택",
    type=["pdf"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    key=f"pdf_uploader_{st.session_state.uploader_key}",
)

@st.cache_data(show_spinner=False)
def _detect_uploaded_pdf_type(filename: str, payload: bytes) -> str:
    """파일명이 일반적이어도 PDF 본문 표식으로 플랫폼을 자동 판별합니다."""
    with tempfile.TemporaryDirectory() as td:
        path = Path(td) / (Path(filename).name or "uploaded.pdf")
        path.write_bytes(payload)
        return detect_pdf_type(str(path))


@st.cache_data(show_spinner=False)
def _parse_uploaded_qoo10(filename: str, payload: bytes):
    with tempfile.TemporaryDirectory() as td:
        path = Path(td) / (Path(filename).name or "qoo10.pdf")
        path.write_bytes(payload)
        return parse_pdf(str(path), forced_type="qoo10")


uploaded_type_choices = {}
if uploaded_files:
    cols = st.columns(2)
    for i, f in enumerate(uploaded_files):
        payload = f.getvalue()
        ptype = _detect_uploaded_pdf_type(f.name, payload)
        icon = {"shopee": "🛍️", "lazada": "🟠", "qoo10": "🇯🇵", "ebay": "🛒", "unknown": "❓"}.get(ptype, "📄")
        label = {"shopee": "쇼피", "lazada": "라자다", "qoo10": "큐텐재팬", "ebay": "이베이", "unknown": "미확인"}.get(ptype, "")
        target_col = cols[i % 2]
        target_col.markdown(f"{icon} `{f.name}` — {label}")
        if ptype == "unknown":
            selected_label = target_col.selectbox(
                "문서 종류",
                UNKNOWN_PDF_TYPE_OPTIONS,
                key=f"pdf_type_{st.session_state.uploader_key}_{i}_{f.name}",
                label_visibility="collapsed",
            )
            uploaded_type_choices[f.name] = UNKNOWN_PDF_TYPE_TO_CODE[selected_label]
        else:
            uploaded_type_choices[f.name] = ptype

    # 큐텐재팬 PDF는 업로드 즉시 파싱하여 STEP 2 입력 목록에 자동 반영합니다.
    qoo10_uploads = []
    for f in uploaded_files:
        if uploaded_type_choices.get(f.name) != "qoo10":
            continue
        payload = f.getvalue()
        file_key = hashlib.sha256(payload).hexdigest()
        qoo10_uploads.append((f, payload, file_key))

    # 업로더에서 제거된 PDF의 자동 입력 행은 STEP 2에서도 함께 제거합니다.
    current_qoo10_keys = {key for _, _, key in qoo10_uploads}
    st.session_state.qoo10_entries = [
        entry for entry in st.session_state.qoo10_entries
        if not entry.get("_auto_imported") or entry.get("_file_key") in current_qoo10_keys
    ]
    st.session_state.qoo10_auto_imported_keys = set(st.session_state.qoo10_auto_imported_keys) & current_qoo10_keys

    existing_auto_keys = {
        entry.get("_file_key") for entry in st.session_state.qoo10_entries
        if entry.get("_auto_imported") and entry.get("_file_key")
    }
    for f, payload, file_key in qoo10_uploads:
        if file_key in existing_auto_keys or file_key in st.session_state.qoo10_auto_imported_keys:
            continue
        result = _parse_uploaded_qoo10(f.name, payload)
        st.session_state.qoo10_auto_imported_keys.add(file_key)
        if not result:
            continue
        st.session_state.qoo10_entries.append({
            "period_start": result.get("period_start", ""),
            "period_end": result.get("period_end", ""),
            "tracking_no": result.get("tracking_no", ""),
            "qty": int(result.get("qty", 0) or 0),
            "amount": float(result.get("amount", 0) or 0),
            "write_date": result.get("write_date", ""),
            "_source_file": f.name,
            "_file_key": file_key,
            "_auto_imported": True,
            "_submitter": result.get("submitter") or {},
        })

st.divider()

# ══════════════════════════════════════════════════════════════════
# STEP 2 — 큐텐재팬 정보 입력
# ══════════════════════════════════════════════════════════════════
st.markdown("### 🇯🇵 STEP 2 — 큐텐재팬 정보 입력")
st.markdown(
    '<div class="warn-box">큐텐재팬 PDF를 업로드하면 아래 입력 목록에 자동 반영됩니다. 필요한 경우 직접 추가할 수도 있습니다.</div>',
    unsafe_allow_html=True,
)

def _fmt_date(v: str) -> str:
    d = re.sub(r"\D", "", str(v or ""))
    if len(d) == 8:
        return f"{d[:4]}-{d[4:6]}-{d[6:8]}"
    return str(v or "").strip()

with st.form("qoo10_add_form", clear_on_submit=True):
    fp1, fp2 = st.columns(2)
    in_ps = fp1.text_input("거래기간 시작일", placeholder="예: 20260101")
    in_pe = fp2.text_input("거래기간 종료일", placeholder="예: 20260131")
    fc1, fc2, fc3, fc4 = st.columns(4)
    in_amount = fc1.number_input("금액(JPY)", min_value=0, value=0, format="%d")
    in_qty = fc2.number_input("건수", min_value=0, value=0, format="%d")
    in_track = fc3.text_input("발송번호", placeholder="예: K2512244647017")
    in_wdate = fc4.text_input("발행일", placeholder="예: 20260205")
    added = st.form_submit_button("➕ 추가", use_container_width=True)

if added:
    if in_amount > 0 or in_qty > 0 or in_track.strip():
        st.session_state.qoo10_entries.append({
            "period_start": _fmt_date(in_ps),
            "period_end": _fmt_date(in_pe),
            "tracking_no": in_track.strip(),
            "qty": int(in_qty),
            "amount": float(in_amount),
            "write_date": _fmt_date(in_wdate),
        })
    else:
        st.warning("금액·건수·발송번호 중 하나는 입력해야 합니다.")

if st.session_state.qoo10_entries:
    visible_cols = ["period_start", "period_end", "tracking_no", "qty", "amount", "write_date"]
    df_show = pd.DataFrame(st.session_state.qoo10_entries)[visible_cols].rename(columns={
        "period_start": "거래기간 시작",
        "period_end": "거래기간 종료",
        "tracking_no": "발송번호",
        "qty": "건수",
        "amount": "금액(JPY)",
        "write_date": "발행일",
    })
    df_show.index = range(1, len(df_show) + 1)
    df_show["금액(JPY)"] = df_show["금액(JPY)"].map(lambda x: f"{int(x):,}")
    df_show["건수"] = df_show["건수"].map(lambda x: f"{int(x):,}")
    st.table(df_show)
    total_amt = sum(e["amount"] for e in st.session_state.qoo10_entries)
    total_qty = sum(e["qty"] for e in st.session_state.qoo10_entries)
    st.caption(f"합계: {len(st.session_state.qoo10_entries)}건 / 수량 {int(total_qty):,} / 금액 {int(total_amt):,} JPY")
    if st.button("🗑️ 큐텐 입력 전체 삭제"):
        st.session_state.qoo10_entries = []
        st.rerun()
else:
    st.caption("아직 추가된 큐텐재팬 건이 없습니다.")

st.divider()

# ══════════════════════════════════════════════════════════════════
# STEP 3 — 생성 문서 선택 및 환율 안내
# ══════════════════════════════════════════════════════════════════
st.markdown("### ✅ STEP 3 — 생성할 문서 선택")
cc1, cc2, cc3 = st.columns(3)
make_sales = cc1.checkbox("매출집계", value=True)
make_zero = cc2.checkbox("영세율첨부서류제출명세서", value=True)
make_export = cc3.checkbox("수출실적명세서", value=False)

zero_doc_mode = "전체"
if make_zero:
    zero_doc_mode = st.radio(
        "영세율첨부서류제출명세서 생성 범위",
        ["전체", "월별"],
        horizontal=True,
        help="전체를 선택하면 전체 통합 파일 1개만, 월별을 선택하면 월별 파일만 생성합니다.",
    )

st.markdown(
    '<div class="info-box">환율은 서울외국환중개 기간별 매매기준율에서 자동 수집합니다. 이미 수집된 환율은 서버 캐시에 저장하고 부족한 구간만 추가 조회합니다.</div>',
    unsafe_allow_html=True,
)

st.divider()

# ══════════════════════════════════════════════════════════════════
# STEP 4 — 처리 시작
# ══════════════════════════════════════════════════════════════════
st.markdown("### ⚡ STEP 4 — 처리 시작")
has_process_input = bool(uploaded_files) or bool(st.session_state.qoo10_entries)
process_btn = st.button(
    "🚀 엑셀 파일 생성하기",
    type="primary",
    use_container_width=True,
    disabled=not has_process_input,
)
if not has_process_input:
    st.caption("PDF를 업로드하거나 큐텐재팬 정보를 입력하면 생성 버튼이 활성화됩니다.")

progress_bar = st.empty()
status_text = st.empty()
log_area = st.empty()


def _needed_currencies(shopee_results, lazada_result, qoo10_result):
    used = set()
    for sd in shopee_results or []:
        if sd.get("currency"):
            used.add(sd["currency"])
    if lazada_result:
        for it in lazada_result.get("items", []):
            if it.get("currency"):
                used.add(it["currency"])
    # 큐텐재팬은 일별/기간평균 환율을 사용하지 않고 반기말(6월/12월)의
    # 서울외국환중개 공식 월평균 매매기준율만 사용합니다.
    # 따라서 큐텐 때문에 JPY 일별 환율을 별도로 수집하지 않습니다.
    return sorted(used)



def _qoo10_reporting_month(entry=None, result=None):
    """큐텐 거래기간 기준 반기말 월(YYYY-06 또는 YYYY-12)을 반환합니다."""
    entry = entry or {}
    result = result or {}

    period_end = entry.get("period_end") or result.get("period_end") or ""
    digits = re.sub(r"\D", "", str(period_end))[:8]
    if len(digits) >= 6:
        year = digits[:4]
        month = int(digits[4:6])
        return f"{year}-06" if month <= 6 else f"{year}-12"

    base = (
        entry.get("period_start") or result.get("period_start")
        or entry.get("write_date") or result.get("write_date") or ""
    )
    digits = re.sub(r"\D", "", str(base))[:8]
    if len(digits) >= 6:
        year = digits[:4]
        month = int(digits[4:6])
        return f"{year}-06" if month <= 6 else f"{year}-12"
    return ""


def _monthly_rate_requests(ebay_results, qoo10_result):
    """통화별로 공식 월평균 환율이 필요한 월 목록을 만듭니다."""
    requests = {}

    # 이베이: PDF의 실제 발행월별 월평균 환율
    for er in ebay_results or []:
        for it in er.get("items", []):
            currency = str(it.get("currency", "")).strip().upper()
            month_value = str(it.get("month", "")).strip()
            if currency and re.fullmatch(r"20\d{2}-\d{2}", month_value):
                requests.setdefault(currency, set()).add(month_value)

    # 큐텐재팬: 거래기간이 속한 반기의 말월(6월 또는 12월) 월평균 환율
    if qoo10_result:
        entries = qoo10_result.get("entries") or [{}]
        for entry in entries:
            month_value = _qoo10_reporting_month(entry, qoo10_result)
            if month_value:
                requests.setdefault("JPY", set()).add(month_value)

    return {currency: sorted(months) for currency, months in requests.items() if months}


def _filter_monthly_rate_data(rate_data, requested_months):
    """조회 구간 중 실제 필요한 월만 환율 시트에 남깁니다."""
    wanted = set(requested_months or [])
    filtered = [
        row for row in (rate_data.get("monthly", []) or [])
        if str(row.get("year_month", "")) in wanted
    ]
    values = [float(row.get("rate", 0) or 0) for row in filtered if float(row.get("rate", 0) or 0) > 0]
    result = dict(rate_data or {})
    result["monthly"] = filtered
    result["period"] = (
        f"{requested_months[0]} ~ {requested_months[-1]}" if requested_months else ""
    )
    result["average"] = round(sum(values) / len(values), 2) if values else 0.0
    result["monthly_average"] = result["average"]
    result["min"] = min(values) if values else 0.0
    result["max"] = max(values) if values else 0.0
    result["range"] = round(max(values) - min(values), 2) if values else 0.0
    return result

def _daily_rate_period_bounds(shopee_results, lazada_result, qoo10_result):
    """실제 신고기간의 시작/종료일을 반환합니다. 작성일은 환율시트 기간에 포함하지 않습니다."""
    starts = []
    ends = []

    def _add(start_value, end_value):
        sdt = pd.to_datetime(str(start_value or "").replace(".", "-"), errors="coerce")
        edt = pd.to_datetime(str(end_value or "").replace(".", "-"), errors="coerce")
        if not pd.isna(sdt):
            starts.append(sdt.normalize())
        if not pd.isna(edt):
            ends.append(edt.normalize())

    for sd in shopee_results or []:
        _add(sd.get("period_start"), sd.get("period_end"))
        if not sd.get("period_start") or not sd.get("period_end"):
            tx_dates = [pd.to_datetime(str(tx.get("date", "")).replace(".", "-"), errors="coerce") for tx in sd.get("transactions", [])]
            tx_dates = [d.normalize() for d in tx_dates if not pd.isna(d)]
            if tx_dates:
                starts.append(min(tx_dates)); ends.append(max(tx_dates))

    if lazada_result:
        _add(lazada_result.get("period_start"), lazada_result.get("period_end"))

    if qoo10_result:
        _add(qoo10_result.get("period_start"), qoo10_result.get("period_end"))
        for entry in qoo10_result.get("entries", []):
            _add(entry.get("period_start"), entry.get("period_end"))

    if not starts and not ends:
        return None, None
    display_start = min(starts or ends)
    display_end = max(ends or starts)
    return display_start, display_end


def _build_qoo10_result():
    entries = list(st.session_state.qoo10_entries)
    if not entries:
        return None
    submitter = next((e.get("_submitter") for e in entries if (e.get("_submitter") or {}).get("name")), {})
    return {
        "submitter": submitter,
        "type": "qoo10",
        "carrier": "국제로지스틱",
        "destination": "JP",
        "currency": "JPY",
        "period_start": min((e.get("period_start") for e in entries if e.get("period_start")), default=""),
        "period_end": max((e.get("period_end") for e in entries if e.get("period_end")), default=""),
        "write_date": max((e.get("write_date") for e in entries if e.get("write_date")), default=""),
        "qty": sum(int(e.get("qty", 0) or 0) for e in entries),
        "amount": sum(float(e.get("amount", 0) or 0) for e in entries),
        "tracking_no": entries[0].get("tracking_no", "") if entries else "",
        "entries": entries,
    }

if process_btn:
    if not (make_sales or make_zero or make_export):
        st.error("생성할 문서를 하나 이상 선택해 주세요.")
        st.stop()

    logs = []
    def log(msg):
        logs.append(str(msg))
        log_area.markdown('<div class="log-box">' + "\n".join(logs[-120:]) + '</div>', unsafe_allow_html=True)

    st.session_state.result_files = []
    progress_bar.progress(3, text="처리 준비 중...")
    status_text.info("PDF 파싱, 환율 수집, 엑셀 생성 중입니다...")

    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        pdf_paths = []
        for uf in (uploaded_files or []):
            p = tmpdir / uf.name
            p.write_bytes(uf.getbuffer())
            pdf_paths.append(p)

        try:
            # PDF 파싱
            t_pdf = time.perf_counter()
            progress_bar.progress(15, text="📄 PDF 분석 중...")
            log("📄 PDF 분석 중...")
            shopee_results = []
            lazada_result = None
            ebay_results = []
            for p in pdf_paths:
                selected_type = uploaded_type_choices.get(p.name, detect_pdf_type(p.name))
                forced_type = selected_type if detect_pdf_type(p.name) == "unknown" else None
                detected_type = forced_type or detect_pdf_type(p.name)
                # 큐텐 PDF는 업로드 단계에서 파싱되어 STEP 2 목록에 자동 반영됩니다.
                if detected_type == "qoo10":
                    matched = [e for e in st.session_state.qoo10_entries if e.get("_source_file") == p.name]
                    if matched:
                        e = matched[0]
                        log(f"[OK] 큐텐재팬: {p.name} / {int(e.get('qty', 0)):,}건 / {int(e.get('amount', 0)):,} JPY")
                    else:
                        log(f"[WARN] 큐텐재팬 PDF 자동입력 실패: {p.name} / STEP 2에서 직접 입력")
                    continue
                buf = io.StringIO()
                with contextlib.redirect_stdout(buf):
                    result = parse_pdf(str(p), forced_type=forced_type)
                if not result:
                    log(f"[WARN] 파싱 실패 또는 미확인: {p.name}")
                    continue
                if result.get("type") == "shopee":
                    shopee_results.append(result)
                    log(f"[OK] 쇼피 {result.get('currency','?')}: {p.name} / {result.get('total_qty',0):,}건")
                elif result.get("type") == "lazada":
                    lazada_result = result
                    log(f"[OK] 라자다: {p.name} / {len(result.get('items', [])):,}건")
                elif result.get("type") == "ebay":
                    ebay_results.append(result)
                    log(f"[OK] 이베이: {p.name} / {len(result.get('items', [])):,}건")

            qoo10_result = _build_qoo10_result()
            if qoo10_result:
                log(f"[OK] 큐텐 STEP 2: {len(qoo10_result.get('entries', [])):,}건 / {int(qoo10_result.get('amount',0)):,} JPY")

            if not shopee_results and not lazada_result and not qoo10_result and not ebay_results:
                raise RuntimeError("처리할 데이터가 없습니다. PDF 또는 큐텐 수동 입력을 확인해 주세요.")
            log(f"✅ PDF 분석 완료 ({time.perf_counter() - t_pdf:.1f}초)")

            # 환율 수집
            daily_needed = _needed_currencies(shopee_results, lazada_result, qoo10_result)
            monthly_requests = _monthly_rate_requests(ebay_results, qoo10_result)
            display_start, display_end = _daily_rate_period_bounds(shopee_results, lazada_result, qoo10_result)
            if display_start is None or display_end is None:
                today = pd.Timestamp.today().normalize()
                display_start = today
                display_end = today
            # 1월 1일 등 휴일의 직전 영업일 환율을 확보하기 위해 7일 앞에서부터 수집합니다.
            rate_start = display_start - pd.Timedelta(days=RATE_LOOKBACK_DAYS)
            rate_end = display_end

            t_rate = time.perf_counter()
            progress_bar.progress(45, text="💱 환율 확인 중...")
            rates = {}
            if daily_needed:
                rates = fetch_all_currencies_for_period(
                    rate_start, rate_end, daily_needed, logger=log,
                    display_start=display_start, display_end=display_end,
                )
            if monthly_requests:
                requested_currencies = sorted(monthly_requests)
                log(f"💱 월평균 환율 확인 중... ({', '.join(requested_currencies)})")
                monthly_rates = {}
                silent_logger = lambda _msg: None
                for currency in requested_currencies:
                    requested_months = monthly_requests[currency]
                    fetched = fetch_monthly_avg_currencies_for_period(
                        requested_months[0], requested_months[-1], [currency], logger=silent_logger
                    )
                    monthly_rates[currency] = _filter_monthly_rate_data(
                        fetched[currency], requested_months
                    )
                rates = merge_monthly_rates(rates, monthly_rates)
                log("✅ 월평균 환율 확인 완료")
            log(f"✅ 환율 확인 완료 ({time.perf_counter() - t_rate:.1f}초)")

            # 출력 라벨/파일명
            year = rate_end.year
            month = rate_end.month
            disp_label, fname_label = period_labels(shopee_results, lazada_result, qoo10_result, ebay_results=ebay_results, fallback=f"{year}년 {month:02d}월")
            fsafe = safe_filename(fname_label or f"{year}{month:02d}")
            company = company_name_from_results(shopee_results, lazada_result, qoo10_result, ebay_results=ebay_results)

            created = []
            t_excel = time.perf_counter()
            progress_bar.progress(75, text="📊 엑셀 생성 중...")
            log("📊 선택한 문서 생성 중...")

            if make_sales:
                sales_path = tmpdir / f"매출집계_{fsafe}.xlsx"
                buf = io.StringIO()
                with contextlib.redirect_stdout(buf):
                    generate_excel(
                        shopee_results=shopee_results,
                        lazada_result=lazada_result,
                        qoo10_result=qoo10_result,
                        rates=rates,
                        output_path=str(sales_path),
                        ebay_results=ebay_results,
                        year=year,
                        month=month,
                    )
                created.append(sales_path)
                log(f"[OK] 매출집계 생성: {sales_path.name}")

            if make_zero or make_export:
                rows = build_declaration_rows(shopee_results, lazada_result, qoo10_result, rates, ebay_results=ebay_results)
                if make_zero:
                    zero_mode_arg = "all" if zero_doc_mode == "전체" else "monthly"
                    zero_files = create_zero_rate_attachments(
                        rows,
                        tmpdir,
                        company,
                        base_dir=BASE_DIR,
                        mode=zero_mode_arg,
                    )
                    created.extend(zero_files)
                    log(f"[OK] 영세율첨부서류제출명세서 생성({zero_doc_mode}): {len(zero_files)}개")
                if make_export:
                    export_file = create_export_performance(rows, tmpdir, company, base_dir=BASE_DIR)
                    created.append(export_file)
                    log(f"[OK] 수출실적명세서 생성: {export_file.name}")

            result_files = []
            for p in created:
                result_files.append({
                    "name": p.name,
                    "bytes": p.read_bytes(),
                    "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                })
            st.session_state.result_files = result_files
            log(f"✅ 문서 생성 완료 ({time.perf_counter() - t_excel:.1f}초)")
            progress_bar.progress(100, text="✅ 완료")
            status_text.success(f"✅ 엑셀 생성 완료! — {disp_label}")
            log("✅ 전체 처리 완료")
        except Exception as e:
            progress_bar.progress(100, text="오류 발생")
            status_text.error(f"❌ 오류: {e}")
            st.exception(e)

if st.session_state.result_files:
    st.divider()
    st.markdown("### 📥 결과 파일 다운로드")
    for i, f in enumerate(st.session_state.result_files):
        st.download_button(
            f"⬇️ {f['name']}",
            data=f["bytes"],
            file_name=f["name"],
            mime=f["mime"],
            key=f"download_{i}_{f['name']}",
            use_container_width=True,
        )

st.divider()
with st.expander("📌 파일명 규칙 안내"):
    st.markdown(
        """
| 파일명 패턴 | 플랫폼 |
|---|---|
| `유엠(UM)_MY_*.pdf` | 쇼피 말레이시아 |
| `유엠(UM)_PH_*.pdf` | 쇼피 필리핀 |
| `유엠(UM)_SG_*.pdf` | 쇼피 싱가폴 |
| `유엠(UM)_TH_*.pdf` | 쇼피 태국 |
| `유엠(UM)_TW_*.pdf` | 쇼피 대만 |
| `유엠(UM)_VN_*.pdf` | 쇼피 베트남 |
| `라자다_*.pdf` | 라자다 |
| `큐텐재팬_*.pdf` | 큐텐재팬 — PDF 자동인식 후 STEP 2에 반영 |

참고: 쇼피는 업체명과 무관하게 `_MY_`, `_PH_`, `_SG_`, `_TH_`, `_TW_`, `_VN_`, `_BR_`, `_MX_` 국가코드 패턴도 함께 인식합니다.
"""
    )
