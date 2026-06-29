# -*- coding: utf-8 -*-
"""
소포수령증 자동화 웹앱 — GitHub 기존버전 기반 + v29 환율/문서선택 반영
실행: streamlit run app.py
"""

from __future__ import annotations

import contextlib
import io
import os
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
from modules.exchange_rate import fetch_all_currencies_for_period, RATE_LOOKBACK_DAYS
from modules.extra_docs import (
    build_declaration_rows,
    company_name_from_results,
    create_export_performance,
    create_zero_rate_attachments,
    safe_filename,
)

CURRENCIES = ["MYR", "PHP", "SGD", "THB", "TWD", "VND", "JPY", "BRL", "MXN"]

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

# ══════════════════════════════════════════════════════════════════
# STEP 1 — PDF 업로드
# ══════════════════════════════════════════════════════════════════
st.markdown("### 📄 STEP 1 — 소포수령증 PDF 업로드")
c_desc, c_reset = st.columns([6, 1])
c_desc.caption("쇼피, 라자다 PDF를 한꺼번에 올려주세요. 큐텐재팬은 STEP 2에서 직접 입력합니다.")
if c_reset.button("🔄 초기화"):
    st.session_state.uploader_key += 1
    st.session_state.qoo10_entries = []
    st.session_state.result_files = []
    st.rerun()

uploaded_files = st.file_uploader(
    "PDF 파일 선택",
    type=["pdf"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    key=f"pdf_uploader_{st.session_state.uploader_key}",
)

if uploaded_files:
    cols = st.columns(2)
    for i, f in enumerate(uploaded_files):
        ptype = detect_pdf_type(f.name)
        icon = {"shopee": "🛍️", "lazada": "🟠", "qoo10": "🇯🇵", "unknown": "❓"}.get(ptype, "📄")
        label = {"shopee": "쇼피", "lazada": "라자다", "qoo10": "큐텐재팬", "unknown": "미확인"}.get(ptype, "")
        cols[i % 2].markdown(f"{icon} `{f.name}` — {label}")

st.divider()

# ══════════════════════════════════════════════════════════════════
# STEP 2 — 큐텐재팬 정보 입력
# ══════════════════════════════════════════════════════════════════
st.markdown("### 🇯🇵 STEP 2 — 큐텐재팬 정보 입력")
st.markdown(
    '<div class="warn-box">큐텐재팬 PDF는 자동 추출이 불안정하므로 기존버전과 동일하게 직접 입력합니다.</div>',
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
    df_show = pd.DataFrame(st.session_state.qoo10_entries).rename(columns={
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
make_export = cc3.checkbox("수출실적명세서", value=True)

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
process_btn = st.button("🚀 엑셀 파일 생성하기", type="primary", use_container_width=True, disabled=not bool(uploaded_files))
if not uploaded_files:
    st.caption("PDF를 업로드하면 생성 버튼이 활성화됩니다.")

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
    if qoo10_result:
        used.add("JPY")
    return sorted(used)


def _date_values_for_rates(shopee_results, lazada_result, qoo10_result):
    values = []
    for sd in shopee_results or []:
        for key in ["period_start", "period_end", "write_date"]:
            if sd.get(key):
                values.append(sd[key])
        for tx in sd.get("transactions", []):
            if tx.get("date"):
                values.append(tx["date"])
    if lazada_result:
        for key in ["period_start", "period_end", "write_date"]:
            if lazada_result.get(key):
                values.append(lazada_result[key])
    if qoo10_result:
        for key in ["period_start", "period_end", "write_date"]:
            if qoo10_result.get(key):
                values.append(qoo10_result[key])
        for e in qoo10_result.get("entries", []):
            for key in ["period_start", "period_end", "write_date"]:
                if e.get(key):
                    values.append(e[key])
    parsed = []
    for v in values:
        dt = pd.to_datetime(str(v).replace(".", "-"), errors="coerce")
        if not pd.isna(dt):
            parsed.append(dt.normalize())
    return parsed


def _build_qoo10_result():
    entries = list(st.session_state.qoo10_entries)
    if not entries:
        return None
    return {
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
        for uf in uploaded_files:
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
            for p in pdf_paths:
                # 큐텐은 STEP 2 수동 입력을 사용하므로 PDF OCR은 생략합니다.
                if detect_pdf_type(p.name) == "qoo10":
                    log(f"[SKIP] 큐텐재팬 PDF: {p.name} / STEP 2 수동 입력 사용")
                    continue
                buf = io.StringIO()
                with contextlib.redirect_stdout(buf):
                    result = parse_pdf(str(p))
                if not result:
                    log(f"[WARN] 파싱 실패 또는 미확인: {p.name}")
                    continue
                if result.get("type") == "shopee":
                    shopee_results.append(result)
                    log(f"[OK] 쇼피 {result.get('currency','?')}: {p.name} / {result.get('total_qty',0):,}건")
                elif result.get("type") == "lazada":
                    lazada_result = result
                    log(f"[OK] 라자다: {p.name} / {len(result.get('items', [])):,}건")

            qoo10_result = _build_qoo10_result()
            if qoo10_result:
                log(f"[OK] 큐텐 수동 입력: {len(qoo10_result.get('entries', [])):,}건 / {int(qoo10_result.get('amount',0)):,} JPY")

            if not shopee_results and not lazada_result and not qoo10_result:
                raise RuntimeError("처리할 데이터가 없습니다. PDF 또는 큐텐 수동 입력을 확인해 주세요.")
            log(f"✅ PDF 분석 완료 ({time.perf_counter() - t_pdf:.1f}초)")

            # 환율 수집
            needed = _needed_currencies(shopee_results, lazada_result, qoo10_result)
            date_values = _date_values_for_rates(shopee_results, lazada_result, qoo10_result)
            if not date_values:
                today = pd.Timestamp.today().normalize()
                rate_start = today - pd.Timedelta(days=RATE_LOOKBACK_DAYS)
                rate_end = today
            else:
                rate_start = min(date_values) - pd.Timedelta(days=RATE_LOOKBACK_DAYS)
                rate_end = max(date_values)

            t_rate = time.perf_counter()
            progress_bar.progress(45, text="💱 환율 확인 중...")
            rates = fetch_all_currencies_for_period(rate_start, rate_end, needed, logger=log)
            log(f"✅ 환율 확인 완료 ({time.perf_counter() - t_rate:.1f}초)")

            # 출력 라벨/파일명
            year = rate_end.year
            month = rate_end.month
            disp_label, fname_label = period_labels(shopee_results, lazada_result, qoo10_result, fallback=f"{year}년 {month:02d}월")
            fsafe = safe_filename(fname_label or f"{year}{month:02d}")
            company = company_name_from_results(shopee_results, lazada_result, qoo10_result)

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
                        year=year,
                        month=month,
                    )
                created.append(sales_path)
                log(f"[OK] 매출집계 생성: {sales_path.name}")

            if make_zero or make_export:
                rows = build_declaration_rows(shopee_results, lazada_result, qoo10_result, rates)
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
| `큐텐재팬_*.pdf` | 큐텐재팬 — STEP 2에서 수동 입력 |

참고: 쇼피는 업체명과 무관하게 `_MY_`, `_PH_`, `_SG_`, `_TH_`, `_TW_`, `_VN_`, `_BR_`, `_MX_` 국가코드 패턴도 함께 인식합니다.
"""
    )
