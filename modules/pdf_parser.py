"""
소포수령증 PDF 파싱 모듈
- 쇼피(Shopee): 두라로지스틱스, 텍스트 기반 PDF
- 라자다(Lazada): 용성종합물류, 텍스트 기반 PDF (요약)
- 큐텐(Qoo10): 국제로지스틱, 표/텍스트 우선 + OCR 보완 자동 추출
"""

import pdfplumber
import re
from pathlib import Path
from typing import Optional


# ── 통화 매핑 ──────────────────────────────────────────────────
COUNTRY_TO_CURRENCY = {
    'MY': 'MYR', 'PH': 'PHP', 'SG': 'SGD',
    'TH': 'THB', 'TW': 'TWD', 'VN': 'VND', 'JP': 'JPY',
    'BR': 'BRL', 'MX': 'MXN',
    'US': 'USD', 'EU': 'EUR', 'GB': 'GBP', 'CA': 'CAD', 'AU': 'AUD',
}

CURRENCY_NAMES_KR = {
    'MYR': '말레이시아 링깃 (MYR)',
    'PHP': '필리핀 페소 (PHP)',
    'SGD': '싱가포르 달러 (SGD)',
    'THB': '태국 바트 (THB)',
    'TWD': '대만 달러 (TWD)',
    'VND': '베트남 동 (VND)',
    'JPY': '일본 엔 (JPY) (100)',
    'USD': '미국 달러 (USD)',
    'EUR': '유로 (EUR)',
    'GBP': '영국 파운드 (GBP)',
    'CAD': '캐나다 달러 (CAD)',
    'AUD': '호주 달러 (AUD)',
}

# 파일명에서 국가코드 감지
SHOPEE_FILE_PATTERNS = {
    '_MY_': 'MY', '_PH_': 'PH', '_SG_': 'SG',
    '_TH_': 'TH', '_TW_': 'TW', '_VN_': 'VN',
    '_BR_': 'BR', '_MX_': 'MX',
}


def _detect_pdf_type_from_text(text: str) -> str:
    """PDF 본문 표식으로 플랫폼을 판별합니다."""
    raw = str(text or "")
    compact = re.sub(r"\s+", "", raw).lower()

    # 큐텐/국제로지스틱 양식: 판매처 Qoo10 + JPY 금액표가 핵심 표식입니다.
    if ("qoo10" in compact and "국제로지스틱" in compact) or (
        "금액(jpy)" in compact and "국제로지스틱" in compact
    ):
        return "qoo10"

    # 이베이/린코스 양식은 발행월별 내역과 린코스 배송사 표식이 함께 있습니다.
    if "린코스" in compact or "lincos" in compact:
        return "ebay"

    if "라자다" in compact or "lazada" in compact or "용성종합물류" in compact:
        return "lazada"

    # 쇼피/두라로지스틱스 양식은 상세 운송장 내역이 존재합니다.
    if "두라로지스틱스" in compact and ("운송장번호" in compact or "수출신고금액" in compact):
        return "shopee"

    return "unknown"


def _extract_pdf_text_for_detection(pdf_path: str, max_pages: int = 2) -> str:
    try:
        path = Path(pdf_path)
        if not path.exists() or not path.is_file():
            return ""
        texts = []
        with pdfplumber.open(str(path)) as pdf:
            for page in pdf.pages[:max_pages]:
                text = page.extract_text()
                if text:
                    texts.append(text)
        return "\n".join(texts)
    except Exception:
        return ""


def detect_pdf_type(pdf_path: str) -> str:
    """파일명과 PDF 본문으로 종류 판단 -> shopee/lazada/qoo10/ebay/unknown."""
    name = Path(str(pdf_path)).name
    lower = name.lower()
    if "라자다" in name or "lazada" in lower:
        return "lazada"
    if "큐텐" in name or "qoo10" in lower:
        return "qoo10"
    if "이베이" in name or "ebay" in lower or "린코스" in name or "lincos" in lower:
        return "ebay"
    # 쇼피: 업체명과 무관하게 파일명의 국가코드 패턴(_MY_, _TW_ 등)으로 인식
    if re.search(r"_(MY|PH|SG|TH|TW|VN|BR|MX|JP)_", name):
        return "shopee"
    if "유엠(UM)_" in name or "유엠_" in name:
        return "shopee"

    # 파일명이 일반적인 경우에는 본문 표식을 확인합니다.
    content_type = _detect_pdf_type_from_text(_extract_pdf_text_for_detection(str(pdf_path)))
    return content_type


# ─────────────────────────────────────────────────────────────────
# 쇼피 PDF 파싱
# ─────────────────────────────────────────────────────────────────

def _extract_submitter(full_text: str) -> dict:
    """소포수령증 '1. 제출자 인적사항'에서 상호·사업자번호·대표자·주소 추출."""
    sub = {'name': '', 'biz_no': '', 'ceo': '', 'address': ''}
    m = re.search(r'사업자등록번호\s+(\d{3}-\d{2}-\d{5})', full_text)
    if m: sub['biz_no'] = m.group(1)
    m = re.search(r'상호\(법인명\)\s+(.+)', full_text)
    if m: sub['name'] = m.group(1).strip()
    m = re.search(r'대표자\s*성명\s+(\S+)', full_text)
    if m: sub['ceo'] = m.group(1).strip()
    # 주소: 제출자 섹션의 비라벨 줄 + 거래기간 줄 끝의 (괄호)
    addr = []
    in_sec = False
    for ln in full_text.splitlines():
        t = ln.strip()
        if t.startswith('1. 제출자'):
            in_sec = True
            continue
        if t.startswith('2.'):
            break
        if not in_sec:
            continue
        if any(t.startswith(k) for k in ('사업자등록번호', '상호(법인명)', '대표자', '거래기간')):
            if t.startswith('거래기간'):
                mm = re.search(r'(\([^)]+\))\s*$', t)
                if mm:
                    addr.append(mm.group(1))
            continue
        if t:
            addr.append(t)
    sub['address'] = ' '.join(addr)
    return sub


def parse_shopee_pdf(pdf_path: str) -> dict:
    """
    쇼피 소포수령증 PDF 파싱
    Returns: {
        'type': 'shopee',
        'carrier': str,       # 해외배송업체
        'country': str,       # 'MY', 'PH', ...
        'currency': str,      # 'MYR', 'PHP', ...
        'period_start': str,  # '2025-12-01'
        'period_end': str,    # '2025-12-31'
        'write_date': str,    # '2026-01-09'
        'total_qty': int,
        'total_amount': float,
        'transactions': [{'carrier','date','tracking_no','country','qty','amount'}]
    }
    """
    # 파일명에서 국가 파악
    filename = Path(pdf_path).name
    country = ''
    for pattern, cc in SHOPEE_FILE_PATTERNS.items():
        if pattern in filename:
            country = cc
            break

    # PDF 전체 텍스트 추출
    all_lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_lines.extend(text.splitlines())

    full_text = '\n'.join(all_lines)

    # ── 헤더 정보 파싱 ──────────────────────────
    carrier = '주)두라로지스틱스'  # 기본값

    # 거래기간
    period_match = re.search(r'(\d{4}-\d{2}-\d{2})\s*~\s*(\d{4}-\d{2}-\d{2})', full_text)
    period_start = period_match.group(1) if period_match else ''
    period_end   = period_match.group(2) if period_match else ''

    # 작성일자
    write_match = re.search(r'작성일자\s+([\d-]+)', full_text)
    write_date = write_match.group(1) if write_match else ''

    # 통화
    currency_match = re.search(r'통화\s+([A-Z]{3})', full_text)
    if currency_match:
        currency = currency_match.group(1)
        if not country:
            # 통화로 국가 역추적
            for cc, cur in COUNTRY_TO_CURRENCY.items():
                if cur == currency:
                    country = cc
                    break
    else:
        currency = COUNTRY_TO_CURRENCY.get(country, '')

    # 합계 수량 + 금액 — 한 줄에 "합계 781 162,001.42" 형식
    total_qty    = 0
    total_amount = 0.0

    # 1차: "합계 <수량> <금액>" 패턴
    summary_match = re.search(r'합계\s+([\d,]+)\s+([\d,]+\.?\d*)', full_text)
    if summary_match:
        total_qty    = int(summary_match.group(1).replace(',', ''))
        total_amount = float(summary_match.group(2).replace(',', ''))
    else:
        # 2차: "MYR <수량> <금액>" 패턴 (통화코드 + 수량 + 금액)
        currency_line = re.search(rf'{currency}\s+([\d,]+)\s+([\d,]+\.?\d*)', full_text)
        if currency_line:
            total_qty    = int(currency_line.group(1).replace(',', ''))
            total_amount = float(currency_line.group(2).replace(',', ''))

    # ── 거래 내역 파싱 ──────────────────────────
    transactions = []
    in_section = False

    # 데이터 행 패턴: 배송업체 날짜(YYYY.MM.DD) 운송장번호 국가 수량 금액
    tx_pattern = re.compile(
        r'^(.+?)\s+'                      # 배송업체
        r'(\d{4}[-\.]\d{2}[-\.]\d{2})\s+'  # 날짜
        r'([A-Z0-9]{10,})\s+'              # 운송장번호
        r'([A-Z]{2})\s+'                   # 국가코드
        r'(\d+)\s+'                        # 수량
        r'([\d,]+\.?\d*)$'                 # 금액
    )

    for line in all_lines:
        line = line.strip()
        if '3.' in line and '해외배송' in line and '내역' in line:
            in_section = True
            continue
        if in_section and '상기 내역' in line:
            break
        if in_section:
            m = tx_pattern.match(line)
            if m:
                # 날짜 정규화 (2025-12-03 → 2025.12.03)
                date_str = m.group(2).replace('-', '.')
                transactions.append({
                    'carrier':     m.group(1).strip(),
                    'date':        date_str,
                    'tracking_no': m.group(3),
                    'country':     m.group(4),
                    'qty':         int(m.group(5)),
                    'amount':      float(m.group(6).replace(',', '')),
                })

    # 합계 금액 보정 (파싱 실패 시 거래 합산)
    if total_amount == 0.0 and transactions:
        total_amount = sum(t['amount'] for t in transactions)
    if total_qty == 0 and transactions:
        total_qty = sum(t['qty'] for t in transactions)

    # 배송업체명 (첫 거래에서)
    if transactions:
        carrier = transactions[0]['carrier']

    submitter = _extract_submitter(full_text)

    return {
        'type':         'shopee',
        'submitter':    submitter,
        'carrier':      carrier,
        'country':      country,
        'currency':     currency,
        'period_start': period_start,
        'period_end':   period_end,
        'write_date':   write_date,
        'total_qty':    total_qty,
        'total_amount': total_amount,
        'transactions': transactions,
    }


# ─────────────────────────────────────────────────────────────────
# 라자다 PDF 파싱
# ─────────────────────────────────────────────────────────────────

def parse_lazada_pdf(pdf_path: str) -> dict:
    """
    라자다 소포수령증 PDF 파싱 (요약 형태)
    Returns: {
        'type': 'lazada',
        'carrier': str,
        'period_start': str,
        'period_end': str,
        'write_date': str,
        'items': [{'service','carrier','origin','destination','tracking_no','qty','amount','currency'}]
    }
    """
    all_lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_lines.extend(text.splitlines())

    full_text = '\n'.join(all_lines)

    # 배송사
    carrier_match = re.search(r'해외배송업체\s+(.+?)\s+출발', full_text)
    carrier = carrier_match.group(1).strip() if carrier_match else '용성종합물류'

    # 거래기간
    period_match = re.search(r'(\d{4}\.\d{2}\.\d{2})\s*[–-]\s*(\d{4}\.\d{2}\.\d{2})', full_text)
    if period_match:
        period_start = period_match.group(1).replace('.', '-')
        period_end   = period_match.group(2).replace('.', '-')
    else:
        period_start = period_end = ''

    # 작성일자
    write_match = re.search(r'작성일자\s+([\d.]+)', full_text)
    write_date = write_match.group(1).replace('.', '-') if write_match else ''

    # 배송 내역 파싱
    items = []
    item_pattern = re.compile(
        r'(라자다)\s+'                        # 서비스
        r'([\w\s가-힣]+?)\s+'                 # 배송업체
        r'([A-Z]{2})\s+'                      # 출발국
        r'([A-Z]{2})\s+'                      # 도착국
        r'([\w\d]+(?:외\d+\n?건)?)\s+'        # 발송번호
        r'([\d,]+)건\s+'                      # 수량
        r'([\d,]+\.?\d*)\(([A-Z]{3})\)'      # 금액(통화)
    )

    for line in all_lines:
        line = line.strip()
        m = item_pattern.search(line)
        if m:
            items.append({
                'service':     m.group(1),
                'carrier':     m.group(2).strip(),
                'origin':      m.group(3),
                'destination': m.group(4),
                'tracking_no': m.group(5).replace('\n', ''),
                'qty':         int(m.group(6).replace(',', '')),
                'amount':      float(m.group(7).replace(',', '')),
                'currency':    m.group(8),
            })

    # 패턴이 너무 엄격하면 간단 방식으로 재시도
    if not items:
        for line in all_lines:
            m = re.search(r'라자다.*?([A-Z]{2})\s+([\w\d]+(?:외\d+건)?)\s+([\d,]+)건\s+([\d,]+\.?\d*)\(([A-Z]{3})\)', line)
            if m:
                items.append({
                    'service':     '라자다',
                    'carrier':     carrier,
                    'origin':      'KR',
                    'destination': m.group(1),
                    'tracking_no': m.group(2),
                    'qty':         int(m.group(3).replace(',', '')),
                    'amount':      float(m.group(4).replace(',', '')),
                    'currency':    m.group(5),
                })

    return {
        'type':         'lazada',
        'carrier':      carrier,
        'period_start': period_start,
        'period_end':   period_end,
        'write_date':   write_date,
        'items':        items,
    }



# ─────────────────────────────────────────────────────────────────
# 이베이/린코스 PDF 파싱
# ─────────────────────────────────────────────────────────────────

def _normalize_korean_date(value: str) -> str:
    """2026년01월01일 / 2026-01-01 / 2026.01.01 → YYYY-MM-DD"""
    d = re.sub(r'\D', '', str(value or ''))
    if len(d) >= 8:
        return f"{d[:4]}-{d[4:6]}-{d[6:8]}"
    return ''


def _normalize_month(value: str) -> str:
    """2026-03 / 2026년 03월 → YYYY-MM"""
    text = str(value or '').strip()
    m = re.search(r'(20\d{2})\D*([01]?\d)', text)
    if not m:
        return ''
    y = m.group(1)
    mo = int(m.group(2))
    if not 1 <= mo <= 12:
        return ''
    return f"{y}-{mo:02d}"


def _month_start_end(month_key: str):
    """YYYY-MM → (YYYY-MM-01, YYYY-MM-last_day)"""
    import calendar
    m = _normalize_month(month_key)
    if not m:
        return '', ''
    y, mo = map(int, m.split('-'))
    last = calendar.monthrange(y, mo)[1]
    return f"{y}-{mo:02d}-01", f"{y}-{mo:02d}-{last:02d}"


def _num(value, default=0.0):
    s = str(value or '').replace(',', '').strip()
    if not s:
        return default
    try:
        return float(s)
    except Exception:
        return default


def parse_ebay_lincos_pdf(pdf_path: str) -> dict:
    """
    이베이 - 린코스 해외배송 소포 수령증 파싱.
    이 양식은 개별 발행일이 아니라 발행월별 합계가 있으므로,
    각 행에 month=YYYY-MM을 저장하고 이후 월평균 매매기준율을 적용합니다.
    """
    submitter = {'name': '', 'biz_no': '', 'ceo': '', 'address': ''}
    carrier = '린코스(주)'
    period_start = ''
    period_end = ''
    write_date = ''
    items = []
    summary_items = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables() or []
            for table in tables:
                if not table:
                    continue
                # 제출자 인적사항 표: 거래기간/작성일자가 함께 있는 상단 표만 사용
                header_text_for_submitter = ' '.join(str(c or '') for row in table for c in row)
                if (any(row and any('사업자등록번호' in str(c or '') for c in row) for row in table)
                        and ('거래기간' in header_text_for_submitter or '작성일자' in header_text_for_submitter)):
                    for row in table:
                        cells = [str(c or '').strip() for c in row]
                        joined = ' '.join(cells)
                        if '사업자등록번호' in joined:
                            for i, c in enumerate(cells):
                                if '사업자등록번호' in c and i + 1 < len(cells):
                                    submitter['biz_no'] = cells[i + 1]
                                if '상호' in c and i + 1 < len(cells):
                                    submitter['name'] = cells[i + 1]
                        if '성명' in joined:
                            for i, c in enumerate(cells):
                                if '성명' in c and i + 1 < len(cells):
                                    submitter['ceo'] = cells[i + 1]
                                if '사업장 소재지' in c and i + 1 < len(cells):
                                    submitter['address'] = cells[i + 1].replace('\n', ' ')
                        if '거래기간' in joined:
                            for i, c in enumerate(cells):
                                if '거래기간' in c and i + 1 < len(cells):
                                    raw = cells[i + 1]
                                    parts = re.split(r'~|–|-{2,}', raw)
                                    if len(parts) >= 2:
                                        period_start = _normalize_korean_date(parts[0])
                                        period_end = _normalize_korean_date(parts[1])
                                if '작성일자' in c and i + 1 < len(cells):
                                    write_date = _normalize_korean_date(cells[i + 1])

                header_text = ' '.join(str(c or '') for row in table for c in row)
                # 2. 소포 수령증 요약 표
                if '현지송장번호' in header_text and '통화단위' in header_text:
                    cur_carrier = carrier
                    cur_service = ''
                    cur_tracking = ''
                    for row in table[1:]:
                        if not row or len(row) < 6:
                            continue
                        r = [(c if c is not None else '') for c in row]
                        if str(r[0]).strip():
                            cur_carrier = str(r[0]).strip()
                        if str(r[1]).strip():
                            cur_service = str(r[1]).strip()
                        if str(r[2]).strip():
                            cur_tracking = str(r[2]).replace('\n', '').strip()
                        currency = str(r[3]).strip().upper()
                        if not re.fullmatch(r'[A-Z]{3}', currency):
                            continue
                        qty = int(_num(r[4], 0))
                        amount = _num(r[5], 0)
                        if qty or amount:
                            summary_items.append({
                                'carrier': cur_carrier or carrier,
                                'service': cur_service,
                                'tracking_no': cur_tracking,
                                'currency': currency,
                                'qty': qty,
                                'amount': amount,
                            })

                # 3. 해외배송 내역서 월별 표
                if '발행월' in header_text and '통화단위' in header_text and '신고금액' in header_text:
                    cur_month = ''
                    cur_carrier = carrier
                    cur_service = ''
                    for row in table:
                        if not row or len(row) < 6:
                            continue
                        r = [(c if c is not None else '') for c in row]
                        first = str(r[0] or '').strip()
                        mkey = _normalize_month(first)
                        if mkey:
                            cur_month = mkey
                        if str(r[1]).strip() and '해외배송업체' not in str(r[1]):
                            cur_carrier = str(r[1]).strip()
                        if str(r[2]).strip() and '배송국가' not in str(r[2]):
                            cur_service = str(r[2]).strip()
                        currency = str(r[3] or '').strip().upper()
                        if not cur_month or not re.fullmatch(r'[A-Z]{3}', currency):
                            continue
                        qty = int(_num(r[4], 0))
                        amount = _num(r[5], 0)
                        if not qty and not amount:
                            continue
                        month_start, month_end = _month_start_end(cur_month)
                        items.append({
                            'carrier': cur_carrier or carrier,
                            'service': cur_service,
                            'country': cur_service,
                            'destination': cur_service,
                            'tracking_no': '',
                            'month': cur_month,
                            'date': month_end,
                            'period_start': month_start,
                            'period_end': month_end,
                            'currency': currency,
                            'qty': qty,
                            'amount': amount,
                            'rate_basis': 'monthly_average',
                        })

    if not period_start or not period_end:
        months = sorted({_normalize_month(it.get('month', '')) for it in items if _normalize_month(it.get('month', ''))})
        if months:
            period_start = _month_start_end(months[0])[0]
            period_end = _month_start_end(months[-1])[1]

    return {
        'type': 'ebay',
        'platform': '이베이',
        'carrier': carrier,
        'submitter': submitter,
        'period_start': period_start,
        'period_end': period_end,
        'write_date': write_date,
        'items': items,
        'summary_items': summary_items,
    }

# ─────────────────────────────────────────────────────────────────
# 큐텐 PDF (이미지 기반 → OCR 자동 추출)
# ─────────────────────────────────────────────────────────────────

def _ocr_pdf(pdf_path: str) -> str:
    """
    pdf2image + pytesseract로 PDF에서 텍스트 추출.
    전체 페이지 OCR + 핵심 영역(합계 행이 있는 중하단부) 별도 OCR을 합쳐서 반환.
    """
    try:
        from pdf2image import convert_from_path
        import pytesseract

        pages = convert_from_path(pdf_path, dpi=200)
        all_texts = []

        for page in pages:
            w, h = page.size

            # 1) 전체 페이지 OCR (기간/날짜 추출)
            full_text = pytesseract.image_to_string(page, lang='eng', config='--psm 6')
            all_texts.append(full_text)

            # 2) 핵심 영역별 별도 OCR (합계 행 등)
            # 합계 행은 보통 50~70% 위치에 있음 — 더 촘촘하게 스캔
            regions = [
                page.crop((0, int(h * 0.50), w, int(h * 0.75))),  # 핵심 영역
                page.crop((0, int(h * 0.45), w, int(h * 0.70))),
                page.crop((0, int(h * 0.40), w, int(h * 0.65))),
                page.crop((0, int(h * 0.55), w, int(h * 0.80))),
                page.crop((0, int(h * 0.35), w, int(h * 0.60))),
            ]
            for region in regions:
                region_text = pytesseract.image_to_string(
                    region, lang='eng',
                    config='--psm 6'
                )
                all_texts.append(region_text)

        return '\n'.join(all_texts)
    except ImportError:
        print("⚠️ pdf2image 또는 pytesseract가 설치되지 않았습니다.")
        return ''
    except Exception as e:
        print(f"⚠️ OCR 오류: {e}")
        return ''


def _clean_tracking_no(value: str) -> str:
    text = str(value or "").replace("\n", " ").strip()
    m = re.search(r"[Kk]\d{13,}", text)
    return m.group(0).upper() if m else ""


def _parse_qoo10_tables(pdf_path: str) -> dict:
    """텍스트형 큐텐 소포수령증의 표를 우선 파싱합니다."""
    submitter = {"name": "", "biz_no": "", "ceo": "", "address": ""}
    period_start = ""
    period_end = ""
    write_date = ""
    carrier = "국제로지스틱"
    tracking_no = ""
    qty = 0
    amount = 0.0

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                for table in page.extract_tables() or []:
                    if not table:
                        continue
                    normalized = [[str(c or "").strip() for c in row] for row in table if row]
                    table_text = " ".join(c for row in normalized for c in row)

                    # 1. 제출자 인적사항
                    if "사업자 등록번호" in table_text and "거래기간" in table_text:
                        for row in normalized:
                            for i, cell in enumerate(row):
                                if "사업자 등록번호" in cell and i + 1 < len(row):
                                    submitter["biz_no"] = row[i + 1]
                                elif "상호" in cell and i + 1 < len(row):
                                    submitter["name"] = row[i + 1]
                                elif "성명" in cell and i + 1 < len(row):
                                    submitter["ceo"] = row[i + 1]
                                elif "사업장 소재지" in cell and i + 1 < len(row):
                                    submitter["address"] = row[i + 1].replace("\n", " ")
                                elif "거래기간" in cell and i + 1 < len(row):
                                    raw = row[i + 1]
                                    m = re.search(
                                        r"(\d{4})[/.-](\d{2})[/.-](\d{2})\s*[~–-]\s*"
                                        r"(\d{4})[/.-](\d{2})[/.-](\d{2})",
                                        raw,
                                    )
                                    if m:
                                        period_start = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
                                        period_end = f"{m.group(4)}-{m.group(5)}-{m.group(6)}"
                                elif "작성일자" in cell and i + 1 < len(row):
                                    m = re.search(r"(\d{4})[-/.](\d{2})[-/.](\d{2})", row[i + 1])
                                    if m:
                                        write_date = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

                    # 3. 해외배송 내역서 상세/합계 행
                    if "금액(JPY)" in table_text and ("발송수량" in table_text or "해외배송업체" in table_text):
                        header_idx = None
                        for idx, row in enumerate(normalized):
                            if any("금액(JPY)" in c for c in row):
                                header_idx = idx
                                break
                        if header_idx is not None:
                            headers = normalized[header_idx]
                            def _idx(keyword):
                                return next((i for i, c in enumerate(headers) if keyword in c), None)
                            i_carrier = _idx("해외배송업체")
                            i_track = _idx("발송번호")
                            i_qty = _idx("발송수량")
                            i_amt = _idx("금액(JPY)")
                            for row in normalized[header_idx + 1:]:
                                if not row or "당기 해외배송 합계" in " ".join(row):
                                    continue
                                if i_carrier is not None and i_carrier < len(row) and row[i_carrier]:
                                    carrier = row[i_carrier]
                                if i_track is not None and i_track < len(row):
                                    tracking_no = _clean_tracking_no(row[i_track]) or tracking_no
                                if i_qty is not None and i_qty < len(row):
                                    m = re.search(r"[\d,]+", row[i_qty])
                                    if m:
                                        qty = int(m.group(0).replace(",", ""))
                                if i_amt is not None and i_amt < len(row):
                                    m = re.search(r"[\d,]+(?:\.\d+)?", row[i_amt])
                                    if m:
                                        amount = float(m.group(0).replace(",", ""))
                                if qty or amount:
                                    break
    except Exception:
        pass

    return {
        "submitter": submitter,
        "carrier": carrier,
        "period_start": period_start,
        "period_end": period_end,
        "write_date": write_date,
        "qty": qty,
        "amount": amount,
        "tracking_no": tracking_no,
    }


def parse_qoo10_pdf(pdf_path: str) -> Optional[dict]:
    """큐텐재팬 소포수령증을 표 우선, 텍스트/OCR 보완 방식으로 파싱합니다."""
    table_data = _parse_qoo10_tables(pdf_path)

    # pdfplumber 텍스트 추출
    all_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text += text + "\n"

    # 이미지 PDF면 OCR 보완
    if not all_text.strip():
        print("  큐텐 PDF 이미지 기반 -> OCR 시도...")
        all_text = _ocr_pdf(pdf_path)
    if not all_text.strip() and not (table_data.get("qty") or table_data.get("amount")):
        print("  큐텐 PDF 텍스트 추출 실패 - 수동 입력 필요")
        return None

    period_start = table_data.get("period_start", "")
    period_end = table_data.get("period_end", "")
    if not period_start or not period_end:
        m = re.search(
            r"(\d{4})[/.-](\d{2})[/.-](\d{2})\s*[~–-]\s*"
            r"(\d{4})[/.-](\d{2})[/.-](\d{2})",
            all_text,
        )
        if m:
            period_start = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
            period_end = f"{m.group(4)}-{m.group(5)}-{m.group(6)}"

    write_date = table_data.get("write_date", "")
    if not write_date:
        m = re.search(r"(\d{4})[-/.](\d{2})[-/.](\d{2})(?:\s+\d{2}:\d{2}:\d{2})?", all_text)
        if m:
            # 거래기간 첫 날짜가 아니라 작성일자 표식 뒤의 날짜를 우선 재탐색
            wm = re.search(r"작성일자\s+(\d{4})[-/.](\d{2})[-/.](\d{2})", all_text)
            src = wm or m
            write_date = f"{src.group(1)}-{src.group(2)}-{src.group(3)}"

    qty = int(table_data.get("qty", 0) or 0)
    amount = float(table_data.get("amount", 0) or 0)

    # 정상 텍스트 표: 상세 행 또는 합계 행의 '23 건 74,170'
    if qty == 0 or amount == 0:
        candidates = re.findall(r"([\d,]+)\s*건\s+([\d,]+(?:\.\d+)?)", all_text)
        if candidates:
            q, a = candidates[-1]
            qty = qty or int(q.replace(",", ""))
            amount = amount or float(a.replace(",", ""))

    # OCR 보완 패턴
    if qty == 0 or amount == 0:
        patterns = [
            r"\b(\d{2,4})\s+[2H건a-zA-Z]\s+([\d,]{5,})\b",
            r"[|｜]\s*([\d,]{2,4})\s*[|｜]\s*([\d,]{5,})",
            r"\b(\d{3,4})\s+\d\s+([\d,]{5,})",
        ]
        for pat in patterns:
            m = re.search(pat, all_text)
            if m:
                qty = qty or int(m.group(1).replace(",", ""))
                amount = amount or float(m.group(2).replace(",", ""))
                break

    tracking_no = table_data.get("tracking_no", "") or _clean_tracking_no(all_text)

    submitter = table_data.get("submitter") or {"name": "", "biz_no": "", "ceo": "", "address": ""}
    if not submitter.get("biz_no"):
        m = re.search(r"사업자\s*등록번호\s+(\d{3}-\d{2}-\d{5})", all_text)
        if m:
            submitter["biz_no"] = m.group(1)
    if not submitter.get("name"):
        m = re.search(r"상호\s*\(법인명\)\s+(.+?)(?:\s+성명\s*\(대표자\)|$)", all_text)
        if m:
            submitter["name"] = m.group(1).strip()
    if not submitter.get("ceo"):
        m = re.search(r"성명\s*\(대표자\)\s+(\S+)", all_text)
        if m:
            submitter["ceo"] = m.group(1).strip()
    if not submitter.get("address"):
        m = re.search(r"사업장\s*소재지\s+(.+?)(?:\n|거래기간)", all_text)
        if m:
            submitter["address"] = m.group(1).strip()

    if qty == 0 and amount == 0.0:
        print("  큐텐 PDF 수량/금액 파싱 실패 - STEP 2에서 직접 입력 필요")
        return None

    print(f"  큐텐 PDF 자동인식 - 기간:{period_start}~{period_end} 수량:{qty}건 JPY:{amount:,.0f}")
    return {
        "type": "qoo10",
        "submitter": submitter,
        "carrier": table_data.get("carrier") or "국제로지스틱",
        "destination": "JP",
        "currency": "JPY",
        "period_start": period_start,
        "period_end": period_end,
        "write_date": write_date,
        "qty": qty,
        "amount": float(amount),
        "tracking_no": tracking_no,
    }


# ─────────────────────────────────────────────────────────────────
# 통합 파싱 함수
# ─────────────────────────────────────────────────────────────────

def parse_pdf(pdf_path: str, forced_type: str = None) -> Optional[dict]:
    """PDF 종류 자동 감지 또는 사용자가 선택한 종류로 파싱"""
    pdf_type = (forced_type or detect_pdf_type(pdf_path) or 'unknown').strip().lower()
    aliases = {
        'ebay_lincos': 'ebay', 'lincos': 'ebay', '이베이': 'ebay',
        '쇼피': 'shopee', '라자다': 'lazada', '큐텐': 'qoo10', '큐텐재팬': 'qoo10',
    }
    pdf_type = aliases.get(pdf_type, pdf_type)
    print(f"  파싱 중: {Path(pdf_path).name} → [{pdf_type}]")

    if pdf_type == 'shopee':
        return parse_shopee_pdf(pdf_path)
    elif pdf_type == 'lazada':
        return parse_lazada_pdf(pdf_path)
    elif pdf_type == 'qoo10':
        return parse_qoo10_pdf(pdf_path)
    elif pdf_type == 'ebay':
        return parse_ebay_lincos_pdf(pdf_path)
    else:
        print(f"  ⚠️  알 수 없는 PDF 형식: {Path(pdf_path).name}")
        return None
