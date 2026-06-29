"""
소포수령증 PDF 파싱 모듈
- 쇼피(Shopee): 두라로지스틱스, 텍스트 기반 PDF
- 라자다(Lazada): 용성종합물류, 텍스트 기반 PDF (요약)
- 큐텐(Qoo10): 국제로지스틱, 이미지 기반 PDF → OCR 자동 추출
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
}

CURRENCY_NAMES_KR = {
    'MYR': '말레이시아 링깃 (MYR)',
    'PHP': '필리핀 페소 (PHP)',
    'SGD': '싱가포르 달러 (SGD)',
    'THB': '태국 바트 (THB)',
    'TWD': '대만 달러 (TWD)',
    'VND': '베트남 동 (VND)',
    'JPY': '일본 엔 (JPY) (100)',
}

# 파일명에서 국가코드 감지
SHOPEE_FILE_PATTERNS = {
    '_MY_': 'MY', '_PH_': 'PH', '_SG_': 'SG',
    '_TH_': 'TH', '_TW_': 'TW', '_VN_': 'VN',
    '_BR_': 'BR', '_MX_': 'MX',
}


def detect_pdf_type(pdf_path: str) -> str:
    """파일명으로 소포수령증 종류 판단 → 'shopee' | 'lazada' | 'qoo10' | 'unknown'"""
    name = Path(pdf_path).name
    lower = name.lower()
    if '라자다' in name or 'lazada' in lower:
        return 'lazada'
    if '큐텐' in name or 'qoo10' in lower:
        return 'qoo10'
    # 쇼피: 업체명과 무관하게 파일명의 국가코드 패턴(_MY_, _TW_ 등)으로 인식
    if re.search(r'_(MY|PH|SG|TH|TW|VN|BR|MX|JP)_', name):
        return 'shopee'
    # 기존 호환: 유엠 키워드
    if '유엠(UM)_' in name or '유엠_' in name:
        return 'shopee'
    return 'unknown'


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


def parse_qoo10_pdf(pdf_path: str) -> Optional[dict]:
    """
    큐텐재팬 소포수령증 파싱
    이미지 기반 PDF → OCR로 자동 추출 시도
    OCR 실패 시 None 반환 (수동 입력 필요)
    """
    # 먼저 pdfplumber로 텍스트 추출 시도
    all_text = ''
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                all_text += t + '\n'

    # 텍스트 없으면 OCR 시도
    if not all_text.strip():
        print("  🔍 큐텐 PDF 이미지 기반 → OCR 시도...")
        all_text = _ocr_pdf(pdf_path)

    if not all_text.strip():
        print("  ⚠️  큐텐 PDF 텍스트 추출 실패 — 수동 입력 필요")
        return None

    # ── 거래기간 파싱 ──
    # 패턴: "2025/12/01 ~ 2025/12/31" 또는 "2025-12-01 ~ 2025-12-31"
    period_match = re.search(
        r'(\d{4})[/.\-](\d{2})[/.\-](\d{2})\s*[~–\-]\s*(\d{4})[/.\-](\d{2})[/.\-](\d{2})',
        all_text
    )
    period_start = ''
    period_end   = ''
    if period_match:
        period_start = f"{period_match.group(1)}-{period_match.group(2)}-{period_match.group(3)}"
        period_end   = f"{period_match.group(4)}-{period_match.group(5)}-{period_match.group(6)}"

    # ── 작성일자 ──
    # 패턴: "2026-01-05 16:17:50"
    write_match = re.search(r'(\d{4})[-/.](\d{2})[-/.](\d{2})\s+\d{2}:\d{2}:\d{2}', all_text)
    write_date = ''
    if write_match:
        write_date = f"{write_match.group(1)}-{write_match.group(2)}-{write_match.group(3)}"

    # ── 발송수량 + 금액 ──
    # OCR 결과 예시: "| S7| HHS Sal 386 2 3,802,685"
    #   386 = 발송수량, 2 = '건'의 OCR 오인식, 3,802,685 = JPY 금액
    qty    = 0
    amount = 0.0

    # 패턴 1: "N (단일문자) large_amount" — 가장 일반적인 OCR 오인식 형태
    # 예: "386 2 3,802,685" 또는 "386 H 3,802,685"
    m1 = re.search(r'\b(\d{2,4})\s+[2H건a-zA-Z]\s+([\d,]{5,})\b', all_text)
    if m1:
        qty    = int(m1.group(1).replace(',', ''))
        amount = float(m1.group(2).replace(',', ''))

    # 패턴 2: "N건 amount" (정상 OCR된 경우)
    if qty == 0:
        m2 = re.search(r'([\d,]+)\s*건\s+([\d,]+)', all_text)
        if m2:
            qty    = int(m2.group(1).replace(',', ''))
            amount = float(m2.group(2).replace(',', ''))

    # 패턴 3: | N | amount | 파이프로 구분된 표 형식
    if qty == 0:
        m3 = re.search(r'[\|｜]\s*([\d,]{2,4})\s*[\|｜]\s*([\d,]{5,})', all_text)
        if m3:
            qty    = int(m3.group(1).replace(',', ''))
            amount = float(m3.group(2).replace(',', ''))

    # 패턴 4: 큰 숫자(6자리+) 앞의 3~4자리 숫자 = qty, 큰 숫자 = amount
    if qty == 0:
        m4 = re.search(r'\b(\d{3,4})\s+\d\s+([\d,]{5,})', all_text)
        if m4:
            qty    = int(m4.group(1).replace(',', ''))
            amount = float(m4.group(2).replace(',', ''))

    # ── 추적번호 ──
    tracking_match = re.search(r'[Kk]\d{13,}', all_text)
    tracking_no = tracking_match.group(0) if tracking_match else ''

    # 주요 데이터 추출 성공 여부
    if qty == 0 and amount == 0.0:
        print(f"  ⚠️  큐텐 OCR 성공했으나 수량/금액 파싱 실패 — 수동 입력 필요")
        print(f"      OCR 텍스트 (일부): {all_text[:300]}")
        return None

    print(f"  ✅ 큐텐 OCR 성공 — 기간:{period_start}~{period_end} 수량:{qty}건 JPY:{amount:,.0f}")

    return {
        'type':         'qoo10',
        'carrier':      '국제로지스틱',
        'destination':  'JP',
        'currency':     'JPY',
        'period_start': period_start,
        'period_end':   period_end,
        'write_date':   write_date,
        'qty':          qty,
        'amount':       float(amount),
        'tracking_no':  tracking_no,
    }


# ─────────────────────────────────────────────────────────────────
# 통합 파싱 함수
# ─────────────────────────────────────────────────────────────────

def parse_pdf(pdf_path: str) -> Optional[dict]:
    """PDF 종류 자동 감지 후 파싱"""
    pdf_type = detect_pdf_type(pdf_path)
    print(f"  파싱 중: {Path(pdf_path).name} → [{pdf_type}]")

    if pdf_type == 'shopee':
        return parse_shopee_pdf(pdf_path)
    elif pdf_type == 'lazada':
        return parse_lazada_pdf(pdf_path)
    elif pdf_type == 'qoo10':
        return parse_qoo10_pdf(pdf_path)
    else:
        print(f"  ⚠️  알 수 없는 PDF 형식: {Path(pdf_path).name}")
        return None
