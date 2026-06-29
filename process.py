#!/usr/bin/env python3
"""
소포수령증 자동 처리 메인 스크립트

사용법:
  python process.py                  # input/ 폴더의 PDF를 처리
  python process.py --year 2025 --month 12   # 연월 지정
  python process.py --pdf path/to/file.pdf   # 특정 PDF만 처리
"""

import sys
import os
import argparse
import shutil
import yaml
import re
from pathlib import Path
from datetime import datetime

# 현재 디렉토리를 경로에 추가
sys.path.insert(0, str(Path(__file__).parent))

from modules.pdf_parser    import parse_pdf, detect_pdf_type
from modules.exchange_rate import fetch_all_currencies
from modules.excel_writer  import generate_excel

BASE_DIR    = Path(__file__).parent
INPUT_DIR   = BASE_DIR / 'input'
OUTPUT_DIR  = BASE_DIR / 'output'
LOGS_DIR    = BASE_DIR / 'logs'
CONFIG_FILE = BASE_DIR / 'config.yaml'

CURRENCIES = ['MYR', 'PHP', 'SGD', 'THB', 'TWD', 'VND', 'JPY']


# ── 설정 파일 로드 ──────────────────────────────────────────────

def load_config() -> dict:
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f) or {}
    return {}


# ── PDF 파일 수집 ───────────────────────────────────────────────

def collect_pdfs(input_dir: Path) -> list:
    pdfs = list(input_dir.glob('*.pdf')) + list(input_dir.glob('*.PDF'))
    # 이미 처리된 폴더 제외
    return [p for p in pdfs if 'processed' not in str(p)]


# ── 연월 추정 (파일명에서) ────────────────────────────────────────

def infer_year_month(pdf_paths: list) -> tuple:
    """파일명에서 연월 자동 감지 (예: 20251201 → 2025, 12)"""
    for path in pdf_paths:
        m = re.search(r'(\d{4})(\d{2})\d{2}', path.name)
        if m:
            return int(m.group(1)), int(m.group(2))
    # 없으면 현재 월 - 1
    today = datetime.today()
    month = today.month - 1 or 12
    year  = today.year if today.month > 1 else today.year - 1
    return year, month


# ── 메인 처리 ───────────────────────────────────────────────────

def process(year: int = None, month: int = None, pdf_paths: list = None):
    config = load_config()
    print(f'\n{"="*60}')
    print(f'  소포수령증 자동화 처리 시작 — {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    print(f'{"="*60}')

    # PDF 파일 목록
    if pdf_paths is None:
        pdf_paths = collect_pdfs(INPUT_DIR)

    if not pdf_paths:
        print('  ⚠️  처리할 PDF 파일이 없습니다. input/ 폴더에 소포수령증 PDF를 넣어주세요.')
        return

    print(f'\n📄 처리할 PDF: {len(pdf_paths)}개')
    for p in pdf_paths:
        print(f'   - {p.name}')

    # 연월 결정
    if year is None or month is None:
        year, month = infer_year_month(pdf_paths)
        print(f'\n📅 처리 연월: {year}년 {month:02d}월 (자동 감지)')
    else:
        print(f'\n📅 처리 연월: {year}년 {month:02d}월')

    # ── PDF 파싱 ──────────────────────────────────────────
    print(f'\n[1/4] PDF 파싱 중...')
    shopee_results = []
    lazada_result  = None
    qoo10_result   = None

    for pdf_path in pdf_paths:
        result = parse_pdf(str(pdf_path))
        if result is None:
            continue
        ptype = result.get('type')
        if ptype == 'shopee':
            shopee_results.append(result)
            print(f'     쇼피 {result.get("currency","?")} — '
                  f'{result.get("total_qty",0)}건 / '
                  f'{result.get("total_amount",0):,.2f} {result.get("currency","")}')
        elif ptype == 'lazada':
            lazada_result = result
            print(f'     라자다 — {len(result.get("items",[]))}개 국가')
        elif ptype == 'qoo10':
            qoo10_result = result
            if qoo10_result:
                print(f'     큐텐 JPY — {qoo10_result.get("qty",0)}건 / {qoo10_result.get("amount",0):,} JPY')
            else:
                print('     큐텐 — 이미지 PDF, 수동 입력 필요')

    # 큐텐 수동 입력 반영 (config.yaml)
    qoo10_manual = config.get('qoo10', {})
    if not qoo10_result and qoo10_manual.get('jpy_amount'):
        qoo10_result = {
            'type':        'qoo10',
            'carrier':     '국제로지스틱',
            'destination': 'JP',
            'currency':    'JPY',
            'qty':         qoo10_manual.get('qty', 0),
            'amount':      float(qoo10_manual.get('jpy_amount', 0)),
            'tracking_no': qoo10_manual.get('tracking_no', ''),
        }
        print(f'     큐텐 (수동입력) — {qoo10_result["qty"]}건 / {qoo10_result["amount"]:,} JPY')
    elif not qoo10_result:
        print('     ⚠️  큐텐 데이터 없음 (config.yaml에 수동 입력하거나 PDF 추가)')

    # ── 환율 수집 ──────────────────────────────────────────
    print(f'\n[2/4] SMBS 환율 수집 중...')

    # config.yaml에 수동 환율이 있으면 사용
    manual_rates = config.get('manual_rates', {})
    rates = {}

    if manual_rates.get('use_manual', False):
        print('  📋 수동 환율 사용 (config.yaml)')
        rates = _build_manual_rates(manual_rates, year, month)
    else:
        rates = fetch_all_currencies(year, month, CURRENCIES)
        # 실패한 통화는 수동 데이터로 보완
        for cur in CURRENCIES:
            if rates.get(cur) is None and manual_rates.get(cur):
                rates[cur] = _build_manual_rates(manual_rates, year, month).get(cur)
                print(f'  📋 {cur} — 수동 환율 사용')

    # ── 엑셀 생성 ──────────────────────────────────────────
    print(f'\n[3/4] 엑셀 생성 중...')
    output_filename = f'매출집계_{year}{month:02d}.xlsx'
    output_path = OUTPUT_DIR / output_filename

    generate_excel(
        shopee_results=shopee_results,
        lazada_result=lazada_result,
        qoo10_result=qoo10_result,
        rates=rates,
        output_path=str(output_path),
        year=year,
        month=month,
    )

    # ── 처리 완료 파일 이동 ────────────────────────────────
    print(f'\n[4/4] 처리 완료 파일 이동 중...')
    processed_dir = INPUT_DIR / 'processed' / f'{year}{month:02d}'
    processed_dir.mkdir(parents=True, exist_ok=True)

    for pdf_path in pdf_paths:
        dest = processed_dir / pdf_path.name
        shutil.move(str(pdf_path), str(dest))
        print(f'  → {pdf_path.name}  →  processed/{year}{month:02d}/')

    print(f'\n{"="*60}')
    print(f'  ✅ 완료! 결과 파일: output/{output_filename}')
    print(f'{"="*60}\n')

    return str(output_path)


def _build_manual_rates(manual_rates: dict, year: int, month: int) -> dict:
    """config.yaml 수동 환율 → rates 딕셔너리 형식으로 변환"""
    import calendar
    last_day = calendar.monthrange(year, month)[1]
    period = f'{year}년 {month:02d}월 01일 ~ {year}년 {month:02d}월 {last_day:02d}일'

    SMBS_NAMES = {
        'MYR': '말레이시아 링깃 (MYR)', 'PHP': '필리핀 페소 (PHP)',
        'SGD': '싱가포르 달러 (SGD)', 'THB': '태국 바트 (THB)',
        'TWD': '대만 달러 (TWD)', 'VND': '베트남 동 (VND)',
        'JPY': '일본 엔 (JPY) (100)', 'BRL': '브라질 헤알 (BRL)',
    }

    result = {}
    for cur, avg in manual_rates.items():
        if cur in ('use_manual',):
            continue
        if isinstance(avg, (int, float)):
            result[cur] = {
                'period':        period,
                'currency':      cur,
                'currency_name': SMBS_NAMES.get(cur, cur),
                'average':       float(avg),
                'min': float(avg), 'min_date': '',
                'max': float(avg), 'max_date': '',
                'range': 0.0, 'cross_rate': 0.0,
                'daily': [],
            }
    return result


# ── CLI 진입점 ──────────────────────────────────────────────────

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='소포수령증 자동 처리')
    parser.add_argument('--year',  type=int, help='처리 연도 (예: 2025)')
    parser.add_argument('--month', type=int, help='처리 월 (예: 12)')
    parser.add_argument('--pdf',   type=str, help='특정 PDF 파일 경로 (복수 지정 시 쉼표 구분)')
    args = parser.parse_args()

    pdf_paths = None
    if args.pdf:
        pdf_paths = [Path(p.strip()) for p in args.pdf.split(',')]

    process(
        year=args.year,
        month=args.month,
        pdf_paths=pdf_paths,
    )
