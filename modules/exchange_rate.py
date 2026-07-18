# -*- coding: utf-8 -*-
"""
서울외국환중개(SMBS) 기간별 매매기준율 수집 + 로컬/서버 캐시 모듈.

- 기존 GitHub excel_writer.py가 요구하는 인터페이스를 유지합니다.
  - fetch_all_currencies(year, month, currencies)
  - get_rate_for_date(rate_data, date_str)
  - avg_rate_for_period(rate_data, start, end)
- v29 방식처럼 SMBS에서 직접 수집하고 data/exchange_rate_cache.csv에 저장합니다.
"""

from __future__ import annotations

import os
import re
import time
import shutil
import bisect
import json
from io import StringIO
from pathlib import Path
from datetime import datetime, date
from urllib.parse import urlencode
from typing import Callable, Iterable, Optional

import pandas as pd
import requests
from bs4 import BeautifulSoup

SMBS_STD_RATE_URL = "http://www.smbs.biz/ExRate/StdExRate.jsp"
SMBS_STD_RATE_PRINT_URL = "http://www.smbs.biz/ExRate/StdExRate_print.jsp"
SMBS_MON_AVG_RATE_URL = "http://www.smbs.biz/ExRate/MonAvgStdExRate.jsp"
SMBS_MON_AVG_RATE_PRINT_URL = "http://www.smbs.biz/ExRate/MonAvgStdExRate_print.jsp"
RATE_CACHE_FILE = Path(__file__).resolve().parents[1] / "data" / "exchange_rate_cache.csv"
MONTHLY_RATE_CACHE_FILE = Path(__file__).resolve().parents[1] / "data" / "monthly_exchange_rate_cache.csv"
MONTHLY_RATE_CACHE_SOURCE = "SMBS_MON_AVG_OFFICIAL"
_FIXED_RATE_JSON_CANDIDATES = [
    Path(__file__).resolve().parents[1] / "data" / "fixed_rates_2025.json",
    Path(__file__).resolve().parents[1] / "fixed_rates_2025.json",
    Path.cwd() / "data" / "fixed_rates_2025.json",
    Path.cwd() / "fixed_rates_2025.json",
]
RATE_LOOKBACK_DAYS = 7

# 서울외국환중개는 아래 통화를 100통화 단위로 고시합니다.
# 계산에는 1통화 단위 원화환율을 사용하되, 환율(통화) 시트에는
# 서울외국환중개 원문과 동일한 100통화 단위 값을 표시합니다.
SMBS_SOURCE_UNIT_DIVISOR = {
    "JPY": 100.0,
    "IDR": 100.0,
    "VND": 100.0,
}

CURRENCY_NAMES = {
    "MYR": "말레이시아 링깃 (MYR)",
    "PHP": "필리핀 페소 (PHP)",
    "SGD": "싱가포르 달러 (SGD)",
    "THB": "태국 바트 (THB)",
    "TWD": "대만 달러 (TWD)",
    "VND": "베트남 동 (VND)",
    "JPY": "일본 엔 (JPY)",
    "IDR": "인도네시아 루피아 (IDR)",
    "BRL": "브라질 헤알 (BRL)",
    "MXN": "멕시코 페소 (MXN)",
    "USD": "미국 달러 (USD)",
    "EUR": "유로 (EUR)",
    "GBP": "영국 파운드 (GBP)",
    "CAD": "캐나다 달러 (CAD)",
    "AUD": "호주 달러 (AUD)",
}

CURRENCY_KOREAN_KEYWORDS = {
    "USD": ["미국", "달러", "USD"],
    "EUR": ["유로", "EUR"],
    "GBP": ["영국", "파운드", "GBP"],
    "CAD": ["캐나다", "달러", "CAD"],
    "AUD": ["호주", "달러", "AUD"],
    "JPY": ["일본", "엔", "JPY"],
    "TWD": ["대만", "달러", "TWD"],
    "THB": ["태국", "밧", "바트", "THB"],
    "SGD": ["싱가포르", "SGD"],
    "MYR": ["말레이시아", "링깃", "MYR"],
    "PHP": ["필리핀", "페소", "PHP"],
    "VND": ["베트남", "동", "VND"],
    "IDR": ["인도네시아", "루피아", "IDR"],
    "MXN": ["멕시코", "페소", "MXN"],
    "BRL": ["브라질", "헤알", "BRL"],
}

_LOGGER: Optional[Callable[[str], None]] = None

def set_logger(logger: Optional[Callable[[str], None]] = None):
    global _LOGGER
    _LOGGER = logger


def _log(message: str):
    if _LOGGER:
        _LOGGER(message)
    else:
        print(message)


def to_number(value):
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace(",", "").replace("원", "").replace("KRW", "")
    if not s:
        return None
    try:
        return float(s)
    except Exception:
        return None


def normalize_smbs_rate(currency: str, value):
    """SMBS 고시단위를 1통화 단위 원화환율로 변환합니다.

    JPY/IDR/VND는 사이트에서 100통화 단위로 고시되므로 100으로 나눕니다.
    그 외 통화는 원본 값을 그대로 사용합니다.
    """
    rate = to_number(value)
    if rate is None:
        return None
    divisor = SMBS_SOURCE_UNIT_DIVISOR.get(str(currency or "").upper(), 1.0)
    return float(rate) / float(divisor)


def smbs_source_rate(currency: str, value):
    """1통화 단위 내부 환율을 서울외국환중개 표시단위로 되돌립니다."""
    rate = to_number(value)
    if rate is None:
        return None
    multiplier = SMBS_SOURCE_UNIT_DIVISOR.get(str(currency or "").upper(), 1.0)
    return float(rate) * float(multiplier)


def applied_rate_precision(currency: str) -> int:
    """100통화 단위 고시 통화는 1통화 환율을 소수점 넷째 자리까지 유지합니다."""
    return 4 if str(currency or "").upper() in SMBS_SOURCE_UNIT_DIVISOR else 2


def round_applied_rate(currency: str, value) -> float:
    rate = to_number(value)
    if rate is None:
        return 0.0
    return round(float(rate), applied_rate_precision(currency))


def _normalize_rate_column(df: pd.DataFrame, currency: str) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out = df.copy()
    out["rate"] = out["rate"].apply(lambda v: normalize_smbs_rate(currency, v))
    return out


def parse_date(value):
    if pd.isna(value):
        return pd.NaT
    if isinstance(value, pd.Timestamp):
        return pd.NaT if pd.isna(value) else value.normalize()
    if isinstance(value, datetime):
        return pd.to_datetime(value).normalize()
    if isinstance(value, date):
        return pd.to_datetime(value).normalize()
    if isinstance(value, (int, float)):
        if 30000 <= float(value) <= 60000:
            dt = pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
            return pd.NaT if pd.isna(dt) else dt.normalize()
        return pd.NaT
    s = str(value).strip()
    if not s:
        return pd.NaT
    if s.upper() in {"날짜", "DATE", "통화명", "환율", "매매기준율", "평균환율", "최저치", "최고치", "기록일", "CROSS RATE"}:
        return pd.NaT
    s = s.replace(".", "-").replace("/", "-")
    if re.fullmatch(r"\d{8}", s):
        dt = pd.to_datetime(s, format="%Y%m%d", errors="coerce")
    else:
        dt = pd.to_datetime(s, errors="coerce")
    return pd.NaT if pd.isna(dt) else dt.normalize()


def _business_days_between(start_date, end_date):
    start_date = pd.to_datetime(start_date).normalize()
    end_date = pd.to_datetime(end_date).normalize()
    if start_date > end_date:
        return []
    return [d for d in pd.date_range(start_date, end_date, freq="D") if d.weekday() < 5]


def _has_weekday(start_date, end_date):
    return bool(_business_days_between(start_date, end_date))


def _previous_weekday(dt):
    dt = pd.to_datetime(dt).normalize() - pd.Timedelta(days=1)
    while dt.weekday() >= 5:
        dt -= pd.Timedelta(days=1)
    return dt


def load_rate_cache(currency=None):
    path = Path(RATE_CACHE_FILE)
    if not path.exists():
        return pd.DataFrame(columns=["currency", "date", "rate", "fetched_at"])
    try:
        df = pd.read_csv(path, dtype={"currency": str})
        if df.empty:
            return pd.DataFrame(columns=["currency", "date", "rate", "fetched_at"])
        df["currency"] = df["currency"].astype(str).str.upper().str.strip()
        df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.normalize()
        df["rate"] = df["rate"].apply(to_number)
        df = df.dropna(subset=["currency", "date", "rate"])
        if currency:
            df = df[df["currency"] == str(currency).upper()].copy()
        return df.drop_duplicates(subset=["currency", "date"], keep="last").sort_values(["currency", "date"])
    except Exception as e:
        _log(f"[WARN] 환율 캐시를 읽지 못했습니다. 새로 수집합니다: {e}")
        return pd.DataFrame(columns=["currency", "date", "rate", "fetched_at"])


def save_rate_cache(currency, data):
    if data is None or data.empty:
        return
    path = Path(RATE_CACHE_FILE)
    path.parent.mkdir(parents=True, exist_ok=True)
    new_df = data[["date", "rate"]].copy()
    new_df["currency"] = str(currency).upper()
    new_df["date"] = pd.to_datetime(new_df["date"], errors="coerce").dt.normalize()
    new_df["rate"] = new_df["rate"].apply(to_number)
    new_df["fetched_at"] = datetime.now().isoformat(timespec="seconds")
    new_df = new_df.dropna(subset=["date", "rate"])[["currency", "date", "rate", "fetched_at"]]
    old_df = load_rate_cache()
    merged = new_df.copy() if old_df.empty else pd.concat([old_df, new_df], ignore_index=True)
    merged["currency"] = merged["currency"].astype(str).str.upper().str.strip()
    merged["date"] = pd.to_datetime(merged["date"], errors="coerce").dt.normalize()
    merged["rate"] = merged["rate"].apply(to_number)
    merged = merged.dropna(subset=["currency", "date", "rate"])
    merged = merged.drop_duplicates(subset=["currency", "date"], keep="last").sort_values(["currency", "date"])
    out = merged.copy()
    out["date"] = out["date"].dt.strftime("%Y-%m-%d")
    out.to_csv(path, index=False, encoding="utf-8-sig")


def row_mentions_currency(row_values, currency):
    text = " ".join(str(v) for v in row_values if not pd.isna(v)).upper()
    if currency.upper() in text:
        return True
    return any(str(kw).upper() in text for kw in CURRENCY_KOREAN_KEYWORDS.get(currency.upper(), []))


def is_currency_name_cell(value, currency):
    if value is None or pd.isna(value):
        return False
    s = str(value).strip().upper()
    if currency.upper() in s:
        return True
    return any(str(kw).upper() in s for kw in CURRENCY_KOREAN_KEYWORDS.get(currency.upper(), []))


def clean_smbs_rate_dataframe(df, currency):
    if df is None or df.empty:
        return pd.DataFrame(columns=["date", "rate"])
    out = df.copy()
    out["date"] = pd.to_datetime(out["date"]).dt.normalize()
    out["rate"] = out["rate"].apply(to_number)
    out = out.dropna(subset=["date", "rate"])
    if currency.upper() in {"VND", "JPY", "IDR"}:
        out = out[out["rate"].round(8) != 100]
    out = out[out["rate"] > 0]
    out = out.drop_duplicates(subset=["date"], keep="last").sort_values("date")
    return out[["date", "rate"]].copy()



def _norm_date_key(value) -> str:
    """환율 색인용 날짜 문자열(YYYY.MM.DD)로 정규화합니다."""
    dt = parse_date(value)
    if pd.isna(dt):
        return ""
    return pd.to_datetime(dt).strftime("%Y.%m.%d")


def load_fixed_rate_json(currency: str, start_date=None, end_date=None) -> pd.DataFrame:
    """기존 GitHub에 fixed_rates_2025.json이 있으면 사이트 접속 없이 사용합니다.
    파일이 없거나 JSON 형식이 아니면 빈 DataFrame을 반환합니다.
    """
    cur = str(currency).upper()
    for path in _FIXED_RATE_JSON_CANDIDATES:
        try:
            if not path.exists() or path.stat().st_size <= 2:
                continue
            with open(path, encoding="utf-8") as f:
                raw = json.load(f)
            rates = raw.get("rates", raw)
            daymap = rates.get(cur)
            if not isinstance(daymap, dict):
                continue
            rows = []
            for d, r in daymap.items():
                dt = parse_date(d)
                val = to_number(r)
                if pd.isna(dt) or val is None:
                    continue
                rows.append({"date": dt, "rate": val})
            if not rows:
                continue
            df = clean_smbs_rate_dataframe(pd.DataFrame(rows), cur)
            if start_date is not None and end_date is not None:
                sdt = pd.to_datetime(start_date).normalize()
                edt = pd.to_datetime(end_date).normalize()
                prev = df[df["date"] < sdt].sort_values("date").tail(1)
                data = df[(df["date"] >= sdt) & (df["date"] <= edt)].copy()
                if not prev.empty:
                    data = pd.concat([prev, data], ignore_index=True)
                if data.empty:
                    continue
                return fill_missing_dates(data[["date", "rate"]], sdt, edt)
            return df
        except Exception:
            continue
    return pd.DataFrame(columns=["date", "rate"])

def build_smbs_std_params(currency, start_date, end_date):
    start = pd.to_datetime(start_date)
    end = pd.to_datetime(end_date)
    return {
        "StrSch_sYear": start.strftime("%Y"),
        "StrSch_sMonth": start.strftime("%m"),
        "StrSch_sDay": start.strftime("%d"),
        "StrSch_eYear": end.strftime("%Y"),
        "StrSch_eMonth": end.strftime("%m"),
        "StrSch_eDay": end.strftime("%d"),
        "StrSchFull": start.strftime("%Y.%m.%d"),
        "StrSchFull2": end.strftime("%Y.%m.%d"),
        "quick_date": "",
        "tongwha_code": currency,
    }


def build_smbs_std_url(currency, start_date, end_date, base_url=SMBS_STD_RATE_URL):
    return f"{base_url}?{urlencode(build_smbs_std_params(currency, start_date, end_date))}"


def parse_rate_table_from_html(html, currency, debug_prefix=None):
    if not html:
        return pd.DataFrame(columns=["date", "rate"])

    try:
        tables = pd.read_html(StringIO(html))
    except Exception:
        tables = []

    candidates = []
    for table in tables:
        if table is None or table.empty:
            continue
        df = table.copy()
        df.columns = [str(c).strip() for c in df.columns]
        date_col = None
        rate_col = None
        for c in df.columns:
            name = str(c).replace(" ", "")
            if date_col is None and ("날짜" in name or "DATE" in name.upper()):
                date_col = c
            if rate_col is None and ("매매기준율" in name or "기준율" in name or "RATE" in name.upper()):
                rate_col = c
        if date_col is not None and rate_col is not None:
            temp = pd.DataFrame({"date": df[date_col].apply(parse_date), "rate": df[rate_col].apply(to_number)}).dropna(subset=["date", "rate"])
            temp = clean_smbs_rate_dataframe(temp, currency)
            if not temp.empty:
                candidates.append(temp)
                continue

        records = []
        for _, row in df.iterrows():
            vals = row.tolist()
            joined = " ".join(str(v) for v in vals if not pd.isna(v))
            if re.search(r"\([A-Z]{3}\)", joined) and not row_mentions_currency(vals, currency):
                continue
            row_date = None
            nums = []
            for v in vals:
                dt = parse_date(v)
                if row_date is None and not pd.isna(dt):
                    row_date = dt
                if is_currency_name_cell(v, currency):
                    continue
                num = to_number(v)
                if num is None:
                    continue
                if 30000 <= num <= 60000 or 1900 <= num <= 2100:
                    continue
                if currency.upper() in {"VND", "JPY", "IDR"} and abs(num - 100) < 1e-9:
                    continue
                nums.append(num)
            if row_date is not None:
                positive = [x for x in nums if x > 0]
                if positive:
                    records.append({"date": row_date, "rate": positive[0]})
        if records:
            temp = clean_smbs_rate_dataframe(pd.DataFrame(records), currency)
            if not temp.empty:
                candidates.append(temp)

    if not candidates:
        text = re.sub(r"<script.*?</script>", " ", html, flags=re.I | re.S)
        text = re.sub(r"<style.*?</style>", " ", text, flags=re.I | re.S)
        text = re.sub(r"<[^>]+>", " ", text).replace("&nbsp;", " ")
        units = [re.sub(r"\s+", " ", x).strip() for x in text.splitlines()]
        units.append(re.sub(r"\s+", " ", text))
        date_pat = re.compile(r"(20\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2})")
        num_pat = re.compile(r"[-+]?\d{1,3}(?:,\d{3})*(?:\.\d+)?|[-+]?\d+(?:\.\d+)?")
        records = []
        for unit in units:
            for dm in date_pat.finditer(unit):
                dt = parse_date(dm.group(1))
                if pd.isna(dt):
                    continue
                tail = unit[dm.end(): dm.end() + 180]
                nums = []
                for nm in num_pat.finditer(tail):
                    num = to_number(nm.group(0))
                    if num is None or num <= 0:
                        continue
                    if 30000 <= num <= 60000 or 1900 <= num <= 2100:
                        continue
                    if currency.upper() in {"VND", "JPY", "IDR"} and abs(num - 100) < 1e-9 and nm.start() < 100:
                        continue
                    nums.append(num)
                if nums:
                    records.append({"date": dt, "rate": nums[0]})
        if records:
            temp = clean_smbs_rate_dataframe(pd.DataFrame(records), currency)
            if not temp.empty:
                candidates.append(temp)

    if not candidates:
        if debug_prefix:
            try:
                Path(debug_prefix).with_suffix(".html").write_text(html, encoding="utf-8", errors="ignore")
            except Exception:
                pass
        return pd.DataFrame(columns=["date", "rate"])
    return max(candidates, key=len)[["date", "rate"]].copy()


def try_fetch_std_rates_by_requests(currency, start_date, end_date, timeout=10):
    session = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
        "Referer": SMBS_STD_RATE_URL,
        "Origin": "http://www.smbs.biz",
    }
    params = build_smbs_std_params(currency, start_date, end_date)
    start_compact = pd.to_datetime(start_date).strftime("%Y%m%d")
    end_compact = pd.to_datetime(end_date).strftime("%Y%m%d")
    legacy_params = {"yyyymmdd1": start_compact, "yyyymmdd2": end_compact, "curCd": currency}
    legacy_post = dict(legacy_params)
    legacy_post["gubun"] = "1"
    try:
        session.get(SMBS_STD_RATE_URL, headers=headers, timeout=min(timeout, 8))
    except Exception:
        pass
    attempts = [
        ("legacy_get", "GET", SMBS_STD_RATE_URL, legacy_params),
        ("legacy_post", "POST", SMBS_STD_RATE_URL, legacy_post),
        ("print_get", "GET", build_smbs_std_url(currency, start_date, end_date, SMBS_STD_RATE_PRINT_URL), None),
        ("direct_get", "GET", build_smbs_std_url(currency, start_date, end_date, SMBS_STD_RATE_URL), None),
        ("get_params", "GET", SMBS_STD_RATE_URL, params),
        ("post", "POST", SMBS_STD_RATE_URL, params),
    ]
    for _, method, url, payload in attempts:
        try:
            if method == "POST":
                h = dict(headers)
                h["Content-Type"] = "application/x-www-form-urlencoded"
                resp = session.post(url, data=payload, headers=h, timeout=timeout)
            else:
                resp = session.get(url, params=payload, headers=headers, timeout=timeout)
            resp.raise_for_status()
            enc = resp.apparent_encoding or resp.encoding or "cp949"
            if str(enc).lower() in ["iso-8859-1", "ascii"]:
                enc = "cp949"
            resp.encoding = enc
            data = parse_rate_table_from_html(resp.text, currency)
            if not data.empty:
                sdt = pd.to_datetime(start_date).normalize()
                edt = pd.to_datetime(end_date).normalize()
                data = data[(data["date"] >= sdt) & (data["date"] <= edt)].copy()
                if not data.empty:
                    return data
        except Exception:
            continue
    return pd.DataFrame(columns=["date", "rate"])


def fetch_std_rates_by_selenium(currency, start_date, end_date, headless=True, timeout=60):
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.chrome.service import Service
        from selenium.webdriver.support.ui import Select
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.common.exceptions import UnexpectedAlertPresentException
    except ImportError as e:
        raise RuntimeError("Selenium이 설치되어 있지 않습니다. pip install selenium 실행 후 다시 실행하세요.") from e

    def make_options(mode):
        opts = Options()
        if headless:
            opts.add_argument("--headless=new" if mode == "new" else "--headless")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--disable-extensions")
        opts.add_argument("--disable-popup-blocking")
        opts.add_argument("--window-size=1500,1100")
        opts.add_argument("--lang=ko-KR")
        for candidate in [os.environ.get("CHROME_BINARY"), os.environ.get("CHROME_BIN"), shutil.which("chromium"), shutil.which("chromium-browser"), shutil.which("google-chrome"), "/usr/bin/chromium", "/usr/bin/chromium-browser", "/usr/bin/google-chrome"]:
            if candidate and Path(candidate).exists():
                opts.binary_location = str(candidate)
                break
        return opts

    def create_driver():
        service = None
        for candidate in [os.environ.get("CHROMEDRIVER"), os.environ.get("CHROMEDRIVER_PATH"), shutil.which("chromedriver"), "/usr/bin/chromedriver", "/usr/lib/chromium/chromedriver"]:
            if candidate and Path(candidate).exists():
                service = Service(str(candidate))
                break
        errors = []
        for mode in (["new", "old"] if headless else ["visible"]):
            try:
                opts = make_options(mode)
                return webdriver.Chrome(service=service, options=opts) if service else webdriver.Chrome(options=opts)
            except Exception as e:
                errors.append(f"{mode}: {e}")
        raise RuntimeError("Chrome/Chromium 실행 실패. packages.txt에 chromium, chromium-driver가 필요할 수 있습니다. " + " | ".join(errors))

    start_dot = pd.to_datetime(start_date).strftime("%Y.%m.%d")
    end_dot = pd.to_datetime(end_date).strftime("%Y.%m.%d")
    keyword_map = {
        "MYR": ["MYR", "말레이", "링깃"], "PHP": ["PHP", "필리핀", "페소"],
        "SGD": ["SGD", "싱가", "싱가포르"], "THB": ["THB", "태국", "바트", "밧"],
        "TWD": ["TWD", "대만"], "VND": ["VND", "베트남", "동"],
        "BRL": ["BRL", "브라질"], "MXN": ["MXN", "멕시코"], "JPY": ["JPY", "일본", "엔"],
    }
    keywords = [x.upper() for x in keyword_map.get(currency.upper(), [currency.upper()])]
    driver = create_driver()
    try:
        wait = WebDriverWait(driver, timeout)
        driver.get(build_smbs_std_url(currency, start_date, end_date))
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
        time.sleep(3)
        data = parse_rate_table_from_html(driver.page_source, currency)
        if not data.empty:
            sdt = pd.to_datetime(start_date).normalize()
            edt = pd.to_datetime(end_date).normalize()
            data = data[(data["date"] >= sdt) & (data["date"] <= edt)].copy()
            if not data.empty:
                return data

        # 직접 URL 실패 시 폼 조작
        for sel in driver.find_elements(By.TAG_NAME, "select"):
            try:
                s = Select(sel)
                for opt in s.options:
                    txt = (opt.text or "").upper()
                    val = (opt.get_attribute("value") or "").upper()
                    if any(k in txt or k in val for k in keywords):
                        s.select_by_visible_text(opt.text)
                        break
            except Exception:
                pass
        try:
            result = driver.execute_script(
                """
                const s = arguments[0], e = arguments[1];
                const startFull = document.querySelector('#startDate, input[name="StrSchFull"]');
                const endFull = document.querySelector('#endDate, input[name="StrSchFull2"]');
                if (startFull && endFull) {
                  startFull.value = s; endFull.value = e;
                  startFull.setAttribute('value', s); endFull.setAttribute('value', e);
                  for (const el of [startFull, endFull]) {
                    el.dispatchEvent(new Event('input', {bubbles:true}));
                    el.dispatchEvent(new Event('change', {bubbles:true}));
                    el.dispatchEvent(new Event('blur', {bubbles:true}));
                  }
                  const p1 = s.split('.'), p2 = e.split('.');
                  const setv = (name, val) => { const el = document.querySelector(`input[name="${name}"]`); if (el) { el.value = val; el.setAttribute('value', val); } };
                  setv('StrSch_sYear', p1[0]); setv('StrSch_sMonth', p1[1]); setv('StrSch_sDay', p1[2]);
                  setv('StrSch_eYear', p2[0]); setv('StrSch_eMonth', p2[1]); setv('StrSch_eDay', p2[2]);
                  return 2;
                }
                return 0;
                """,
                start_dot, end_dot,
            )
        except UnexpectedAlertPresentException:
            try:
                driver.switch_to.alert.accept()
            except Exception:
                pass
            result = 0
        if int(result or 0) >= 2:
            driver.execute_script("if (typeof doSearch === 'function') { doSearch('frm_SearchDate'); } else { const f=document.forms['frm_SearchDate']; if(f) f.submit(); }")
            time.sleep(5)
            data = parse_rate_table_from_html(driver.page_source, currency)
            if not data.empty:
                return data
        return pd.DataFrame(columns=["date", "rate"])
    finally:
        driver.quit()


def fetch_smbs_period_rates(currency, start_date, end_date):
    start_date = pd.to_datetime(start_date).normalize()
    end_date = pd.to_datetime(end_date).normalize()
    if not _has_weekday(start_date, end_date):
        return pd.DataFrame(columns=["date", "rate"])
    data = try_fetch_std_rates_by_requests(currency, start_date, end_date)
    if not data.empty:
        return data
    return fetch_std_rates_by_selenium(currency, start_date, end_date, headless=True)


def _has_full_daily_cache(cache_df, start_date, end_date):
    """요청 기간의 모든 달력 날짜가 캐시에 있는지 확인합니다.
    주말/공휴일도 직전 환율로 채워 캐시에 저장해두면 다음 실행 때 재조회하지 않습니다.
    """
    if cache_df is None or cache_df.empty:
        return False
    start_date = pd.to_datetime(start_date).normalize()
    end_date = pd.to_datetime(end_date).normalize()
    days = set(pd.date_range(start_date, end_date, freq="D"))
    have = set(pd.to_datetime(cache_df["date"], errors="coerce").dropna().dt.normalize())
    return days.issubset(have)


def _cache_requested_daily_range(currency, start_date, end_date):
    """캐시에 있는 원시 영업일 환율을 이용해 요청 기간 전체를 일자별로 채워 저장합니다.
    이렇게 해야 2025-12-13~2025-12-14 같은 주말 시작 구간을 매번 다시 조회하지 않습니다.
    """
    start_date = pd.to_datetime(start_date).normalize()
    end_date = pd.to_datetime(end_date).normalize()
    cache = load_rate_cache(currency)
    if cache.empty:
        return
    data = cache[(cache["date"] >= start_date) & (cache["date"] <= end_date)].copy()
    prev = cache[cache["date"] < start_date].sort_values("date").tail(1)
    if not prev.empty:
        data = pd.concat([prev, data], ignore_index=True)
    if data.empty:
        return
    filled = fill_missing_dates(data[["date", "rate"]], start_date, end_date)
    save_rate_cache(currency, filled)


def get_cached_or_fetch_smbs_period_rates(currency, start_date, end_date, quiet=False):
    currency = str(currency).upper()
    start_date = pd.to_datetime(start_date).normalize()
    end_date = pd.to_datetime(end_date).normalize()
    cache = load_rate_cache(currency)

    if _has_full_daily_cache(cache, start_date, end_date):
        if not quiet:
            _log(f"  - {currency}: 저장된 환율 사용")
        data = cache[(cache["date"] >= start_date) & (cache["date"] <= end_date)].copy()
        return data[["date", "rate"]].drop_duplicates(subset=["date"], keep="last").sort_values("date")

    fixed = load_fixed_rate_json(currency, start_date, end_date)
    if fixed is not None and not fixed.empty and _has_full_daily_cache(fixed.assign(currency=currency), start_date, end_date):
        if not quiet:
            _log(f"  - {currency}: 저장된 고정환율 사용")
        save_rate_cache(currency, fixed)
        return fixed[["date", "rate"]].drop_duplicates(subset=["date"], keep="last").sort_values("date")

    segments = []
    if cache.empty:
        segments.append((start_date, end_date))
    else:
        min_cached = cache["date"].min().normalize()
        max_cached = cache["date"].max().normalize()
        if min_cached > start_date:
            segments.append((start_date, min_cached - pd.Timedelta(days=1)))
        if max_cached < end_date:
            segments.append((max_cached + pd.Timedelta(days=1), end_date))

        # 캐시 범위 안에 있지만 달력 일자가 비어 있는 경우는 기존 데이터로 먼저 채워 저장합니다.
        # 그래도 양 끝 범위가 부족한 경우에만 사이트 조회를 수행합니다.
        if not segments:
            if not quiet:
                _log(f"  - {currency}: 저장된 환율 보정")
            _cache_requested_daily_range(currency, start_date, end_date)

    if segments and not quiet:
        _log(f"  - {currency}: 부족한 환율 수집")

    for seg_start, seg_end in segments:
        if seg_start > seg_end:
            continue
        query_start, query_end = seg_start, seg_end
        if not _has_weekday(seg_start, seg_end):
            query_start = _previous_weekday(seg_start)
            query_end = query_start
        fetched = fetch_smbs_period_rates(currency, query_start, query_end)
        save_rate_cache(currency, fetched)

    # 이번 요청 기간 전체를 직전 영업일 환율로 채워 캐시에 다시 저장합니다.
    # 다음 실행부터는 같은 기간을 사이트에 다시 조회하지 않습니다.
    _cache_requested_daily_range(currency, start_date, end_date)

    cache = load_rate_cache(currency)
    data = cache[(cache["date"] >= start_date) & (cache["date"] <= end_date)].copy()
    prev = cache[cache["date"] < start_date].sort_values("date").tail(1)
    if not prev.empty:
        data = pd.concat([prev, data], ignore_index=True)
    return data[["date", "rate"]].drop_duplicates(subset=["date"], keep="last").sort_values("date")


def fill_missing_dates(rate_df, start_date=None, end_date=None):
    rate_df = rate_df.copy()
    rate_df["date"] = pd.to_datetime(rate_df["date"]).dt.normalize()
    rate_df = rate_df.drop_duplicates(subset=["date"], keep="last").sort_values("date")
    start_date = pd.to_datetime(start_date or rate_df["date"].min()).normalize()
    end_date = pd.to_datetime(end_date or rate_df["date"].max()).normalize()
    reindex_start = min(start_date, rate_df["date"].min()) if not rate_df.empty else start_date
    idx = pd.date_range(start=reindex_start, end=end_date, freq="D")
    filled = rate_df.set_index("date").reindex(idx)
    filled["rate"] = filled["rate"].ffill().bfill()
    filled = filled.reset_index().rename(columns={"index": "date"})
    filled = filled[filled["date"] >= start_date].reset_index(drop=True)
    return filled


def _to_rate_entry(currency, raw_df, start_date, end_date, display_start=None, display_end=None):
    fetch_start = pd.to_datetime(start_date).normalize()
    fetch_end = pd.to_datetime(end_date).normalize()
    display_start = pd.to_datetime(display_start if display_start is not None else fetch_start).normalize()
    display_end = pd.to_datetime(display_end if display_end is not None else fetch_end).normalize()
    if display_end < display_start:
        display_start, display_end = display_end, display_start

    empty = {
        "period": f"{display_start.strftime('%Y.%m.%d')} ~ {display_end.strftime('%Y.%m.%d')}",
        "display_start": display_start.strftime("%Y.%m.%d"),
        "display_end": display_end.strftime("%Y.%m.%d"),
        "currency": currency,
        "currency_name": CURRENCY_NAMES.get(currency, currency),
        "average": 0.0,
        "min": 0.0,
        "min_date": "",
        "max": 0.0,
        "max_date": "",
        "range": 0.0,
        "cross_rate": 0.0,
        "daily": [],
    }
    if raw_df is None or raw_df.empty:
        return empty

    # 내부 계산용으로는 직전 영업일을 포함한 수집기간 전체를 보관합니다.
    filled = fill_missing_dates(raw_df, fetch_start, fetch_end)
    filled = _normalize_rate_column(filled, currency)
    stats = filled[(filled["date"] >= display_start) & (filled["date"] <= display_end)].copy()
    if stats.empty:
        stats = filled.copy()

    vals = [float(x) for x in stats["rate"].dropna().tolist()]
    min_idx = stats["rate"].idxmin() if vals else None
    max_idx = stats["rate"].idxmax() if vals else None
    daily = []
    prev = None
    for _, row in filled.iterrows():
        rate = float(row["rate"])
        change = 0.0 if prev is None else round(rate - prev, 6)
        daily.append({
            "date": pd.to_datetime(row["date"]).strftime("%Y.%m.%d"),
            "rate": rate,
            "change": change,
            "cross": 0,
        })
        prev = rate

    return {
        **empty,
        "average": round_applied_rate(currency, sum(vals) / len(vals)) if vals else 0.0,
        "min": round_applied_rate(currency, min(vals)) if vals else 0.0,
        "min_date": pd.to_datetime(stats.loc[min_idx, "date"]).strftime("%Y.%m.%d") if min_idx is not None else "",
        "max": round_applied_rate(currency, max(vals)) if vals else 0.0,
        "max_date": pd.to_datetime(stats.loc[max_idx, "date"]).strftime("%Y.%m.%d") if max_idx is not None else "",
        "range": round_applied_rate(currency, max(vals) - min(vals)) if vals else 0.0,
        "daily": daily,
    }


def fetch_all_currencies_for_period(start_date, end_date, currencies: Iterable[str], logger: Optional[Callable[[str], None]] = None, display_start=None, display_end=None):
    previous_logger = _LOGGER
    set_logger(logger or previous_logger)
    try:
        start_date = pd.to_datetime(start_date).normalize()
        end_date = pd.to_datetime(end_date).normalize()
        out = {}
        used = [str(c).upper() for c in currencies if c]
        if used:
            _log(f"💱 환율 수집 중... ({', '.join(used)})")
        for cur in used:
            raw = get_cached_or_fetch_smbs_period_rates(cur, start_date, end_date)
            out[cur] = _to_rate_entry(cur, raw, start_date, end_date, display_start=display_start, display_end=display_end)
        if used:
            _log("✅ 환율 수집 완료")
        return out
    finally:
        set_logger(previous_logger)


def fetch_all_currencies(year: int, month: int, currencies: Iterable[str]):
    display_start = pd.Timestamp(year=year, month=month, day=1)
    last = display_start + pd.offsets.MonthEnd(0)
    first = display_start - pd.Timedelta(days=RATE_LOOKBACK_DAYS)
    return fetch_all_currencies_for_period(
        first, last, currencies, display_start=display_start, display_end=last
    )



def _daily_frame(rate_data):
    if not rate_data or not rate_data.get("daily"):
        return pd.DataFrame(columns=["date", "rate"])
    cached = rate_data.get("_daily_frame_cache")
    if cached is not None:
        return cached
    df = pd.DataFrame(rate_data.get("daily", []))
    df["date"] = pd.to_datetime(df["date"].apply(parse_date), errors="coerce").dt.normalize()
    df["rate"] = df["rate"].apply(to_number)
    df = df.dropna(subset=["date", "rate"]).drop_duplicates(subset=["date"], keep="last").sort_values("date")
    df = df[["date", "rate"]]
    rate_data["_daily_frame_cache"] = df
    return df


def _date_index(rate_data):
    """대량 거래 환율 조회용 색인. 한 번 만들고 계속 재사용합니다."""
    if not rate_data or not rate_data.get("daily"):
        return {"exact": {}, "dates": [], "rates": []}
    idx = rate_data.get("_date_index")
    if idx is not None:
        return idx
    exact = {}
    for d in rate_data.get("daily", []):
        key = _norm_date_key(d.get("date"))
        rate = to_number(d.get("rate"))
        if key and rate is not None:
            exact[key] = float(rate)
    dates = sorted(exact.keys())
    idx = {"exact": exact, "dates": dates, "rates": [exact[x] for x in dates]}
    rate_data["_date_index"] = idx
    return idx


def get_rate_for_date(rate_data: dict, date_str: str) -> float:
    """특정 날짜 환율 반환. 없으면 가장 가까운 이전 영업일 환율 사용.
    기존 GitHub처럼 날짜 색인을 캐싱해서 대량 거래 계산 속도를 높입니다.
    """
    if not rate_data:
        return 0.0
    if not rate_data.get("daily"):
        return float(rate_data.get("average", 0.0) or 0.0)
    target = _norm_date_key(date_str)
    if not target:
        return float(rate_data.get("average", 0.0) or 0.0)
    idx = _date_index(rate_data)
    if target in idx["exact"]:
        return float(idx["exact"][target])
    dates = idx["dates"]
    rates = idx["rates"]
    if not dates:
        return float(rate_data.get("average", 0.0) or 0.0)
    pos = bisect.bisect_right(dates, target)
    if pos <= 0:
        return float(rates[0])
    return float(rates[pos - 1])


def avg_rate_for_period(rate_data: dict, start: str, end: str) -> float:
    """기간 평균환율. pandas DataFrame 재생성을 피하고 날짜 색인을 재사용합니다."""
    if not rate_data:
        return 0.0
    if not rate_data.get("daily"):
        return float(rate_data.get("average", 0.0) or 0.0)
    s = _norm_date_key(start)
    e = _norm_date_key(end)
    if not s and not e:
        return float(rate_data.get("average", 0.0) or 0.0)
    if not s:
        s = e
    if not e:
        e = s
    if e < s:
        s, e = e, s
    idx = _date_index(rate_data)
    dates, rates = idx["dates"], idx["rates"]
    if not dates:
        return float(rate_data.get("average", 0.0) or 0.0)
    left = bisect.bisect_left(dates, s)
    right = bisect.bisect_right(dates, e)
    vals = [r for r in rates[left:right] if r and r > 0]
    if vals:
        return round_applied_rate(rate_data.get("currency", ""), sum(vals) / len(vals))
    return get_rate_for_date(rate_data, e) or float(rate_data.get("average", 0.0) or 0.0)


# ────────────────────────────────────────────────────────────────
# 월평균 매매기준율 수집/캐시 — 이베이/린코스처럼 발행월만 있는 자료용
# ────────────────────────────────────────────────────────────────

def parse_month_key(value) -> str:
    """문자열/날짜를 YYYY-MM으로 정규화합니다."""
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, (pd.Timestamp, datetime, date)):
        dt = pd.to_datetime(value, errors="coerce")
        return "" if pd.isna(dt) else dt.strftime("%Y-%m")
    s = str(value).strip()
    # 2026-03, 2026.03, 2026년 03월, 202603
    m = re.search(r"(20\d{2})\D*([01]?\d)", s)
    if m:
        y = int(m.group(1)); mo = int(m.group(2))
        if 1 <= mo <= 12:
            return f"{y:04d}-{mo:02d}"
    d = re.sub(r"\D", "", s)
    if len(d) >= 6:
        y = int(d[:4]); mo = int(d[4:6])
        if 1 <= mo <= 12:
            return f"{y:04d}-{mo:02d}"
    return ""


def _strip_smbs_scripts(value) -> str:
    """SMBS 페이지의 d1('...'); 난독화 스크립트를 제거하고 화면 표시 텍스트만 남깁니다."""
    text = str(value or "")
    text = re.sub(r"d\d?\(\s*['\"].*?['\"]\s*\);", "", text)
    return text.strip()


def _month_start(month_key):
    mk = parse_month_key(month_key)
    if not mk:
        return pd.NaT
    return pd.Timestamp(year=int(mk[:4]), month=int(mk[5:7]), day=1)


def _month_end(month_key):
    ms = _month_start(month_key)
    if pd.isna(ms):
        return pd.NaT
    return ms + pd.offsets.MonthEnd(0)


def _months_between_keys(start_month, end_month):
    start = _month_start(start_month)
    end = _month_start(end_month)
    if pd.isna(start) or pd.isna(end):
        return []
    if end < start:
        start, end = end, start
    return [d.strftime("%Y-%m") for d in pd.date_range(start, end, freq="MS")]


def load_monthly_rate_cache(currency=None):
    """서울외국환중개 월평균 페이지에서 직접 수집한 캐시만 불러옵니다.

    v41 이하에서 일별 환율을 자체 평균하여 저장한 구형 캐시는 source 표기가
    없으므로 사용하지 않습니다. 잘못된 값이 재사용되는 것을 막기 위한 조치입니다.
    """
    path = Path(MONTHLY_RATE_CACHE_FILE)
    columns = ["currency", "year_month", "rate", "source", "source_url", "fetched_at"]
    if not path.exists():
        return pd.DataFrame(columns=columns)
    try:
        df = pd.read_csv(path, dtype={"currency": str, "year_month": str, "source": str})
        if df.empty or "source" not in df.columns:
            return pd.DataFrame(columns=columns)
        df["source"] = df["source"].astype(str).str.strip()
        df = df[df["source"] == MONTHLY_RATE_CACHE_SOURCE].copy()
        if df.empty:
            return pd.DataFrame(columns=columns)
        if "source_url" not in df.columns:
            df["source_url"] = SMBS_MON_AVG_RATE_URL
        if "fetched_at" not in df.columns:
            df["fetched_at"] = ""
        df["currency"] = df["currency"].astype(str).str.upper().str.strip()
        df["year_month"] = df["year_month"].apply(parse_month_key)
        df["rate"] = df["rate"].apply(to_number)
        df = df.dropna(subset=["currency", "year_month", "rate"])
        df = df[df["year_month"].astype(bool)]
        if currency:
            df = df[df["currency"] == str(currency).upper()].copy()
        return (
            df[columns]
            .drop_duplicates(subset=["currency", "year_month"], keep="last")
            .sort_values(["currency", "year_month"])
        )
    except Exception as e:
        _log(f"[WARN] 월평균 환율 캐시를 읽지 못했습니다. 서울외국환중개에서 다시 수집합니다: {e}")
        return pd.DataFrame(columns=columns)


def save_monthly_rate_cache(currency, data):
    """서울외국환중개 월평균 페이지에서 직접 읽은 값만 캐시에 저장합니다."""
    if data is None or data.empty:
        return
    path = Path(MONTHLY_RATE_CACHE_FILE)
    path.parent.mkdir(parents=True, exist_ok=True)
    new_df = data[["year_month", "rate"]].copy()
    new_df["currency"] = str(currency).upper()
    new_df["year_month"] = new_df["year_month"].apply(parse_month_key)
    new_df["rate"] = new_df["rate"].apply(to_number)
    new_df["source"] = MONTHLY_RATE_CACHE_SOURCE
    new_df["source_url"] = SMBS_MON_AVG_RATE_URL
    new_df["fetched_at"] = datetime.now().isoformat(timespec="seconds")
    columns = ["currency", "year_month", "rate", "source", "source_url", "fetched_at"]
    new_df = new_df.dropna(subset=["year_month", "rate"])[columns]
    new_df = new_df[new_df["year_month"].astype(bool)]
    old_df = load_monthly_rate_cache()
    merged = pd.concat([old_df, new_df], ignore_index=True)
    merged["currency"] = merged["currency"].astype(str).str.upper().str.strip()
    merged["year_month"] = merged["year_month"].apply(parse_month_key)
    merged["rate"] = merged["rate"].apply(to_number)
    merged["source"] = MONTHLY_RATE_CACHE_SOURCE
    merged["source_url"] = SMBS_MON_AVG_RATE_URL
    merged = merged.dropna(subset=["currency", "year_month", "rate"])
    merged = merged[merged["year_month"].astype(bool)]
    merged = (
        merged[columns]
        .drop_duplicates(subset=["currency", "year_month"], keep="last")
        .sort_values(["currency", "year_month"])
    )
    merged.to_csv(path, index=False, encoding="utf-8-sig")

def build_smbs_month_avg_params(currency, start_month, end_month):
    sdt = _month_start(start_month)
    edt = _month_start(end_month)
    return {
        "StrSch_sYear": sdt.strftime("%Y"),
        "StrSch_sMonth": sdt.strftime("%m"),
        "StrSch_sDay": "01",
        "StrSch_eYear": edt.strftime("%Y"),
        "StrSch_eMonth": edt.strftime("%m"),
        "StrSch_eDay": "01",
        "quick_date": "",
        "tongwha_code": currency,
    }


def build_smbs_month_avg_url(currency, start_month, end_month, base_url=SMBS_MON_AVG_RATE_URL):
    return f"{base_url}?{urlencode(build_smbs_month_avg_params(currency, start_month, end_month))}"


def _decode_smbs_obfuscated_text(payload: str) -> str:
    """SMBS d1/d2/d3/d4 스크립트 안의 %_Z32, %u_Ac6d4 형태를 복원합니다."""
    text = str(payload or "")
    text = re.sub(r"%u_[A-Za-z]([0-9A-Fa-f]{4})", lambda m: chr(int(m.group(1), 16)), text)
    text = re.sub(r"%_[A-Za-z]([0-9A-Fa-f]{2})", lambda m: chr(int(m.group(1), 16)), text)
    text = text.replace("%a", "\n").replace("%A", "\n")
    return text.strip()


def _smbs_cell_text(cell) -> str:
    """브라우저 DOM/requests 원문 양쪽에서 셀 표시문자를 얻습니다."""
    visible_parts = []
    for node in cell.find_all(string=True, recursive=True):
        if getattr(node.parent, "name", "") == "script":
            continue
        value = str(node).strip()
        if value:
            visible_parts.append(value)
    visible = " ".join(visible_parts).strip()
    if visible:
        return visible

    decoded = []
    for script in cell.find_all("script"):
        raw = script.string or script.get_text(" ", strip=True)
        m = re.search(r"d\d?\(\s*['\"](.*?)['\"]\s*\)", raw or "", flags=re.S)
        if m:
            value = _decode_smbs_obfuscated_text(m.group(1))
            if value:
                decoded.append(value)
    return " ".join(decoded).strip()


def parse_month_avg_table_from_html(html, currency):
    """월평균 결과표만 정확히 읽습니다.

    페이지 전체 텍스트를 훑지 않고, caption이 "월평균 매매기준율 결과 표"인
    표의 행만 사용합니다. 요청 통화가 아닌 기본 USD 화면이 돌아온 경우에는
    잘못된 환율을 저장하지 않고 빈 결과로 처리합니다.
    """
    if not html:
        return pd.DataFrame(columns=["year_month", "rate"])

    currency = str(currency or "").upper().strip()
    records = []

    try:
        soup = BeautifulSoup(html, "lxml")
        target_tables = []
        for table in soup.find_all("table"):
            caption = table.find("caption")
            caption_text = caption.get_text(" ", strip=True) if caption else ""
            headers = " ".join(th.get_text(" ", strip=True) for th in table.find_all("th"))
            if "월평균 매매기준율 결과 표" in caption_text or (
                "날짜" in headers and "통화명" in headers and "월평균" in headers
            ):
                target_tables.append(table)

        for table in target_tables:
            for tr in table.find_all("tr"):
                cells = tr.find_all("td")
                if len(cells) < 3:
                    continue
                texts = [_smbs_cell_text(c) for c in cells]
                row_text = " ".join(texts)

                # 요청 통화 행만 허용합니다. 예: 미국 달러 (USD)
                if not re.search(rf"\b{re.escape(currency)}\b", row_text, flags=re.I):
                    continue

                month_match = re.search(r"(20\d{2})[.\-/년\s]+(0?[1-9]|1[0-2])", texts[0])
                if not month_match:
                    continue
                month_key = f"{int(month_match.group(1)):04d}-{int(month_match.group(2)):02d}"

                rate = to_number(texts[-1])
                if rate is None or rate <= 0:
                    continue
                records.append({"year_month": month_key, "rate": float(rate)})
    except Exception:
        records = []

    if not records:
        return pd.DataFrame(columns=["year_month", "rate"])
    return (
        pd.DataFrame(records)
        .drop_duplicates(subset=["year_month"], keep="last")
        .sort_values("year_month")[["year_month", "rate"]]
        .reset_index(drop=True)
    )

def try_fetch_month_avg_rates_by_requests(currency, start_month, end_month, timeout=10):
    session = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
        "Referer": SMBS_MON_AVG_RATE_URL,
        "Origin": "http://www.smbs.biz",
    }
    params = build_smbs_month_avg_params(currency, start_month, end_month)
    attempts = [
        ("print_get", "GET", build_smbs_month_avg_url(currency, start_month, end_month, SMBS_MON_AVG_RATE_PRINT_URL), None),
        ("direct_get", "GET", build_smbs_month_avg_url(currency, start_month, end_month, SMBS_MON_AVG_RATE_URL), None),
        ("get_params", "GET", SMBS_MON_AVG_RATE_URL, params),
        ("post", "POST", SMBS_MON_AVG_RATE_URL, params),
    ]
    try:
        session.get(SMBS_MON_AVG_RATE_URL, headers=headers, timeout=min(timeout, 8))
    except Exception:
        pass
    needed = set(_months_between_keys(start_month, end_month))
    for _, method, url, payload in attempts:
        try:
            if method == "POST":
                h = dict(headers)
                h["Content-Type"] = "application/x-www-form-urlencoded"
                resp = session.post(url, data=payload, headers=h, timeout=timeout)
            else:
                resp = session.get(url, params=payload, headers=headers, timeout=timeout)
            resp.raise_for_status()
            enc = resp.apparent_encoding or resp.encoding or "cp949"
            if str(enc).lower() in ["iso-8859-1", "ascii"]:
                enc = "cp949"
            resp.encoding = enc
            data = parse_month_avg_table_from_html(resp.text, currency)
            if not data.empty:
                data = data[data["year_month"].isin(needed)].copy()
                if not data.empty:
                    return data
        except Exception:
            continue
    return pd.DataFrame(columns=["year_month", "rate"])


def fetch_month_avg_rates_by_selenium(currency, start_month, end_month, headless=True, timeout=60):
    """서울외국환중개 월평균 매매기준율 화면을 브라우저로 조회합니다.

    반환값은 MonAvgStdExRate.jsp 결과표에 표시된 공식 월평균 값뿐입니다.
    일별 환율을 평균내어 보완하지 않습니다.
    """
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.chrome.service import Service
        from selenium.webdriver.support.ui import Select
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
    except ImportError as e:
        raise RuntimeError("Selenium이 설치되어 있지 않습니다. pip install selenium 실행 후 다시 실행하세요.") from e

    def make_options(mode):
        opts = Options()
        if headless:
            opts.add_argument("--headless=new" if mode == "new" else "--headless")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--disable-extensions")
        opts.add_argument("--disable-popup-blocking")
        opts.add_argument("--window-size=1500,1100")
        opts.add_argument("--lang=ko-KR")
        for candidate in [
            os.environ.get("CHROME_BINARY"), os.environ.get("CHROME_BIN"),
            shutil.which("chromium"), shutil.which("chromium-browser"),
            shutil.which("google-chrome"), "/usr/bin/chromium",
            "/usr/bin/chromium-browser", "/usr/bin/google-chrome",
        ]:
            if candidate and Path(candidate).exists():
                opts.binary_location = str(candidate)
                break
        return opts

    def create_driver():
        service = None
        for candidate in [
            os.environ.get("CHROMEDRIVER"), os.environ.get("CHROMEDRIVER_PATH"),
            shutil.which("chromedriver"), "/usr/bin/chromedriver",
            "/usr/lib/chromium/chromedriver",
        ]:
            if candidate and Path(candidate).exists():
                service = Service(str(candidate))
                break
        errors = []
        for mode in (["new", "old"] if headless else ["visible"]):
            try:
                opts = make_options(mode)
                return webdriver.Chrome(service=service, options=opts) if service else webdriver.Chrome(options=opts)
            except Exception as e:
                errors.append(f"{mode}: {e}")
        raise RuntimeError(
            "Chrome/Chromium 실행 실패. packages.txt에 chromium, chromium-driver가 필요할 수 있습니다. "
            + " | ".join(errors)
        )

    currency = str(currency).upper().strip()
    needed = set(_months_between_keys(start_month, end_month))
    sdt = _month_start(start_month)
    edt = _month_start(end_month)
    driver = create_driver()
    try:
        wait = WebDriverWait(driver, timeout)

        # 직접 URL 조회
        driver.get(build_smbs_month_avg_url(currency, start_month, end_month))
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
        time.sleep(2)
        data = parse_month_avg_table_from_html(driver.page_source, currency)
        data = data[data["year_month"].isin(needed)].copy() if not data.empty else data
        if set(data["year_month"].tolist()) >= needed:
            return data.sort_values("year_month").reset_index(drop=True)

        # 직접 URL이 기본 USD 화면으로 돌아오는 환경에서는 폼을 명시적으로 조작합니다.
        driver.get(SMBS_MON_AVG_RATE_URL)
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
        Select(driver.find_element(By.NAME, "tongwha_code")).select_by_value(currency)
        Select(driver.find_element(By.NAME, "StrSch_sYear")).select_by_value(sdt.strftime("%Y"))
        Select(driver.find_element(By.NAME, "StrSch_sMonth")).select_by_value(sdt.strftime("%m"))
        Select(driver.find_element(By.NAME, "StrSch_eYear")).select_by_value(edt.strftime("%Y"))
        Select(driver.find_element(By.NAME, "StrSch_eMonth")).select_by_value(edt.strftime("%m"))
        driver.execute_script(
            "if (typeof doSearch === 'function') { doSearch('frm_SearchDate'); } "
            "else { document.forms['frm_SearchDate'].submit(); }"
        )
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
        time.sleep(3)
        data = parse_month_avg_table_from_html(driver.page_source, currency)
        if data.empty:
            return pd.DataFrame(columns=["year_month", "rate"])
        return (
            data[data["year_month"].isin(needed)]
            .drop_duplicates(subset=["year_month"], keep="last")
            .sort_values("year_month")
            .reset_index(drop=True)
        )
    finally:
        driver.quit()


def get_cached_or_fetch_month_avg_rates(currency, start_month, end_month):
    """공식 월평균 매매기준율만 반환합니다.

    조회 순서:
    1) source=SMBS_MON_AVG_OFFICIAL로 검증된 캐시
    2) 서울외국환중개 MonAvgStdExRate.jsp 직접 요청
    3) 같은 페이지를 숨김 브라우저로 조회

    세 방법으로도 공식 값을 확인하지 못하면 엑셀 생성을 중단합니다.
    일별 환율의 산술평균으로 대체하지 않습니다.
    """
    currency = str(currency).upper()
    months = _months_between_keys(start_month, end_month)
    cache = load_monthly_rate_cache(currency)
    have = set(cache["year_month"].tolist()) if not cache.empty else set()
    missing = [m for m in months if m not in have]

    if not missing:
        _log(f"  - {currency}: 저장된 공식 월평균 환율 사용")
        return (
            cache[cache["year_month"].isin(months)][["year_month", "rate"]]
            .drop_duplicates(subset=["year_month"], keep="last")
            .sort_values("year_month")
        )

    _log(f"  - {currency}: 서울외국환중개 월평균 환율 수집")

    fetched = try_fetch_month_avg_rates_by_requests(currency, missing[0], missing[-1])
    if fetched is None:
        fetched = pd.DataFrame(columns=["year_month", "rate"])
    fetched = fetched[fetched["year_month"].isin(missing)].copy() if not fetched.empty else fetched

    fetched_have = set(fetched["year_month"].tolist()) if not fetched.empty else set()
    unresolved = [m for m in missing if m not in fetched_have]
    if unresolved:
        browser_data = fetch_month_avg_rates_by_selenium(
            currency, unresolved[0], unresolved[-1], headless=True
        )
        if browser_data is not None and not browser_data.empty:
            browser_data = browser_data[browser_data["year_month"].isin(unresolved)].copy()
            fetched = (
                pd.concat([fetched, browser_data], ignore_index=True)
                if not fetched.empty else browser_data
            )

    if fetched is not None and not fetched.empty:
        fetched = (
            fetched.drop_duplicates(subset=["year_month"], keep="last")
            .sort_values("year_month")
        )
        save_monthly_rate_cache(currency, fetched)

    cache = load_monthly_rate_cache(currency)
    result = (
        cache[cache["year_month"].isin(months)][["year_month", "rate"]]
        .drop_duplicates(subset=["year_month"], keep="last")
        .sort_values("year_month")
    )
    remaining = [m for m in months if m not in set(result["year_month"].tolist())]
    if remaining:
        raise RuntimeError(
            "서울외국환중개 월평균 매매기준율을 직접 확인하지 못했습니다. "
            f"통화: {currency}, 미확인 월: {', '.join(remaining)}. "
            "일별 환율의 자체 평균값은 사용하지 않았습니다. 잠시 후 다시 실행해 주세요."
        )

    _log(f"  - {currency}: 공식 월평균 환율 수집 완료")
    return result

def _to_month_rate_entry(currency, raw_df, start_month, end_month):
    months = _months_between_keys(start_month, end_month)
    raw_df = raw_df.copy() if raw_df is not None else pd.DataFrame(columns=["year_month", "rate"])
    raw_df["year_month"] = raw_df.get("year_month", pd.Series(dtype=str)).apply(parse_month_key)
    raw_df["rate"] = raw_df.get("rate", pd.Series(dtype=float)).apply(to_number)
    raw_df = _normalize_rate_column(raw_df, currency)
    raw_df = raw_df.dropna(subset=["year_month", "rate"])
    monthly = []
    vals = []
    lookup = {r["year_month"]: float(r["rate"]) for _, r in raw_df.iterrows() if r.get("year_month")}
    missing = []
    for m in months:
        rate = lookup.get(m, 0.0)
        if not rate:
            missing.append(m)
        monthly.append({"year_month": m, "rate": rate})
        if rate:
            vals.append(rate)
    if missing:
        raise RuntimeError(
            f"{currency} 공식 월평균 매매기준율이 없습니다: {', '.join(missing)}"
        )
    return {
        "period": f"{months[0] if months else ''} ~ {months[-1] if months else ''}",
        "currency": currency,
        "currency_name": CURRENCY_NAMES.get(currency, currency),
        "average": round_applied_rate(currency, sum(vals) / len(vals)) if vals else 0.0,
        "min": min(vals) if vals else 0.0,
        "max": max(vals) if vals else 0.0,
        "min_date": "",
        "max_date": "",
        "range": round_applied_rate(currency, max(vals) - min(vals)) if vals else 0.0,
        "cross_rate": 0.0,
        "daily": [],
        "monthly": monthly,
    }


def fetch_monthly_avg_currencies_for_period(start_month, end_month, currencies: Iterable[str], logger: Optional[Callable[[str], None]] = None):
    previous_logger = _LOGGER
    set_logger(logger or previous_logger)
    try:
        out = {}
        used = [str(c).upper() for c in currencies if c]
        if used:
            _log(f"💱 월평균 환율 확인 중... ({', '.join(used)})")
        for cur in used:
            raw = get_cached_or_fetch_month_avg_rates(cur, start_month, end_month)
            out[cur] = _to_month_rate_entry(cur, raw, start_month, end_month)
        if used:
            _log("✅ 월평균 환율 확인 완료")
        return out
    finally:
        set_logger(previous_logger)


def merge_monthly_rates(base_rates: dict, monthly_rates: dict) -> dict:
    """기존 일별 rate_data에 monthly 목록을 병합합니다."""
    out = dict(base_rates or {})
    for cur, mdata in (monthly_rates or {}).items():
        if cur in out and out[cur]:
            out[cur]["monthly"] = mdata.get("monthly", [])
            out[cur]["monthly_average"] = mdata.get("average", 0.0)
        else:
            out[cur] = mdata
    return out


def monthly_avg_rate_for_month(rate_data: dict, month_key: str) -> float:
    """서울외국환중개에서 직접 수집한 해당 월의 공식 월평균 환율을 반환합니다."""
    mk = parse_month_key(month_key)
    if not rate_data:
        raise RuntimeError(f"공식 월평균 환율 데이터가 없습니다: {mk}")
    for row in rate_data.get("monthly", []) or []:
        if parse_month_key(row.get("year_month")) == mk:
            rate = float(row.get("rate") or 0.0)
            if rate > 0:
                return round_applied_rate(rate_data.get("currency", ""), rate)
    raise RuntimeError(
        f"서울외국환중개 공식 월평균 매매기준율이 없습니다: {mk}. "
        "일별 환율 평균으로 대체하지 않습니다."
    )
