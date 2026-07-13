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

SMBS_STD_RATE_URL = "http://www.smbs.biz/ExRate/StdExRate.jsp"
SMBS_STD_RATE_PRINT_URL = "http://www.smbs.biz/ExRate/StdExRate_print.jsp"
SMBS_MON_AVG_RATE_URL = "http://www.smbs.biz/ExRate/MonAvgStdExRate.jsp"
SMBS_MON_AVG_RATE_PRINT_URL = "http://www.smbs.biz/ExRate/MonAvgStdExRate_print.jsp"
RATE_CACHE_FILE = Path(__file__).resolve().parents[1] / "data" / "exchange_rate_cache.csv"
MONTHLY_RATE_CACHE_FILE = Path(__file__).resolve().parents[1] / "data" / "monthly_exchange_rate_cache.csv"
_FIXED_RATE_JSON_CANDIDATES = [
    Path(__file__).resolve().parents[1] / "data" / "fixed_rates_2025.json",
    Path(__file__).resolve().parents[1] / "fixed_rates_2025.json",
    Path.cwd() / "data" / "fixed_rates_2025.json",
    Path.cwd() / "fixed_rates_2025.json",
]
RATE_LOOKBACK_DAYS = 7

CURRENCY_NAMES = {
    "MYR": "말레이시아 링깃 (MYR)",
    "PHP": "필리핀 페소 (PHP)",
    "SGD": "싱가포르 달러 (SGD)",
    "THB": "태국 바트 (THB)",
    "TWD": "대만 달러 (TWD)",
    "VND": "베트남 동 (VND) (100)",
    "JPY": "일본 엔 (JPY) (100)",
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
    merged = pd.concat([old_df, new_df], ignore_index=True)
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


def get_cached_or_fetch_smbs_period_rates(currency, start_date, end_date):
    currency = str(currency).upper()
    start_date = pd.to_datetime(start_date).normalize()
    end_date = pd.to_datetime(end_date).normalize()
    cache = load_rate_cache(currency)

    if _has_full_daily_cache(cache, start_date, end_date):
        _log(f"  - {currency}: 저장된 환율 사용")
        data = cache[(cache["date"] >= start_date) & (cache["date"] <= end_date)].copy()
        return data[["date", "rate"]].drop_duplicates(subset=["date"], keep="last").sort_values("date")

    fixed = load_fixed_rate_json(currency, start_date, end_date)
    if fixed is not None and not fixed.empty and _has_full_daily_cache(fixed.assign(currency=currency), start_date, end_date):
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
            _log(f"  - {currency}: 저장된 환율 보정")
            _cache_requested_daily_range(currency, start_date, end_date)

    if segments:
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


def _to_rate_entry(currency, raw_df, start_date, end_date):
    if raw_df is None or raw_df.empty:
        return {
            "period": f"{pd.to_datetime(start_date).strftime('%Y.%m.%d')} ~ {pd.to_datetime(end_date).strftime('%Y.%m.%d')}",
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
    filled = fill_missing_dates(raw_df, start_date, end_date)
    vals = [float(x) for x in filled["rate"].dropna().tolist()]
    min_idx = filled["rate"].idxmin() if vals else None
    max_idx = filled["rate"].idxmax() if vals else None
    daily = []
    prev = None
    for _, r in filled.iterrows():
        rate = float(r["rate"])
        change = 0.0 if prev is None else round(rate - prev, 6)
        daily.append({"date": pd.to_datetime(r["date"]).strftime("%Y.%m.%d"), "rate": rate, "change": change, "cross": 0})
        prev = rate
    return {
        "period": f"{pd.to_datetime(start_date).strftime('%Y.%m.%d')} ~ {pd.to_datetime(end_date).strftime('%Y.%m.%d')}",
        "currency": currency,
        "currency_name": CURRENCY_NAMES.get(currency, currency),
        "average": round(sum(vals) / len(vals), 4) if vals else 0.0,
        "min": min(vals) if vals else 0.0,
        "min_date": pd.to_datetime(filled.loc[min_idx, "date"]).strftime("%Y.%m.%d") if min_idx is not None else "",
        "max": max(vals) if vals else 0.0,
        "max_date": pd.to_datetime(filled.loc[max_idx, "date"]).strftime("%Y.%m.%d") if max_idx is not None else "",
        "range": round(max(vals) - min(vals), 4) if vals else 0.0,
        "cross_rate": 0.0,
        "daily": daily,
    }


def fetch_all_currencies_for_period(start_date, end_date, currencies: Iterable[str], logger: Optional[Callable[[str], None]] = None):
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
            out[cur] = _to_rate_entry(cur, raw, start_date, end_date)
        if used:
            _log("✅ 환율 수집 완료")
        return out
    finally:
        set_logger(previous_logger)


def fetch_all_currencies(year: int, month: int, currencies: Iterable[str]):
    last = pd.Timestamp(year=year, month=month, day=1) + pd.offsets.MonthEnd(0)
    first = pd.Timestamp(year=year, month=month, day=1) - pd.Timedelta(days=RATE_LOOKBACK_DAYS)
    return fetch_all_currencies_for_period(first, last, currencies)



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
        return round(sum(vals) / len(vals), 4)
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
    path = Path(MONTHLY_RATE_CACHE_FILE)
    if not path.exists():
        return pd.DataFrame(columns=["currency", "year_month", "rate", "fetched_at"])
    try:
        df = pd.read_csv(path, dtype={"currency": str, "year_month": str})
        if df.empty:
            return pd.DataFrame(columns=["currency", "year_month", "rate", "fetched_at"])
        df["currency"] = df["currency"].astype(str).str.upper().str.strip()
        df["year_month"] = df["year_month"].apply(parse_month_key)
        df["rate"] = df["rate"].apply(to_number)
        df = df.dropna(subset=["currency", "year_month", "rate"])
        df = df[df["year_month"].astype(bool)]
        if currency:
            df = df[df["currency"] == str(currency).upper()].copy()
        return df.drop_duplicates(subset=["currency", "year_month"], keep="last").sort_values(["currency", "year_month"])
    except Exception as e:
        _log(f"[WARN] 월평균 환율 캐시를 읽지 못했습니다. 새로 수집합니다: {e}")
        return pd.DataFrame(columns=["currency", "year_month", "rate", "fetched_at"])


def save_monthly_rate_cache(currency, data):
    if data is None or data.empty:
        return
    path = Path(MONTHLY_RATE_CACHE_FILE)
    path.parent.mkdir(parents=True, exist_ok=True)
    new_df = data[["year_month", "rate"]].copy()
    new_df["currency"] = str(currency).upper()
    new_df["year_month"] = new_df["year_month"].apply(parse_month_key)
    new_df["rate"] = new_df["rate"].apply(to_number)
    new_df["fetched_at"] = datetime.now().isoformat(timespec="seconds")
    new_df = new_df.dropna(subset=["year_month", "rate"])[["currency", "year_month", "rate", "fetched_at"]]
    new_df = new_df[new_df["year_month"].astype(bool)]
    old_df = load_monthly_rate_cache()
    merged = pd.concat([old_df, new_df], ignore_index=True)
    merged["currency"] = merged["currency"].astype(str).str.upper().str.strip()
    merged["year_month"] = merged["year_month"].apply(parse_month_key)
    merged["rate"] = merged["rate"].apply(to_number)
    merged = merged.dropna(subset=["currency", "year_month", "rate"])
    merged = merged[merged["year_month"].astype(bool)]
    merged = merged.drop_duplicates(subset=["currency", "year_month"], keep="last").sort_values(["currency", "year_month"])
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


def parse_month_avg_table_from_html(html, currency):
    if not html:
        return pd.DataFrame(columns=["year_month", "rate"])
    candidates = []
    try:
        tables = pd.read_html(StringIO(html))
    except Exception:
        tables = []
    for table in tables:
        if table is None or table.empty:
            continue
        if table.shape[1] < 3:
            continue
        records = []
        for _, row in table.iterrows():
            vals = row.tolist()
            joined = " ".join(str(v) for v in vals if not pd.isna(v))
            if re.search(r"\([A-Z]{3}\)", joined) and not row_mentions_currency(vals, currency):
                continue
            mkey = ""
            for v in vals:
                mk = parse_month_key(_strip_smbs_scripts(v))
                if mk:
                    mkey = mk
                    break
            if not mkey:
                continue
            nums = []
            for v in vals:
                if is_currency_name_cell(v, currency):
                    continue
                raw = _strip_smbs_scripts(v)
                # 연월 셀 숫자는 환율 후보에서 제외
                if parse_month_key(raw) == mkey and re.search(r"20\d{2}", raw):
                    continue
                num = to_number(raw)
                if num is None:
                    matches = re.findall(r"[-+]?\d{1,3}(?:,\d{3})*(?:\.\d+)?|[-+]?\d+(?:\.\d+)?", raw)
                    for candidate in reversed(matches):
                        num = to_number(candidate)
                        if num is not None:
                            break
                if num is None or num <= 0:
                    continue
                if 1 <= num <= 12 or 1900 <= num <= 2100:
                    continue
                if currency.upper() in {"VND", "JPY", "IDR"} and abs(num - 100) < 1e-9:
                    continue
                nums.append(num)
            if nums:
                records.append({"year_month": mkey, "rate": float(nums[0])})
        if records:
            df = pd.DataFrame(records).drop_duplicates(subset=["year_month"], keep="last").sort_values("year_month")
            if not df.empty:
                candidates.append(df)
    if not candidates:
        # HTML 텍스트 백업 파싱
        text = re.sub(r"<script.*?</script>", " ", html, flags=re.I | re.S)
        text = re.sub(r"<style.*?</style>", " ", text, flags=re.I | re.S)
        text = re.sub(r"<[^>]+>", " ", text).replace("&nbsp;", " ")
        pat = re.compile(r"(20\d{2})[.\-/년\s]*(0?[1-9]|1[0-2])\D{0,50}?([-+]?\d{1,3}(?:,\d{3})*(?:\.\d+)?|[-+]?\d+(?:\.\d+)?)")
        records = []
        for m in pat.finditer(text):
            mkey = f"{int(m.group(1)):04d}-{int(m.group(2)):02d}"
            rate = to_number(m.group(3))
            if rate and rate > 0:
                records.append({"year_month": mkey, "rate": rate})
        if records:
            candidates.append(pd.DataFrame(records).drop_duplicates(subset=["year_month"], keep="last").sort_values("year_month"))
    if not candidates:
        return pd.DataFrame(columns=["year_month", "rate"])
    return max(candidates, key=len)[["year_month", "rate"]].copy()


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


def get_cached_or_fetch_month_avg_rates(currency, start_month, end_month):
    currency = str(currency).upper()
    months = _months_between_keys(start_month, end_month)
    cache = load_monthly_rate_cache(currency)
    have = set(cache["year_month"].tolist()) if not cache.empty else set()
    missing = [m for m in months if m not in have]
    if not missing:
        _log(f"  - {currency}: 저장된 월평균 환율 사용")
        return cache[cache["year_month"].isin(months)][["year_month", "rate"]].drop_duplicates(subset=["year_month"], keep="last").sort_values("year_month")
    _log(f"  - {currency}: 부족한 월평균 환율 수집")
    fetched = try_fetch_month_avg_rates_by_requests(currency, missing[0], missing[-1])
    # 월평균 페이지 수집 실패 시 기간별 일별 환율 평균으로 보완합니다.
    if fetched.empty or not set(missing).issubset(set(fetched.get("year_month", []))):
        fallback_rows = []
        fetched_have = set(fetched.get("year_month", [])) if not fetched.empty else set()
        for m in missing:
            if m in fetched_have:
                continue
            raw = get_cached_or_fetch_smbs_period_rates(currency, _month_start(m), _month_end(m))
            if raw is not None and not raw.empty:
                fallback_rows.append({"year_month": m, "rate": round(float(raw["rate"].mean()), 4)})
        if fallback_rows:
            fetched = pd.concat([fetched, pd.DataFrame(fallback_rows)], ignore_index=True) if not fetched.empty else pd.DataFrame(fallback_rows)
    save_monthly_rate_cache(currency, fetched)
    cache = load_monthly_rate_cache(currency)
    return cache[cache["year_month"].isin(months)][["year_month", "rate"]].drop_duplicates(subset=["year_month"], keep="last").sort_values("year_month")


def _to_month_rate_entry(currency, raw_df, start_month, end_month):
    months = _months_between_keys(start_month, end_month)
    raw_df = raw_df.copy() if raw_df is not None else pd.DataFrame(columns=["year_month", "rate"])
    raw_df["year_month"] = raw_df.get("year_month", pd.Series(dtype=str)).apply(parse_month_key)
    raw_df["rate"] = raw_df.get("rate", pd.Series(dtype=float)).apply(to_number)
    raw_df = raw_df.dropna(subset=["year_month", "rate"])
    monthly = []
    vals = []
    lookup = {r["year_month"]: float(r["rate"]) for _, r in raw_df.iterrows() if r.get("year_month")}
    for m in months:
        rate = lookup.get(m, 0.0)
        monthly.append({"year_month": m, "rate": rate})
        if rate:
            vals.append(rate)
    return {
        "period": f"{months[0] if months else ''} ~ {months[-1] if months else ''}",
        "currency": currency,
        "currency_name": CURRENCY_NAMES.get(currency, currency),
        "average": round(sum(vals) / len(vals), 4) if vals else 0.0,
        "min": min(vals) if vals else 0.0,
        "max": max(vals) if vals else 0.0,
        "min_date": "",
        "max_date": "",
        "range": round(max(vals) - min(vals), 4) if vals else 0.0,
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
    mk = parse_month_key(month_key)
    if not rate_data:
        return 0.0
    for row in rate_data.get("monthly", []) or []:
        if parse_month_key(row.get("year_month")) == mk:
            return float(row.get("rate") or 0.0)
    # 월평균 데이터가 없으면 일별 환율 평균으로 보완
    s = _month_start(mk)
    e = _month_end(mk)
    if pd.isna(s) or pd.isna(e):
        return float(rate_data.get("average", 0.0) or 0.0)
    return avg_rate_for_period(rate_data, s.strftime("%Y-%m-%d"), e.strftime("%Y-%m-%d"))
