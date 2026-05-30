#!/usr/bin/env python3
"""Build a static ChinaMoney monthly exchange-rate cache for GitHub Pages.

The website cannot reliably call ChinaMoney directly from the browser because of
CORS. This script runs in GitHub Actions, fetches RMB central parity data, and
writes site/rates.json. The browser then reads that static JSON.
"""
from __future__ import annotations

import datetime as dt
import json
import time
from decimal import Decimal
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import requests

URL = "https://www.chinamoney.com.cn/ags/ms/cm-u-bk-ccpr/CcprHisNew"
REFERER = "https://www.chinamoney.com.cn/chinese/bkccpr/"
CURRENCIES = ["USD", "EUR", "GBP", "JPY", "CAD", "AUD", "MXN", "PLN", "SEK", "TRY", "SGD"]
OUTPUT = Path("site/rates.json")


def month_starts(start: dt.date, end: dt.date) -> List[dt.date]:
    out = []
    cur = dt.date(start.year, start.month, 1)
    last = dt.date(end.year, end.month, 1)
    while cur <= last:
        out.append(cur)
        if cur.month == 12:
            cur = dt.date(cur.year + 1, 1, 1)
        else:
            cur = dt.date(cur.year, cur.month + 1, 1)
    return out


def record_date(record: Dict[str, Any]) -> Optional[dt.date]:
    for key in ["date", "showDate", "publishDate", "tradeDate"]:
        value = record.get(key)
        if value:
            try:
                return dt.datetime.strptime(str(value)[:10], "%Y-%m-%d").date()
            except Exception:
                pass
    return None


def record_value(record: Dict[str, Any]) -> Optional[Decimal]:
    if record.get("values"):
        try:
            return Decimal(str(record["values"][0]).replace(",", ""))
        except Exception:
            pass
    for key in ["value", "rate", "price", "middlePrice", "centralParity"]:
        value = record.get(key)
        if value not in (None, ""):
            try:
                return Decimal(str(value).replace(",", ""))
            except Exception:
                pass
    return None


def pair_for(currency: str) -> str:
    return "JPY/CNY" if currency == "JPY" else f"{currency}/CNY"


def fetch_records(currency: str, start: dt.date, end: dt.date) -> List[Tuple[dt.date, Decimal]]:
    params = {
        "startDate": start.strftime("%Y-%m-%d"),
        "endDate": end.strftime("%Y-%m-%d"),
        "currency": pair_for(currency),
        "pageNum": "1",
        "pageSize": "5000",
    }
    headers = {
        "Referer": REFERER,
        "Origin": "https://www.chinamoney.com.cn",
        "X-Requested-With": "XMLHttpRequest",
        "User-Agent": "Mozilla/5.0",
    }
    last_error = None
    for attempt in range(3):
        try:
            resp = requests.post(URL, params=params, headers=headers, timeout=30)
            if resp.status_code >= 400:
                resp = requests.get(URL, params=params, headers=headers, timeout=30)
            resp.raise_for_status()
            data = resp.json()
            records = data.get("records") or data.get("data") or []
            if isinstance(records, dict):
                records = records.get("records", [])
            out = []
            for record in records:
                d = record_date(record)
                v = record_value(record)
                if d and v is not None:
                    if currency == "JPY":
                        # ChinaMoney commonly quotes 100JPY/CNY. Store 1 JPY/CNY.
                        v = v / Decimal("100")
                    out.append((d, v))
            return sorted(out, key=lambda x: x[0])
        except Exception as exc:
            last_error = exc
            time.sleep(2 + attempt)
    raise RuntimeError(f"Failed to fetch {currency}: {last_error}")


def latest_on_or_before(records: List[Tuple[dt.date, Decimal]], target: dt.date) -> Optional[Tuple[dt.date, Decimal]]:
    candidates = [item for item in records if item[0] <= target]
    if not candidates:
        return None
    return candidates[-1]


def main() -> int:
    today = dt.date.today()
    start = dt.date(today.year - 3, 1, 1)
    # Include enough future months for settlement files dated slightly ahead.
    if today.month >= 10:
        end = dt.date(today.year + 1, 3, 1)
    else:
        end = dt.date(today.year, today.month + 3, 1)
    months = month_starts(start, end)
    result: Dict[str, Any] = {
        "updated_at": dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z",
        "source": "ChinaMoney/CFETS RMB central parity rate",
        "rule": "For each currency and report month, use the RMB central parity rate on the first calendar day of the month; if unavailable, use the latest trading day before that date.",
        "rates": {},
        "errors": {},
    }
    for currency in CURRENCIES:
        try:
            records = fetch_records(currency, start - dt.timedelta(days=14), end)
            result["rates"][currency] = {}
            for month in months:
                found = latest_on_or_before(records, month)
                if found:
                    source_date, value = found
                    result["rates"][currency][month.isoformat()] = {
                        "rate": float(value),
                        "source_date": source_date.isoformat(),
                        "pair": pair_for(currency),
                    }
            time.sleep(0.8)
        except Exception as exc:
            result["errors"][currency] = str(exc)
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8")
    print(f"Wrote {OUTPUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
