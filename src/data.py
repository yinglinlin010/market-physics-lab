from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from datetime import date as date_type, timedelta
from pathlib import Path
import hashlib
import json
import time
import requests
from io import BytesIO

import numpy as np
import pandas as pd


@dataclass
class PriceData:
    prices: pd.DataFrame
    source: str
    retrieved_at: str
    adjustment: str
    version: str
    synthetic: bool = False


def make_synthetic_prices(n: int = 900, seed: int = 42, phi: float = 0.92) -> PriceData:
    """Generate two positive price paths with a known, mean-reverting log spread."""
    if n < 80:
        raise ValueError("合成資料至少需要 80 個交易日。")
    rng = np.random.default_rng(seed)
    dates = pd.bdate_range("2021-01-04", periods=n, name="date")
    log_b = np.log(100) + np.cumsum(rng.normal(0.00025, 0.009, n))
    spread = np.zeros(n)
    spread[0] = 0.0
    for i in range(1, n):
        spread[i] = phi * spread[i - 1] + rng.normal(0, 0.012)
    log_a = 0.15 + 1.08 * log_b + spread
    prices = pd.DataFrame({"A_SYN": np.exp(log_a), "B_SYN": np.exp(log_b)}, index=dates)
    prices.index.name = "date"
    return PriceData(prices, "本機合成資料（非真實股票）", "由固定種子即時產生", "未調整模擬收盤價", f"synthetic-v1-seed-{seed}", True)


def load_local_csv(path: str | Path, tickers: list[str]) -> PriceData:
    raw = Path(path).read_bytes()
    return load_local_csv_bytes(raw, Path(path).name, tickers)


def load_bundled_twse_sample(path: str | Path) -> PriceData:
    """Load the versioned, openly licensed TWSE sample shipped with the project."""
    path = Path(path)
    raw = path.read_bytes()
    data = load_local_csv_bytes(raw, path.name, ["2330", "2454"])
    data.source = "臺灣證券交易所官方固定樣本（本機資料）"
    data.retrieved_at = "2026-11-01T13:34:02+00:00"
    data.adjustment = "原始日收盤價（未做股利或分割還原）"
    data.version = f"csv-sha256:{file_sha256(raw)}"
    return data


def load_local_csv_bytes(raw: bytes, name: str, tickers: list[str]) -> PriceData:
    frame = pd.read_csv(BytesIO(raw), parse_dates=["date"])
    required = {"date", "ticker", "close"}
    if not required.issubset(frame.columns):
        raise ValueError("CSV 必須包含 date,ticker,close 欄位。")
    frame["ticker"] = frame["ticker"].astype(str)
    frame = frame[frame.ticker.astype(str).isin(tickers)]
    if frame.empty:
        raise ValueError("CSV 中找不到所選股票代碼。")
    if frame.duplicated(["date", "ticker"]).any():
        raise ValueError("CSV 存在重複的日期與股票代碼。")
    if (~np.isfinite(pd.to_numeric(frame.close, errors="coerce"))).any() or (pd.to_numeric(frame.close, errors="coerce") <= 0).any():
        raise ValueError("收盤價必須是有限的正數。")
    prices = frame.pivot(index="date", columns="ticker", values="close").sort_index()
    prices = prices.dropna(subset=tickers)
    return PriceData(prices[tickers], f"使用者 CSV：{name}", "使用者提供", "依 CSV 說明", file_sha256(raw), False)


def fetch_twse(tickers: list[str], start: str, end: str, delay_seconds: float = 0.5,
               cache_dir: Path | None = None, progress_callback=None) -> PriceData:
    """Fetch daily closes through TWSE's official per-stock monthly history endpoint.

    The matching daily-close data is listed on data.gov.tw under Open Government
    Data License 1.0. Requests are sequential and modest; raw source rows are kept
    local unless a user chooses to export them.
    """
    start_date, end_date = pd.Timestamp(start).date(), pd.Timestamp(end).date()
    if start_date >= end_date:
        raise ValueError("開始日期必須早於結束日期。")
    final_day = end_date - timedelta(days=1)
    months = pd.period_range(start_date.strftime("%Y-%m"), final_day.strftime("%Y-%m"), freq="M")
    session = requests.Session()
    session.headers.update({"User-Agent": "market-physics-lab/0.1 (educational research; sequential low-volume requests)"})
    records = []
    cache_dir = cache_dir or Path("data/.twse-cache")
    total_requests = len(tickers) * len(months)
    completed = 0
    for ticker in tickers:
        for month in months:
            month_start = month.start_time
            query_date = month_start.strftime("%Y%m%d")
            cache_path = cache_dir / ticker / f"{month}.json"
            payload = None
            requested = False
            if cache_path.exists():
                try:
                    payload = json.loads(cache_path.read_text(encoding="utf-8"))
                except (json.JSONDecodeError, OSError):
                    payload = None
            if payload is None:
                requested = True
                response = session.get(
                    "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY",
                    params={"date": query_date, "stockNo": ticker, "response": "json"}, timeout=20,
                )
                response.raise_for_status()
                try:
                    payload = response.json()
                except requests.exceptions.JSONDecodeError as exc:
                    raise RuntimeError(
                        f"TWSE 回應不是 JSON（股票 {ticker}、月份 {month}、HTTP {response.status_code}）。"
                        "已保存的月份可在重試時沿用；也可改用本機 CSV。"
                    ) from exc
                cache_path.parent.mkdir(parents=True, exist_ok=True)
                cache_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
            if payload.get("stat") != "OK":
                raise RuntimeError(f"TWSE 回應異常（{ticker}、{query_date}）：{payload.get('stat', '未知狀態')}")
            fields = payload.get("fields", [])
            try:
                date_col, close_col = fields.index("日期"), fields.index("收盤價")
            except ValueError as exc:
                raise RuntimeError("TWSE 回應欄位與預期格式不同，請檢查最新 API 規格。") from exc
            for row in payload.get("data", []):
                roc = row[date_col].strip()
                year, mon, day = map(int, roc.split("/"))
                day_date = date_type(year + 1911, mon, day)
                if start_date <= day_date < end_date:
                    close = float(row[close_col].replace(",", ""))
                    if np.isfinite(close) and close > 0:
                        records.append({"date": pd.Timestamp(day_date), "ticker": ticker, "close": close})
            completed += 1
            if progress_callback:
                progress_callback(completed, total_requests, f"{ticker}｜{month}")
            if requested and delay_seconds and completed < total_requests:
                time.sleep(delay_seconds)
    if not records:
        raise ValueError("TWSE 在所選期間沒有回傳有效共同交易日。")
    long = pd.DataFrame(records).drop_duplicates(["date", "ticker"])
    prices = long.pivot(index="date", columns="ticker", values="close").reindex(columns=tickers).dropna(how="any").sort_index()
    if len(prices) < 80:
        raise ValueError(f"TWSE 共同交易日只有 {len(prices)} 筆；請選擇至少約四個月的期間。")
    prices.index.name = "date"
    return PriceData(prices, "臺灣證券交易所官方 STOCK_DAY 歷史月資料", datetime.now(timezone.utc).isoformat(timespec="seconds"),
                     "原始日收盤價（未做股利或分割還原）", file_sha256(prices.to_csv().encode()), False)


def file_sha256(path_or_bytes: str | Path | bytes) -> str:
    data = path_or_bytes if isinstance(path_or_bytes, bytes) else Path(path_or_bytes).read_bytes()
    return hashlib.sha256(data).hexdigest()


def persist_manifest(directory: Path, metadata: dict) -> None:
    directory.mkdir(parents=True, exist_ok=True)
    (directory / "run_manifest.json").write_text(json.dumps(metadata, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
