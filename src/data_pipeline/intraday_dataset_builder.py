"""
Intraday-Style Dictionary Builder (Daily Bars)

Purpose:
- Builds DICTIONARY_TABLE (passing tickers) and SEEDLING_TABLE (very low-price “seedlings”)
  using the most recent completed daily bars from Alpaca.

Inputs:
- Alpaca daily bars (IEX or SIP feed).
- Local SQLite database file.

Outputs (SQLite):
- DICTIONARY_TABLE: 53-row blocks per ticker (31 most recent completed days).
- SEEDLING_TABLE: 53-row blocks per seedling ticker (31 most recent completed days of tickers under $2).

Notes:
- This file is provided for reference; running requires local dependencies + Alpaca credentials.
- API keys, base URL, and DB path are redacted in the public repo.
"""

# Alpaca intraday dictionary builder

from typing import List, Tuple
import requests
from datetime import datetime, timezone, timedelta
import random
import sqlite3
from queue import Queue, Empty
import threading

from alpaca_trade_api import REST
from alpaca_trade_api.rest import TimeFrame
import pandas as pd 

# Config

# ===== ALPACA MARKET DATA CONFIG FOR MP1 =====
ALPACA_API_KEY = ""
ALPACA_API_SECRET = ""
ALPACA_BASE_URL = ""

# How many worker threads to use for bar fetching
N_WORKERS = 4

DB_PATH = ""

USE_SIP = True
LIMIT_TESTED = False
TESTED_CAP = 100

PRICE_MIN = 10.0
PRICE_MAX = 90.0
VOLUME_MIN = 250_000
VOLUME_MAX = 100_000_000
SEEDLING_CUTOFF = 2.0

from zoneinfo import ZoneInfo
NY_TZ = ZoneInfo("America/New_York")

# Helpers

def get_alpaca_rest() -> REST:
    return REST(ALPACA_API_KEY, ALPACA_API_SECRET, ALPACA_BASE_URL)

def fetch_alpaca_assets():
    url = f"{ALPACA_BASE_URL}/v2/assets?status=active"
    headers = {
        "APCA-API-KEY-ID": ALPACA_API_KEY,
        "APCA-API-SECRET-KEY": ALPACA_API_SECRET,
    }
    try:
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        return resp.json()
    except Exception as e:
        print(f"[ERROR] fetch_alpaca_assets failed: {e}")
        return []

# Keyword filter for non-stocks
NAME_EXCLUDE_KEYWORDS = [
    " etf ",
    " etn ",
    " fund",
    " trust",
    " income",
    " dividend",
    " municipal",
    " muni",
    " bond",
    " notes",
    " note",
    " preferred",
    " preference",
    " depositary",
    " depository",
    " portfolio",
    " index",
    " allocation",
    " target",
    " strategy",
    " capital",
    " real estate",
    " properties",
    " property",
    " reit",
    " closed-end",
    " closed end",
    " unit ",
    " units ",
    " warrant",
    " rights",
]


def is_clean_common_stock(asset: dict) -> bool:
    sym = asset.get("symbol")
    exch = asset.get("exchange")
    status = asset.get("status")
    asset_class = asset.get("class")  # 'us_equity', 'crypto', etc.
    name = (asset.get("name") or "").strip()

    if not sym or not isinstance(sym, str):
        return False

    # Only US common equities
    if asset_class != "us_equity":
        return False

    # Only active listings
    if status != "active":
        return False

    # Drop OTC / pink sheets
    if exch == "OTC":
        return False

    # Only alphabetic tickers
    if not sym.isalpha():
        return False

    # Name-based nuking of non-stock garbage
    name_lower = " " + name.lower() + " "
    name_upper = name.upper()

    for kw in NAME_EXCLUDE_KEYWORDS:
        if kw in name_lower:
            return False

    # Kill obvious ETF families and similar wrapper products
    if name_upper.startswith(
        ("SPDR", "ISHARES", "VANGUARD", "INVESCO", "PROSHARES", "GLOBAL X", "DIREXION")
    ):
        return False
    sym_len = len(sym)
    if sym_len > 5 and (
        "FUND" in name_upper or "TRUST" in name_upper or "ETF" in name_upper
    ):
        return False

    return True

def chunk_list(lst: List[Tuple[str, str]], n_chunks: int) -> List[List[Tuple[str, str]]]:
    if n_chunks <= 1 or len(lst) == 0:
        return [lst]
    chunk_size = (len(lst) + n_chunks - 1) // n_chunks
    return [lst[i : i + chunk_size] for i in range(0, len(lst), chunk_size)]

def drop_incomplete_today_daily_bar(bars: pd.DataFrame) -> pd.DataFrame:
    if bars is None or bars.empty:
        return bars

    first_ts = bars.index[0]
    if isinstance(first_ts, datetime):
        if first_ts.tzinfo is None:
            first_ts = first_ts.replace(tzinfo=timezone.utc)
    else:
        return bars

    first_day_ny = first_ts.astimezone(NY_TZ).date()
    now_ny = datetime.now(NY_TZ)
    today_ny = now_ny.date()

    if now_ny.weekday() < 5:
        if now_ny.hour < 16 and first_day_ny == today_ny:
            return bars.iloc[1:]

    return bars
  
def _make_53row_block(sym: str, name: str, window: pd.DataFrame) -> list:
    block_rows = []
    block_rows.append(("", "", "", "", "", "", "", ""))
    use_name = name or sym
    block_rows.append((use_name, "Date", "Open", "High", "Low", "Close", "Volume", "Volume_Again"))

    first_data_row = True
    for ts, row in window.iterrows():
        ts_dt = ts.to_pydatetime() if hasattr(ts, "to_pydatetime") else ts
        date_str = ts_dt.strftime("%a %b %d %Y %H:%M:%S %Z")

        o = float(row["open"])
        h = float(row["high"])
        l = float(row["low"])
        c = float(row["close"])
        v = int(row["volume"])

        ticker_col = sym if first_data_row else ""
        first_data_row = False

        block_rows.append((ticker_col, date_str, f"{o:.2f}", f"{h:.2f}", f"{l:.2f}", f"{c:.2f}", str(v), str(v)))

    while len(block_rows) < 53:
        block_rows.append(("", "", "", "", "", "", "", ""))

    return block_rows

# Tables + builder

def _create_dictionary_tables(conn: sqlite3.Connection):
    cur = conn.cursor()
    cur.execute("DROP TABLE IF EXISTS DICTIONARY_TABLE")
    cur.execute(
        """
        CREATE TABLE DICTIONARY_TABLE (
            Ticker_name   TEXT,
            Date          TEXT,
            Open          TEXT,
            High          TEXT,
            Low           TEXT,
            Close         TEXT,
            Volume        TEXT,
            Volume_Again  TEXT
        )
        """
    )

    cur.execute("DROP TABLE IF EXISTS SEEDLING_TABLE")
    cur.execute(
        """
        CREATE TABLE SEEDLING_TABLE (
            Ticker_name   TEXT,
            Date          TEXT,
            Open          TEXT,
            High          TEXT,
            Low           TEXT,
            Close         TEXT,
            Volume        TEXT,
            Volume_Again  TEXT
        )
        """
    )
    conn.commit()
def build_alpaca_intraday_dictionary(
    db_path: str = DB_PATH,
):

    data_feed = "sip" if USE_SIP else "iex"
    print(
        f"=== build_alpaca_intraday_dictionary START ===\n"
        f"[CFG] feed={data_feed}, LIMIT_TESTED={LIMIT_TESTED}, TESTED_CAP={TESTED_CAP}"
    )

    conn = sqlite3.connect(db_path)
    _create_dictionary_tables(conn)
    cur = conn.cursor()

    # UNIVERSE
    assets = fetch_alpaca_assets()
    if not assets:
        print("[ERROR] No assets returned by Alpaca; aborting.")
        conn.close()
        return

    universe_symbols: List[Tuple[str, str]] = []
    seen = set()
    for a in assets:
        if not is_clean_common_stock(a):
            continue
        sym = a.get("symbol")
        name = a.get("name") or ""
        if sym in seen:
            continue
        seen.add(sym)
        universe_symbols.append((sym, name))

    print(f"[UNIVERSE] Clean US stocks: {len(universe_symbols)}")
    if not universe_symbols:
        print("[ERROR] Universe empty after filtering; aborting.")
        conn.close()
        return

    random.shuffle(universe_symbols)

    CALENDAR_LOOKBACK_INTRADAY = 70 
    MIN_DAYS = 31
    OFFSET = 0 

    end_dt = datetime.now(timezone.utc)
    start_dt = end_dt - timedelta(days=CALENDAR_LOOKBACK_INTRADAY)

    shared = {
        "tested": 0,
        "passing": 0,
        "seedlings": 0,
        "stop": False,
    }
    progress_lock = threading.Lock()
    q: Queue = Queue()

    def worker(sym_slice: List[Tuple[str, str]]):
        api = get_alpaca_rest()
        for sym, name in sym_slice:
            with progress_lock:
                if shared["stop"]:
                    break
                shared["tested"] += 1
                t_no = shared["tested"]
                if LIMIT_TESTED and t_no > TESTED_CAP:
                    shared["stop"] = True
                    break
                if t_no % 50 == 0:
                    print(
                        f"[PROGRESS] tested={shared['tested']} "
                        f"passing={shared['passing']} seedlings={shared['seedlings']}"
                    )

            try:
                bars = api.get_bars(
                    sym,
                    TimeFrame.Day,
                    start_dt.isoformat(),
                    end_dt.isoformat(),
                    adjustment="raw",
                    feed=data_feed,
                ).df
            except Exception as e:
                print(f"[FAIL] {sym}: GET_BARS_ERROR ({e})")
                continue

            if bars is None or bars.empty:
                print(f"[FAIL] {sym}: NO_BARS")
                continue

            bars = bars.sort_index(ascending=False)
            bars = drop_incomplete_today_daily_bar(bars)

            if len(bars) < MIN_DAYS + OFFSET:
                print(f"[FAIL] {sym}: LESS_THAN_{MIN_DAYS + OFFSET}_BARS ({len(bars)})")
                continue

            latest = bars.iloc[OFFSET]
            try:
                open_val  = float(latest["open"])
                last_close = float(latest["close"])
                last_vol   = float(latest["volume"])
            except Exception:
                print(f"[FAIL] {sym}: BAD_LATEST_BAR")
                continue

            # Seedling logic first
            if open_val < SEEDLING_CUTOFF:
                with progress_lock:
                    shared["seedlings"] += 1
                q.put((sym, name, bars, "seedling"))
                continue

            price_ok = (PRICE_MIN < last_close < PRICE_MAX)
            vol_ok   = (VOLUME_MIN <= last_vol <= VOLUME_MAX)
            if not price_ok or not vol_ok:
                print(
                    f"[FAIL] {sym}: PRICE_VOL_FILTER_FAIL "
                    f"(open={open_val}, close={last_close}, vol={last_vol})"
                )
                continue

            with progress_lock:
                shared["passing"] += 1
            q.put((sym, name, bars, "normal"))

    # Launch workers
    slices = chunk_list(universe_symbols, N_WORKERS)
    workers: List[threading.Thread] = []
    for i, sym_slice in enumerate(slices):
        if not sym_slice:
            continue
        t = threading.Thread(
            target=worker,
            args=(sym_slice,),
            name=f"intra_worker_{i+1}",
            daemon=True,
        )
        t.start()
        workers.append(t)

    # Writer: dictionary + seedlings
    try:
        while True:
            try:
                sym, name, bars, kind = q.get(timeout=1.0)
            except Empty:
                if not any(t.is_alive() for t in workers):
                    break
                continue

            # Use the most recent MIN_DAYS days
            window = bars.iloc[OFFSET:OFFSET + MIN_DAYS]
            if len(window) < MIN_DAYS:
                q.task_done()
                continue

            block_rows = _make_53row_block(sym, name, window)
            if len(block_rows) != 53:
                print(f"[WARN] {sym}/{kind}: block size != 53 ({len(block_rows)})")
                q.task_done()
                continue

            target_table = "SEEDLING_TABLE" if kind == "seedling" else "DICTIONARY_TABLE"
            cur.executemany(
                f"""
                INSERT INTO {target_table}
                (Ticker_name, Date, Open, High, Low, Close, Volume, Volume_Again)
                VALUES (?,?,?,?,?,?,?,?)
                """,
                block_rows,
            )
            conn.commit()
            print(f"[WRITE] {sym}: wrote 53-row block into {target_table}")
            q.task_done()
    finally:
        for t in workers:
            t.join(timeout=5.0)
        conn.close()

    print(
        f"=== build_alpaca_intraday_dictionary COMPLETE ===\n"
        f"tested={shared['tested']} passing={shared['passing']} seedlings={shared['seedlings']}"
    )

if __name__ == "__main__":
    import time
    start_t = time.time()
    build_alpaca_intraday_dictionary()
    end_t = time.time()
    print(f"[TIMER] Program took {end_t - start_t:.2f} seconds.")
