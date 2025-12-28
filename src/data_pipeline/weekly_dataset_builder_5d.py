"""
Weekly Dataset Builder (5-Day)

Purpose:
- Builds 5 rolling 31-trading-day windows (DICT_DAY1..DICT_DAY5) from Alpaca daily bars.
- Creates a stable dataset by intersecting tickers present in all 5 windows:
  DATASET_TICKERS + DS_DAY1..DS_DAY5.

Inputs:
- Alpaca daily bars (IEX or SIP feed).
- Local SQLite database file.

Outputs (SQLite):
- DICT_DAY1..DICT_DAY5: raw 53-row blocks per ticker per day-window.
- DATASET_TICKERS: intersection tickers + fixed ordering.
- DS_DAY1..DS_DAY5: final dataset blocks in consistent order.

Privacy:
- API keys, base URL, and DB path are redacted in the public repo.
"""

# Alpaca config / helpers

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
from zoneinfo import ZoneInfo 

NY_TZ = ZoneInfo("America/New_York")

# ========= GLOBAL TOGGLES =========

ALPACA_API_KEY    = ""
ALPACA_API_SECRET = ""
ALPACA_BASE_URL   = ""
DB_PATH           = ""

USE_SIP = True
LIMIT_TESTED = False
TESTED_CAP = 100

NUM_DAYS = 5
N_WORKERS = 4

BARS_REQUIRED = 31 + (NUM_DAYS - 1)

CALENDAR_LOOKBACK = max(65, BARS_REQUIRED + 30)

PRICE_MIN   = 10.0
PRICE_MAX   = 90.0
VOLUME_MIN  = 250_000
VOLUME_MAX  = 100_000_000

# ========= NUCLEAR NON-STOCK FILTER =========

NAME_EXCLUDE_KEYWORDS = [
    " etf ", " etn ", " fund", " trust", " income", " dividend",
    " municipal", " muni", " bond", " notes", " note", " preferred",
    " preference", " depositary", " depository", " portfolio",
    " index", " allocation", " target", " strategy", " capital",
    " real estate", " properties", " property", " reit",
    " closed-end", " closed end", " unit ", " units ",
    " warrant", " rights",
]

ETF_FAMILIES = (
    "SPDR", "ISHARES", "VANGUARD", "INVESCO",
    "PROSHARES", "GLOBAL X", "DIREXION"
)

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

def is_clean_common_stock(asset: dict) -> bool:
    # Stock universe filter
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

    # Only active
    if status != "active":
        return False

    # Drop OTC / pink sheets
    if exch == "OTC":
        return False

    # Only alphabetic tickers (no dots/hyphens)
    if not sym.isalpha():
        return False

    name_lower = " " + name.lower() + " "
    name_upper = name.upper()

    for kw in NAME_EXCLUDE_KEYWORDS:
        if kw in name_lower:
            return False

    if name_upper.startswith(ETF_FAMILIES):
        return False

    if len(sym) > 5 and ("FUND" in name_upper or "ETF" in name_upper or "TRUST" in name_upper):
        return False

    return True


def chunk_list(lst: List[Tuple[str, str]], n_chunks: int) -> List[List[Tuple[str, str]]]:
    if n_chunks <= 1 or len(lst) == 0:
        return [lst]
    chunk_size = (len(lst) + n_chunks - 1) // n_chunks
    return [lst[i:i + chunk_size] for i in range(0, len(lst), chunk_size)]

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
    block_rows.append(
        (
            use_name,
            "Date",
            "Open",
            "High",
            "Low",
            "Close",
            "Volume",
            "Volume_Again",
        )
    )

    # rows 2–32: 31 days of data (latest at top for this window)
    first_data_row = True
    for ts, row in window.iterrows():
        if hasattr(ts, "to_pydatetime"):
            ts_dt = ts.to_pydatetime()
        else:
            ts_dt = ts
        date_str = ts_dt.strftime("%a %b %d %Y %H:%M:%S %Z")

        o = float(row["open"])
        h = float(row["high"])
        l = float(row["low"])
        c = float(row["close"])
        v = int(row["volume"])

        ticker_col = sym if first_data_row else ""
        first_data_row = False

        block_rows.append(
            (
                ticker_col,
                date_str,
                f"{o:.2f}",
                f"{h:.2f}",
                f"{l:.2f}",
                f"{c:.2f}",
                str(v),
                str(v),
            )
        )

    while len(block_rows) < 53:
        block_rows.append(("", "", "", "", "", "", "", ""))

    return block_rows

# Table creation / dataset build

def _create_dict_and_ds_tables(conn: sqlite3.Connection, num_days: int):
    cur = conn.cursor()

    # DICT_DAY tables
    for i in range(1, num_days + 1):
        tbl = f"DICT_DAY{i}"
        cur.execute(f"DROP TABLE IF EXISTS {tbl}")
        cur.execute(
            f"""
            CREATE TABLE {tbl} (
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

    # DS_DAY tables
    for i in range(1, num_days + 1):
        ds_tbl = f"DS_DAY{i}"
        cur.execute(f"DROP TABLE IF EXISTS {ds_tbl}")
        cur.execute(
            f"""
            CREATE TABLE {ds_tbl} (
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

    # DATASET_TICKERS
    cur.execute("DROP TABLE IF EXISTS DATASET_TICKERS")
    cur.execute(
        """
        CREATE TABLE DATASET_TICKERS (
            ticker    TEXT PRIMARY KEY,
            order_idx INTEGER
        )
        """
    )

    conn.commit()

def _extract_block_ticker(blk: pd.DataFrame) -> str:
    ser = blk["Ticker_name"].fillna("").astype(str).str.strip()
    for t in ser:
        if t:
            return t
    return ""

def _build_datasets_from_dicts(conn: sqlite3.Connection, num_days: int, dataset_size: int):
    cur = conn.cursor()
    dict_tables = [f"DICT_DAY{i}" for i in range(1, num_days + 1)]

    # 1) Build per-day ticker sets
    sets = []
    for tbl in dict_tables:
        df = pd.read_sql_query(f"SELECT * FROM {tbl}", conn)
        tickers = set()
        total_rows = len(df)
        for start in range(0, total_rows, 53):
            blk = df.iloc[start:start + 53]
            if len(blk) < 53:
                continue
            t = _extract_block_ticker(blk)
            if t:
                tickers.add(t)
        sets.append(tickers)
        print(f"[{tbl}] tickers present (no extra filters): {len(tickers)}")

    if not sets:
        print("[WARN] No DICT_DAY tables had data; skipping DS build.")
        return

    inter = set.intersection(*sets)
    print(f"[INTERSECTION] tickers present in ALL {num_days} DICT_DAY tables: {len(inter)}")

    tickers_list = list(inter)
    random.shuffle(tickers_list)
    if dataset_size is not None:
        tickers_list = tickers_list[:dataset_size]

    if len(tickers_list) < dataset_size:
        print(f"[WARN] Intersection < {dataset_size}. Using {len(tickers_list)} tickers.")

    # 2) DATASET_TICKERS
    rows = [(t, i + 1) for i, t in enumerate(tickers_list)]
    cur.executemany("INSERT INTO DATASET_TICKERS(ticker, order_idx) VALUES (?,?)", rows)
    conn.commit()
    print("[DATASET_TICKERS] saved.")

    # 3) DS_DAY tables
    for day_idx, tbl in enumerate(dict_tables, start=1):
        ds_name = f"DS_DAY{day_idx}"
        df = pd.read_sql_query(f"SELECT * FROM {tbl}", conn)
      
        idx_map = {}
        total_rows = len(df)
        for start in range(0, total_rows, 53):
            blk = df.iloc[start:start + 53]
            if len(blk) < 53:
                continue
            t = _extract_block_ticker(blk)
            if t and t not in idx_map:
                idx_map[t] = start

        written_blocks = 0
        for t in tickers_list:
            start = idx_map.get(t)
            if start is None:
                continue
            blk = df.iloc[start:start + 53]
            if len(blk) == 53:
                blk.to_sql(ds_name, conn, if_exists="append", index=False)
                written_blocks += 1

        print(f"[{ds_name}] wrote {written_blocks} blocks ({written_blocks * 53} rows).")

    conn.commit()


def build_alpaca_5day_dicts_and_datasets(
    db_path: str = DB_PATH,
    dataset_size: int = 1000,
):
  
    data_feed = "sip" if USE_SIP else "iex"
    print(
        f"=== build_alpaca_5day_dicts_and_datasets START ===\n"
        f"[CFG] feed={data_feed}, NUM_DAYS={NUM_DAYS}, "
        f"LIMIT_TESTED={LIMIT_TESTED}, TESTED_CAP={TESTED_CAP}"
    )

    conn = sqlite3.connect(db_path)
    _create_dict_and_ds_tables(conn, NUM_DAYS)
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

    # Time range
    end_dt = datetime.now(timezone.utc)
    start_dt = end_dt - timedelta(days=CALENDAR_LOOKBACK)

    shared = {
        "tested": 0,
        "passing": 0,
        "saved": 0,
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
                        f"passing={shared['passing']} saved={shared['saved']}"
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
            bars = drop_incomplete_today_daily_bar(bars)  # NEW

            if len(bars) < BARS_REQUIRED:
                print(f"[FAIL] {sym}: LESS_THAN_{BARS_REQUIRED}_BARS ({len(bars)})")
                continue

            latest = bars.iloc[0]
            try:
                last_close = float(latest["close"])
                last_vol = float(latest["volume"])
            except Exception:
                print(f"[FAIL] {sym}: BAD_LATEST_BAR")
                continue

            price_ok = (PRICE_MIN < last_close < PRICE_MAX)
            vol_ok = (VOLUME_MIN <= last_vol <= VOLUME_MAX)

            if not price_ok or not vol_ok:
                print(
                    f"[FAIL] {sym}: PRICE_VOL_FILTER_FAIL "
                    f"(close={last_close}, vol={last_vol})"
                )
                continue

            with progress_lock:
                shared["passing"] += 1

            q.put((sym, name, bars))

    # Launch workers
    slices = chunk_list(universe_symbols, N_WORKERS)
    workers: List[threading.Thread] = []
    for i, sym_slice in enumerate(slices):
        if not sym_slice:
            continue
        t = threading.Thread(
            target=worker,
            args=(sym_slice,),
            name=f"five_day_worker_{i+1}",
            daemon=True,
        )
        t.start()
        workers.append(t)

    # Writer: build DICT_DAY blocks
    try:
        while True:
            try:
                sym, name, bars = q.get(timeout=1.0)
            except Empty:
                if not any(t.is_alive() for t in workers):
                    break
                continue

            with progress_lock:
                shared["saved"] += 1

            for day_idx in range(1, NUM_DAYS + 1):
                tbl = f"DICT_DAY{day_idx}"
                offset = NUM_DAYS - day_idx 
                window = bars.iloc[offset: offset + 31]
                if len(window) < 31:
                    continue
                block_rows = _make_53row_block(sym, name, window)
                if len(block_rows) != 53:
                    print(f"[WARN] {sym}/{tbl}: block size != 53 ({len(block_rows)})")
                    continue
                cur.executemany(
                    f"""
                    INSERT INTO {tbl}
                    (Ticker_name, Date, Open, High, Low, Close, Volume, Volume_Again)
                    VALUES (?,?,?,?,?,?,?,?)
                    """,
                    block_rows,
                )
            conn.commit()
            print(f"[WRITE] {sym}: wrote 53-row block into DICT_DAY1..{NUM_DAYS}")
            q.task_done()
    finally:
        for t in workers:
            t.join(timeout=5.0)

    print(
        f"=== 5-DAY BLOCK BUILD DONE ===\n"
        f"tested={shared['tested']} passing={shared['passing']} saved={shared['saved']}"
    )

    _build_datasets_from_dicts(conn, NUM_DAYS, dataset_size)

    conn.close()
    print("=== build_alpaca_5day_dicts_and_datasets COMPLETE ===")

# ALPACA DICTIONARY FILLER END

if __name__ == "__main__":
    import time
    start_5day = time.time()
    build_alpaca_5day_dicts_and_datasets()
    end_5day = time.time()
    print(f"[TIMER] Program took {end_5day - start_5day:.2f} seconds.")
