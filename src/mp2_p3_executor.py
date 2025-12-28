"""
MP2 — Part 3: Daily group executor + candidate pruner

Purpose:
- Use MP2 Part 2's ranked groups (extra_big_groups / EBG) to generate today's tradeable ticker list.
- Run the selected group against the latest completed-day dictionary in Excel, intersect BUYs across algos,
  then prune candidates using recent daily history.

Reads:
- stocks_data.db: extra_big_groups (EBG), DICTIONARY_TABLE

Writes:
- stocks_data.db: live_candidates (raw intersection list, then overwritten with pruned scores)
- WB2 ControlSheet (FZ block): group header + raw candidates, then [PRUNED] output

Execution flow:
1) Select best group from EBG (score/pp ordering; can invert for testing).
2) Paste each algo's parameters into a non-colliding variation slot (1..8).
3) Cycle dictionary blocks through Excel, harvest BUYs per algo slot, intersect common tickers.
4) Save intersection to live_candidates and print to WB2.
5) Prune candidates using Alpaca daily bars, score + rank to top N, rewrite WB2 and DB.

Notes:
- This script assumes the Excel workbook is the execution engine (cycler/reset macros and ControlSheet geometry).
- Real trade execution occurs after this stage (Real Trader is downstream).
"""

import re
import time
import math
import json
import sqlite3
import pandas as pd
import xlwings as xw

# mp2 p3 start

EXCEL_FILE = r"<REDACTED_PATH>/WB1.xlsm"
WB2_FILE   = r"<REDACTED_PATH>/WB2.xlsm"
DB_FILE    = r"<REDACTED_PATH>/stocks_data.db"

wb2 = xw.Book(WB2_FILE)
wb2_sheet = wb2.sheets["ControlSheet"]

# toggles
USE_LOWEST_SCORE_GROUP = False 
MIN_COMMON_TICKERS     = 4     
CONTROL_SHEET_NAME = "ControlSheet"
WB2_GROUP_START_CELL = "FZ6"

TRIGGER_MACRO_CELL = "AA1"
RESET_MACRO_CELL = "AC1"
CYCLER_MACRO_CELL = "S1" 

PARAM_START_COL = 56
PP_COL = 48
BASE_START_ROW = 55

START_DATA_ROW = 55
COL_I = 9   # Name
COL_J = 10  # Ticker

VAR_BLOCK_START = {
    1: 11,  # K:P
    2: 17,  # Q:V
    3: 23,  # W:AB
    4: 29,  # AC:AH
    5: 35,  # AI:AN
    6: 41,  # AO:AT
    7: 47,  # AU:AZ
    8: 53,  # BA:BF
}

PRICE_START_COL = 59  # BG
PRICE_WIDTH     = 6   # BG..BL

SHEET_MAP = {
    1: "1 MACD", 2: "2 EMA", 3: "3 RSI", 4: "4 Breakout", 5: "5 ADX",
    6: "6 Volatility Measure", 7: "7 SMA", 8: "8 Bollinger Bands",
    9: "9 EMA & MACD", 10: "10 RSI Bollinger"
}

def _sheet_for_algo(algo_num:int):
    return SHEET_MAP.get(algo_num) or SHEET_MAP.get(str(algo_num))

def paste_group_parameters_for_slot(sheet, algo_num:int, slot:int, param_list):
    base_row = get_algo_base_row(algo_num) + (slot - 1)
    for j, val in enumerate(param_list):
        sheet.range((base_row, PARAM_START_COL + j)).value = val

_TICKER_RE = re.compile(r'^[A-Z0-9.\-]{1,6}$')

def _normalize_name_ticker(name, ticker):
    name_s = "" if name is None else str(name).strip()
    tck_s  = "" if ticker is None else str(ticker).strip()

    if (" " in tck_s) and (" " not in name_s) and (1 <= len(name_s) <= 6):
        name_s, tck_s = tck_s, name_s
    tck_s = tck_s.upper()
    return name_s, tck_s
  
def _looks_like_ticker(s: str) -> bool:
    if not s:
        return False
    if " " in s:
        return False
    return bool(_TICKER_RE.match(s))

def read_algo_buys_for_slot(wb, algo_num:int, slot:int):
    sheet_name = _sheet_for_algo(algo_num)
    if not sheet_name or slot not in VAR_BLOCK_START:
        return {}, 0

    sh = wb.sheets[sheet_name]
    start_col    = VAR_BLOCK_START[slot]
    block_width  = 6
    decision_col = start_col

    tickers_col = sh.range((START_DATA_ROW, COL_J), (1056, COL_J)).value
    if isinstance(tickers_col, list):
        last_idx = 0
        for i, v in enumerate(tickers_col):
            if not v:
                break
            last_idx = i + 1
        n_rows = last_idx
    else:
        n_rows = 0
    if n_rows == 0:
        return {}, 0

    top = START_DATA_ROW
    bot = START_DATA_ROW + n_rows - 1

    names     = sh.range((top, COL_I), (bot, COL_I)).value
    tickers   = sh.range((top, COL_J), (bot, COL_J)).value
    decisions = sh.range((top, decision_col), (bot, decision_col)).value
    blocks    = sh.range((top, start_col), (bot, start_col + block_width - 1)).value
    prices    = sh.range((top, PRICE_START_COL), (bot, PRICE_START_COL + PRICE_WIDTH - 1)).value

    if n_rows == 1:
        names     = [names]
        tickers   = [tickers]
        decisions = [decisions]
        blocks    = [blocks]
        prices    = [prices]

    buys_unique = {}
    raw_count   = 0

    for i in range(n_rows):
        t_raw = tickers[i]
        if not t_raw:
            break
        dec = str(decisions[i]).strip().upper() if decisions[i] is not None else ""
        if dec != "BUY":
            continue
        raw_count += 1
        name_i = names[i]
        block6 = blocks[i]
        price6 = prices[i]
        name_n, tck_n = _normalize_name_ticker(name_i, t_raw)
        if not _looks_like_ticker(tck_n):
            continue

        buys_unique[tck_n] = {
            "name":  name_n,
            "block": block6,
            "price": price6,
        }

    return buys_unique, raw_count

def extract_group_parameters(group_id):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    order_dir = "ASC" if USE_LOWEST_SCORE_GROUP else "DESC"

    cursor.execute(f"""
        SELECT identifier, pp, score, parameters, algo_ticker_counts
        FROM extra_big_groups
        ORDER BY score {order_dir}, pp {order_dir}
    """)

    all_groups = cursor.fetchall()
    conn.close()

    if group_id >= len(all_groups):
        return None
    identifier, pp, score, param_str, algo_counts = all_groups[group_id]
    param_map = json.loads(param_str) if param_str else {}

    algo_list = []
    param_chunks = []

    for key, vals in param_map.items():
        m = re.match(r"No\.(\d+)", str(key))
        if m:
            algo_list.append(int(m.group(1)))
            param_chunks.append(vals)

    if not algo_list:
        nums = [int(n) for n in re.findall(r"No\.(\d+):", algo_counts or "")]
        if nums:
            ordered_params = list(param_map.values()) 
            take = min(len(nums), len(ordered_params))
            algo_list = nums[:take]
            param_chunks = ordered_params[:take]

    return {
        "identifier": identifier,
        "pp": pp,
        "score": score,
        "algorithms": algo_list,
        "parameters": param_chunks
    }

def get_algo_base_row(algo_num):
    return 4 + (int(algo_num) - 1) * 11

def paste_group_parameters(sheet, group_data):
    for algo_entry, param_list in zip(group_data["algorithms"], group_data["parameters"]):
        algo_num = int(algo_entry)
        base_row = get_algo_base_row(algo_num)

        for j, val in enumerate(param_list):
            sheet.range((base_row, PARAM_START_COL + j)).value = val

# trigger a macro in Excel
def trigger_macro(sheet, cell, delay=1.0):
    """
    Triggers the Excel macro tied to a specific control cell by
    writing '100' into that cell.
    """
    sheet.range(cell).value = 100
    sheet.book.save()
    time.sleep(delay)
  
def collect_group_buys(*args, **kwargs):
    return []

def run_mp2_part3_group_executor_real():
    wb = xw.Book(EXCEL_FILE)
    control_sheet = wb.sheets[CONTROL_SHEET_NAME]

    trigger_macro(control_sheet, TRIGGER_MACRO_CELL, delay=15)

    rows_per_dataset = 53
    batch_size    = int(control_sheet.range("AG4").value)
    result_limit  = int(control_sheet.range("AG3").value)  
    batches_per_cycle = max(1, result_limit // max(1, batch_size))

    conn = sqlite3.connect(DB_FILE)
    cur  = conn.cursor()

    cur.execute("DROP TABLE IF EXISTS live_candidates")
    cur.execute("""
        CREATE TABLE live_candidates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT,
            group_identifier TEXT,
            ticker TEXT,
            name TEXT,
            raw_score REAL,
            rank_score INTEGER
        )
    """)

    conn.commit()

    cur.execute("SELECT COUNT(*) FROM DICTIONARY_TABLE")
    total_rows = cur.fetchone()[0]
    total_datasets = total_rows // rows_per_dataset

    group_index = 0
    CYCLER = CYCLER_MACRO_CELL 

    while True: 
        control_sheet.range("Q1").value = BASE_START_ROW
        group_data = extract_group_parameters(group_index)
        if not group_data:
            print("No more groups.")
            break

        print(f" Group {group_index + 1}: {group_data['identifier']}")

        algo_slot_counter = {}   
        slot_map = {}   
        for algo_num in group_data["algorithms"]:
            nxt = algo_slot_counter.get(algo_num, 1)
            if nxt > 8:
                raise RuntimeError(f"Algo {algo_num} needs more than 8 slots.")
            slot_map[algo_num] = nxt
            algo_slot_counter[algo_num] = nxt + 1

        for algo_num, plist in zip(group_data["algorithms"], group_data["parameters"]):
            paste_group_parameters_for_slot(control_sheet, int(algo_num), slot_map[int(algo_num)], plist)

        control_sheet.range("K1").value = batch_size

        per_algo_buys = {int(a): {} for a in group_data["algorithms"]}
        raw_counts    = {int(a): 0   for a in group_data["algorithms"]}

        dataset_index = 0
        window_idx = 1

        while dataset_index < total_datasets:
            result_row = BASE_START_ROW
            remaining_total = total_datasets - dataset_index
            window_target = min(result_limit, remaining_total)
            batches_this_window = math.ceil(window_target / batch_size)

            print(f"\n Window {window_idx}: target={window_target}, batches={batches_this_window}")

            for _b in range(batches_this_window):
                if dataset_index >= total_datasets:
                    break
                remaining_total = total_datasets - dataset_index
                limit_datasets = min(batch_size, remaining_total)
                if limit_datasets != batch_size:
                    control_sheet.range("K1").value = limit_datasets

                offset = dataset_index * rows_per_dataset
                limit  = limit_datasets * rows_per_dataset
                print(f" offset={offset} datasets={limit_datasets} rows={limit}")

                cur.execute(f"SELECT * FROM DICTIONARY_TABLE LIMIT {limit} OFFSET {offset}")
                rows = cur.fetchall()
                if not rows:
                    break

                df = pd.DataFrame(rows)
                app = wb.app
                _u,_a = app.screen_updating, app.display_alerts
                app.screen_updating, app.display_alerts = False, False
                try:
                    control_sheet.range("A:I").clear_contents()
                    control_sheet.range("A1").value = df.values
                finally:
                    app.screen_updating, app.display_alerts = _u,_a

                control_sheet.range("Q1").value = result_row
                wb.save()

                # trigger cycler and wait for completion
                original_o1 = control_sheet.range("O1").value
                control_sheet.range(CYCLER).value = 100

                start_time = time.time()
                while True:
                    if control_sheet.range("O1").value != original_o1:
                        break
                    if time.time() - start_time > 900:
                        raise TimeoutError("Macro timed out.")
                    time.sleep(0.5)

                control_sheet.range("O1").value = 0
                wb.save()

                result_row    += limit_datasets
                dataset_index += limit_datasets

            total_for_window = 0
            for algo_num in group_data["algorithms"]:
                a = int(algo_num)
                slot = slot_map[a]
                buys_dict, raw = read_algo_buys_for_slot(wb, a, slot)
                raw_counts[a] += int(raw or 0)
                agg = per_algo_buys[a]
                for tck, rec in buys_dict.items():
                    agg[tck] = rec
                total_for_window += len(buys_dict)
            print(f"Window {window_idx} → unique buys (sum over algos): {total_for_window}")

            # reset helper cell
            trigger_macro(control_sheet, "AE1", delay=1.5)
            control_sheet.range("Q1").value = BASE_START_ROW
            wb.save()
            window_idx += 1

        common = None
        for algo_num in group_data["algorithms"]:
            tset = set(per_algo_buys[int(algo_num)].keys())
            common = tset if common is None else (common & tset)
        common = sorted(list(common)) if common else []
      
        for algo_num in group_data["algorithms"]:
            a = int(algo_num)
            start_col = VAR_BLOCK_START[slot_map[a]]
            end_col   = start_col + 5
            print(f"  • Algo {a} slot {slot_map[a]} block {start_col}-{end_col} → raw={raw_counts[a]} unique={len(per_algo_buys[a])}")

        if not common:
            print("No common tickers for this group. Trying next group...")
            group_index += 1
            continue

        if len(common) < MIN_COMMON_TICKERS:
            print(
                f"Only {len(common)} common tickers (< {MIN_COMMON_TICKERS}) "
                "for this group. Trying next group..."
            )
            group_index += 1
            continue

        ts_str = time.strftime("%Y-%m-%d")
        ref_algo = int(group_data["algorithms"][0])
        for tck in common:
            rec = per_algo_buys[ref_algo].get(tck, {})

            ticker = str(tck).strip().upper()
            name   = str(rec.get("name") or "").strip()

            cur.execute(
                "INSERT INTO live_candidates (ts, group_identifier, ticker, name) "
                "VALUES (?, ?, ?, ?)",
                (ts_str, group_data["identifier"], ticker, name)
            )

        conn.commit()
        print(f"Saved {len(common)} live candidates to DB.")

        wb2_sheet.range("M3").value = 100
        time.sleep(1.0)
        paste_row = 6
        paste_col = 182  # FZ
        wb2_sheet.range((paste_row, paste_col)).value = group_data["identifier"]
        wb2_sheet.range((paste_row, paste_col + 1)).value = group_data["pp"]
        wb2_sheet.range((paste_row, paste_col + 2)).value = group_data["score"]
        paste_row += 2

        headers = ["Ticker","Name","Decision","Date","Open","High","Low","Close","Volume"]
        wb2_sheet.range((paste_row, paste_col)).value = headers
        paste_row += 1

        ref_algo = int(group_data["algorithms"][0])
        rows_out = []
        for tck in common:
            rec = per_algo_buys[ref_algo][tck]
            name = rec["name"]
            price6 = rec["price"][:PRICE_WIDTH] if rec["price"] else [None]*PRICE_WIDTH
            rows_out.append([tck, name, "BUY"] + price6)

        app = wb2.app
        _u,_a = app.screen_updating, app.display_alerts
        app.screen_updating, app.display_alerts = False, False
        try:
            if rows_out:
                wb2_sheet.range((paste_row, paste_col)).value = rows_out
        finally:
            app.screen_updating, app.display_alerts = _u,_a
        wb2.save()
        print("Group buys pasted into WB2.")
        break
    trigger_macro(control_sheet, RESET_MACRO_CELL, delay=10)

    try: cur.close()
    except: pass
    try: conn.close()
    except: pass
      
# mp2 p2 end

from datetime import datetime, timezone, timedelta
from typing import Dict, Any, List, Tuple

from alpaca_trade_api import REST
from alpaca_trade_api.rest import TimeFrame

# Alpaca config
ALPACA_API_KEY = ""
ALPACA_API_SECRET = ""
ALPACA_BASE_URL = ""
ALPACA_DATA_FEED  = "sip"   # "sip" or "iex"

def _get_alpaca_rest() -> REST:
    return REST(ALPACA_API_KEY, ALPACA_API_SECRET, ALPACA_BASE_URL)

# Config knobs
MAX_FINAL_TICKERS      = 20    
HISTORY_DAYS           = 30   
MIN_DAYS_REQUIRED      = 20   
MAX_DRAWDOWN_FROM_OPEN = 0.03
PROFIT_RATE_MIN        = 0.70 
SAFE_HIT_RATE_MIN      = 0.30  
MAX_LOSS_FLOOR         = -0.07 
Z_WILSON               = 1.0   
CALENDAR_LOOKBACK_DAYS = 60   

def _pick_target_pct_from_open(price: float) -> float:
    """
    Mirror your RealTrader buckets but use OPEN as the reference.
    """
    if price >= 10.0 and price < 30.0:
        return 0.02
    if price >= 30.0 and price < 60.0:
        return 0.0275
    if price >= 60.0:
        return 0.015
    return 0.02

def _classify_day(open_p: float,
                  high_p: float,
                  low_p: float,
                  close_p: float,
                  target_pct: float,
                  max_dd: float) -> Tuple[str, float]:

    if open_p <= 0:
        return "skip", 0.0

    drawdown = (low_p - open_p) / open_p 
    move_high = (high_p - open_p) / open_p
    move_close = (close_p - open_p) / open_p

    safe_dd = drawdown >= -max_dd
    target_price = open_p * (1.0 + target_pct)
    hit_target = high_p >= target_price

    if hit_target and safe_dd:
        return "safe_hit", target_pct
    if (not hit_target) and (move_close > 0.0) and safe_dd:
        return "green_close", move_close
    return "loss", move_close

def _wilson_lower_bound(p: float, n: int, z: float = Z_WILSON) -> float:
    if n <= 0:
        return 0.0
    p = max(0.0, min(1.0, p))
    z2 = z * z
    denom = 1.0 + z2 / n
    centre = p + z2 / (2.0 * n)
    adj = z * math.sqrt((p * (1.0 - p) + z2 / (4.0 * n)) / n)
    lb = (centre - adj) / denom
    return max(0.0, lb)

# core pruner
def run_live_candidates_pruner(
    max_final: int = MAX_FINAL_TICKERS,
    history_days: int = HISTORY_DAYS,
    max_dd: float = MAX_DRAWDOWN_FROM_OPEN,
    min_days_required: int = MIN_DAYS_REQUIRED,
    profit_rate_min: float = PROFIT_RATE_MIN,
    safe_hit_min: float = SAFE_HIT_RATE_MIN,
    max_loss_floor: float = MAX_LOSS_FLOOR,
    z_wilson: float = Z_WILSON,
) -> None:
  
    print("=== Live Candidates Historical Pruner START ===")

    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    cur.execute("SELECT MAX(ts) FROM live_candidates")
    row = cur.fetchone()
    if not row or row[0] is None:
        print("No entries in live_candidates; nothing to prune.")
        cur.close()
        conn.close()
        return

    latest_ts = row[0]
    print(f"Using latest ts from live_candidates: {latest_ts}")
    cur.execute(
        "SELECT ticker, group_identifier, name FROM live_candidates WHERE ts = ?",
        (latest_ts,)
    )
    rows = cur.fetchall()
    ...

    ticker_to_group: Dict[str, str] = {}
    ticker_to_name: Dict[str, str] = {}
    for t, gid, name in rows:
        if not t:
            continue
        t_clean = str(t).strip().upper()
        if not t_clean:
            continue
        ticker_to_group[t_clean] = gid
        ticker_to_name[t_clean] = str(name).strip() if name is not None else ""

    all_tickers = sorted(ticker_to_group.keys())
    print(f"Found {len(all_tickers)} unique tickers in live_candidates.")

    if not all_tickers:
        print("No valid tickers parsed from live_candidates.")
        cur.close()
        conn.close()
        return

    group_identifier = list(ticker_to_group.values())[0]
    print(f"Group identifier (from live_candidates): {group_identifier}")

    api = _get_alpaca_rest()
    end_dt = datetime.now(timezone.utc)
    start_dt = end_dt - timedelta(days=CALENDAR_LOOKBACK_DAYS)

    ticker_metrics: List[Dict[str, Any]] = []

    for idx, symbol in enumerate(all_tickers, start=1):
        print(f"[{idx}/{len(all_tickers)}] Fetching history for {symbol}...")
        try:
            bars = api.get_bars(
                symbol,
                TimeFrame.Day,
                start_dt.isoformat(),
                end_dt.isoformat(),
                adjustment="raw",
                feed=ALPACA_DATA_FEED,
            ).df
        except Exception as e:
            print(f"  -> FAILED: Alpaca get_bars error for {symbol}: {e}")
            continue

        if bars is None or bars.empty:
            print(f"  -> SKIP: no bars returned for {symbol}.")
            continue

        bars = bars.sort_index(ascending=False)
        window = bars.iloc[:history_days]
        if window.shape[0] < min_days_required:
            print(
                f"  -> SKIP: only {window.shape[0]} days for {symbol}, "
                f"need at least {min_days_required}."
            )
            continue

        n_days = 0
        n_safe_hit = 0
        n_green = 0
        n_loss = 0

        win_returns: List[float] = []
        loss_returns: List[float] = []
        ranges: List[float] = []

        latest_bar = window.iloc[0]
        latest_date_ts = window.index[0]

        for ts, bar in window.iterrows():
            try:
                o = float(bar["open"])
                h = float(bar["high"])
                l = float(bar["low"])
                c = float(bar["close"])
            except Exception:
                continue

            if o <= 0:
                continue

            target_pct = _pick_target_pct_from_open(o)
            label, r = _classify_day(o, h, l, c, target_pct, max_dd)

            if label == "skip":
                continue

            n_days += 1

            # intraday range
            day_range = (h - l) / o if o > 0 else 0.0
            ranges.append(day_range)

            if label == "safe_hit":
                n_safe_hit += 1
                n_green += 1   
                win_returns.append(r)
            elif label == "green_close":
                n_green += 1
                win_returns.append(r)
            else:  # loss
                n_loss += 1
                loss_returns.append(r)

        if n_days < min_days_required:
            print(
                f"  -> SKIP: after cleaning, only {n_days} valid days for {symbol}, "
                f"need {min_days_required}."
            )
            continue

        p_profit = n_green / float(n_days)
        p_safe_hit = n_safe_hit / float(n_days)
        p_loss = n_loss / float(n_days)

        avg_profit_win = sum(win_returns) / len(win_returns) if win_returns else 0.0
        avg_loss_abs = (
            -sum(loss_returns) / len(loss_returns) if loss_returns else 0.0
        ) 
        max_loss = min(loss_returns) if loss_returns else 0.0

        avg_range = sum(ranges) / len(ranges) if ranges else 0.0

        try:
            latest_open = float(latest_bar["open"])
        except Exception:
            latest_open = o 
        target_pct_avg = _pick_target_pct_from_open(latest_open)

        wl = _wilson_lower_bound(p_profit, n_days, z=z_wilson)

        # guardrails
        passes = True

        if p_profit < profit_rate_min:
            passes = False
        if p_safe_hit < safe_hit_min:
            passes = False
        if max_loss < max_loss_floor:
            passes = False

        if target_pct_avg > 0:
            if avg_range < target_pct_avg * 1.2:
                passes = False
            if avg_range > target_pct_avg * 5.0:
                passes = False

        if target_pct_avg <= 0:
            target_pct_avg = 0.01

        payoff_ratio = avg_profit_win / target_pct_avg if target_pct_avg > 0 else 0.0
        loss_ratio = avg_loss_abs / target_pct_avg if target_pct_avg > 0 else 0.0

        score = (
            100.0 * (
                0.6 * wl +
                0.2 * p_safe_hit +
                0.2 * payoff_ratio
            )
            - 50.0 * loss_ratio
            - 30.0 * max(0.0, -max_loss - 0.05)
            - 10.0 * p_loss
        )

        if hasattr(latest_date_ts, "to_pydatetime"):
            latest_date_dt = latest_date_ts.to_pydatetime()
        else:
            latest_date_dt = latest_date_ts
        latest_date_str = latest_date_dt.strftime("%Y-%m-%d")

        latest_open = float(latest_bar["open"])
        latest_high = float(latest_bar["high"])
        latest_low = float(latest_bar["low"])
        latest_close = float(latest_bar["close"])
        latest_vol = float(latest_bar["volume"])

        ticker_metrics.append(
            {
                "ticker": symbol,
                "group_identifier": group_identifier,
                "n_days": n_days,
                "n_profit": n_green,
                "n_safe_hit": n_safe_hit,
                "n_loss": n_loss,
                "p_profit": p_profit,
                "p_safe_hit": p_safe_hit,
                "p_loss": p_loss,
                "avg_profit_win": avg_profit_win,
                "avg_loss_abs": avg_loss_abs,
                "max_loss": max_loss,
                "avg_range": avg_range,
                "target_pct_avg": target_pct_avg,
                "wilson_lb": wl,
                "score": score,
                "passes_guardrails": passes,
                "latest": {
                    "date": latest_date_str,
                    "open": latest_open,
                    "high": latest_high,
                    "low": latest_low,
                    "close": latest_close,
                    "volume": latest_vol,
                },
            }
        )

        print(
            f"  -> {symbol}: n_days={n_days}, p_profit={p_profit:.3f}, "
            f"p_safe_hit={p_safe_hit:.3f}, wl={wl:.3f}, score={score:.2f}, "
            f"passes_guardrails={passes}"
        )

    cur.close()
    conn.close()

    if not ticker_metrics:
        print("No tickers with usable history; pruner stops here.")
        print("=== Live Candidates Historical Pruner END (no metrics) ===")
        return

    total_candidates = len(ticker_metrics)
    good = [m for m in ticker_metrics if m["passes_guardrails"]]
    bad  = [m for m in ticker_metrics if not m["passes_guardrails"]]

    print(
        f"{len(good)} tickers passed guardrails out of {total_candidates} "
        f"({len(bad)} did not)."
    )

    good_sorted = sorted(good, key=lambda m: m["score"], reverse=True)
    bad_sorted  = sorted(bad,  key=lambda m: m["score"], reverse=True)

    top: List[Dict[str, Any]] = []

    # soft guardrails
    for m in good_sorted:
        if len(top) >= max_final:
            break
        top.append(m)

    if len(top) < max_final:
        for m in bad_sorted:
            if len(top) >= max_final:
                break
            top.append(m)

    if not top:
        print("No tickers selected after scoring; pruner stops here.")
        print("=== Live Candidates Historical Pruner END (no selection) ===")
        return
    if len(top) < max_final and total_candidates >= max_final:
        print(
            f"Warning: only {len(top)} tickers selected out of requested {max_final}; "
            "guardrails and data quality were very restrictive."
        )

    top = sorted(top, key=lambda m: m["score"], reverse=True)

    # 1–20 rank score
    max_rank_score = 20
    scores = [m["score"] for m in top]
    max_score = max(scores)
    min_score = min(scores)
    span = max_score - min_score

    if span <= 1e-9:
        for m in top:
            m["rank_score"] = max_rank_score
    else:
        gamma = 2.0
        for m in top:
            z = (m["score"] - min_score) / span 
            z = max(0.0, min(1.0, z))
            rank = 1 + int(round((z ** gamma) * (max_rank_score - 1)))
            m["rank_score"] = rank

    print(f"Selected top {len(top)} tickers (max requested = {max_final}):")
    for m in top:
        print(
            f"  {m['ticker']}: score={m['score']:.2f}, "
            f"rank_score={m['rank_score']}, "
            f"p_profit={m['p_profit']:.3f}, p_safe_hit={m['p_safe_hit']:.3f}, "
            f"max_loss={m['max_loss']:.3f}"
        )
    try:
        wb2_local = xw.Book(WB2_FILE)

        ws2 = wb2_local.sheets["ControlSheet"]
    except Exception as e:
        print(f"Could not open WB2 for output: {e}")
        print("=== Live Candidates Historical Pruner END (no Excel write) ===")
        return

    # Trigger macro to clear that area in Excel
    try:
        ws2.range("M3").value = 100
        time.sleep(1.0)
    except Exception as e:
        print(f"Warning: could not trigger clear macro M3: {e}")

    paste_row = 6
    paste_col = 182  # FZ

    if top:
        gid_label = top[0]["group_identifier"]
    else:
        gid_label = group_identifier

    ws2.range((paste_row, paste_col)).value = f"{gid_label} [PRUNED]"
    ws2.range((paste_row, paste_col + 1)).value = None
    ws2.range((paste_row, paste_col + 2)).value = None 
    paste_row += 2
    headers = [
        "Ticker",
        "Name",
        "Decision",
        "Date",
        "Open",
        "High",
        "Low",
        "Close",
        "Volume",
        "Score",
        "RankScore",
        "ProfitRate",
        "SafeHitRate",
        "AvgWin%",
        "AvgLoss%",
        "MaxLoss%",
    ]

    ws2.range((paste_row, paste_col)).value = headers
    paste_row += 1
    rows_out = []
    for m in top:
        latest = m["latest"]
        tck = m["ticker"]
        name_for_ticker = ticker_to_name.get(tck, tck)

        rows_out.append(
            [
                tck,
                name_for_ticker,
                "BUY",
                latest["date"],
                latest["open"],
                latest["high"],
                latest["low"],
                latest["close"],
                latest["volume"],
                round(m["score"], 2),
                int(m.get("rank_score", 0)),
                round(m["p_profit"] * 100.0, 2),
                round(m["p_safe_hit"] * 100.0, 2),
                round(m["avg_profit_win"] * 100.0, 2),
                round(m["avg_loss_abs"] * 100.0, 2),
                round(m["max_loss"] * 100.0, 2),
            ]
        )


    if rows_out:
        ws2.range((paste_row, paste_col)).value = rows_out
        wb2_local.save()
        print("Excel WB2 updated with pruned top tickers.")
    else:
        print("No top rows to write to Excel.")
    conn2 = sqlite3.connect(DB_FILE)
    cur2 = conn2.cursor()
    cur2.execute("DELETE FROM live_candidates")
    for m in top:
        tck = m["ticker"]
        name_val = ticker_to_name.get(tck, "")
        cur2.execute(
            """
            INSERT INTO live_candidates (
                ts,
                group_identifier,
                ticker,
                name,
                raw_score,
                rank_score
            )
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                latest_ts,
                m["group_identifier"],
                tck,
                name_val,
                float(m["score"]),
                int(m.get("rank_score", 0)),
            )
        )
    conn2.commit()
    cur2.close()
    conn2.close()
    print("live_candidates table overwritten with pruned tickers (with scores).")
    print("=== Live Candidates Historical Pruner END ===")


if __name__ == "__main__":
    run_mp2_part3_group_executor_real()
    run_live_candidates_pruner() 
