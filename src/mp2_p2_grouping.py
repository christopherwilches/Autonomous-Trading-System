"""
MP2 — Part 2: Group Synergy Combination Tester (threaded + bounded)

Loads pruned_* variations (5-day buy lists) into RAM maps, then exhaustively constructs
multi-algorithm groups by intersecting “consensus tickers” per day (tickers all members bought).
Groups are scored using pooled PP + conservative bounds + stability + coverage + repeat-control,
and the strongest size-3/size-4 candidates are retained, persisted, and exported to WB2.

Input: pruned_* tables + DS_DAY1..DS_DAY5 baseline tables (SQLite), WB2 export workbook
Output: best groups (SIZE3/SIZE4) stored in extra_big_groups + exported to WB2, plus algo_run_stats
"""

import math
import threading
import time
from collections import defaultdict
import sys
sys.stdout.reconfigure(line_buffering=True)
from threading import Event
class ThreadTimeout(Exception):
    pass
topk_adds_lock = threading.RLock()
topk_adds_per_thread = defaultdict(lambda: {3: 0, 4: 0})

import os
import json
import psutil
import concurrent.futures
import sqlite3
import xlwings as xw
thread_local_data = threading.local()
heartbeat_stop = threading.Event()
DB_FILE = r"C:/REDACTED/stocks_data.db"
HEALTH_LOG_PATH = r"C:/REDACTED/health_mp2_p2.txt"
WB2_FILE = r"C:/REDACTED/Combination Grouper MME3.xls.xlsm"

HEALTH_INTERVAL_SEC = 600

from statistics import stdev
import statistics

# Group-size counters
group_size_counter = defaultdict(int) 
top_size_counter = defaultdict(int)  
TOTAL_WALL_BUDGET = 86400 
_base_per_task_limit = 0.0 
_delta_per_task = 0.0 
_finished_tasks = 0
tested_groups_total = 0
tested_groups_lock = threading.Lock()

_limit_lock = threading.RLock()

_global_start = time.time()
_target_end = _global_start + TOTAL_WALL_BUDGET    
_global_end  = _global_start + (TOTAL_WALL_BUDGET * 1.10) 

def _to_float_safe(x):
    try:
        return float(x)
    except Exception:
        return None

def _profit_from_open_close(r):
    o = _to_float_safe(r.get("open"))
    c = _to_float_safe(r.get("close"))
    if o is None or c is None:
        return 0
    return 1 if (c - o) > 0 else 0

def _slim_stock_dict_from_buys(s, day_idx):
    return {
        "ticker": s.get("ticker"),
        "name":   s.get("name"),
        "open":   s.get("open"),
        "close":  s.get("close"),
        "day":    day_idx,
        "is_profit": _profit_from_open_close(s),
    }

def _is_profit_row(r):
    try:
        return 1 if int(r.get("is_profit", 0)) == 1 else 0
    except Exception:
        return 0

_algo_stats = {
    # algo_idx: {
    #   "threads": 0,
    #   "timeouts": 0,
    #   "runtime_sec": 0.0,
    #   "sum_pp": 0.0,          # averages over starters (variations)
    #   "sum_tickers": 0.0      # averages over starters (variations)
    # }
}
_algo_stats_lock = threading.RLock()


def run_mp2_general_group_tester():
    global group_size_counter, top_size_counter
    dedupe_count = 0
    group_size_counter.clear()
    top_size_counter.clear()
    global tested_groups_total
    with tested_groups_lock:
        tested_groups_total = 0

    with _algo_stats_lock:
        _algo_stats.clear()
        for _ai in range(10): 
            _algo_stats[_ai] = {
                "threads": 0,
                "timeouts": 0,
                "runtime_sec": 0.0,
                "sum_pp": 0.0,
                "sum_tickers": 0.0,
            }


    global _global_start, _target_end, _global_end
    with _limit_lock:
        _global_start = time.time()
        _target_end  = _global_start + TOTAL_WALL_BUDGET
        _global_end  = _global_start + (TOTAL_WALL_BUDGET * 2.0)

    total_start_time = _global_start
    pre_thread_start_time = None

    from threading import Event
    exit_event = Event()

    _health_lock = threading.Lock()
    _health_start_ts = time.time()
    _prev_used_mb = None
    _checkpoint_counter = 0  
    _health_path_active = HEALTH_LOG_PATH

    try:
        for _p in (HEALTH_LOG_PATH, HEALTH_LOG_FALLBACK):
            try:
                os.makedirs(os.path.dirname(_p), exist_ok=True)
            except Exception:
                pass
        for _p in (HEALTH_LOG_PATH, HEALTH_LOG_FALLBACK):
            try:
                with open(_p, "w", encoding="utf-8") as _f:
                    _f.write("")
            except Exception:
                pass
        _health_path_active = HEALTH_LOG_PATH

    except Exception:
        _health_path_active = None 

    def _fmt_hms(seconds):
        h = int(seconds // 3600)
        m = int((seconds % 3600) // 60)
        s = int(seconds % 60)
        return f"{h}:{m:02d}:{s:02d}"
    def _fmt_mem(x_mb, signed=False):
        try:
            if x_mb is None:
                return "NA"
            sign = ""
            if signed:
                sign = "+" if x_mb >= 0 else "-"
                x_mb = abs(x_mb)
            if x_mb >= 1024:
                return f"{sign}{(x_mb/1024):.2f}GB"
            if isinstance(x_mb, float):
                return f"{sign}{x_mb:.1f}MB"
            return f"{sign}{int(x_mb)}MB"
        except Exception:
            return "NA"

    def _health_write_block(text):
        nonlocal _health_path_active
        block = f"\n{text}\n"
        print(block, end="", flush=True)

        if _health_path_active is None:
            return
        try:
            with _health_lock:
                with open(_health_path_active, "a", encoding="utf-8") as _f:
                    _f.write(block)
        except Exception:
            try:
                with _health_lock:
                    with open(HEALTH_LOG_FALLBACK, "a", encoding="utf-8") as _f:
                        _f.write(block)
                _health_path_active = HEALTH_LOG_FALLBACK
            except Exception:
                return 
        else:
            if _health_path_active == HEALTH_LOG_FALLBACK:
                try:
                    with open(HEALTH_LOG_PATH, "a", encoding="utf-8") as _pf:
                        _pf.write("") 
                    _health_path_active = HEALTH_LOG_PATH
                    with open(_health_path_active, "a", encoding="utf-8") as _pf2:
                        _pf2.write("\n[HEALTH NOTICE] primary log re-acquired; continuing here.\n")
                except Exception:
                    pass

    def _snapshot_ram():
        try:
            vm = psutil.virtual_memory()
            total_mb = (vm.total or 0) / (1024 * 1024)
            free_mb  = (vm.available or 0) / (1024 * 1024)
            used_mb  = total_mb - free_mb
            return int(total_mb), used_mb, free_mb, vm.percent
        except Exception:
            return None, None, None, None
    _baseline_total_mb, _baseline_used_mb, _baseline_free_mb, _baseline_pct = _snapshot_ram()
    if _baseline_used_mb is None:
        _baseline_used_mb = 0.0
    _prev_used_mb = _baseline_used_mb
    def _log_ram_checkpoint(tag):
        nonlocal _prev_used_mb, _checkpoint_counter
        total_mb, used_mb, free_mb, pct = _snapshot_ram()
        if used_mb is None:
            return
        num = _checkpoint_counter
        _checkpoint_counter += 1

        now = time.time()
        since = now - _health_start_ts
        wall = time.strftime("%Y-%m-%d %H:%M:%S")

        try:
            cur_limit = effective_thread_limit()

        except Exception:
            cur_limit = None
        try:
            remaining = max(0, (total_tasks - _finished_tasks))
            finished = _finished_tasks
        except Exception:
            remaining = finished = None

        _prev_used_mb = used_mb

        # Example:
        # [HB 10] 2025-03-01 13:40:00 | +0:10:00 since start | TOTAL_RAM=31.9GB USED=12.4GB FREE=19.5GB | LIMIT=123.45s | DONE=56 LEFT=142 | TAG=CHECKPOINT
        msg = (
            f"[HB {num:02d}] {wall} | +{_fmt_hms(since)} since start | "
            f"TOTAL_RAM={_fmt_mem(total_mb)} USED={_fmt_mem(used_mb)} FREE={_fmt_mem(free_mb)}"
        )
        if cur_limit is not None:
            msg += f" | LIMIT={cur_limit:.2f}s"
        if finished is not None and remaining is not None:
            msg += f" | DONE={finished} LEFT={remaining}"
        msg += f" | TAG={tag}"

        _health_write_block(msg)
    _checkpoint_interval = max(1.0, float(TOTAL_WALL_BUDGET) / 24.0)
    def _ram_checkpoints():
        _log_ram_checkpoint("BASELINE")
        return

    threading.Thread(target=_ram_checkpoints, daemon=True).start()

    WATCHDOG_SECONDS = int(TOTAL_WALL_BUDGET * 1.1) 

    def watchdog():
        if WATCHDOG_SECONDS is None:
            return
        time.sleep(WATCHDOG_SECONDS)
        exit_event.set()
        _health_write_block("[Watchdog] Timeout reached. Stopping threads...")
        try:
            _log_ram_checkpoint("TIMEOUT")
        except Exception:
            pass

    if WATCHDOG_SECONDS:
        threading.Thread(target=watchdog, daemon=True).start()

    WB2_FILE = r"C:/REDACTED/WB2.xlsm"

    wb = xw.Book(WB2_FILE)
    sheet = wb.sheets[0]

    try:
        app = wb.app
        app.api.ScreenUpdating = True
        app.api.EnableEvents = True
        app.api.DisplayStatusBar = True
        app.api.Calculation = -4105 
    except Exception:
        pass


    conn = sqlite3.connect(DB_FILE, check_same_thread=False) 
    cursor = conn.cursor()
    db_lock = threading.RLock()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS extra_big_groups (
            identifier TEXT PRIMARY KEY,
            pp REAL,
            score REAL,
            parameters TEXT,
            ticker_count REAL,                -- avg shared buys per day
            algo_ticker_counts TEXT,
            median_pp REAL,
            second_worst_pp REAL,
            wilson_lb REAL,
            mad_pp REAL,
            iqr_pp REAL,
            cv_buys REAL,
            daily_pp_json TEXT,
            daily_buys_json TEXT,
            shared_tickers_by_day_json TEXT,  -- explicit lists of shared tickers per day
            group_size INTEGER,
            members_json TEXT,
            repeat_factor REAL                -- how aggressively consensus names are reused (>=1.0)
        )
    ''')
    conn.commit()

    cursor.execute("PRAGMA table_info(extra_big_groups)")
    _ebg_cols = [r[1] for r in cursor.fetchall()]
    if "repeat_factor" not in _ebg_cols:
        cursor.execute("ALTER TABLE extra_big_groups ADD COLUMN repeat_factor REAL")
        conn.commit()

    identifier_cols = [3, 19, 35, 51, 67, 83, 99, 115, 131, 147]
    start_row = 6
    param_offset = 5
    result_offset = 4
    paste_col = 169
    paste_row = 6

    param_headers_map = {}
    variation_map = {}
    full_stocks_map = {}

    from itertools import combinations

    # Variation Map
    param_headers_map = {}
    variation_map = {}    
    full_stocks_map = {}   
    params_map = {}        
    variation_pp_map = {}  

    PRUNED_TABLES = [
        "pruned_MACD", "pruned_EMA", "pruned_RSI", "pruned_Breakout", "pruned_ADX",
        "pruned_Volatility", "pruned_SMA", "pruned_Bollinger_Bands",
        "pruned_EMA_MACD_Combo", "pruned_RSI_Bollinger"
    ]

    algo_idx_to_table = {i: t for i, t in enumerate(PRUNED_TABLES)}
    def _resolve_pruned_schema(cursor, table):
        cursor.execute(f"PRAGMA table_info({table})")
        cols_info = cursor.fetchall()  
        cols = [r[1] for r in cols_info]

        def pick(cands, default=None):
            for c in cands:
                if c in cols:
                    return c
            return default

        mapping = {
            "identifier":    pick(["identifier", "id", "variation_id"]),
            "pp":            pick(["pp", "pooled_pp", "overall_pp", "median_pp", "true_pp"]),
            "uscore":        pick(["universal_score", "u_score", "score", "rank_score"]),
            "params_json":   pick(["params_json", "params", "parameters_json"]),
            "headers_json":  pick(["param_headers_json", "headers_json", "param_headers"]),
            "stocks_json":   pick(["stocks_json", "stocks", "stocks_blob", "shared_stocks_json"]),
        }
        return mapping, set(cols)

    def _to_float_or_neg_inf(x):
        try:
            return float(x)
        except Exception:
            return float("-inf")
    for algo_idx in range(len(PRUNED_TABLES)):
        table = algo_idx_to_table[algo_idx]
        cursor.execute(
            f"""
            SELECT
                identifier,
                final_order_score,
                pooled_pp,
                params_json,
                param_headers_json,
                day1_buys_json, day2_buys_json, day3_buys_json, day4_buys_json, day5_buys_json
            FROM {table}
            """
        )
        rows = cursor.fetchall()
        if not rows:
            variation_map[algo_idx] = []
            continue

        first_headers = []
        try:
            if rows[0][4]:
                first_headers = json.loads(rows[0][4])
        except Exception:
            first_headers = []
        param_headers_map[algo_idx] = first_headers

        def _tofloat(x, default=-1e18):
            try:
                return float(x)
            except Exception:
                return default
        rows.sort(key=lambda r: (_tofloat(r[1]), _tofloat(r[2])), reverse=True)

        variation_ids = []
        for identifier, final_score, pooled_pp, params_json, _headers_json, d1, d2, d3, d4, d5 in rows:
            def _load_day(blob, day_idx):
                try:
                    arr = json.loads(blob) if blob else []
                except Exception:
                    arr = []
                out = []
                for s in arr:
                    if not s:
                        continue
                    t = s.get("ticker")
                    if not t:
                        continue
                    out.append(_slim_stock_dict_from_buys(s, day_idx))
                return out

            stocks_all = []
            stocks_all.extend(_load_day(d1, 1))
            stocks_all.extend(_load_day(d2, 2))
            stocks_all.extend(_load_day(d3, 3))
            stocks_all.extend(_load_day(d4, 4))
            stocks_all.extend(_load_day(d5, 5))

            variation_ids.append(identifier)

            try:
                params = json.loads(params_json) if params_json else []
            except Exception:
                params = []

            params_map[(algo_idx, identifier)] = params
            full_stocks_map[(algo_idx, identifier)] = stocks_all

            try:
                variation_pp_map[(algo_idx, identifier)] = float(pooled_pp)
            except Exception:
                variation_pp_map[(algo_idx, identifier)] = None

        variation_map[algo_idx] = variation_ids

    day_buys_map = {} 
    day_profit_map = {}  
    for (a, ident), rows in full_stocks_map.items():
        for d in range(1, 6):
            day_buys_map[(a, ident, d)] = set()
            day_profit_map[(a, ident, d)] = {}

        for r in rows:    
            t = r.get("ticker")
            if not t:
                continue
            d = int(r.get("day", 0) or 0)
            if d < 1 or d > 5:
                continue

            day_buys_map[(a, ident, d)].add(t)
            prof = _is_profit_row(r) 
            prev = day_profit_map[(a, ident, d)].get(t, 0)
            if prof > prev:
                day_profit_map[(a, ident, d)][t] = prof

    DAY_DIFFICULTY_WEIGHTS = [1.0] * 5
    base_wr = [0.0] * 5
    try:
        DAY_TABLES = ["DS_DAY1", "DS_DAY2", "DS_DAY3", "DS_DAY4", "DS_DAY5"]
        BLOCK_ROWS = 53
        FIRST_DATA_ROW = 3 

        for d_idx, day_table in enumerate(DAY_TABLES):
            try:
                cursor.execute(f"SELECT rowid, Open, Close FROM {day_table}")
                rows = cursor.fetchall()
            except Exception as _e:
                print(f"[warn] could not read {day_table} for baseline WR: {_e}", flush=True)
                continue

            total = 0
            wins = 0
            for rowid, o, c in rows:
                if (rowid - FIRST_DATA_ROW) % BLOCK_ROWS != 0:
                    continue

                o_f = _to_float_safe(o)
                c_f = _to_float_safe(c)
                if o_f is None or c_f is None:
                    continue

                total += 1
                if c_f > o_f:
                    wins += 1

            if total > 0:
                base_wr[d_idx] = 100.0 * wins / total
            else:
                base_wr[d_idx] = 0.0

        if any(base_wr):
            avg_wr = sum(base_wr) / 5.0
            k = 0.02 
            weights = []
            for d_idx in range(5):
                diff = avg_wr - base_wr[d_idx]
                w = 1.0 + k * diff
                # clamp to a sane band
                w = max(0.5, min(1.5, w))
                weights.append(w)
            DAY_DIFFICULTY_WEIGHTS = weights
    except Exception as _e:
        print(f"[warn] DAY_DIFFICULTY_WEIGHTS fallback to 1.0 due to: {_e}", flush=True)
        base_wr = [0.0] * 5

    print("\n[Global day baseline from DS_DAY# first rows]")
    for idx in range(5):
        bw = base_wr[idx]
        w = DAY_DIFFICULTY_WEIGHTS[idx]
        print(f"  Day {idx + 1}: base_wr={bw:.2f}%   weight={w:.3f}")
    print("")

    def extract_params(row_or_identifier, algo_idx):
        return params_map.get((algo_idx, row_or_identifier), [])

    def intersect_full_stock_data(s1, s2):
        t1 = {s["ticker"]: s for s in s1}
        t2 = {s["ticker"]: s for s in s2}
        return [t1[t] for t in t1 if t in t2]
    def calc_pp(stocks):
        if not stocks:
            return 0.0
        n = len(stocks)
        h = sum(int(s.get("is_profit", 0)) for s in stocks)
        return round(100.0 * h / n, 2)


    def group_score(pp, group_size):
        return round((math.log(group_size + 1) * (pp ** 1.15)), 4)

    def score_group(pp, group_size, shared_count):
        return round(pp * math.log1p(group_size) * math.log1p(shared_count), 4)

    best_by_size = {3: [], 4: []}
    def compute_group_stability(daily_pp_list, lambda_factor=1.5):
        if not daily_pp_list:
            return 0.0, 0.0, -1e9

        try:
            n = len(daily_pp_list)
            pp_avg = sum(daily_pp_list) / float(n)
            if n < 2:
                pp_std = 0.0
            else:
                pp_std = float(statistics.pstdev(daily_pp_list))
        except Exception:
            return 0.0, 0.0, -1e9

        stab_score = pp_avg - lambda_factor * pp_std
        return pp_avg, pp_std, stab_score


    MAX_GROUP_SIZE = 4
    MAX_BUYS_PER_DAY = 20
    MIN_POOLED_PP_BC = 50.0
  
    # Top-K storage per size 3 and 4
    TOPK_LIMIT_PER_SIZE = 1000
    topk_by_size = {3: [], 4: []}
    stats_by_size = {
        3: {"n": 0, "mean": 0.0, "M2": 0.0},
        4: {"n": 0, "mean": 0.0, "M2": 0.0},
    }

    SOFT_THREAD_CAP_SEC = 10000.0

    def _thread_time_up(deadline):
        try:
            return (time.time() - deadline) >= min(current_per_task_limit(), SOFT_THREAD_CAP_SEC)
        except Exception:
            return True

    def effective_thread_limit():
        try:
            return min(current_per_task_limit(), SOFT_THREAD_CAP_SEC)
        except Exception:
            return SOFT_THREAD_CAP_SEC
    def _upsert_best_group(entry, size_now):
        try:
            _g_members = entry["group"]       
            _tickers   = entry["tickers"]   
            _pp        = entry.get("pp")
            _sc        = entry.get("score")
            _consensus_sets = entry.get("consensus_sets_by_day") or [set()] * 5
            _repeat_factor = entry.get("repeat_factor")

            _members = []
            for (a, ident_mem) in _g_members:
                all_rows = full_stocks_map.get((a, ident_mem), [])
                daily = []
                for d in range(1, 6):
                    day_set = _consensus_sets[d - 1]
                    day_rows = [
                        r for r in all_rows
                        if r.get("ticker") in day_set and int(r.get("day", 0) or 0) == d
                    ]
                    buys_d = len(day_rows)
                    hits_d = sum(1 for r in day_rows if int(r.get("is_profit", 0)) == 1)
                    pp_d = (100.0 * hits_d / buys_d) if buys_d else 0.0
                    daily.append({
                        "day": d,
                        "date": (
                            target_dates_list[d - 1]
                            if 'target_dates_list' in locals() and len(target_dates_list) >= d
                            else f"Day{d}"
                        ),
                        "buys": buys_d,
                        "hits": hits_d,
                        "pp": round(pp_d, 2),
                        "tickers": [
                            {
                                "ticker": r.get("ticker"),
                                "name": r.get("name"),
                                "open": r.get("open"),
                                "close": r.get("close"),
                                "is_profit": int(r.get("is_profit", 0)),
                                "day": r.get("day"),
                            }
                            for r in day_rows
                        ],
                    })
                _members.append({
                    "algo_index": a,
                    "algo_table": PRUNED_TABLES[a],
                    "identifier": ident_mem,
                    "params": params_map.get((a, ident_mem), []),
                    "total_unique_tickers": len(set().union(*_consensus_sets)),
                    "daily": daily,
                })

            _params = {ident: params_map.get((_a, ident), []) for (_a, ident) in _g_members}
            _counts = []
            for (a, ident) in _g_members:
                _cnt = len(full_stocks_map.get((a, ident), []))
                _counts.append(f"No.{a+1}:{_cnt}")
            _counts_text = "|".join(_counts) if _counts else ""
            _ticker_count = len(_tickers)
            _shared_tickers_by_day = []
            for d in range(1, 6):
                day_set = _consensus_sets[d - 1]
                day_list = []
                for t in sorted(day_set):
                    prof_any = 0
                    for (a, ident_mem) in _g_members:
                        if day_profit_map[(a, ident_mem, d)].get(t, 0) == 1:
                            prof_any = 1
                            break
                    day_list.append({
                        "day": d,     
                        "ticker": t,
                        "profit": prof_any
                    })
                _shared_tickers_by_day.append(day_list)

            size_key = f"SIZE{size_now}"

            with db_lock:
                cursor.execute(
                    '''REPLACE INTO extra_big_groups
                       (identifier, pp, score, parameters, ticker_count, algo_ticker_counts,
                        median_pp, second_worst_pp, wilson_lb, mad_pp, iqr_pp, cv_buys,
                        daily_pp_json, daily_buys_json, shared_tickers_by_day_json,
                        group_size, members_json, repeat_factor)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                    (
                        size_key,
                        _pp,
                        _sc,
                        json.dumps(_params),
                        _ticker_count,
                        _counts_text,
                        entry.get("median_pp"),
                        entry.get("second_worst_pp"),
                        entry.get("wilson_lb"),
                        entry.get("mad_pp"),
                        entry.get("iqr_pp"),
                        entry.get("cv_buys"),
                        json.dumps(entry.get("daily_stats")),
                        json.dumps(entry.get("buy_stats")),
                        json.dumps(_shared_tickers_by_day),
                        size_now,
                        json.dumps(_members),
                        _repeat_factor,
                    )
                )
                conn.commit()
        except Exception as _e:
            try:
                print(f"[EBG upsert warn] size={size_now} err={_e}", flush=True)
            except Exception:
                pass
    # Bell-curve Top-K accumulator
    def consider_best_group(size_now, entry):
        if size_now not in (3, 4):
            return

        pp_val = float(entry.get("pp", 0.0) or 0.0)
        if pp_val < MIN_POOLED_PP_BC:
            return

        day_buys_list = entry.get("buy_stats") or []
        if any(b > MAX_BUYS_PER_DAY for b in day_buys_list):
            return

        mad_pp    = float(entry.get("mad_pp", 0.0))
        iqr_pp    = float(entry.get("iqr_pp", 0.0))
        cv_buys   = float(entry.get("cv_buys", 0.0))
        wilson_lb = float(entry.get("wilson_lb", 0.0))

        # mad_pp: average absolute deviation of daily PP from median
        # iqr_pp: spread of middle 60% of daily PP
        # cv_buys: relative volatility in buy counts across days
        # wilson_lb: conservative PP floor
      
        if (
            mad_pp > 25.0 or      
            iqr_pp > 45.0 or 
            cv_buys > 0.80 or 
            wilson_lb < 10.0
        ):
            return

        score_val = float(entry.get("score", 0.0))

        stats = stats_by_size[size_now]
        n = stats["n"] + 1
        delta = score_val - stats["mean"]
        mean_new = stats["mean"] + delta / n
        M2_new = stats["M2"] + delta * (score_val - mean_new)
        stats["n"] = n
        stats["mean"] = mean_new
        stats["M2"] = M2_new

        bucket = topk_by_size[size_now]
        bucket.append(entry)
        bucket.sort(key=lambda e: e["score"], reverse=True)
        if len(bucket) > TOPK_LIMIT_PER_SIZE:
            bucket.pop() 

        if entry in bucket:
            tid = getattr(thread_local_data, "thread_id", None)
            if tid is not None:
                with topk_adds_lock:
                    counts = topk_adds_per_thread.setdefault(tid, {3: 0, 4: 0})
                    counts[size_now] = counts.get(size_now, 0) + 1

        cur = best_by_size.setdefault(size_now, [])
        cur.append(entry)
        cur.sort(key=lambda e: e["score"], reverse=True)
        if len(cur) > 1:
            del cur[1:]

        _upsert_best_group(cur[0], size_now)

    def _sweet_spot_bonus(mean_buys: float) -> float:
        if mean_buys <= 0.0:
            return -15.0

        if mean_buys < 7.0:
            return -3.0 * (7.0 - mean_buys)

        if mean_buys <= 14.0:
            return 2.0 * (mean_buys - 7.0)

        if mean_buys <= 18.0:
            return 2.0 * (18.0 - mean_buys)

        return -4.0 * (mean_buys - 18.0)

    def build(group, consensus_by_day, used_algos, hits, deadline=None):
        global tested_groups_total
        if exit_event.is_set() or time.time() >= _global_end:
            return
        # time check
        if deadline is not None:
            if _thread_time_up(deadline):
                return
        hits[0] += 1

        if not any(consensus_by_day.get(d) for d in range(1, 6)):
            return

        size_now = len(group)

        day_buys = []
        day_hits = []
        for d in range(1, 6):
            tickers_d = consensus_by_day.get(d, set())
            buys_d = len(tickers_d)
            if buys_d == 0:
                day_buys.append(0)
                day_hits.append(0)
                continue

            hits_d = 0
            for t in tickers_d:
                prof_any = 0
                for (a, ident) in group:
                    if day_profit_map[(a, ident, d)].get(t, 0) == 1:
                        prof_any = 1
                        break
                hits_d += prof_any

            day_buys.append(buys_d)
            day_hits.append(hits_d)

        # daily floor
        if size_now >= 2 and any(h < 8 for h in day_hits):
            return

        weekly_set = set().union(*(consensus_by_day.get(d, set()) for d in range(1, 6)))
        # repeat_factor
        total_consensus_buys = sum(day_buys)
        unique_consensus = len(weekly_set)
        repeat_factor = (total_consensus_buys / max(1, unique_consensus)) if unique_consensus else 0.0
        if size_now >= 2:
            global tested_groups_total
            with tested_groups_lock:
                tested_groups_total += 1

            daily_pp = [
                (100.0 * day_hits[i] / day_buys[i]) if day_buys[i] > 0 else 0.0
                for i in range(5)
            ]

            weighted_pp = sum(
                daily_pp[i] * DAY_DIFFICULTY_WEIGHTS[i]
                for i in range(5)
            ) / 5.0

            pooled_buys = sum(day_buys)
            pooled_hits = sum(day_hits)
            pooled_pp = 100.0 * pooled_hits / max(1, pooled_buys)

            sorted_pp = sorted(daily_pp)
            median_pp = sorted_pp[2]
            second_worst_pp = sorted_pp[1]
            mad_pp = sum(abs(p - median_pp) for p in daily_pp) / 5.0
            iqr_pp = sorted_pp[3] - sorted_pp[1]

            mean_buys = sum(day_buys) / 5.0
            nonzero = [b for b in day_buys if b > 0]
            cv_buys = (stdev(nonzero) / mean_buys) if mean_buys and len(nonzero) > 1 else 0.0

            if pooled_buys == 0:
                wilson_lb = 0.0
            else:
                phat = pooled_hits / pooled_buys
                z = 1.96
                wilson_lb = 100.0 * (
                    (phat + z*z/(2*pooled_buys) -
                     z * math.sqrt((phat*(1-phat)+z*z/(4*pooled_buys))/pooled_buys))
                    / (1+z*z/pooled_buys)
                )

            # 1) Trend of PP across the week
            xs = [1, 2, 3, 4, 5]
            mean_x = sum(xs) / 5.0
            mean_y = sum(daily_pp) / 5.0
            num = sum((x - mean_x) * (y - mean_y) for x, y in zip(xs, daily_pp))
            den = sum((x - mean_x) ** 2 for x in xs) or 1.0
            pp_slope = num / den

            # 2) Profit base
            profit_base = (
                0.60 * weighted_pp +
                0.25 * median_pp +
                0.15 * second_worst_pp
            )

            # 3) Stability penalties
            trend_penalty = 6.0 * max(0.0, -pp_slope)
            stability_penalty = (
                0.25 * mad_pp +
                0.18 * iqr_pp +
                0.10 * cv_buys * 100.0 +
                trend_penalty
            )
            # 4) Coverage
            capped_buys = min(mean_buys, 25.0)
            coverage_core = 6.0 * math.log1p(capped_buys)

            overflow_buys = max(0.0, mean_buys - 25.0)
            coverage_penalty = 0.5 * overflow_buys

            sweet_spot = _sweet_spot_bonus(mean_buys)

            max_buys_day = max(day_buys) if day_buys else 0
            hard_over_penalty = 6.0 * max(0.0, max_buys_day - 20.0) ** 2

            coverage_term = coverage_core + sweet_spot - coverage_penalty - hard_over_penalty

            # 5) Repeat-factor penalty
            rf_excess = max(0.0, repeat_factor - 1.20)
            repeat_penalty = 25.0 * (rf_excess ** 2)

            # 6) reward for upward PP slope
            trend_boost = 3.0 * max(0.0, pp_slope)

            # 7) Final score
            score_val = (
                profit_base +
                0.80 * wilson_lb +
                coverage_term +
                trend_boost -
                stability_penalty -
                repeat_penalty
            )

            if 3 <= size_now <= 4:
                entry = {
                    "group": [(a, identifier_map[(a, ident)]) for (a, ident) in group],
                    "tickers": frozenset(weekly_set),
                    "pp": round(pooled_pp, 2),
                    "score": round(score_val, 4),
                    "daily_stats": [round(x, 2) for x in daily_pp],
                    "buy_stats": list(day_buys),
                    "wilson_lb": round(wilson_lb, 2),
                    "median_pp": round(median_pp, 2),
                    "second_worst_pp": round(second_worst_pp, 2),
                    "mad_pp": round(mad_pp, 2),
                    "iqr_pp": round(iqr_pp, 2),
                    "cv_buys": round(cv_buys, 3),
                    "repeat_factor": round(repeat_factor, 3),
                    "consensus_sets_by_day": [
                        set(consensus_by_day.get(d, set())) for d in range(1, 6)
                    ],
                }
                consider_best_group(size_now, entry)

        # Expansion rails
        if size_now >= MAX_GROUP_SIZE:
            return

        max_algo_in_group = max(a for a, _ in group)
        for next_algo in range(max_algo_in_group + 1, len(PRUNED_TABLES)):
            if exit_event.is_set():
                return
            if _thread_time_up(deadline):
                return

            if next_algo in used_algos:
                continue
            for ident in variation_map.get(next_algo, []):
                if _thread_time_up(deadline):
                    return

                new_consensus = {}
                has_any = False
                for d in range(1, 6):
                    base_set = consensus_by_day.get(d, set())
                    add_set  = day_buys_map[(next_algo, ident, d)]
                    inter    = base_set & add_set
                    new_consensus[d] = inter
                    if inter:
                        has_any = True

                if not has_any:
                    continue

                group.append((next_algo, ident))
                used_algos.add(next_algo)
                build(group, new_consensus, used_algos, hits, deadline=deadline)
                used_algos.remove(next_algo)
                group.pop()

    counter = 0
    total_variations = sum(len(v) for v in variation_map.values())
    checkpoint = max(1, total_variations // 10)
    print("\n[] Selecting ALL variations as thread starters (exhaustive coverage)...")

    identifier_map = {}
    for algo_idx, id_list in variation_map.items():
        for ident in id_list:
            identifier_map[(algo_idx, ident)] = ident

    special_algos = {0, 1, 3}

    normal_items = []
    special_items = []

    for ai in range(len(identifier_cols)):
        for ident in variation_map.get(ai, []):
            tcount = len(full_stocks_map[(ai, ident)])
            if ai in special_algos:
                special_items.append((ai, ident, tcount))
            else:
                normal_items.append((ai, ident, tcount))

    normal_items.sort(key=lambda x: (x[2], x[0], x[1]))
    special_items.sort(key=lambda x: (x[2], x[0], x[1]))

    thread_starters = [(ai, ident) for (ai, ident, _t) in normal_items]
    thread_starters += [(ai, ident) for (ai, ident, _t) in special_items]

    if thread_starters:
        first_len = len(full_stocks_map[(thread_starters[0][0], thread_starters[0][1])])
        last_len  = len(full_stocks_map[(thread_starters[-1][0], thread_starters[-1][1])])
        print(f" Mixed ordering applied. First tickers={first_len}, Last tickers={last_len} | specials at end: {sorted(special_algos)}")

    total_tasks = len(thread_starters)
    max_workers = min(4, max(1, total_tasks))
    batches = math.ceil(total_tasks / max_workers) or 1
    with _limit_lock:
        _base_per_task_limit = float(TOTAL_WALL_BUDGET) / float(batches)
        _delta_per_task = 0.0
        _finished_tasks = 0
    print(f"[init] tasks={total_tasks} | workers={max_workers} | batches={batches} | base_limit={_base_per_task_limit:.3f}s")
    _bank_seconds = 0.0 

    def _apply_bank_locked():
        nonlocal _delta_per_task, _finished_tasks, _bank_seconds
        now = time.time()
        time_left = max(0.0, _target_end - now)
        remaining = max(1, total_tasks - _finished_tasks)
        batches_left = max(1, math.ceil(remaining / float(max_workers)))
        cap_per_task = time_left / float(batches_left)

        current_limit = _base_per_task_limit + _delta_per_task
        per_task_headroom = max(0.0, cap_per_task - current_limit)
        if per_task_headroom <= 0.0 or _bank_seconds <= 0.0:
            return 0.0, 0.0, batches_left, current_limit, current_limit, _bank_seconds

        per_task_bump_from_bank = _bank_seconds / float(remaining)
        per_task_bump = min(per_task_bump_from_bank, per_task_headroom)

        applied_total = per_task_bump * remaining
        _delta_per_task += per_task_bump
        _bank_seconds = max(0.0, _bank_seconds - applied_total)

        new_limit = _base_per_task_limit + _delta_per_task
        return applied_total, per_task_bump, batches_left, current_limit, new_limit, _bank_seconds
    _bank_seconds = 0.0 

    def current_per_task_limit():
        with _limit_lock:
            return max(0.01, _base_per_task_limit + _delta_per_task)

    def donate_leftover(seconds, donor_tag=""):
        nonlocal _bank_seconds, _delta_per_task, _finished_tasks
        with _limit_lock:
            if seconds <= 0:
                cur = _base_per_task_limit + _delta_per_task
                print(
                    f"[limit] banked 0.000s from {donor_tag or 'thread'} | "
                    f"applied_total 0.000s (per-task +0.000s) | "
                    f"limit {cur:.3f}→{cur:.3f} | bank_left={_bank_seconds:.3f}s"
                )
                return

            _bank_seconds += seconds
            remaining = max(1, total_tasks - _finished_tasks)
            per_task_bump = _bank_seconds / float(remaining) * 0.6

            old_limit = _base_per_task_limit + _delta_per_task
            _delta_per_task += per_task_bump
            new_limit = _base_per_task_limit + _delta_per_task

            applied_total = per_task_bump * remaining
            _bank_seconds = max(0.0, _bank_seconds - applied_total)

            print(
                f"[limit] banked {seconds:.3f}s from {donor_tag or 'thread'} | "
                f"applied_total {applied_total:.3f}s (per-task +{per_task_bump:.3f}s) | "
                f"limit {old_limit:.3f}→{new_limit:.3f} | bank_left={_bank_seconds:.3f}s"
            )

    _rss_baseline = None
    _sys_used_baseline = None
    _proc = None
    try:
        _proc = psutil.Process(os.getpid())
        mi = _proc.memory_info()
        _rss_baseline = (mi.rss or 0) / (1024 * 1024)
        vm0 = psutil.virtual_memory()
        _sys_used_baseline = ((vm0.total or 0) - (vm0.available or 0)) / (1024 * 1024)
    except Exception:
        _proc = None

    def _hb_once(prefix="HB"):
        try:
            rss_mb = vms_mb = p_cpu = sys_cpu = None
            total_mb = avail_mb = mem_pct = swap_used_mb = None
            if _proc:
                try:
                    mem = _proc.memory_info()
                    rss_mb = (mem.rss or 0) / (1024 * 1024)
                    vms_mb = (mem.vms or 0) / (1024 * 1024)
                    p_cpu = _proc.cpu_percent(interval=None)
                except Exception:
                    pass
            try:
                vm = psutil.virtual_memory()
                sm = psutil.swap_memory()
                total_mb = (vm.total or 0) / (1024 * 1024)
                avail_mb = (vm.available or 0) / (1024 * 1024)
                mem_pct = vm.percent
                swap_used_mb = (sm.used or 0) / (1024 * 1024)
                sys_cpu = psutil.cpu_percent(interval=0.1)
            except Exception:
                pass

            # workload stats
            try:
                variations_loaded = len(full_stocks_map)
            except Exception:
                variations_loaded = None
            try:
                total_tickers = sum(len(v) for v in full_stocks_map.values()) if full_stocks_map else 0
                avg_tickers = (total_tickers / max(1, variations_loaded)) if variations_loaded else None
            except Exception:
                avg_tickers = None
            try:
                s3 = 1 if best_by_size.get(3) else 0
                s4 = 1 if best_by_size.get(4) else 0
            except Exception:
                s3 = s4 = None

            try:
                cur_limit = effective_thread_limit()

            except Exception:
                cur_limit = None
            try:
                remaining = max(0, total_tasks - _finished_tasks)
                finished = _finished_tasks
            except Exception:
                remaining = finished = None
            sys_used_mb = None
            sys_used_delta_mb = None
            rss_delta_mb = None
            if total_mb is not None and avail_mb is not None:
                sys_used_mb = total_mb - avail_mb
            try:
                if _sys_used_baseline is not None and sys_used_mb is not None:
                    sys_used_delta_mb = sys_used_mb - _sys_used_baseline
            except Exception:
                pass
            try:
                if _rss_baseline is not None and rss_mb is not None:
                    rss_delta_mb = rss_mb - _rss_baseline
            except Exception:
                pass
            line = []
            line.append(f"[{prefix}]")
            if total_mb is not None:
                line.append(f"SYS={_fmt_mem(total_mb)}")
            if sys_used_mb is not None and mem_pct is not None:
                line.append(f"USED={_fmt_mem(sys_used_mb)}({int(mem_pct)}%)")
                line.append(f"FREE={_fmt_mem(avail_mb)}")

            if sys_used_delta_mb is not None:
                line.append(f"USEDΔ={_fmt_mem(sys_used_delta_mb, signed=True)}")
            if rss_mb is not None:
                line.append(f"RSS={_fmt_mem(rss_mb)}")
            if rss_delta_mb is not None:
                line.append(f"RSSΔ={_fmt_mem(rss_delta_mb, signed=True)}")
            if vms_mb is not None:
                line.append(f"VMS={_fmt_mem(vms_mb)}")
            if swap_used_mb is not None:
                line.append(f"SWAP={_fmt_mem(swap_used_mb)}")

            if p_cpu is not None:
                line.append(f"P_CPU={p_cpu:.1f}%")
            if sys_cpu is not None:
                line.append(f"S_CPU={sys_cpu:.1f}%")
            if variations_loaded is not None:
                line.append(f"VAR={variations_loaded}")
            if avg_tickers is not None:
                line.append(f"AVG_TICK={avg_tickers:.1f}")
            if s3 is not None:
                line.append(f"S3={s3} S4={s4}")
            if cur_limit is not None:
                line.append(f"LIMIT={cur_limit:.2f}s")
            if finished is not None and remaining is not None:
                line.append(f"DONE={finished} LEFT={remaining}")
            elapsed = time.time() - _global_start
            until_tgt = max(0.0, _target_end - time.time())
            until_end = max(0.0, _global_end - time.time())
            line.append(f"t+{elapsed:.0f}s tgt-{until_tgt:.0f}s end-{until_end:.0f}s")
            _health_write_block(" ".join(line))

        except Exception as _e:
            try:
                print(f"[HB] err {type(_e).__name__}: {_e}", flush=True)
            except Exception:
                pass
    def _start_heartbeat():
        def _hb():
            while not exit_event.is_set() and not heartbeat_stop.is_set():
                time.sleep(HEALTH_INTERVAL_SEC)
                _log_ram_checkpoint("CHECKPOINT")
        threading.Thread(target=_hb, daemon=True).start()

    def mark_finished():
        nonlocal _finished_tasks
        with _limit_lock:
            _finished_tasks += 1

    pre_thread_start_time = time.time()

    print(f"[Timer] Pre-thread logic done. Time elapsed: {pre_thread_start_time - total_start_time:.2f} seconds")
    _start_heartbeat()

    print(f"\n[] Starting {len(thread_starters)} starters with a 4-worker pool...\n")

    def run_thread(base_id, algo, row):
        start = time.time()
        limit_at_start = effective_thread_limit()
        thread_local_data.thread_id = base_id + 1
        var_id = identifier_map[(algo, row)]
        base_stocks = full_stocks_map[(algo, row)]
        base_pp = variation_pp_map.get((algo, var_id), None)
        base_tickers = len(base_stocks)
        pp_str = f"{base_pp:.2f}%" if isinstance(base_pp, (int, float)) else "NA"
        _health_write_block(
            f"[Thread {base_id+1} START] {var_id} | PP={pp_str} | T={base_tickers} | limit_start={limit_at_start:.2f}s"
        )

        deadline = start
        # seed per-day consensus
        seed_consensus = {d: set(day_buys_map[(algo, row, d)]) for d in range(1, 6)}
        hits = [0]
        try:
            build([(algo, row)], seed_consensus, {algo}, hits, deadline=deadline)

        except ThreadTimeout as e:
            print(f"[Thread {base_id+1}] Exited: {e} | Attempts: {hits[0]}", flush=True)
        finally:
            dur = time.time() - start
            limit_at_end = effective_thread_limit()

            timed_out = (
                dur >= (limit_at_end - 0.01)
                or time.time() >= _global_end
                or exit_event.is_set()
            )

            status = "timeout" if timed_out else "done"
            end_pp_str = f"{base_pp:.2f}%" if isinstance(base_pp, (int, float)) else "NA"
            with topk_adds_lock:
                counts = topk_adds_per_thread.get(base_id + 1, {3: 0, 4: 0})
                adds3 = counts.get(3, 0)
                adds4 = counts.get(4, 0)

            _health_write_block(
                f"[Thread {base_id+1} END] {var_id} | {status} | "
                f"dur={dur:.2f}s | attempts={hits[0]} | "
                f"limit_start={limit_at_start:.2f}s | limit_end={limit_at_end:.2f}s | "
                f"PP={end_pp_str} | T={base_tickers} | "
                f"TopK_adds(size3={adds3}, size4={adds4})"
            )

            with _algo_stats_lock:
                st = _algo_stats.get(algo)
                if st is not None:
                    st["threads"] += 1
                    st["runtime_sec"] += float(dur)
                    if isinstance(base_pp, (int, float)):
                        st["sum_pp"] += float(base_pp)
                    st["sum_tickers"] += float(base_tickers)
                    if timed_out:
                        st["timeouts"] += 1

            leftover = 0.0 if timed_out else max(0.0, limit_at_end - dur)
            mark_finished()
            if leftover > 0:
                donate_leftover(leftover, donor_tag=f"thread {base_id+1}")
    pre_thread_start_time = time.time()
    from concurrent.futures import ThreadPoolExecutor
    try:
        with ThreadPoolExecutor(max_workers=max_workers) as ex:
            iterator = iter(enumerate(thread_starters))
            in_flight = set()

            for _ in range(max_workers):
                try:
                    i, (algo, row) = next(iterator)
                except StopIteration:
                    break
                in_flight.add(ex.submit(run_thread, i, algo, row))

            while in_flight:
                try:
                    done, in_flight = concurrent.futures.wait(
                        in_flight,
                        return_when=concurrent.futures.FIRST_COMPLETED
                    )
                    if exit_event.is_set() or time.time() >= _global_end:
                        exit_event.set()
                        try:
                            ex.shutdown(wait=False, cancel_futures=True)
                        except Exception:
                            pass
                        break

                    for _ in done:
                        try:
                            i, (algo, row) = next(iterator)
                            in_flight.add(ex.submit(run_thread, i, algo, row))
                        except StopIteration:
                            pass
                except KeyboardInterrupt:
                    print("[] KeyboardInterrupt received — cancelling all threads now...", flush=True)
                    exit_event.set()
                    try:
                        ex.shutdown(wait=False, cancel_futures=True)
                    except Exception:
                        pass
                    break

    except KeyboardInterrupt:
        print("[] KeyboardInterrupt received (outer) — stopping...", flush=True)
        exit_event.set()

    thread_end_time = time.time()
    print(f"[Timer] Threads finished. Time elapsed: {thread_end_time - pre_thread_start_time:.2f} seconds")
    print(f"[Timer] TOTAL Time (before + during threads): {thread_end_time - total_start_time:.2f} seconds")
    _log_ram_checkpoint("FINAL")

    heartbeat_stop.set()
    try:
        time.sleep(0.1)
        _hb_once(prefix="HB_END")
    except Exception:
        pass
    # Bell-curve selection
    print("\n[] Bell-curve + stability selection for size-3 and size-4 groups...")
    SIGMA_MULT = 1.0   
    STAB_PERCENTILE = 0.60  

    for sz in (3, 4):
        top_list = topk_by_size.get(sz, [])
        stats = stats_by_size.get(sz, {})
        n_raw = int(stats.get("n", 0)) 

        if not top_list or n_raw == 0:
            print(f"[] size {sz}: no eligible groups.")
            best_by_size[sz] = []
            continue

        # stability metrics
        stab_scores = []
        for e in top_list:
            daily_pp_list = e.get("daily_stats") or []
            pp_avg_g, pp_std_g, stab_g = compute_group_stability(daily_pp_list, lambda_factor=1.5)
            e["pp_avg_stab"] = round(pp_avg_g, 3)
            e["pp_std_stab"] = round(pp_std_g, 3)
            e["stab_score"] = round(stab_g, 3)
            stab_scores.append(stab_g)

        if stab_scores:
            stab_scores_sorted = sorted(stab_scores)
            idx = int(STAB_PERCENTILE * (len(stab_scores_sorted) - 1))
            stab_threshold = stab_scores_sorted[idx]
            stable_pool = [e for e in top_list if e.get("stab_score", -1e9) >= stab_threshold]
            if not stable_pool:
                stable_pool = top_list
        else:
            stab_threshold = -1e9
            stable_pool = top_list

        scores = [float(e["score"]) for e in stable_pool]
        n = len(scores)
        if n == 0:
            print(f"[] size {sz}: stability pool empty, falling back to raw Top-K max.")
            chosen = max(top_list, key=lambda e: e["score"])
            best_by_size[sz] = [chosen]
            _upsert_best_group(chosen, sz)
            continue

        mean = sum(scores) / float(n)
        if n > 1:
            variance = sum((s - mean) ** 2 for s in scores) / float(n - 1)
            sigma = math.sqrt(max(0.0, variance))
        else:
            sigma = 0.0

        target = mean + SIGMA_MULT * sigma
        candidates = [e for e in stable_pool if float(e["score"]) >= target]
        if not candidates:
            chosen = max(stable_pool, key=lambda e: e["score"])
        else:
            chosen = min(candidates, key=lambda e: abs(float(e["score"]) - target))

        best_by_size[sz] = [chosen]

        total_buys = sum(chosen.get("buy_stats", []))
        pp_list = chosen.get("daily_stats", [])
        pp_avg_chosen = sum(pp_list) / float(len(pp_list)) if pp_list else 0.0
        pp_std_chosen = 0.0
        if len(pp_list) > 1:
            try:
                pp_std_chosen = float(statistics.pstdev(pp_list))
            except Exception:
                pp_std_chosen = 0.0

        print(
            f"[] size {sz}: n_raw={n_raw}, TopK={len(top_list)}, "
            f"stab_pool={len(stable_pool)}, stab_thresh={stab_threshold:.3f}, "
            f"mean_score={mean:.3f}, sigma={sigma:.3f}, target={target:.3f}, "
            f"chosen_score={float(chosen['score']):.3f}, "
            f"pp={float(chosen['pp']):.2f}, buys={total_buys}, "
            f"pp_avg={pp_avg_chosen:.2f}, pp_std={pp_std_chosen:.2f}, "
            f"stab_score={chosen.get('stab_score', 0):.3f}"
        )
        _upsert_best_group(chosen, sz)

    print("\n[Top-K Algo Usage (by size)]")
    for sz in (3, 4):
        bucket = topk_by_size.get(sz, [])
        if not bucket:
            print(f" size {sz}: no Top-K groups.")
            continue

        algo_counts = defaultdict(int)
        for e in bucket:
            for (algo_idx, _ident) in e["group"]:
                algo_counts[algo_idx] += 1

        print(f" size {sz}:")
        for algo_idx in range(len(PRUNED_TABLES)):
            c = algo_counts.get(algo_idx, 0)
            print(f"   Algo {algo_idx + 1}: {c} groups in Top-{TOPK_LIMIT_PER_SIZE}")

    _rows_for_db = []
    print("\n[ Algo Run Stats]")
    print("Algo | Variations | Timeouts | Total(s) | Avg_PP | Avg_Tickers")
    print("-----+------------+----------+----------+--------+------------")
    with _algo_stats_lock:
        for _ai in range(len(PRUNED_TABLES)):
            st = _algo_stats.get(_ai) or {"threads":0,"timeouts":0,"runtime_sec":0.0,"sum_pp":0.0,"sum_tickers":0.0}
            th = int(st["threads"])
            to = int(st["timeouts"])
            rt = float(st["runtime_sec"])
            ap = (st["sum_pp"]/th) if th > 0 else None
            at = (st["sum_tickers"]/th) if th > 0 else None
            ap_str = f"{ap:.2f}" if ap is not None else "NA"
            at_str = f"{at:.2f}" if at is not None else "NA"
            print(f"{_ai+1:>4} | {th:>10} | {to:>8} | {rt:>8.2f} | {ap_str:>6} | {at_str:>10}")
            # variations == threads
            _rows_for_db.append((_ai+1, PRUNED_TABLES[_ai], th, th, to, rt, ap, at))
          
    # DB table: algo_run_stats
    from datetime import datetime, timedelta
    try:
        from zoneinfo import ZoneInfo
        _tz = ZoneInfo("America/New_York")
    except Exception:
        _tz = None

    def _fmt_mdy(d):
        return f"{d.month}/{d.day}/{d.year}"

    def _fmt_run_dt(dt):
        m = dt.month
        d = dt.day
        y = dt.year
        h24 = dt.hour
        minute = dt.minute
        ampm = "AM" if h24 < 12 else "PM"
        h12 = h24 % 12
        if h12 == 0:
            h12 = 12
        return f"({m}/{d}/{y}_{h12}:{minute:02d} {ampm})"

    now_dt = datetime.now(_tz) if _tz else datetime.now()
    run_dt_text = _fmt_run_dt(now_dt)

    weekday_idx = now_dt.weekday()
    days_to_next_monday = (7 - weekday_idx) % 7
    if days_to_next_monday == 0:
        days_to_next_monday = 7
    monday = (now_dt + timedelta(days=days_to_next_monday)).date()
    target_dates_list = [_fmt_mdy(monday + timedelta(days=i)) for i in range(5)]
    target_dates_text = " | ".join(target_dates_list)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS algo_run_stats (
            week_number        INTEGER,
            run_dt             TEXT,
            target_dates       TEXT,
            algo_number        INTEGER,
            algo_table         TEXT,
            threads            INTEGER,
            variations         INTEGER,
            timeouts           INTEGER,
            total_runtime_sec  REAL,
            avg_pp             REAL,
            avg_tickers        REAL,
            PRIMARY KEY (week_number, algo_number)
        )
    """)

    cursor.execute(
        "SELECT week_number FROM algo_run_stats WHERE target_dates = ? ORDER BY week_number DESC LIMIT 1",
        (target_dates_text,)
    )
    _wk_row = cursor.fetchone()

    if _wk_row and _wk_row[0] is not None:
        next_week_number = int(_wk_row[0])
    else:
        cursor.execute("SELECT MAX(week_number) FROM algo_run_stats")
        _max_row = cursor.fetchone()
        next_week_number = (int(_max_row[0]) + 1) if (_max_row and _max_row[0] is not None) else 1

    _rows_with_meta = []
    for (_algo_number, _algo_table, th, vr, to, rt, ap, at) in _rows_for_db:
        _rows_with_meta.append((
            next_week_number,    
            run_dt_text,           
            target_dates_text,      
            _algo_number,
            _algo_table,
            th,                      # threads
            vr,                      # variations
            to,                      # timeouts
            rt,                      # total_runtime_sec
            ap,                      # avg_pp
            at                       # avg_tickers
        ))

    cursor.executemany(
        """
        INSERT OR REPLACE INTO algo_run_stats (
            week_number, run_dt, target_dates,
            algo_number, algo_table, threads, variations, timeouts,
            total_runtime_sec, avg_pp, avg_tickers
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        _rows_with_meta
    )
    conn.commit()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS extra_big_groups (
            identifier TEXT PRIMARY KEY,
            pp REAL,
            score REAL,
            parameters TEXT,
            ticker_count REAL,                -- avg shared buys per day
            algo_ticker_counts TEXT,
            median_pp REAL,
            second_worst_pp REAL,
            wilson_lb REAL,
            mad_pp REAL,
            iqr_pp REAL,
            cv_buys REAL,
            daily_pp_json TEXT,
            daily_buys_json TEXT,
            shared_tickers_by_day_json TEXT,  -- explicit lists of shared tickers per day
            group_size INTEGER,
            members_json TEXT,
            repeat_factor REAL                -- how aggressively consensus names are reused (>=1.0)
        )
    ''')

    def _rehydrate_share_for_db(group_members, ticker_set):
        if not group_members or not ticker_set:
            return []
        a0, ident0 = group_members[0]
        stocks0 = full_stocks_map[(a0, ident0)]
        tset = set(ticker_set)
        return [s for s in stocks0 if s.get("ticker") in tset]


    # Export bounded winners by size
    side_row = 7
    side_col = 166  # FJ

    try:
        sheet.range("J3").value = 100
        time.sleep(3)
    except Exception:
        pass
    def _rehydrate_shared_rows(group_members, ticker_set):
        if not group_members or not ticker_set:
            return []
        a0, ident0 = group_members[0]
        stocks0 = full_stocks_map[(a0, ident0)]
        tset = set(ticker_set)
        return [s for s in stocks0 if s.get("ticker") in tset]
    def _rehydrate_shared_rows_by_day(group_members, day_ticker_set, day_idx):
        if not group_members or not day_ticker_set:
            return []
        a0, ident0 = group_members[0]
        rows0 = full_stocks_map[(a0, ident0)]
        tset = set(day_ticker_set)
        return [r for r in rows0 if r.get("ticker") in tset and int(r.get("day", 0) or 0) == day_idx]


    def _paste_one_group_at(col, row, group_members, ticker_set, entry):
        ident = "_".join([ident for (_a, ident) in group_members])
        sheet.range((row + 1, col - 3)).value = ident
        sheet.range((row + 1, col - 2)).value = entry.get("pp")
        sheet.range((row + 1, col - 1)).value = entry.get("score")

        rc = row
        for a, ident in group_members:
            headers = param_headers_map.get(a, [])
            values  = params_map.get((a, ident), [])
            if headers:
                sheet.range((rc, col)).value = headers
            if values:
                sheet.range((rc + 1, col)).value = values
            rc += 3
        result_headers = ["Ticker", "Name", "Open", "Close", "ProfitFlag(1/0)"]
        sheet.range((rc, col)).value = result_headers

        # 5-day export section
        daily_pps = entry.get("daily_stats", [])
        daily_buys = entry.get("buy_stats", [])
        dates = target_dates_list

        for i in range(5):
            day_label = dates[i]
            pp_day = round(daily_pps[i], 2) if i < len(daily_pps) else 0
            buys_day = int(daily_buys[i]) if i < len(daily_buys) else 0
            day_set = (entry.get("consensus_sets_by_day") or [set()]*5)[i]
            shared_rows_day = _rehydrate_shared_rows_by_day(group_members, day_set, i+1)
            profits_day = sum(1 for s in shared_rows_day if int(s.get("is_profit", 0)) == 1)
            pp_row = round(100 * profits_day / buys_day, 2) if buys_day else 0

            sheet.range((rc + 1, col - 3)).value = f"{ident} - {day_label}"
            sheet.range((rc + 1, col - 2)).value = pp_day
            sheet.range((rc + 1, col - 1)).value = entry.get("score")
            sheet.range((rc + 1, col)).value = [f"Buys={buys_day}", f"Profits={profits_day}", f"PP={pp_row}"]
            rc += 1

            day_headers = ["Ticker", "Name", "Open", "Close", "ProfitFlag(1=close>open)"]
            sheet.range((rc, col)).value = day_headers

            if shared_rows_day:
                vals = []
                for s in shared_rows_day:
                    vals.append([
                        s.get("ticker", ""),
                        s.get("name", ""),
                        s.get("open", ""),
                        s.get("close", ""),
                        int(s.get("is_profit", 0)),
                    ])
                sheet.range((rc + 1, col)).value = vals
                rc += len(vals) + 1
        return rc

    print("\n[] Exporting best-by-size (3..4) to GE7...")
    for size in (3, 4):
        entries = best_by_size.get(size, [])
        if not entries:
            continue
        for entry in entries:
            end_row = _paste_one_group_at(
                side_col, side_row,
                entry["group"], entry["tickers"],
                entry
            )
            side_row = end_row + 2
    for k in list(best_by_size.keys()):
        best_by_size[k] = None
    wb.save()
  
    print(f"[] Total groups evaluated (counter) = {tested_groups_total}")
    print(f"[] Near-duplicate groups skipped by Jaccard = {dedupe_count}")

    try:
        pass
    finally:
        try:
            wb.save()
        except Exception as _e:
            print(f"[save warn] { _e }", flush=True)
        try:
            conn.commit()
        except Exception as _e:
            print(f"[db commit warn] { _e }", flush=True)

    print("\n[BUCKET SUMMARY] (disabled in exhaustive mode)")

# mp2 p2 end
if __name__ == "__main__":
    run_mp2_general_group_tester()
