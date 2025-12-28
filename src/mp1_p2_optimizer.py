"""
MP1 — Part 2: 5-day rank and parameter recommendation

Reads recent 5-day results from the Optuna database, scores the current 8 variations
per algorithm, and proposes the next 8 parameter sets using exploitation + exploration.
Writes the new parameter grid back into the Excel ControlSheet for the next run.

Input: optuna_data.db
Output: updated parameter rows in Excel
"""

import os
import json
import math
import random
import sqlite3
import xlwings as xw
import statistics as stats
from collections import Counter

# File paths
EXCEL_FILE = r"<REDACTED_PATH>/MakeMoneyExcel3.xlsm"
optuna_db_path = r"<REDACTED_PATH>/optuna_data.db"

import uuid
import time

# Worksheet Name
SHEET_NAME = "ControlSheet"

# Exponential weighting
EW_SCHEME_DEFAULT = "0.50,0.25,0.15,0.07,0.03"

# Gates
CONSISTENCY_HITS_GATE = 4          # min hits on each day
EXPORT_LB_MIN = 27.5               # Wilson LB floor for export
EXPORT_POOLED_BUYS_MIN = 600       # evidence floor for export

PRICE_START_COL = 59          # BG: [Date, Open, High, Low, Close, Volume]
PRICE_WIDTH     = 6
VAR_BLOCK_FIRST_COLS = [11, 17, 23, 29, 35, 41, 47, 53] 

# Parameter settings
PARAM_RANGES = {
    "MACD": [(5, 10, 1), (12, 18, 1), (5, 10, 1), (-2.5, 6, 0.05), (1, 3, 1), ["None", "Vol"], (3, 25, 0.1), (1, 10, 1), (0.00, 0.20, 0.01)],
    "EMA": [(6, 15, 1), (1, 3, 1), ["None", "Vol"], (1, 30, 1), (1, 10, 1)],
    "RSI": [(8, 14, 1), (20, 35, 1), ["None", "Vol"], (1, 30, 1), (1, 10, 1), (1, 1.2, 0.05), (1, 1.3, 0.05)],
    "Breakout": [(-50, -5, 2.5), ["None", "Vol"], (1, 30, 1), (1, 10, 1), (8, 18, 1)],
    "ADX": [(10, 14, 1), (10, 30, 1), (60, 100, 1), ["None", "Vol"], (1, 30, 1), (1, 10, 1), ["ADX Only", "ADX + DI+ > DI-"]],
    "Volatility": [(10, 19, 1), (1, 3, 0.1), (3, 7, 0.1), ["None", "Vol"], (1, 30, 1), (1, 10, 1), (0, 1, 0.1), (1, 4, 0.2)],
    "SMA": [(6, 12, 1), (2, 8, 1), ["None", "Vol"], (1, 30, 1), (1, 10, 1)],
    "Bollinger_Bands": [(6, 17, 1), (-7.5, 2.5, 0.05), ["None", "Vol"], (1, 30, 1), (1, 10, 1), (1, 3, 1)],
    "EMA_MACD_Combo": [(6, 15, 1), (16, 30, 1), (5, 12, 1), (-0.8, 1, 0.05), ["None", "Vol"], (1, 30, 1), (1, 10, 1), (-0.8, 2, 0.05), (2, 6, 1)],
    "RSI_Bollinger": [(8, 14, 1), (-7.5, 2.5, 0.05), (1, 3, 1), ["None", "Vol"], (1, 30, 1), (1, 10, 1), ["None", "RSI"], (20, 35, 1), (1, 1.2, 0.05), (1, 1.3, 0.05)]
}

CONTROL_ORDER = [
    "MACD","EMA","RSI","Breakout",
    "ADX","Volatility",
    "SMA","Bollinger_Bands","EMA_MACD_Combo","RSI_Bollinger"
]

SHEET_FOR_ALGO = {
    "MACD": "MACD",
    "EMA": "EMA",
    "RSI": "RSI",
    "Breakout": "Breakout",
    "ADX": "ADX",
    "Volatility": "Volatility",
    "SMA": "SMA",
    "Bollinger_Bands": "Bollinger_Bands",
    "EMA_MACD_Combo": "EMA_MACD_Combo",
    "RSI_Bollinger": "RSI_Bollinger",
}

ALGO_PARAM_COUNTS = {
    "MACD": 9,
    "EMA": 5,
    "RSI": 7,
    "Breakout": 5,
    "ADX": 6,
    "Volatility": 6,
    "SMA": 5,
    "Bollinger_Bands": 6,
    "EMA_MACD_Combo": 9,
    "RSI_Bollinger": 10,
}

def _row_for_algo(algo_name: str) -> int:
    """Return the starting row for this algo in ControlSheet (row 4, 15, 26, ...)."""
    idx = CONTROL_ORDER.index(algo_name)  
    return 4 + idx * 11

def _is_profit_cellblock(block6):
    for v in block6:
        if isinstance(v,(int,float)) and float(v) > 0: return True
        s = str(v).strip().upper()
        if "PROFIT" in s: return True
    return False

def ensure_algo_tables_schema():
    conn = sqlite3.connect(optuna_db_path)
    try:
        cur = conn.cursor()

        cur.execute("PRAGMA journal_mode=WAL;")
        cur.execute("PRAGMA synchronous=NORMAL;")
        cur.execute("PRAGMA temp_store=MEMORY;")
        cur.execute("PRAGMA mmap_size=268435456;") 
        cur.execute("PRAGMA cache_size=-131072;")   
        cur.execute("PRAGMA page_size=4096;")
      
        base_cols = [
            ("id","INTEGER PRIMARY KEY"),
            ("run_id","TEXT"),
            ("params","TEXT"),
            ("trial_number","INTEGER"),
            ("variation_number","INTEGER"),

            ("day1_buys","INTEGER"),("day1_hits","INTEGER"),("day1_pp","REAL"),("day1_buys_json","TEXT"),
            ("day2_buys","INTEGER"),("day2_hits","INTEGER"),("day2_pp","REAL"),("day2_buys_json","TEXT"),
            ("day3_buys","INTEGER"),("day3_hits","INTEGER"),("day3_pp","REAL"),("day3_buys_json","TEXT"),
            ("day4_buys","INTEGER"),("day4_hits","INTEGER"),("day4_pp","REAL"),("day4_buys_json","TEXT"),
            ("day5_buys","INTEGER"),("day5_hits","INTEGER"),("day5_pp","REAL"),("day5_buys_json","TEXT"),

            ("pooled_buys","INTEGER"),("pooled_hits","INTEGER"),("pooled_pp","REAL"),
            ("median_pp_5d","REAL"),("pp_mad_5d","REAL"),
            ("buycount_med_5d","REAL"),("buycount_mad_5d","REAL"),
            ("pp_iqr_5d","REAL"),("buycount_cv_5d","REAL"),

            ("wilson_lb_5d","REAL"),

            ("ew_scheme","TEXT"),("ew_pp_5d","REAL"),("ew_hits_5d","REAL"),

            ("repeat_ticker_rate_5d","REAL"),("top_10_ticker_share_5d","REAL"),
            ("avg_buy_price_5d","REAL"),("median_buy_price_5d","REAL"),("avg_buy_volume_5d","REAL"),

            ("last_day_pp","REAL"),("last_day_buys","INTEGER"),("last_day_hits","INTEGER"),

            ("min_daily_hits","INTEGER"),("passed_consistency_gate","INTEGER"),("passed_export_gate","INTEGER")
        ]

        # tables to create
        for algo in CONTROL_ORDER:
            cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (algo,))
            exists = cur.fetchone() is not None
            if not exists:
                cols_sql = ", ".join([f"{n} {t}" for n,t in base_cols])
                cur.execute(f"CREATE TABLE {algo} ({cols_sql})")
            else:
                cur.execute(f"PRAGMA table_info({algo})")
                have = {row[1] for row in cur.fetchall()}
                want = [n for n,_ in base_cols]

                legacy = [c for c in ("profit_percentage","stocks_bought") if c in have]
                missing = [n for n,_ in base_cols if n not in have]

                if legacy:
                    tmp = f"{algo}__new"
                    cols_sql = ", ".join([f"{n} {t}" for n,t in base_cols])
                    cur.execute(f"DROP TABLE IF EXISTS {tmp}")
                    cur.execute(f"CREATE TABLE {tmp} ({cols_sql})")
                    overlap = [n for n,_ in base_cols if n in have]
                    if overlap:
                        col_list = ", ".join(overlap)
                        cur.execute(f"INSERT INTO {tmp} ({col_list}) SELECT {col_list} FROM {algo}")
                    cur.execute(f"DROP TABLE {algo}")
                    cur.execute(f"ALTER TABLE {tmp} RENAME TO {algo}")
                    have = set(want)

                if missing and not legacy:
                    for col, ctype in base_cols:
                        if col not in have:
                            cur.execute(f"ALTER TABLE {algo} ADD COLUMN {col} {ctype}")

            cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{algo}_params_var ON {algo}(params, variation_number)")
            cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{algo}_params_only ON {algo}(params)")
            cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{algo}_id_desc ON {algo}(id DESC)")
            cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{algo}_var_id ON {algo}(variation_number, id)")
            cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{algo}_params_id ON {algo}(params, id)")
            cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{algo}_pv_idd ON {algo}(params, variation_number, id DESC)")
            cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{algo}_params_var_legacy ON {algo}(variation_number, params)")
        conn.commit()
        cur.execute("ANALYZE;")
        conn.commit()
    finally:
        conn.close()
      
# P2 Logic
from statistics import StatisticsError

RND = random.Random(int.from_bytes(os.urandom(16), "little"))

# Internal schema indices (parameter positions)
ADX_PERIOD, ADX_THRESH, ADX_VOL, ADX_LOOK, ADX_W, ADX_STRAT = range(6)

VOL_F1, VOL_A, VOL_B, VOL_VOL, VOL_LB, VOL_W, VOL_T = range(7)
def _internal_ranges_for(algo_name):
    """
    Return ranges aligned to the INTERNAL parameter order used in this module.
    Only used for diversity/normalization/dedup logic.
    """
    if algo_name == "ADX":
        (p_lo,p_hi,p_step) = PARAM_RANGES["ADX"][0]
        thr_lo, thr_hi, thr_step = 10, 100, 1
        vol_choices   = PARAM_RANGES["ADX"][3]
        look_lo,look_hi,_ = PARAM_RANGES["ADX"][4]
        w_lo,w_hi,_  = PARAM_RANGES["ADX"][5]
        strat_choices = PARAM_RANGES["ADX"][6]
        return [
            (p_lo,p_hi,p_step),        # period
            (thr_lo,thr_hi,thr_step),  # single threshold axis for internals
            vol_choices,               # Vol/None
            (look_lo,look_hi,1),       # lookback
            (w_lo,w_hi,1),             # weight
            strat_choices              # strategy
        ]
    else:
        return PARAM_RANGES[algo_name]

# Adapters: Excel <-> internal param shapes
def _to_internal_params(algo_name, p):
    if not isinstance(p, list):
        p = [] if p is None else [p]

    if algo_name == "Volatility":
        out = [None]*7
        out[VOL_F1] = _safe_float(p[0], 12)
        ab = p[1] if len(p) > 1 else None
        a_val, b_val = None, None
        if isinstance(ab, str) and "-" in ab:
            try:
                s1, s2 = ab.split("-", 1)
                a_val, b_val = float(s1), float(s2)
            except:
                a_val, b_val = None, None
        elif isinstance(ab, (list, tuple)) and len(ab) == 2:
            a_val, b_val = _safe_float(ab[0], None), _safe_float(ab[1], None)
        else:
            a_val = _safe_float(p[1] if len(p) > 1 else None, None)
            b_val = _safe_float(p[2] if len(p) > 2 else None, None)
        # fallback to midpoints if missing
        a_val = a_val if a_val is not None else 2.0
        b_val = b_val if b_val is not None else 5.0
        if a_val > b_val:
            a_val, b_val = b_val, a_val
        out[VOL_A], out[VOL_B] = round(a_val, 1), round(b_val, 1)

        # vol flag
        out[VOL_VOL] = p[2] if len(p) > 2 and str(p[2]) in ("Vol", "None") else \
                       (p[3] if len(p) > 3 and str(p[3]) in ("Vol", "None") else "None")
        out[VOL_LB] = int(_safe_float(p[3] if len(p) > 3 and isinstance(p[3], (int,float)) else (p[4] if len(p) > 4 else 10), 10))
        out[VOL_W]  = int(_safe_float(p[4] if len(p) > 4 and isinstance(p[4], (int,float)) else (p[5] if len(p) > 5 else 5), 5))
      
        # Trend threshold
        out[VOL_T]  = round(_safe_float(p[5] if len(p) > 5 else (p[6] if len(p) > 6 else 1.0), 1.0), 2)
        return out

    if algo_name == "ADX":
        out = [None]*6
        out[ADX_PERIOD] = int(_safe_float(p[0] if len(p) > 0 else None, 12))
        out[ADX_THRESH] = int(_safe_float(p[1] if len(p) > 1 else None, 30))
        out[ADX_VOL]    = p[2] if len(p) > 2 and str(p[2]) in ("Vol", "None") else "None"
        out[ADX_LOOK]   = int(_safe_float(p[3] if len(p) > 3 else None, 10))
        out[ADX_W]      = int(_safe_float(p[4] if len(p) > 4 else None, 5))
        out[ADX_STRAT]  = p[5] if len(p) > 5 and str(p[5]) in ("ADX Only","ADX + DI+ > DI-") else "ADX Only"
        return out
      
    return list(p)

def _to_excel_params(algo_name, pi):
    if algo_name == "Volatility":
        F1  = int(_safe_float(pi[VOL_F1], 12))
        a   = round(_safe_float(pi[VOL_A], 2.0), 1)
        b   = round(_safe_float(pi[VOL_B], 5.0), 1)
        if a > b: a, b = b, a
        ab  = f"{a:.1f}-{b:.1f}"
        vol = pi[VOL_VOL] if str(pi[VOL_VOL]) in ("Vol","None") else "None"
        LB  = int(_safe_float(pi[VOL_LB], 10))
        W   = int(_safe_float(pi[VOL_W], 5))
        T   = round(_safe_float(pi[VOL_T], 1.0), 2)
        return [F1, ab, vol, LB, W, T]
    if algo_name == "ADX":
        # [period, threshold, vol_flag, lookback, weight, strategy]
        period = int(_safe_float(pi[ADX_PERIOD], 12))
        th     = int(_safe_float(pi[ADX_THRESH], 30))
        vol    = pi[ADX_VOL] if str(pi[ADX_VOL]) in ("Vol","None") else "None"
        look = max(1, min(30, int(_safe_float(pi[ADX_LOOK], 10))))
        w    = max(1, min(10,  int(_safe_float(pi[ADX_W], 5))))

        strat  = pi[ADX_STRAT] if str(pi[ADX_STRAT]) in ("ADX Only","ADX + DI+ > DI-") else "ADX Only"
        return [period, th, vol, look, w, strat]
    return list(pi)
  
def _db_param_json_for_query(algo_name, internal_params):
    _pi = _to_internal_params(algo_name, internal_params)
    _excel_row = _to_excel_params(algo_name, _pi)
    return json.dumps(_excel_row)


def _canonical_params_json_for_db(algo_name, internal_params):
    _pi = _to_internal_params(algo_name, internal_params)
    _excel_row = _to_excel_params(algo_name, _pi)
    return json.dumps(_excel_row)

def _grid_L1_for_internal(a, b, ranges):
    s = 0.0; n = 0
    for (va, vb, spec) in zip(a, b, ranges):
        if isinstance(spec, list) and not all(isinstance(x,(int,float)) for x in spec):
            s += 0.0 if va == vb else 1.0
            n += 1
        else:
            lo,hi,step = spec
            try:
                fa = float(va); fb = float(vb)
            except:
                continue
            ran = (hi - lo) if hi > lo else 1.0
            s += abs(fa - fb)/ran
            n += 1
    return s / max(1, n)

_P2_L1_MIN = {
    "ADX": 0.25,
    "Volatility": 0.25,
    "_default": 0.10
}

def _dedup_and_diversify_batch(algo_name, batch_internal):
    # Remove duplicates
    ranges = _internal_ranges_for(algo_name)
    keep = []
    seen_canon = set()
    l1_min = _P2_L1_MIN.get(algo_name, _P2_L1_MIN["_default"])

    for cand in batch_internal:
        canon = _canonical_params_json_for_db(algo_name, cand)
        if canon in seen_canon:
            continue
        if any(_grid_L1_for_internal(prev, cand, ranges) < l1_min for prev in keep):
            continue
        seen_canon.add(canon)
        keep.append(_pad_params_to_ranges(cand, ranges))

    while len(keep) < min(8, len(batch_internal)):
        fallback = []
        for spec in ranges:
            if isinstance(spec, list) and not all(isinstance(x,(int,float)) for x in spec):
                fallback.append(random.choice(spec))
            else:
                lo,hi,step = spec
                if step:
                    steps = int(round((hi - lo)/step))
                    v = lo + step * random.randint(0, max(0, steps))
                    if step == 1 and abs(v - round(v)) < 1e-9:
                        v = int(round(v))
                else:
                    v = lo + (hi - lo)*random.random()
                fallback.append(v)
        canon = _canonical_params_json_for_db(algo_name, fallback)
        if canon in seen_canon or any(_grid_L1_for_internal(prev, fallback, ranges) < l1_min for prev in keep):
            continue
        seen_canon.add(canon)
        keep.append(_pad_params_to_ranges(fallback, ranges))

    return keep[:8]

def _safe_float(v, default=None):
    try:
        return float(v)
    except:
        return default


def _seed_from_ranges(ranges):
    seed = []
    for spec in ranges:
        if isinstance(spec, list) and not all(isinstance(x,(int, float)) for x in spec):
            seed.append(spec[0] if spec else None)
        else:
            lo, hi, step = spec
            mid = (lo + hi)/2.0
            if step == 1:
                mid = int(round(mid))
            seed.append(mid)
    return seed

def _pad_params_to_ranges(params, ranges):
    out = []
    for i, spec in enumerate(ranges):
        v = params[i] if i < len(params) else None
        if isinstance(spec, list) and not all(isinstance(x,(int,float)) for x in spec):
            choices = spec
            out.append(v if v in choices else (choices[0] if choices else v))
        else:
            lo, hi, step = spec
            try:
                fv = float(v)
            except:
                fv = (lo + hi)/2.0
            fv = max(lo, min(hi, fv))
            if step:
                k = round((fv - lo)/step)
                fv = lo + k*step
            if step == 1 and abs(fv - round(fv)) < 1e-9:
                fv = int(round(fv))
            out.append(fv)
    return out

def _normalize_for_key(vals, ranges):
    out = []
    for v, spec in zip(vals, ranges):
        if isinstance(spec, list):
            out.append(v)
        else:
            lo, hi, step = spec
            try:
                fv = float(v)
                if step:
                    k = round((fv - lo)/step)
                    fv = lo + k*step
                out.append(int(round(fv)) if step == 1 else float(fv))
            except:
                out.append(v)
    return out

def _normalize_history_set(history_json_set, ranges):
    out = set()
    for s in history_json_set:
        try:
            vals = json.loads(s)
            vals_n = _normalize_for_key(vals, ranges)
            out.add(json.dumps(vals_n))
        except:
            out.add(s)
    return out

def _db_current_params_for_algo(conn, algo_table, algo_name):
    ranges = PARAM_RANGES[algo_name]

    def _seed():
        return _seed_from_ranges(ranges)

    out = [None]*8
    cur = conn.cursor()
    cur.execute(f"""
        SELECT variation_number, params
        FROM {algo_table}
        WHERE variation_number BETWEEN 1 AND 8
        ORDER BY id DESC
    """)
    seen = set()
    for var, params_json in cur.fetchall():
        if var in seen:
            continue
        try:
            p = json.loads(params_json)
        except Exception:
            p = None
        out[var-1] = p
        seen.add(var)
        if len(seen) == 8:
            break

    for i in range(8):
        if not isinstance(out[i], list):
            out[i] = _seed()
        out[i] = _pad_params_to_ranges(out[i], ranges)
        out[i] = _to_internal_params(algo_name, out[i])

    return out
def _fetch_5d_metrics_for(conn, algo_table, params_list, variation_number):
    K = 48
    cur = conn.cursor()
    params_json = _db_param_json_for_query(algo_table, params_list)

    cur.execute(f"""
        SELECT
            id,
            pooled_buys, pooled_hits, pooled_pp,
            median_pp_5d, pp_mad_5d, pp_iqr_5d,
            buycount_med_5d, buycount_mad_5d, buycount_cv_5d,
            wilson_lb_5d,
            ew_pp_5d, ew_hits_5d,
            repeat_ticker_rate_5d, top_10_ticker_share_5d,
            avg_buy_price_5d, median_buy_price_5d, avg_buy_volume_5d,
            last_day_pp, last_day_buys, last_day_hits,
            min_daily_hits, passed_consistency_gate, passed_export_gate
        FROM {algo_table}
        WHERE params = ? AND variation_number = ?
        ORDER BY id DESC
        LIMIT {K}
    """,(params_json, variation_number))
    rows = cur.fetchall()
    if not rows:
        return None
    w = 1.0
    decay = 0.85
    tot_w = 0.0
    agg = {
        "pooled_buys":0.0, "pooled_hits":0.0, "pooled_pp":0.0,
        "median_pp_5d":[], "pp_mad_5d":[], "pp_iqr_5d":[],
        "buycount_med_5d":[], "buycount_mad_5d":[], "buycount_cv_5d":[],
        "wilson_lb_5d":0.0, "ew_pp_5d":0.0, "ew_hits_5d":0.0,
        "repeat_ticker_rate_5d":0.0, "top_10_ticker_share_5d":0.0,
        "avg_buy_price_5d":0.0, "median_buy_price_5d":0.0, "avg_buy_volume_5d":0.0,
        "last_day_pp":0.0, "last_day_buys":0.0, "last_day_hits":0.0,
        "min_daily_hits":0.0, "pass_gate":0.0, "pass_export":0.0
    }
    def f(x):
        try: return float(x or 0.0)
        except: return 0.0

    for r in rows:
        (_id,
         pb, ph, pp,
         mpp, madp, iqrp,
         mbu, mabu, cvb,
         wlb,
         ewpp, ewh,
         rep, top10,
         apx, mpx, avol,
         lpp, lbu, lhi,
         minh, pgate, pexp) = r

        agg["pooled_buys"] += w * f(pb)
        agg["pooled_hits"] += w * f(ph)
        agg["pooled_pp"]   += w * f(pp)

        agg["median_pp_5d"].append(f(mpp))
        agg["pp_mad_5d"].append(f(madp))
        agg["pp_iqr_5d"].append(f(iqrp))

        agg["buycount_med_5d"].append(f(mbu))
        agg["buycount_mad_5d"].append(f(mabu))
        agg["buycount_cv_5d"].append(f(cvb))

        agg["wilson_lb_5d"] += w * f(wlb)
        agg["ew_pp_5d"]     += w * f(ewpp)
        agg["ew_hits_5d"]   += w * f(ewh)

        agg["repeat_ticker_rate_5d"] += w * f(rep)
        agg["top_10_ticker_share_5d"] += w * f(top10)

        agg["avg_buy_price_5d"]    += w * f(apx)
        agg["median_buy_price_5d"] += w * f(mpx)
        agg["avg_buy_volume_5d"]   += w * f(avol)

        agg["last_day_pp"]   += w * f(lpp)
        agg["last_day_buys"] += w * f(lbu)
        agg["last_day_hits"] += w * f(lhi)

        agg["min_daily_hits"] += w * f(minh)
        agg["pass_gate"]      += w * (1 if pgate else 0)
        agg["pass_export"]    += w * (1 if pexp else 0)

        tot_w += w
        w *= decay

    import statistics as stats
    def rmed(a):
        return float(stats.median(a)) if a else 0.0
    def rmad(a):
        if not a: return 0.0
        m = stats.median(a)
        return float(stats.median([abs(x - m) for x in a]))
    def riqr(a):
        if not a: return 0.0
        try:
            q = stats.quantiles(a, n=4, method="inclusive")
            return float(q[2] - q[0])
        except: return 0.0

    out = {
        "pooled_buys": agg["pooled_buys"]/tot_w if tot_w else 0.0,
        "pooled_hits": agg["pooled_hits"]/tot_w if tot_w else 0.0,
        "pooled_pp":   0.0,

        "median_pp_5d":   rmed(agg["median_pp_5d"]),
        "pp_mad_5d":      rmad(agg["pp_mad_5d"]),
        "pp_iqr_5d":      riqr(agg["pp_iqr_5d"]),
        "buycount_med_5d":rmed(agg["buycount_med_5d"]),
        "buycount_mad_5d":rmad(agg["buycount_mad_5d"]),
        "buycount_cv_5d": rmed(agg["buycount_cv_5d"]),

        "wilson_lb_5d": agg["wilson_lb_5d"]/tot_w if tot_w else 0.0,
        "ew_pp_5d":     agg["ew_pp_5d"]/tot_w     if tot_w else 0.0,
        "ew_hits_5d":   agg["ew_hits_5d"]/tot_w   if tot_w else 0.0,

        "repeat_ticker_rate_5d": agg["repeat_ticker_rate_5d"]/tot_w if tot_w else 0.0,
        "top_10_ticker_share_5d":agg["top_10_ticker_share_5d"]/tot_w if tot_w else 0.0,

        "avg_buy_price_5d":    agg["avg_buy_price_5d"]/tot_w if tot_w else 0.0,
        "median_buy_price_5d": agg["median_buy_price_5d"]/tot_w if tot_w else 0.0,
        "avg_buy_volume_5d":   agg["avg_buy_volume_5d"]/tot_w if tot_w else 0.0,

        "last_day_pp":   agg["last_day_pp"]/tot_w     if tot_w else 0.0,
        "last_day_buys": agg["last_day_buys"]/tot_w   if tot_w else 0.0,
        "last_day_hits": agg["last_day_hits"]/tot_w   if tot_w else 0.0,

        "min_daily_hits": agg["min_daily_hits"]/tot_w if tot_w else 0.0,
        "pass_gate":      1 if agg["pass_gate"]/tot_w >= 0.5 else 0,
        "pass_export":    1 if agg["pass_export"]/tot_w >= 0.5 else 0
    }
    ph, pb = out["pooled_hits"], out["pooled_buys"]
    out["pooled_pp"] = (100.0 * ph / pb) if pb > 0 else 0.0
    return out

def _lexi_score(m):
    if not m: return None

    gate = 1 if m["pass_gate"] else 0
    if gate == 0:
        return -1e30

    lb   = float(m["wilson_lb_5d"])
    med  = float(m["median_pp_5d"])
    mad  = float(m.get("pp_mad_5d", 0.0))
    iqr  = float(m.get("pp_iqr_5d", 0.0))
    bcv  = float(m.get("buycount_cv_5d", 0.0))
    ev   = int(m.get("pooled_buys", 0))
    ewpp = float(m.get("ew_pp_5d", 0.0))
    lpp  = float(m.get("last_day_pp", 0.0))
    rep  = float(m.get("repeat_ticker_rate_5d", 0.0))
    top10= float(m.get("top_10_ticker_share_5d", 0.0))

    stab_penalty   = (mad + iqr) * 0.5 + bcv * 0.25
    # concentration penalty
    conc_penalty   = (max(0.0, rep - 85.0) * 0.05) + (top10 * 0.35)
    recency_bonus  = (ewpp * 0.6) + (lpp * 0.4)
    evidence_bonus = min(10000, ev) * 0.01
    base = (1e12) + (lb * 1e9) + (med * 1e6)
    soft = (recency_bonus * 1e3) + (evidence_bonus * 1e2) - (stab_penalty * 1e2) - (conc_penalty * 1e2)
    return base + soft

def _ablation_around(best_params, algo_name, k=8, rnd=None):
    if rnd is None: rnd = RND
    ranges = PARAM_RANGES[algo_name]
    base = _pad_params_to_ranges(best_params, ranges)
    out = []

    idxs = list(range(len(ranges)))
    rnd.shuffle(idxs)
    idxs = (idxs * ((k + len(idxs) - 1)//len(idxs)))[:k]

    for j in idxs:
        cand = base[:]
        spec = ranges[j]
        if isinstance(spec, list) and not all(isinstance(x,(int,float)) for x in spec):
            choices = [c for c in spec if c != cand[j]] or spec
            cand[j] = rnd.choice(choices)
        else:
            lo, hi, step = spec
            if step:
                jumps = rnd.choice([1,2,3])
                direction = rnd.choice([-1,1])
                try: v = float(cand[j])
                except: v = (lo + hi)/2.0
                v = v + direction * jumps * step
                nsteps = round((v - lo)/step)
                v = lo + nsteps*step
                v = max(lo, min(hi, v))
                if step == 1: v = int(v)
                cand[j] = v
            else:
                span = (hi - lo)
                cand[j] = max(lo, min(hi, float(cand[j]) + 0.15*span*(2*rnd.random()-1)))

        out.append(cand)
    return out[:k]
  
def _all_historical_params_for_algo(conn, algo_table):
    cur = conn.cursor()
    cur.execute(f"SELECT params FROM {algo_table} ORDER BY id DESC LIMIT 2000")
    return {row[0] for row in cur.fetchall()}

def _collect_historical_scored_params(conn, algo_table, algo_name):
    MAX_DISTINCT = 600
    cur = conn.cursor()

    cur.execute(f"""
        SELECT p.params
        FROM (
            SELECT params, MAX(id) AS mx
            FROM {algo_table}
            GROUP BY params
            ORDER BY mx DESC
            LIMIT {MAX_DISTINCT}
        ) AS p
        ORDER BY p.mx DESC
    """)
    param_jsons = [row[0] for row in cur.fetchall()]

    out = []
    for js in param_jsons:
        try:
            p_ext = json.loads(js)
            p_int = _to_internal_params(algo_name, p_ext)
        except:
            continue

        best = None
        for v in range(1, 9):
            m = _fetch_5d_metrics_for(conn, algo_table, p_int, v)
            if m is None:
                continue
            s = _lexi_score(m)
            if best is None or s > best:
                best = s

        if best is not None:
            out.append((p_int, float(best)))

    out.sort(key=lambda t: t[1], reverse=True)
    return out

def _clip(val, lo, hi, step=None):
    v = max(lo, min(hi, val))
    if step:
        k = round((v - lo)/step)
        v = lo + k*step
        v = max(lo, min(hi, v))
    return v

def _mutate_numeric(val, lo, hi, step, scale=0.25):
    width = (hi - lo) * float(scale)
    if step is not None:
        width = max(step, width)
    cand = float(val) + (width * (2*RND.random() - 1))
    return _clip(cand, lo, hi, step)

def _coerce_for_excel(algo_name, params_list):
    target_len = ALGO_PARAM_COUNTS.get(algo_name, len(params_list))
    out = list(params_list[:])

    if algo_name == "ADX":
        if len(out) < 6:
            out += [None] * (6 - len(out))
        try:
            out[1] = int(round(float(out[1])))
        except:
            out[1] = 30
        if str(out[2]).strip() not in ("Vol", "None"):
            out[2] = "None"
        try:
            if out[3] is None:
                out[3] = 10
            else:
                out[3] = int(round(float(out[3])))
        except:
            pass

        try: out[4] = int(round(float(out[4])))
        except: out[4] = 5
        if str(out[5]).strip() not in ("ADX Only","ADX + DI+ > DI-"):
            out[5] = "ADX Only"

    elif algo_name == "Volatility":
        if len(out) < 6:
            out += [None] * (6 - len(out))
        try:
            out[0] = int(round(float(out[0])))
        except:
            out[0] = 12
        ab = out[1]
        try:
            if isinstance(ab, str) and "-" in ab:
                a_str, b_str = ab.split("-", 1)
                a = float(a_str); b = float(b_str)
            else:
                a = float(ab[0]); b = float(ab[1])
            if a > b: a, b = b, a
            out[1] = f"{a:.1f}-{b:.1f}"
        except:
            if out[1] is None or out[1] in ("", [], [None, None]):
                out[1] = "2.0-5.0"

        if str(out[2]).strip() not in ("Vol", "None"):
            out[2] = "None"
        try:
            if out[3] is None:
                out[3] = 10
            else:
                out[3] = int(round(float(out[3])))
        except:
            pass

        try: out[4] = int(round(float(out[4])))
        except: out[4] = 5
        try:
            tt = float(out[5])
            out[5] = round(tt, 1)
        except:
            out[5] = 1.0

    out = out[:target_len]
    if len(out) < target_len:
        out += [None] * (target_len - len(out))
    return out

def _ensure_min_uniques(batch, ranges, min_uniques=3, rnd=None):
    if not batch: return
    if rnd is None:
        rnd = RND
    rows = len(batch); cols = len(batch[0])
    for j in range(cols):
        spec = ranges[j] if j < len(ranges) else None
        col_vals = [row[j] for row in batch]
        uniq = {json.dumps(v, sort_keys=True) for v in col_vals}
        if isinstance(spec, list) and not all(isinstance(x,(int,float)) for x in spec):
            if len(uniq) < 2 and len(spec) > 1:
                picks = rnd.sample(range(rows), k=min(2, rows))
                for i in picks:
                    choices = [c for c in spec if c != batch[i][j]]
                    if choices:
                        batch[i][j] = rnd.choice(choices)
        else:
            if spec and len(uniq) < min_uniques:
                lo,hi,step = spec
                need = min_uniques - len(uniq)
                idxs = rnd.sample(range(rows), k=min(need, rows))
                for i in idxs:
                    try:
                        v = float(batch[i][j])
                    except:
                        v = (lo + hi)/2.0
                    if step:
                        delta = step * rnd.choice([1,2]) * rnd.choice([-1,1])
                        nv = max(lo, min(hi, v + delta))
                        if step == 1:
                            nv = int(nv)
                    else:
                        nv = max(lo, min(hi, v + 0.07*(hi-lo)*(2*rnd.random()-1)))
                    batch[i][j] = nv
def _propose_for_ADX(best_params, k=8):
    (p_lo,p_hi,p_step) = PARAM_RANGES["ADX"][0]
    (t_lo1,t_hi1,_)    = PARAM_RANGES["ADX"][1]
    (t_lo2,t_hi2,_)    = PARAM_RANGES["ADX"][2]
    look_lo,look_hi,_  = PARAM_RANGES["ADX"][4]
    w_lo,w_hi,_        = PARAM_RANGES["ADX"][5]
    strat_choices      = PARAM_RANGES["ADX"][6]
    vol_choices        = PARAM_RANGES["ADX"][3]

    base = best_params[:]
    out = []
    bands = [("low",(t_lo1,t_hi1)),("high",(t_lo2,t_hi2))]
    for _, (lo, hi) in bands:    
        for _ in range(max(1, k // 2)):
            period = int(_clip(round(_mutate_numeric(base[0], p_lo, p_hi, p_step, scale=0.6)), p_lo, p_hi))
            thresh = int(_clip(round(_mutate_numeric(base[1], lo, hi, 1, scale=0.6)), lo, hi))
            vol    = RND.choice(vol_choices)  # alternate Vol/None when needed
            look   = int(_clip(round(_mutate_numeric(_safe_float(base[3], 15), look_lo, look_hi, 1, scale=0.8)), look_lo, look_hi))
            weight = int(_clip(round(_mutate_numeric(_safe_float(base[4], 5), w_lo, w_hi, 1, scale=0.8)), w_lo, w_hi))
            strat  = RND.choice(strat_choices)  # alternate both strategies
            out.append([period, thresh, vol, look, weight, strat])
    _ensure_min_uniques(out, _internal_ranges_for("ADX"), min_uniques=3, rnd=RND)

    if len({p[ADX_LOOK] for p in out}) < 3:
        for i in range(min(4, len(out))):
            out[i][ADX_LOOK] = RND.randint(look_lo, look_hi)
    if len({p[ADX_W] for p in out}) < 3:
        for i in range(min(4, len(out))):
            out[-1 - i][ADX_W] = RND.randint(w_lo, w_hi)
    return out[:k]

def _force_batch_diversity(algo_name, batch):
    rnd = RND
    if not batch:
        return

    if algo_name == "ADX":
        def band(t):
            try:
                x = int(round(float(t)))
                if 10 <= x <= 30: return "low"
                if 60 <= x <= 100: return "high"
            except: pass
            return "other"

        lows  = [i for i,p in enumerate(batch) if band(p[ADX_THRESH]) == "low"]
        highs = [i for i,p in enumerate(batch) if band(p[ADX_THRESH]) == "high"]
        if not lows and highs:
            for i in highs[:max(1, len(highs)//2)]:
                batch[i][ADX_THRESH] = rnd.randint(10, 30)
        if not highs and lows:
            for i in lows[:max(1, len(lows)//2)]:
                batch[i][ADX_THRESH] = rnd.randint(60, 100)

        vols = [i for i,p in enumerate(batch) if str(p[ADX_VOL]) == "Vol"]
        nones= [i for i,p in enumerate(batch) if str(p[ADX_VOL]) == "None"]
        if not vols and nones:
            for i in nones[:max(1, len(nones)//2)]: batch[i][ADX_VOL] = "Vol"
        if not nones and vols:
            for i in vols[:max(1, len(vols)//2)]:  batch[i][ADX_VOL] = "None"

        s1 = [i for i,p in enumerate(batch) if str(p[ADX_STRAT]) == "ADX Only"]
        s2 = [i for i,p in enumerate(batch) if str(p[ADX_STRAT]) == "ADX + DI+ > DI-"]
        if not s1 and s2:
            for i in s2[:max(1, len(s2)//2)]: batch[i][ADX_STRAT] = "ADX Only"
        if not s2 and s1:
            for i in s1[:max(1, len(s1)//2)]: batch[i][ADX_STRAT] = "ADX + DI+ > DI-"

        for p in batch:
            p[ADX_PERIOD] = int(_safe_float(p[ADX_PERIOD], 12))
            p[ADX_THRESH] = int(_safe_float(p[ADX_THRESH], 30))
            p[ADX_LOOK]   = int(_safe_float(p[ADX_LOOK],   10))
            p[ADX_W]      = int(_safe_float(p[ADX_W],      5))
            if str(p[ADX_VOL]) not in ("Vol","None"):
                p[ADX_VOL] = "None"
            if str(p[ADX_STRAT]) not in ("ADX Only","ADX + DI+ > DI-"):
                p[ADX_STRAT] = "ADX Only"
        if len({p[ADX_LOOK] for p in batch}) < 3:
            idxs = rnd.sample(range(len(batch)), k=min(4, len(batch)))
            for i in idxs:
                batch[i][ADX_LOOK] = rnd.randint(1, 30)
        if len({p[ADX_W] for p in batch}) < 3:
            idxs = rnd.sample(range(len(batch)), k=min(4, len(batch)))
            for i in idxs:
                batch[i][ADX_W] = rnd.randint(1, 10)

    elif algo_name == "Volatility":
        a_vals = [p[VOL_A] for p in batch]
        b_vals = [p[VOL_B] for p in batch]
        a_min, a_max = 1.0, 3.0
        b_min, b_max = 3.0, 7.0
        terc = lambda lo, hi, k: lo + k*(hi-lo)/3.0
        for idxs, lo, hi, slot in [(a_vals, a_min, a_max, VOL_A), (b_vals, b_min, b_max, VOL_B)]:
            pass

        vols = [i for i,p in enumerate(batch) if str(p[VOL_VOL]) == "Vol"]
        nones= [i for i,p in enumerate(batch) if str(p[VOL_VOL]) == "None"]
        if not vols and nones:
            for i in nones[:max(1, len(nones)//2)]: batch[i][VOL_VOL] = "Vol"
        if not nones and vols:
            for i in vols[:max(1, len(vols)//2)]:  batch[i][VOL_VOL] = "None"

        lows  = [i for i,p in enumerate(batch) if _safe_float(p[VOL_T], 1.0) <= 1.0]
        highs = [i for i,p in enumerate(batch) if _safe_float(p[VOL_T], 1.0)  > 1.0]
        if not lows and highs:
            for i in highs[:max(1, len(highs)//2)]:
                p = batch[i]; p[VOL_T] = round(max(0.0, min(1.0, _safe_float(p[VOL_T],1.8) - 0.6)), 2)
        if not highs and lows:
            for i in lows[:max(1, len(lows)//2)]:
                p = batch[i]; p[VOL_T] = round(min(4.0, max(1.2, _safe_float(p[VOL_T],0.8) + 0.8)), 2)

        for p in batch:
            p[VOL_F1] = int(_safe_float(p[VOL_F1], 12))
            a = round(_safe_float(p[VOL_A], 2.0), 1)
            b = round(_safe_float(p[VOL_B], 5.0), 1)
            if a > b: a, b = b, a
            p[VOL_A], p[VOL_B] = a, b
            p[VOL_LB] = int(_safe_float(p[VOL_LB], 10))
            p[VOL_W]  = int(_safe_float(p[VOL_W],  5))
            p[VOL_T]  = round(_safe_float(p[VOL_T], 1.0), 2)
            if str(p[VOL_VOL]) not in ("Vol","None"):
                p[VOL_VOL] = rnd.choice(["Vol","None"])

    else:
        cols = len(batch[0])
        uniq_counts = [len({row[j] for row in batch}) for j in range(cols)]
        deficit = 4 - sum(1 for c in uniq_counts if c > 1)
        if deficit > 0:
            for _ in range(deficit*2):
                i = rnd.randrange(len(batch))
                j = rnd.randrange(cols)
                v = batch[i][j]
                try:
                    vv = float(v)
                    batch[i][j] = vv + (0.1 if vv == 0 else 0.05*vv)
                except:
                    if isinstance(v, str) and v.lower() in ("yes","no"):
                        batch[i][j] = "yes" if v.lower()=="no" else "no"


def _propose_for_Volatility(best_params, k=8):
    F1_lo,F1_hi,F1_st = PARAM_RANGES["Volatility"][0]
    a_lo,a_hi,a_st    = PARAM_RANGES["Volatility"][1]
    b_lo,b_hi,b_st    = PARAM_RANGES["Volatility"][2]
    vol_choices       = PARAM_RANGES["Volatility"][3]
    LB_lo,LB_hi,_     = PARAM_RANGES["Volatility"][4]
    W_lo,W_hi,_       = PARAM_RANGES["Volatility"][5]
    t_lo1,t_hi1,t_st1 = PARAM_RANGES["Volatility"][6]
    t_lo2,t_hi2,t_st2 = PARAM_RANGES["Volatility"][7]

    base = _to_internal_params("Volatility", best_params)

    a0 = _safe_float(base[VOL_A], 2.0)
    b0 = _safe_float(base[VOL_B], 5.0)
    if a0 > b0: a0, b0 = b0, a0

    out = []
    bands = [(t_lo1,t_hi1,t_st1),(t_lo2,t_hi2,t_st2)]
    for (tl,th,ts) in bands:
        for _ in range(max(1, k//2)):
            F1 = int(_clip(round(_mutate_numeric(_safe_float(base[VOL_F1], (F1_lo+F1_hi)/2), F1_lo,F1_hi,F1_st, scale=0.6)), F1_lo,F1_hi))
            a  = round(_clip(_mutate_numeric(a0, a_lo,a_hi,a_st, scale=0.6), a_lo,a_hi,a_st), 1)
            b  = round(_clip(_mutate_numeric(b0, b_lo,b_hi,b_st, scale=0.6), b_lo,b_hi,b_st), 1)
            if a > b: a, b = b, a
            vol = base[VOL_VOL] if str(base[VOL_VOL]) in ("Vol","None") else RND.choice(vol_choices)
            LB  = int(_clip(round(_mutate_numeric(_safe_float(base[VOL_LB], (LB_lo+LB_hi)/2), LB_lo,LB_hi,1, scale=0.6)), LB_lo,LB_hi))
            W   = int(_clip(round(_mutate_numeric(_safe_float(base[VOL_W], (W_lo+W_hi)/2), W_lo,W_hi,1, scale=0.6)), W_lo,W_hi))
            T   = _clip(_mutate_numeric(_safe_float(base[VOL_T], 1.0), tl,th,ts, scale=0.6), tl,th,ts)
            out.append([F1, a, b, vol, LB, W, round(T, 2 if ts==0.2 else 1)])
    return out[:k]


def _propose_generic(best_params, algo_name, k=8):
    ranges = PARAM_RANGES[algo_name]
    best_params = _pad_params_to_ranges(best_params, ranges)
    out = []
    for _ in range(k):
        cand = []
        for j,param in enumerate(ranges):
            if isinstance(param, list) and not all(isinstance(x,(int,float)) for x in param):
                choices = param
                v = best_params[j] if best_params[j] in choices else RND.choice(choices)
                if RND.random() < 0.35:
                    v = RND.choice(choices)
                cand.append(v)
            else:
                lo,hi,step = param
                base = best_params[j]
                try:
                    base_num = float(base)
                except:
                    base_num = (lo + hi)/2.0
                v = _clip(_mutate_numeric(base_num, lo,hi,step, scale=0.5), lo,hi,step)
                if step == 1 and (abs(v-round(v)) < 1e-9):
                    v = int(round(v))
                cand.append(v)
        out.append(cand)
    return out
def _propose_new_params_for_algo(algo_name, ranked_current, already_tested_set_json,
                                 need=8, explore_phase=False, eps_cat=0.25,
                                 min_l1_exploit=0.10, min_l1_explore=0.25):

    conn = sqlite3.connect(optuna_db_path)
    try:
        table = algo_name
        external_ranges = PARAM_RANGES[algo_name]
        internal_ranges = _internal_ranges_for(algo_name)

        # Builds historical scored set
        hist = _collect_historical_scored_params(conn, table, algo_name)

        # Builds dedupe set from DB
        raw_hist = _all_historical_params_for_algo(conn, table)
        hist_internal_norm = set()
        for js in raw_hist:
            try:
                p_ext = json.loads(js)
                p_int = _to_internal_params(algo_name, p_ext)
                hist_internal_norm.add(json.dumps(_normalize_for_key(p_int, internal_ranges)))
            except:
                pass

        hist_canonical = set()
        for js in raw_hist:
            try:
                p_ext = json.loads(js)
                p_int = _to_internal_params(algo_name, p_ext)
                hist_canonical.add(_db_param_json_for_query(algo_name, p_int))
            except:
                hist_canonical.add(js)

        current_canonical = set()
        for (_p_int, _score, _var_idx) in ranked_current:
            try:
                current_canonical.add(_db_param_json_for_query(algo_name, _p_int))
            except:
                pass

        if not hist:
            seed_ext = _seed_from_ranges(external_ranges)
            seed_int = _to_internal_params(algo_name, seed_ext)

            def _uniform_internal_one():
                cand = []
                for spec in internal_ranges:
                    if isinstance(spec, list) and not all(isinstance(x,(int,float)) for x in spec):
                        cand.append(RND.choice(spec))
                    else:
                        lo,hi,step = spec
                        if step:
                            nsteps = int(round((hi - lo)/step))
                            v = lo + step * RND.randint(0, max(0,nsteps))
                            if step == 1: v = int(v)
                        else:
                            v = lo + (hi - lo)*RND.random()
                        cand.append(v)
                return _pad_params_to_ranges(cand, internal_ranges)

            pool = []
            for _ in range(max(need, 8)):
                pool.append(seed_int[:])
            for _ in range(max(need*2, 16)):
                pool.append(_uniform_internal_one())
            pool.extend(_ablation_around(seed_int, algo_name, k=max(need,8)))

            out, seen_int, seen_canon = [], set(), set()
            for c in pool:
                key_int = json.dumps(_normalize_for_key(c, internal_ranges))
                canon   = _db_param_json_for_query(algo_name, c)  # exact DB JSON after snapping
                if key_int in hist_internal_norm or key_int in seen_int:
                    continue
                if canon in hist_canonical or canon in current_canonical or canon in seen_canon:
                    continue
                seen_int.add(key_int)
                seen_canon.add(canon)
                out.append(_pad_params_to_ranges(c, internal_ranges))
                if len(out) >= need:
                    break

            _force_batch_diversity(algo_name, out)
            _ensure_min_uniques(out, internal_ranges, min_uniques=3, rnd=RND)
            return out[:need]

        gamma = 0.2
        kq = max(1, int(math.ceil(gamma * len(hist))))
        good = hist[:kq]
        bad  = hist[kq:] if len(hist) > kq else hist[-1:]

        # Phase knobs
        if explore_phase:
            n_model  = max(need*2, 16)
            n_ablate = max(need*2, 16)
            n_global = max(need*3, 24)
            min_l1   = float(min_l1_explore)
            eps      = max(0.40, float(eps_cat))
        else:
            n_model  = max(need*3, 24)
            n_ablate = max(need*1,  8)
            n_global = max(need*2, 16)
            min_l1   = float(min_l1_exploit)
            eps      = float(eps_cat)

        # Per-parameter models
        cat_models = []
        num_models = []
        def _bandwidth(vals, lo, hi):
            if len(vals) < 2:
                return max((hi - lo)*0.10, 1e-9)
            sd = float(stats.pstdev(vals)) if len(vals) > 1 else (hi - lo)*0.10
            return max(sd*0.9, (hi - lo)*0.02)

        for j, spec in enumerate(internal_ranges):
            if isinstance(spec, list) and not all(isinstance(x,(int,float)) for x in spec):
                counts = {c: 1.0 for c in spec}
                for (p,_s) in good:
                    if j < len(p) and p[j] in counts:
                        counts[p[j]] += 1.0
                cat_models.append(counts); num_models.append(None)
            else:
                lo, hi, step = spec
                gvals = []
                for (p,_s) in good:
                    if j < len(p):
                        try:
                            v = float(p[j])
                            if step:
                                kst = round((v - lo)/step)
                                v = max(lo, min(hi, lo + kst*step))
                            gvals.append(v)
                        except:
                            pass
                if gvals:
                    mu = float(stats.mean(gvals))
                    bw = _bandwidth(gvals, lo, hi)
                    num_models.append((lo,hi,step,mu,bw))
                    cat_models.append(None)
                else:
                    num_models.append(None); cat_models.append(None)

        def approx_score_internal(p):
            score = 0.0
            for j, spec in enumerate(internal_ranges):
                if isinstance(spec, list) and not all(isinstance(x,(int,float)) for x in spec):
                    counts = cat_models[j]
                    if counts:
                        total = sum(counts.values())
                        pj = (counts.get(p[j], 1.0)/total)
                        score += math.log(max(pj, 1e-12))
                else:
                    model = num_models[j]
                    if not model:
                        continue
                    lo,hi,step,mu,bw = model
                    try:
                        x = float(p[j])
                        dens = math.exp(-0.5*((x-mu)/max(bw,1e-9))**2) / max(bw,1e-9)
                        score += math.log(max(dens, 1e-12))
                    except:
                        pass
            return score

        def model_sample_internal_one():
            cand = []
            for j, spec in enumerate(internal_ranges):
                if isinstance(spec, list) and not all(isinstance(x,(int,float)) for x in spec):
                    counts = cat_models[j]
                    choices = spec
                    if counts:
                        total = sum(counts.values())
                        probs = [(counts[c]/total)*(1.0-eps) + (eps/len(choices)) for c in choices]
                        r = RND.random(); acc = 0.0; pick = choices[-1]
                        for c, p in zip(choices, probs):
                            acc += p
                            if r <= acc:
                                pick = c; break
                        cand.append(pick)
                    else:
                        cand.append(RND.choice(choices))
                else:
                    lo,hi,step = spec
                    model = num_models[j]
                    if model:
                        lo,hi,step,mu,bw = model
                        v = RND.gauss(mu, bw)
                        v = max(lo, min(hi, v))
                        if step:
                            kst = round((v - lo)/step); v = lo + kst*step
                        if step == 1 and abs(v - round(v)) < 1e-9:
                            v = int(round(v))
                        cand.append(v)
                    else:
                        if step:
                            nsteps = int(round((hi - lo)/step))
                            v = lo + step * RND.randint(0, max(0,nsteps))
                            if step == 1: v = int(v)
                        else:
                            v = lo + (hi - lo)*RND.random()
                        cand.append(v)
            return _pad_params_to_ranges(cand, internal_ranges)

        def uniform_internal_one():
            cand = []
            for spec in internal_ranges:
                if isinstance(spec, list) and not all(isinstance(x,(int,float)) for x in spec):
                    cand.append(RND.choice(spec))
                else:
                    lo,hi,step = spec
                    if step:
                        nsteps = int(round((hi - lo)/step))
                        v = lo + step * RND.randint(0, max(0,nsteps))
                        if step == 1: v = int(v)
                    else:
                        v = lo + (hi - lo)*RND.random()
                    cand.append(v)
            return _pad_params_to_ranges(cand, internal_ranges)

        def grid_L1_internal(a, b):
            s = 0.0; n = 0
            for (va, vb, spec) in zip(a, b, internal_ranges):
                if isinstance(spec, list) and not all(isinstance(x,(int,float)) for x in spec):
                    s += 0.0 if va == vb else 1.0
                    n += 1
                else:
                    lo,hi,step = spec
                    try:
                        fa = float(va); fb = float(vb)
                    except:
                        continue
                    ran = (hi - lo) if hi > lo else 1.0
                    s += abs(fa - fb)/ran
                    n += 1
            return s / max(1, n)

        # Pools
        pool_model = [model_sample_internal_one() for _ in range(n_model)]

        best_anchor_int = good[0][0] if good else _to_internal_params(algo_name, _seed_from_ranges(external_ranges))
        if algo_name == "ADX":
            pool_ablate = _propose_for_ADX(best_anchor_int, k=n_ablate)
        elif algo_name == "Volatility":
            pool_ablate = _propose_for_Volatility(best_anchor_int, k=n_ablate)
        else:
            pool_ablate = _ablation_around(best_anchor_int, algo_name, k=n_ablate, rnd=RND)

        pool_global = [uniform_internal_one() for _ in range(n_global)]
        candidates = pool_model + pool_ablate + pool_global
        scored = []
        for c in candidates:
            try:
                s = approx_score_internal(c)
            except:
                s = 0.0
            scored.append((s, c))
        scored.sort(key=lambda t: t[0], reverse=True)
        out, seen_int, seen_canon = [], set(), set()
        for _, cand in scored:
            key_int = json.dumps(_normalize_for_key(cand, internal_ranges))
            canon   = _db_param_json_for_query(algo_name, cand)
            if key_int in hist_internal_norm or key_int in seen_int:
                continue
            if canon in hist_canonical or canon in current_canonical or canon in seen_canon:
                continue
            if any(grid_L1_internal(prev, cand) < min_l1 for prev in out):
                continue
            seen_int.add(key_int)
            seen_canon.add(canon)
            out.append(_pad_params_to_ranges(cand, internal_ranges))
            if len(out) >= need:
                break
        if len(out) < need:
            for c in pool_global:
                key_int = json.dumps(_normalize_for_key(c, internal_ranges))
                canon   = _db_param_json_for_query(algo_name, c)
                if key_int in hist_internal_norm or key_int in seen_int:
                    continue
                if canon in hist_canonical or canon in current_canonical or canon in seen_canon:
                    continue

                if any(grid_L1_internal(prev, c) < (min_l1 * 0.6) for prev in out):
                    continue
                seen_int.add(key_int)
                seen_canon.add(canon)
                out.append(_pad_params_to_ranges(c, internal_ranges))
                if len(out) >= need:
                    break

        _force_batch_diversity(algo_name, out)
        _ensure_min_uniques(out, internal_ranges, min_uniques=3, rnd=RND)
        return out[:need]
    finally:
        conn.close()

def _write_params_to_excel(control_sheet, algo_name, params8):
    base_row = _row_for_algo(algo_name)
    width_hint = ALGO_PARAM_COUNTS.get(algo_name, max(len(p) for p in params8) if params8 else 12)

    for i, raw in enumerate(params8):
        row = base_row + i
        clear_width = max(12, width_hint)
        control_sheet.range((row, 56), (row, 56 + clear_width - 1)).value = [[""] * clear_width]

        internal = _to_internal_params(algo_name, raw)
        excel_row = _to_excel_params(algo_name, internal)
        control_sheet.range((row, 56)).value = [excel_row]

def run_p2_rank_and_recommend():
    wb = xw.Book(EXCEL_FILE)
    control_sheet = wb.sheets["ControlSheet"]

    ensure_algo_tables_schema()
    conn = sqlite3.connect(optuna_db_path)
    try:
        # 1) current params per algo
        all_current = []
        for algo_name in CONTROL_ORDER:
            cur8 = _db_current_params_for_algo(conn, algo_name, algo_name)
            all_current.append((algo_name, cur8))

        # 2) score each variation from 5-day metrics
        algo_scores = {}
        for algo_name, cur8 in all_current:
            algo_scores[algo_name] = []
            for var_idx in range(8):
                params_list = cur8[var_idx]
                metrics = _fetch_5d_metrics_for(conn, algo_name, params_list, var_idx+1)
                score = -1e20 if metrics is None else _lexi_score(metrics)
                algo_scores[algo_name].append((params_list, float(score), var_idx+1))
            algo_scores[algo_name].sort(key=lambda t: t[1], reverse=True)

        # 3) history set for dedupe per algo
        history = {}
        for algo_name in CONTROL_ORDER:
            seen_json = _all_historical_params_for_algo(conn, algo_name)
            history[algo_name] = set(seen_json)

        # 4) propose next params and paste by name
        for algo_name, _cur8 in all_current:
            ranked = algo_scores[algo_name] 
            explore_now = len(history[algo_name]) < 80
            raw8_internal = _propose_new_params_for_algo(
                algo_name, ranked, history[algo_name], need=8,
                explore_phase=explore_now, eps_cat=0.25, min_l1_exploit=0.10, min_l1_explore=0.25
            )
            new8 = _dedup_and_diversify_batch(algo_name, raw8_internal)
            _write_params_to_excel(control_sheet, algo_name, new8)

        wb.save()
        print("P2 complete: DB-scored (5-day) → next 8 params pasted to Excel (UI order).")
    finally:
        conn.close()

# --- end of P2 ---
if __name__ == "__main__":
    run_p2_rank_and_recommend()
