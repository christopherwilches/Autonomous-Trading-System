"""
MP1 — Part 1: 5-day variation execution runner

Runs each algorithm variation over 5 trading days using Excel as the execution engine.
Python feeds stock data into the workbook, triggers the cycle macro, extracts BUY/HIT
results, and writes daily + 5-day aggregated metrics to the Optuna database.

Input: stock datasets (DB), Excel workbook
Output: per-variation performance records in optuna_data.db
"""

import json
import math
import sqlite3
import time
from collections import Counter
import pandas as pd
import xlwings as xw
import statistics as stats

# Paths / constants
EXCEL_FILE = r"<REDACTED_PATH>/WB1"
DB_FILE = r"<REDACTED_PATH>/stocks_data.db"
optuna_db_path = r"<REDACTED_PATH>/optuna_data.db"

# Excel Cell Locations
TOTAL_STOCKS_CELL = "AG3"  # Total stocks count
STOCKS_PER_CYCLE_CELL = "AG4"  # Batch size per cycle
START_ROW_CELL = "Q1"  # Row to start processing (cycling results)
CYCLING_TRIGGER_CELL = "S1"  # Cell that triggers the cycling program
CYCLING_COMPLETION_CELL = "O1"  # Cell that updates when cycling is finished

# Worksheet Name
SHEET_NAME = "ControlSheet"

DATASET_ROWS = 53  # **Each dataset is 53 rows**
BASE_START_ROW = 55  # **First row for stock printing in ControlSheet

# Exponential weighting
EW_SCHEME_DEFAULT = "0.50,0.25,0.15,0.07,0.03"

# Gates
CONSISTENCY_HITS_GATE = 4          # min hits on each day
EXPORT_LB_MIN = 27.5               # Wilson LB floor for export
EXPORT_POOLED_BUYS_MIN = 600       

PRICE_START_COL = 59          # BG: [Date, Open, High, Low, Close, Volume]
PRICE_WIDTH     = 6
VAR_BLOCK_FIRST_COLS = [11, 17, 23, 29, 35, 41, 47, 53]  # 8 variation blocks, 6 cols each

CONTROL_ORDER = [
    "MACD","EMA","RSI","Breakout",
    "ADX","Volatility",
    "SMA","Bollinger_Bands","EMA_MACD_Combo","RSI_Bollinger"
]

def _parse_ew_scheme(s):
    try:
        w = [float(x) for x in str(s or EW_SCHEME_DEFAULT).split(",")]
        return (w + [0,0,0,0,0])[:5] 
    except:
        return [0.50,0.25,0.15,0.07,0.03]

def _iqr(vals):
    vals = [float(x) for x in vals if x is not None]
    if len(vals) < 2: return 0.0
    try:
        q = stats.quantiles(vals, n=4, method="inclusive")
        return float(q[2] - q[0])
    except:
        return 0.0

def _cv(vals):
    vals = [float(x) for x in vals if x is not None]
    if not vals: return 0.0
    m = stats.mean(vals)
    if m == 0: return 0.0
    try:
        return float(stats.pstdev(vals) / m * 100.0)
    except:
        return 0.0

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

def _excel_params_for_algo_variations(control_sheet, algo_index):
    base_row = 4 + (algo_index * 11)
    out = []
    for v in range(8):
        row = base_row + v
        vals = []
        for j in range(20):
            val = control_sheet.range((row, 56 + j)).value
            if val in (None, ""):
                break
            vals.append(val)
        out.append(vals)
    return out  # length 8

def _is_profit_cellblock(block6):
    for v in block6:
        if isinstance(v,(int,float)) and float(v) > 0: return True
        s = str(v).strip().upper()
        if "PROFIT" in s: return True
    return False

def _read_variation_day(wb, algo_sheet_name, var_slot):
    sh = wb.sheets[algo_sheet_name]
    tickers = sh.range((55,10),(1056,10)).value
    if not isinstance(tickers, list): return 0,0,[]

    n_rows = 0
    for i,v in enumerate(tickers):
        if not v: break
        n_rows = i+1
    if n_rows == 0: return 0,0,[]

    start_col = VAR_BLOCK_FIRST_COLS[var_slot-1]
    block_vals = sh.range((55, start_col),(54+n_rows, start_col+5)).value
    decisions  = sh.range((55, start_col),(54+n_rows, start_col)).value
    tickers    = sh.range((55,9),(54+n_rows,9)).value   # I = Ticker
    names      = sh.range((55,10),(54+n_rows,10)).value # J = Name

    prices     = sh.range((55, PRICE_START_COL),(54+n_rows, PRICE_START_COL+PRICE_WIDTH-1)).value

    buys, hits, out = 0,0,[]
    for i in range(n_rows):
        dec = str(decisions[i]).strip().upper() if decisions[i] is not None else ""
        if dec == "BUY":
            buys += 1
            blk6 = block_vals[i] if isinstance(block_vals[i], list) else [block_vals[i]]
            if _is_profit_cellblock(blk6): hits += 1
            pr6 = prices[i] if isinstance(prices[i], list) else [prices[i]]
            out.append({
                "ticker": tickers[i],
                "name":   names[i],
                "open":   pr6[1] if len(pr6)>1 else None,
                "high":   pr6[2] if len(pr6)>2 else None,
                "low":    pr6[3] if len(pr6)>3 else None,
                "close":  pr6[4] if len(pr6)>4 else None,
                "volume": pr6[5] if len(pr6)>5 else None
            })
    return buys, hits, out
def _find_or_create_row_for_params(conn, algo_table, params_list, variation_number, run_id, anchor_only_insert=False):
    cur = conn.cursor()
    params_json = json.dumps(params_list)

    # Day 1
    if anchor_only_insert:
        cur.execute(f"SELECT COALESCE(MAX(trial_number), 0) FROM {algo_table}")
        last_trial = int(cur.fetchone()[0] or 0)
        cur.execute(f"SELECT trial_number FROM {algo_table} WHERE run_id=? ORDER BY id DESC LIMIT 1", (run_id,))
        existing_trial = cur.fetchone()
        if existing_trial:
            trial_number = int(existing_trial[0])
        else:
            trial_number = last_trial + 1
        cur.execute(f"""
            INSERT INTO {algo_table} (run_id, params, trial_number, variation_number)
            VALUES (?,?,?,?)
        """, (run_id, params_json, trial_number, variation_number))
        conn.commit()
        return cur.lastrowid, trial_number

    # Days 2–5
    cur.execute(f"""
        SELECT id, trial_number FROM {algo_table}
        WHERE run_id=? AND variation_number=? AND params=?
        ORDER BY id DESC LIMIT 1
    """, (run_id, variation_number, params_json))
    row = cur.fetchone()

    if row:
        return row[0], int(row[1])

    cur.execute(f"SELECT COALESCE(MAX(trial_number), 0) FROM {algo_table}")
    trial_number = int(cur.fetchone()[0]) + 1

    cur.execute(f"""
        INSERT INTO {algo_table} (run_id, params, trial_number, variation_number)
        VALUES (?,?,?,?)
    """, (run_id, params_json, trial_number, variation_number))
    conn.commit()
    return cur.lastrowid, trial_number

def _update_day_result(conn, algo_table, row_id, day_idx, buys, hits, buys_json):
    pp = (100.0 * hits / buys) if buys > 0 else 0.0
    cur = conn.cursor()
    cur.execute(f"""
        UPDATE {algo_table}
        SET day{day_idx}_buys=?, day{day_idx}_hits=?, day{day_idx}_pp=?, day{day_idx}_buys_json=?
        WHERE id=?
    """,(buys, hits, pp, json.dumps(buys_json), row_id))
    conn.commit()

def _wilson_lower_bound(successes, trials, z=1.96):
    if trials <= 0: return 0.0
    phat = successes / trials
    denom = 1 + z*z/trials
    centre = phat + z*z/(2*trials)
    margin = z*math.sqrt((phat*(1-phat)+z*z/(4*trials))/trials)
    return 100.0 * (centre - margin)/denom
def _finalize_5d_metrics(conn, algo_table, row_id, min_hits_gate=CONSISTENCY_HITS_GATE):
    cur = conn.cursor()

    # pull all per-day fields
    cur.execute(f"""
        SELECT
            day1_buys, day1_hits, day1_pp, day1_buys_json,
            day2_buys, day2_hits, day2_pp, day2_buys_json,
            day3_buys, day3_hits, day3_pp, day3_buys_json,
            day4_buys, day4_hits, day4_pp, day4_buys_json,
            day5_buys, day5_hits, day5_pp, day5_buys_json,
            ew_scheme
        FROM {algo_table} WHERE id=?
    """,(row_id,))
    r = cur.fetchone()
    if not r: return

    # unpack
    b = [r[0] or 0, r[4] or 0, r[8] or 0, r[12] or 0, r[16] or 0]
    h = [r[1] or 0, r[5] or 0, r[9] or 0, r[13] or 0, r[17] or 0]
    pp= [float(r[2] or 0.0), float(r[6] or 0.0), float(r[10] or 0.0), float(r[14] or 0.0), float(r[18] or 0.0)]
    jsons = [r[3], r[7], r[11], r[15], r[19]]
    ew_s = r[20] or EW_SCHEME_DEFAULT

    # pooled metrics
    pooled_buys = int(sum(b))
    pooled_hits = int(sum(h))
    pooled_pp   = float((100.0*pooled_hits/pooled_buys) if pooled_buys>0 else 0.0)

    # robust day statistics
    median_pp = float(stats.median(pp))
    mad_pp    = float(stats.median([abs(x - median_pp) for x in pp]))
    med_buys  = float(stats.median(b))
    mad_buys  = float(stats.median([abs(x - med_buys) for x in b]))
    pp_iqr    = _iqr(pp)
    buy_cv    = _cv(b)

    # Wilson LB
    wilson_lb = float(_wilson_lower_bound(pooled_hits, pooled_buys))

    # EW metrics
    weights = _parse_ew_scheme(ew_s)
    ew_pp   = float(sum(w * v for w,v in zip(weights, pp)))
    ew_hits = float(sum(w * v for w,v in zip(weights, h)))

    # ticker-level features (from buys_json)
    all_tickers = []
    all_prices  = []
    all_vols    = []
    for js in jsons:
        if not js: continue
        try:
            rows = json.loads(js)
            for it in rows:
                tk = it.get("ticker")
                if tk: all_tickers.append(str(tk))
                p = it.get("open", None)
                if p is None: p = it.get("close", None)
                if isinstance(p,(int,float)): all_prices.append(float(p))
                v = it.get("volume", None)
                if isinstance(v,(int,float)): all_vols.append(float(v))
        except Exception:
            pass

    uniq = set(all_tickers)
    repeat_rate = 0.0
    top10_share = 0.0
    if all_tickers:
        c = Counter(all_tickers)
        repeated = sum(1 for t,cnt in c.items() if cnt >= 2)
        repeat_rate = float(repeated / max(1, len(uniq)) * 100.0)
        top10 = sum(x for _,x in c.most_common(10))
        top10_share = float(top10 / len(all_tickers) * 100.0)

    avg_price   = float(stats.mean(all_prices)) if all_prices else 0.0
    med_price   = float(stats.median(all_prices)) if all_prices else 0.0
    avg_volume  = float(stats.mean(all_vols)) if all_vols else 0.0

    # recency aliases (day 5)
    last_pp   = float(pp[4] if len(pp)>=5 else 0.0)
    last_buys = int(b[4] if len(b)>=5 else 0)
    last_hits = int(h[4] if len(h)>=5 else 0)

    # gates
    min_hits = int(min(h) if h else 0)
    passed_consistency = 1 if min_hits >= int(min_hits_gate) else 0
    passed_export = 1 if (passed_consistency == 1 and wilson_lb >= EXPORT_LB_MIN and pooled_buys >= EXPORT_POOLED_BUYS_MIN) else 0

    # write back
    cur.execute(f"""
        UPDATE {algo_table}
        SET
            pooled_buys=?, pooled_hits=?, pooled_pp=?,
            median_pp_5d=?, pp_mad_5d=?, buycount_med_5d=?, buycount_mad_5d=?,
            pp_iqr_5d=?, buycount_cv_5d=?,
            wilson_lb_5d=?,
            ew_scheme=?, ew_pp_5d=?, ew_hits_5d=?,
            repeat_ticker_rate_5d=?, top_10_ticker_share_5d=?,
            avg_buy_price_5d=?, median_buy_price_5d=?, avg_buy_volume_5d=?,
            last_day_pp=?, last_day_buys=?, last_day_hits=?,
            min_daily_hits=?, passed_consistency_gate=?, passed_export_gate=?
        WHERE id=?
    """,(
        pooled_buys, pooled_hits, pooled_pp,
        median_pp, mad_pp, med_buys, mad_buys,
        pp_iqr, buy_cv,
        wilson_lb,
        ew_s, ew_pp, ew_hits,
        repeat_rate, top10_share,
        avg_price, med_price, avg_volume,
        last_pp, last_buys, last_hits,
        min_hits, passed_consistency, passed_export,
        row_id
    ))
    conn.commit()
def record_day_results_and_finalize_if_needed(day_idx, run_id, anchor_day1=False):
    wb = xw.Book(EXCEL_FILE)
    control_sheet = wb.sheets["ControlSheet"]
    ensure_algo_tables_schema()
    conn = sqlite3.connect(optuna_db_path)
    try:
        for algo_index, algo_name in enumerate(CONTROL_ORDER):
            try:
                sheet_name = SHEET_FOR_ALGO[algo_name]
                params_8 = _excel_params_for_algo_variations(control_sheet, algo_index)
                for v in range(1, 9):
                    params_list = params_8[v-1]
                    row_id, _trial = _find_or_create_row_for_params(
                        conn, algo_name, params_list, v, run_id,
                        anchor_only_insert=anchor_day1
                    )
                    buys, hits, buys_json = _read_variation_day(wb, sheet_name, v)
                    _update_day_result(conn, algo_name, row_id, day_idx, buys, hits, buys_json)
                    if day_idx == 5:
                        _finalize_5d_metrics(conn, algo_name, row_id, min_hits_gate=CONSISTENCY_HITS_GATE)
            except Exception as e:
                print(f"[WARN] record_day_results: algo '{algo_name}' failed on day {day_idx}: {e}")
                continue
    finally:
        conn.close()

# Excel sheet map
SHEET_FOR_ALGO = {
    "MACD": "1 MACD",
    "EMA": "2 EMA",
    "RSI": "3 RSI",
    "Breakout": "4 Breakout",
    "ADX": "5 ADX",
    "Volatility": "6 Volatility Measure",
    "SMA": "7 SMA",
    "Bollinger_Bands": "8 Bollinger Bands",
    "EMA_MACD_Combo": "9 EMA & MACD",
    "RSI_Bollinger": "10 RSI Bollinger",
}

# run_id helpers
def _ensure_runs_table():
    try:
        conn = sqlite3.connect(optuna_db_path)
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS RUNS (
                id INTEGER PRIMARY KEY,
                run_id TEXT UNIQUE,
                created_at TEXT
            )
        """)
        conn.commit()
    finally:
        try:
            conn.close()
        except:
            pass


def _new_run_id():
    _ensure_runs_table()
    rid_base = f"RUN{int(time.time()*1000)}"
    suffix = 0
    while True:
        run_id = rid_base if suffix == 0 else f"{rid_base}_{suffix}"
        try:
            conn = sqlite3.connect(optuna_db_path)
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO RUNS (run_id, created_at) VALUES (?, datetime('now'))",
                (run_id,),
            )
            conn.commit()
            return run_id
        except Exception:
            suffix += 1
            if suffix > 99:
                raise
        finally:
            try:
                conn.close()
            except:
                pass
              
# main runner
def run_cycling_program():
    try:
        current_run_id = _new_run_id()

        wb = xw.Book(EXCEL_FILE)
        sheet = wb.sheets[SHEET_NAME]
        app = wb.app
      
        try:
            app.Interactive = True; app.EnableEvents = True; app.DisplayAlerts = True
            app.DisplayStatusBar = True; app.ScreenUpdating = True
            app.Calculation = -4105
            try:
                for _sh in wb.sheets:
                    try: _sh.api.ScrollArea = ""
                    except: pass
            except: pass
            app.api.DoEvents()
        except: pass

        total_stocks = int(sheet.range(TOTAL_STOCKS_CELL).value)
        batch_size   = int(sheet.range(STOCKS_PER_CYCLE_CELL).value)
        sheet.range("K1").value = batch_size
        num_cycles = (total_stocks // batch_size) + (1 if total_stocks % batch_size != 0 else 0)

        wb.sheets["ControlSheet"].range("AE1").value = 100
        time.sleep(2)

        conn = sqlite3.connect(DB_FILE)

        for day_idx in range(1, 6):
            dataset_table = f"DS_DAY{day_idx}"
            print(f"\n=== DAY {day_idx}/5 with {dataset_table} ===")
            sheet.range(START_ROW_CELL).value = BASE_START_ROW

            for cycle in range(num_cycles):
                total_rows = batch_size * DATASET_ROWS
                start = cycle * total_rows + 1
                end   = start + total_rows - 1
                df = pd.read_sql_query(
                    f'SELECT * FROM "{dataset_table}" WHERE rowid BETWEEN {start} AND {end} ORDER BY rowid', conn
                )

                if not df.empty:
                    prev_upd, prev_alerts = app.screen_updating, app.display_alerts
                    try:
                        app.screen_updating, app.display_alerts = False, False
                        sheet.range("A1").value = df.values
                    finally:
                        app.screen_updating, app.display_alerts = prev_upd, prev_alerts
                else:
                    print(f"[WARN] {dataset_table} window {start}-{end} returned 0 rows.")

                start_row = BASE_START_ROW + (cycle * batch_size)
                sheet.range(START_ROW_CELL).value = start_row

                initial_o1 = sheet.range(CYCLING_COMPLETION_CELL).value
                sheet.range(CYCLING_TRIGGER_CELL).value = 100

                timeout_limit = time.time() + 600
                while True:
                    if time.time() > timeout_limit:
                        raise TimeoutError(f"Timeout waiting for macro day {day_idx}, cycle {cycle+1}.")
                    if sheet.range(CYCLING_COMPLETION_CELL).value != initial_o1:
                        break
                    time.sleep(1.5)

                print(f"[DONE] Day {day_idx} — Batch {cycle+1}/{num_cycles}")

            try:
                sheet.range(START_ROW_CELL).value = BASE_START_ROW
                print(f"[DEBUG] Day {day_idx}: Q1 reset OK")
            except Exception as e:
                print(f"[ERROR] Day {day_idx}: failed to reset Q1 → {e}")
                raise

            try:
                record_day_results_and_finalize_if_needed(
                    day_idx,
                    current_run_id,
                    anchor_day1=(day_idx == 1)
                )
                print(f"[DEBUG] Day {day_idx}: results recorded OK")
            except Exception as e:
                print(f"[ERROR] Day {day_idx}: record_day_results failed → {e}")
                try:
                    wb.sheets['ControlSheet'].range('AE1').value = 100
                except: pass
                raise

            try:
                wb.sheets["ControlSheet"].range("AE1").value = 100
                print(f"[DEBUG] Day {day_idx}: AE1 clear OK")
            except Exception as e:
                print(f"[ERROR] Day {day_idx}: AE1 clear failed → {e}")
                raise
            time.sleep(1)
        conn.close()
        wb.save()
        print("P1 complete: 5-day run finished and saved.")

    except Exception as e:
        print(f"[ERROR] P1 failed: {e}")

# MP1 - PART 1 END

if __name__ == "__main__":
    run_cycling_program()
