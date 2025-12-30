import time
import sqlite3
import pandas as pd
import xlwings as xw

import optuna
import threading
import pythoncom
from optuna.distributions import FloatDistribution, IntDistribution, CategoricalDistribution

import json
import ssl
import yfinance as yf
from datetime import datetime
import random
import math
from concurrent.futures import ThreadPoolExecutor
import itertools
import signal
from optuna._experimental import ExperimentalWarning
import warnings
warnings.filterwarnings("ignore", category=optuna.exceptions.ExperimentalWarning)

# file paths 

EXCEL_FILE = r"C:/REDACTED/WB1.xlsm"

DB_FILE = r"C:/REDACTED/stocks_data.db"

optuna_db_path = r"C:/REDACTED/optuna_10A_data.db"

# Excel Cell Locations
TOTAL_STOCKS_CELL = "AG3" 
STOCKS_PER_CYCLE_CELL = "AG4" 
START_ROW_CELL = "Q1"  
CYCLING_TRIGGER_CELL = "S1"  
CYCLING_COMPLETION_CELL = "O1" 


SHEET_NAME = "ControlSheet"

# Constants
DATASET_ROWS = 53  
BASE_START_ROW = 55 

ALGO_NAMES = [
    "MACD", "EMA", "RSI", "Breakout", "ADX", "Volatility",
    "SMA", "Bollinger_Bands", "EMA_MACD_Combo", "RSI_Bollinger"
]

SHEET_NAMES = [
    "1 MACD", "2 EMA", "3 RSI", "4 Breakout", "5 ADX",
    "6 Volatility Measure", "7 SMA", "8 Bollinger Bands",
    "9 EMA & MACD", "10 RSI Bollinger"
]

conn = sqlite3.connect(optuna_db_path, check_same_thread=False) 
cursor = conn.cursor()

for algo in ALGO_NAMES:
    cursor.execute(f"CREATE TABLE IF NOT EXISTS {algo} (id INTEGER PRIMARY KEY, params TEXT, profit_percentage REAL)")
conn.commit()
conn.close()

# Parameter search ranges are omitted for privacy
PARAM_RANGES: dict = {}

ssl_context = ssl._create_unverified_context()
target_tables = ["DS1_SET", "DS2_SET", "DS3_SET", "DS4_SET"]


batch_size = 5       
batch_pause = 0          
pause_every = 1000       
pause_duration = 180    

# mp1 p1

def run_cycling_program():
    """Loops through all sets needed based on AG3, pasting into A1 and cycling through."""
    try:
        wb = xw.Book(EXCEL_FILE)
        sheet = wb.sheets[SHEET_NAME]

        # Read helper cells
        total_stocks = int(sheet.range(TOTAL_STOCKS_CELL).value)  
        batch_size = int(sheet.range(STOCKS_PER_CYCLE_CELL).value)  
        sheet.range("K1").value = batch_size  
        num_cycles = total_stocks // batch_size
        if total_stocks % batch_size != 0:
            num_cycles += 1

        print(f"[INFO] Running {num_cycles} cycles with {batch_size} datasets per cycle.")

        # Trigger macro to clear data blocks
        wb.sheets["ControlSheet"].range("AE1").value = 100
        time.sleep(3)  

        conn = sqlite3.connect(DB_FILE)

        for cycle in range(num_cycles):
            offset = cycle * batch_size
            total_rows = batch_size * DATASET_ROWS  

            print(f"[INFO] Cycle {cycle + 1}/{num_cycles}: Fetching rows {offset} to {offset + total_rows}.")

            # Extract stock data
            query = f'SELECT * FROM "DATASET_ONE" LIMIT {total_rows} OFFSET {offset * DATASET_ROWS}'
            df = pd.read_sql_query(query, conn)

            sheet.range("A:H").clear_contents()

            if not df.empty:
                try:
                    df = df.reset_index(drop=True) 
                    df = df.iloc[:, :]  
                    sheet.range("A1").value = df.values
                    print(f"[SUCCESS] {batch_size} datasets ({total_rows} rows) written to A1.")

                except Exception as e:
                    print(f"[ERROR] Failed to paste data: {e}")
            else:
                print("[ERROR] DataFrame is empty, nothing to paste!")

            start_row = BASE_START_ROW + (cycle * batch_size) 
            sheet.range(START_ROW_CELL).value = start_row  

            print(f"[INFO] Start row for cycle {cycle + 1}: {start_row}")
            # Trigger cycling program
            initial_end_time = sheet.range(CYCLING_COMPLETION_CELL).value  
            sheet.range(CYCLING_TRIGGER_CELL).value = 100  
            wb.save()

            try:
                timeout_limit = time.time() + 600  
                polling_interval = 2 

                print(f"[INFO] Polling every {polling_interval}s for macro completion...")

                while True:
                    if time.time() > timeout_limit:
                        raise TimeoutError("  Timeout! Batch polling exceeded 10 minutes.")

                    current_end_time = sheet.range(CYCLING_COMPLETION_CELL).value
                    if current_end_time != initial_end_time:
                        break

                    print(f"[INFO] Still waiting... (Polling every {polling_interval}s)")
                    time.sleep(polling_interval)

                print(f"[ ] Cycle {cycle + 1} completed.")

            except TimeoutError as te:
                print(f"[ ] Timeout during cycle {cycle + 1}: {te}")

        conn.close()
        wb.save()
        print("Excel Cycling Program Completed Successfully.")

    except Exception as e:
        print(f"[ERROR] Excel Automation or DB Extraction Failed: {e}")

# mp1 p1 end

used_sets_by_algo = set()
def run_optuna_optimization():
    try:
        wb = xw.Book(EXCEL_FILE)
        sheet = wb.sheets[SHEET_NAME]
        import optuna
        optuna.logging.set_verbosity(optuna.logging.ERROR)

        row = 4
        save_all_variations_to_db()
        save_excel_params_to_temp_table()
        first_half = ['MACD', 'EMA', 'RSI', 'Breakout', 'ADX', 'Volatility']
        second_half = ['SMA', 'Bollinger_Bands', 'EMA_MACD_Combo', 'RSI_Bollinger']
        for algo_name in first_half + second_half:

            def make_objective(start_row, i):
                def objective(trial):
                    params = []

                    if algo_name == "Volatility":
                        val1 = trial.suggest_float(f"{algo_name}_v{i}_param1", 10, 19, step=1)  # Volatility F1

                        # Merge two values into one parameter 
                        pmerge1 = trial.suggest_float(f"{algo_name}_v{i}_pmerge1", 1.0, 3.0, step=0.1)
                        pmerge2 = trial.suggest_float(f"{algo_name}_v{i}_pmerge2", 3.0, 7.0, step=0.1)
                        merged_param = f"{min(pmerge1, pmerge2):.1f}-{max(pmerge1, pmerge2):.1f}"

                        val3 = trial.suggest_categorical(f"{algo_name}_v{i}_param3", ["Vol", "None"])
                        val4 = trial.suggest_float(f"{algo_name}_v{i}_param4", 1, 30, step=1)
                        val5 = trial.suggest_int(f"{algo_name}_v{i}_param5", 1, 10)

                        # Choose one of two ranges for trend threshold
                        trend_range = trial.suggest_categorical(f"{algo_name}_v{i}_trend_range", ["low", "high"])
                        if trend_range == "low":
                            t_val = trial.suggest_float(f"{algo_name}_v{i}_trend_val", 0.0, 1.0, step=0.1)
                        else:
                            t_val = trial.suggest_float(f"{algo_name}_v{i}_trend_val", 1.0, 4.0, step=0.2)

                        params = [val1, merged_param, val3, val4, val5, round(t_val, 1)]

                        param_str = str(params)

                    elif algo_name == "ADX":
                        val1 = trial.suggest_int(f"{algo_name}_v{i}_param1", 10, 14)  # ADX Period

                        # Choose one of two ranges for threshold
                        adx_range = trial.suggest_categorical(f"{algo_name}_v{i}_range", ["low", "high"])
                        if adx_range == "low":
                            val2 = trial.suggest_float(f"{algo_name}_v{i}_param2", 10, 30, step=1)
                        else:
                            val2 = trial.suggest_float(f"{algo_name}_v{i}_param2", 60, 100, step=1)

                        val3 = trial.suggest_categorical(f"{algo_name}_v{i}_param4", ["Vol", "None"])
                        val4 = trial.suggest_int(f"{algo_name}_v{i}_param3", 1, 30)
                        val5 = trial.suggest_int(f"{algo_name}_v{i}_param5", 1, 10)

                        # Strategy Type as final param
                        strategy = trial.suggest_categorical(f"{algo_name}_v{i}_param6", ["ADX Only", "ADX + DI+ > DI-"])

                        params = [val1, int(val2), val3, val4, val5, strategy]

                        param_str = str(params)

                    else:
                        ranges = PARAM_RANGES[algo_name]
                        for j, param in enumerate(ranges):
                            name = f"{algo_name}_v{i}_param{j+1}"
                            if isinstance(param, (list, tuple)) and not all(isinstance(x, (int, float)) for x in param):
                                val = trial.suggest_categorical(name, param)
                            else:
                                try:
                                    low, high, step = param
                                except ValueError:
                                        raise ValueError(f"[ERROR] PARAM_RANGES for '{algo_name}' has invalid param format at index {j}: {param}")
                                val = trial.suggest_float(name, low, high, step=step)

                            params.append(val)


                        param_str = str(params)

                    # Write parameters to Excel with safe formatting
                    row_values = [val if isinstance(val, (int, float)) else str(val) for val in params]
                    sheet.range((start_row + i, 56)).value = [row_values]

                    # Read profit percentage from Excel 
                    pp = sheet.range((start_row + i, 48)).value or 0.0

                    trial.set_user_attr("variation_index", i)
                    trial.set_user_attr("params", str(params))
                    return float(pp)
                return objective
            for i in range(8):
                try:
                    from optuna.samplers import TPESampler
                    # TPE sampler with advanced control
                    sampler = TPESampler(multivariate=True, group=True, warn_independent_sampling=False)
                    study = optuna.create_study(
                        study_name=f"{algo_name}_run_{i}",
                        direction="maximize",
                        storage=f"sqlite:///{optuna_db_path}",
                        load_if_exists=True,
                        sampler=sampler
                    )

                    study.optimize(make_objective(row, i), n_trials=1)

                except optuna.exceptions.TrialPruned:
                    pass
                except Exception as e:
                    print(f"[ERROR] Algorithm {algo_name} v{i} failed: {e}")

            row += 11

        print(" MP1 - Part 2 (Optuna Optimization) Completed.")

    except Exception as e:
        print(f"[ERROR] Optuna Optimization Failed: {e}")

    overwrite_trial_params_with_temp_values()

# mp2 p2 end

# p2 save params helper
def save_excel_params_to_temp_table():
    DB_PATH = r"C:/REDACTED/optuna_10A_data.db"
    EXCEL_FILE = r"C:/REDACTED/WB1.xlsm" 
    SHEET_NAME = "ControlSheet"  

    VARIATIONS_PER_ALGO = 8 
    ALGO_LIST = ['MACD', 'EMA', 'RSI', 'Breakout', 'ADX', 'Volatility',
                 'SMA', 'Bollinger_Bands', 'EMA_MACD_Combo', 'RSI_Bollinger']

    START_ROW = 4
    VARIATION_HEIGHT = 11
    PARAM_START_COL = 56 
    MAX_PARAM_COLS = 10   

    import sqlite3
    import xlwings as xw

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Reset temp table
    cursor.execute("DROP TABLE IF EXISTS temp_params")
    cursor.execute("CREATE TABLE temp_params (param_value TEXT)")

    # Open Excel
    wb = xw.Book(EXCEL_FILE)
    sheet = wb.sheets[SHEET_NAME]

    total_params = 0
    row = START_ROW

    for algo in ALGO_LIST:
        for i in range(VARIATIONS_PER_ALGO):
            variation_row = row + i
            param_count = 0

            for j in range(MAX_PARAM_COLS):
                val = sheet.range((variation_row, PARAM_START_COL + j)).value
                if val is None or val == "":
                    break

                # Handle Volatility hyphen-split case
                if algo == "Volatility" and isinstance(val, str) and '-' in val:
                    parts = val.split('-')
                    for part in parts:
                        cursor.execute("INSERT INTO temp_params (param_value) VALUES (?)", (part.strip(),))
                        total_params += 1
                        param_count += 1

                elif algo == "Volatility" and param_count == 6:
                    try:
                        num = float(val)
                        range_type = "low" if num < 1 else "high"
                        cursor.execute("INSERT INTO temp_params (param_value) VALUES (?)", (range_type,))
                        total_params += 1
                        param_count += 1
                    except:
                        pass 

                    cursor.execute("INSERT INTO temp_params (param_value) VALUES (?)", (val,))
                    total_params += 1
                    param_count += 1

                elif algo == "ADX" and j == 1:
                    try:
                        threshold_val = float(val)
                        highlow = "high" if threshold_val >= 50 else "low"
                        cursor.execute("INSERT INTO temp_params (param_value) VALUES (?)", (highlow,))
                        total_params += 1
                        param_count += 1
                    except:
                        pass

                    cursor.execute("INSERT INTO temp_params (param_value) VALUES (?)", (str(val),))
                    total_params += 1
                    param_count += 1

                else:
                    cursor.execute("INSERT INTO temp_params (param_value) VALUES (?)", (str(val),))
                    total_params += 1
                    param_count += 1

        row += VARIATION_HEIGHT

    conversion_dict = {
        "None": "0",
        "Vol": "1",
        "RSI": "1",
        "low": "0",
        "high": "1",
        "ADX Only": "0",
        "ADX + DI+ > DI-": "1"
    }

    cursor.execute("SELECT rowid, param_value FROM temp_params")
    all_rows = cursor.fetchall()

    for rowid, val in all_rows:
        new_val = conversion_dict.get(val.strip(), val)
        cursor.execute("UPDATE temp_params SET param_value = ? WHERE rowid = ?", (new_val, rowid))

    conn.commit()
    conn.close()

    print(f"\n  Done! Total parameters saved: {total_params}")

# overwrite old params with new ones          

def overwrite_trial_params_with_temp_values():
    DB_PATH = r"C:/REDACTED/optuna_10A_data.db"
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Load temp values
    cursor.execute("SELECT param_value FROM temp_params")
    temp_values = [row[0] for row in cursor.fetchall()]

    # Get bottom N param_ids from trial_params
    cursor.execute("SELECT param_id FROM trial_params ORDER BY param_id DESC LIMIT ?", (len(temp_values),))
    param_ids = [row[0] for row in cursor.fetchall()]

    if len(param_ids) != len(temp_values):
        print(f"[ERROR] Count mismatch: {len(param_ids)} trial_params vs {len(temp_values)} Excel values")
        conn.close()
        return

    for param_id, new_val in zip(param_ids, reversed(temp_values)):
        cursor.execute("UPDATE trial_params SET param_value = ? WHERE param_id = ?", (str(new_val), param_id))

    conn.commit()
    conn.close()
    print("  trial_params updated with reversed Excel values.")

# clean all optuna tables start
def wipe_optuna_db():
    conn = sqlite3.connect(optuna_db_path)
    cur = conn.cursor()

    for algo in ALGO_NAMES:
        cur.execute(f"DELETE FROM {algo}")

    # Delete Optuna tracking tables too
    optuna_tables = [
        "trials", "trial_values", "trial_params", "trial_user_attributes",
        "trial_system_attributes", "trial_heartbeats", "trial_intermediate_values", "studies",
        "study_directions", "study_user_attributes", "study_system_attributes"
    ]
    for tbl in optuna_tables:
        cur.execute(f"DELETE FROM {tbl}")

    conn.commit()
    conn.close()
    print("[CLEANUP] All Optuna trials and logs deleted.")

# Save parameter info into DB
def save_all_variations_to_db():
    wb = xw.Book(EXCEL_FILE)
    control_sheet = wb.sheets["ControlSheet"]

    # Constants
    start_data_row = 55
    ticker_col = 9   # Column I
    name_col = 10    # Column J
    result_col_starts = [11, 17, 23, 29, 35, 41, 47, 53] 


    conn = sqlite3.connect(optuna_db_path)
    cur = conn.cursor()
    for algo_index, (sheet_name, algo_name) in enumerate(zip(SHEET_NAMES, ALGO_NAMES)):

        # Match correct DB table name based on sheet_name
        possible_tables = ["ADX", "Bollinger_Bands", "Breakout", "EMA", "EMA_MACD_Combo", "MACD",
                           "RSI", "RSI_Bollinger", "SMA", "Volatility"]
        sheet_lower = sheet_name.lower()

        if "rsi" in sheet_lower and "bollinger" not in sheet_lower:
            matched_table_name = "RSI"
        elif "ema" in sheet_lower and "macd" not in sheet_lower:
            matched_table_name = "EMA"
        elif "macd" in sheet_lower and "ema" not in sheet_lower:
            matched_table_name = "MACD"
        elif "ema" in sheet_lower and "macd" in sheet_lower:
            matched_table_name = "EMA_MACD_Combo"
        elif "rsi" in sheet_lower and "bollinger" in sheet_lower:
            matched_table_name = "RSI_Bollinger"
        elif "bollinger" in sheet_lower:
            matched_table_name = "Bollinger_Bands"
        elif "adx" in sheet_lower:
            matched_table_name = "ADX"
        elif "breakout" in sheet_lower:
            matched_table_name = "Breakout"
        elif "volatility" in sheet_lower:
            matched_table_name = "Volatility"
        elif "sma" in sheet_lower:
            matched_table_name = "SMA"
        else:
            matched_table_name = algo_name 
          
        print(f"\n[INFO] Saving for {matched_table_name} using sheet '{sheet_name}'")
        # Get matching sheet
        try:
            algo_sheet = wb.sheets[sheet_name]
        except:
            print(f"[WARN] Sheet '{sheet_name}' not found. Skipping.")
            continue

        cur.execute(f"""
            CREATE TABLE IF NOT EXISTS {matched_table_name} (
                id INTEGER PRIMARY KEY,
                params TEXT,
                profit_percentage REAL
            )
        """)
        conn.commit()

        required_cols = {
            "trial_number": "INTEGER",
            "variation_number": "INTEGER",
            "stocks_bought": "TEXT"
        }

        cur.execute(f"PRAGMA table_info({matched_table_name})")
        existing_cols = {col[1] for col in cur.fetchall()}

        for col, col_type in required_cols.items():
            if col not in existing_cols:
                cur.execute(f"ALTER TABLE {matched_table_name} ADD COLUMN {col} {col_type}")
                print(f"[INFO] Column '{col}' added to '{matched_table_name}' table.")

        # Count existing rows to find trial number
        cur.execute(f"SELECT COUNT(*) FROM {matched_table_name}")
        count = cur.fetchone()[0]
        base_trial = (count // 8) + 1

        # Get the starting row in ControlSheet
        control_start_row = 4 + (algo_index * 11)
        # Read entire data block
        data_block = algo_sheet.range((55, 9), (1054, 57 + 1)).value 

        for var_num in range(1, 9):
            row_offset = var_num - 1
            variation_col = result_col_starts[row_offset]

            # Fetch parameters
            params = []
            for j in range(20):
                val = control_sheet.range((control_start_row + row_offset, 56 + j)).value
                if val is None:
                    break
                params.append(val)

            # Fetch PP
            pp = control_sheet.range((control_start_row + row_offset, 48)).value or 0.0

            # Process each row from cached data_block
            stocks = []
            for row_data in data_block:
                ticker = row_data[0]  
                name = row_data[1]    
                if not ticker:
                    continue

                decision = row_data[variation_col - 9] 
                if str(decision).strip().upper() == "BUY":
                    r_val  = row_data[variation_col - 9 + 1]
                    p_val  = row_data[variation_col - 9 + 2]
                    vr_val = row_data[variation_col - 9 + 3]
                    vd_val = row_data[variation_col - 9 + 4]

                    def is_number(x):
                        try:
                            return float(x) == float(x)
                        except:
                            return False

                    def is_verbal(x):
                        return isinstance(x, str) and str(x).startswith("A")

                    result_val = r_val
                    profit_val = p_val
                    verbal_result_val = vr_val
                    verbal_decision_val = vd_val

                    if is_verbal(r_val) and is_number(vr_val):
                        result_val = vr_val
                        verbal_result_val = r_val
                    if is_verbal(p_val) and is_number(vd_val):
                        profit_val = vd_val
                        verbal_decision_val = p_val

                    stocks.append({
                        "ticker": ticker,
                        "name": name,
                        "decision": decision,
                        "result": result_val,
                        "profit": profit_val,
                        "verbal_result": verbal_result_val,
                        "verbal_decision": verbal_decision_val,
                        "symbol_selected": row_data[variation_col - 9 + 5] or ticker
                    })

            # Store everything in DB
            cur.execute(f"""
                INSERT INTO {matched_table_name} (params, profit_percentage, trial_number, variation_number, stocks_bought)
                VALUES (?, ?, ?, ?, ?)
            """, (json.dumps(params), pp, base_trial, var_num, json.dumps(stocks)))

        conn.commit()

    conn.close()
    print("\n All variations saved to DB.")

# mp1 p3 start
def export_variations_to_excel():
    wb1 = xw.Book(EXCEL_FILE)
    control_sheet = wb1.sheets["ControlSheet"]

    WB2_FILE = r"C:/REDACTED/WB2.xlsm"
    wb2 = xw.Book(WB2_FILE)
    output_sheet = wb2.sheets["ControlSheet"]

    import time
    output_sheet.range("P3").value = "wipe_trigger"
    wb2.app.calculate()
    time.sleep(5)

    conn = sqlite3.connect(optuna_db_path)
    cur = conn.cursor()
    STOCKS_DB_FILE = r"C:/REDACTED/stocks_data.db"
    sd_conn = sqlite3.connect(STOCKS_DB_FILE)
    sd_cur = sd_conn.cursor()

    start_cell_row = 6
    start_cell_col = 3
    param_start_col = 56
    row_buffer_between_variations = 2
    col_buffer_between_algorithms = 16

    ALGO_DISPLAY_NAMES = [
        "MACD", "EMA", "RSI", "Breakout", "ADX", "Volatility",
        "SMA", "Bollinger_Bands", "EMA_MACD_Combo", "RSI_Bollinger"
    ]

    for algo_index, algo_name in enumerate(ALGO_DISPLAY_NAMES):

        cur.execute(f"SELECT * FROM {algo_name}")
        rows = cur.fetchall()

        if not rows:
            continue

        algo_col_start = start_cell_col + (algo_index * col_buffer_between_algorithms)
        param_col_start = algo_col_start + 5

        header_row = 3 + (algo_index * 11)
        param_headers = []
        for j in range(10):
            val = control_sheet.range((header_row, param_start_col + j)).value
            if val:
                param_headers.append(val)
            else:
                break
              
        # choose PP floor, round-down-to-50 target, 1.5× oversample, dedup, cap to T
        # Build a flat list of all variations for this algo             
        all_vars = []
        for row_data in rows:
            _, param_json, pp, trial_num, var_num, stock_json = row_data
            try:
                pp = float(pp)
                if not stock_json:
                    continue
                params = json.loads(param_json)
                stocks = json.loads(stock_json)
                all_vars.append({
                    "pp": pp,
                    "trial_num": trial_num,
                    "var_num": var_num,
                    "params": params,
                    "stocks": [s for s in stocks if s.get("profit") not in [None, "", "null"]],
                })
            except:
                continue

        if not all_vars:
            print(f"[{algo_name}] No variations available in DB. Skipping.")
            continue
        FLOORS = [50, 45, 40, 35, 30, 25]
        ref_floor = None
        ref_pool = []
        for f in FLOORS:
            pool = [v for v in all_vars if v["pp"] >= f]
            if pool:
                ref_floor = f
                ref_pool = sorted(pool, key=lambda x: -x["pp"])
                break

        if ref_floor is None:
            print(f"[{algo_name}] No variations at ≥25% PP. Skipping.")
            continue

        S_ref = len(ref_pool)

        T = 50 if S_ref > 0 and S_ref < 50 else ( (S_ref // 50) * 50 if S_ref >= 50 else 0 )
        if T == 0:
            print(f"[{algo_name}] No supply at reference floor. Skipping.")
            continue

        # Build candidate pool from TOP PP OVERALL 
        all_sorted = sorted(all_vars, key=lambda x: -x["pp"])
        total_avail = len(all_sorted)
        if T == 50:
            want = min(75, total_avail)
        else:
            want = min(int(round(1.5 * T)), total_avail)
        candidates = all_sorted[:want]

        print(f"[{algo_name}] RefFloor={ref_floor}% | S_ref={S_ref} | Target T={T} | Candidates(from ALL)={len(candidates)}")

        #  Deduplicate
        def variation_similarity(p1, p2, algo_name):
            same = 0
            total = min(len(p1), len(p2), len(PARAM_RANGES[algo_name]))
            for i in range(total):
                base = PARAM_RANGES[algo_name][i]
                try:
                    if algo_name == "Volatility" and i == 1:
                        a_min, a_max = map(float, str(p1[1]).split("-"))
                        b_min, b_max = map(float, str(p2[1]).split("-"))
                        base_min = PARAM_RANGES["Volatility"][1]
                        base_max = PARAM_RANGES["Volatility"][2]
                        low, high, step = base_min
                        tol_min = step
                        if abs(a_min - b_min) <= tol_min:
                            same += 1
                        low, high, step = base_max
                        tol_max = step
                        if abs(a_max - b_max) <= tol_max:
                            same += 1
                        continue
                    if algo_name == "Volatility" and i == 2:
                        continue
                    if isinstance(base, list):
                        if str(p1[i]) == str(p2[i]):
                            same += 1
                        continue
                    low, high, step = base
                    options = round((high - low) / step) + 1
                    tol = 0 if options <= 10 else step if options <= 40 else 2 * step if options <= 150 else 3 * step
                    if abs(p1[i] - p2[i]) <= tol:
                        same += 1
                except:
                    continue
            return same > (total // 2)

        def ticker_overlap(s1, s2):
            t1 = set(s.get("ticker") for s in s1 if s.get("ticker"))
            t2 = set(s.get("ticker") for s in s2 if s.get("ticker"))
            return len(t1 & t2) / max(len(t1), len(t2)) if t1 and t2 else 0

        deduped = []
        for v in candidates:
            dup = False
            for d in deduped:
                if algo_name == "EMA":
                    t1 = set(s.get("ticker") for s in v["stocks"] if s.get("ticker"))
                    t2 = set(s.get("ticker") for s in d["stocks"] if s.get("ticker"))
                    if abs(len(t1.symmetric_difference(t2))) <= 1:
                        # keep the one with more tickers or higher PP
                        if (len(t1), v["pp"]) > (len(t2), d["pp"]):
                            deduped.remove(d)
                            deduped.append(v)
                        dup = True
                        break
                else:
                    if variation_similarity(v["params"], d["params"], algo_name) and ticker_overlap(v["stocks"], d["stocks"]) >= 0.6:
                        # keep higher PP; if tied, keep more tickers
                        if (v["pp"], len(v["stocks"])) > (d["pp"], len(d["stocks"])):
                            deduped.remove(d)
                            deduped.append(v)
                        dup = True
                        break
            if not dup:
                deduped.append(v)

        kept = len(deduped)
        print(f"[{algo_name}] Dedup kept {kept} / {len(candidates)}")

        # Final selection size
        selected = sorted(deduped, key=lambda x: -x["pp"])
        if kept > T:
            selected = selected[:T]

        if not selected:
            print(f"[{algo_name}] Nothing left after dedup. Skipping.")
            continue

        # score for ordering
        import math
        e = 1.15
        scored = []
        for entry in selected:
            pp = entry["pp"]
            tickers = len(entry["stocks"])
            entry["tickers"] = tickers
            entry["universal_score"] = math.log(tickers + 1) * (pp ** e)
            scored.append(entry)

        scored.sort(key=lambda x: x["universal_score"], reverse=True)

        print(f"[INFO] Exporting {len(scored)} variations to WB2 (post‑dedup, capped to T)")
        print(f"Universal Score Formula → ln(T+1) * (PP^{e})\n")
        for i, entry in enumerate(scored[:min(25, len(scored))], 1):  # print sample
            print(f"[{i}] Trial {entry['trial_num']} | Var {entry['var_num']} | PP {entry['pp']:.2f}% | "
                  f"Tickers: {entry['tickers']} | Score: {round(entry['universal_score'], 2)}")

        table_name = f"pruned_{algo_name}"
        sd_cur.execute(f"""
            CREATE TABLE IF NOT EXISTS {table_name} (
                identifier TEXT PRIMARY KEY,
                algo_idx INTEGER,
                algo_name TEXT,
                trial_num INTEGER,
                var_num INTEGER,
                pp REAL,
                universal_score REAL,
                params_json TEXT,
                param_headers_json TEXT,
                stocks_json TEXT
            )
        """)
        sd_cur.execute(f"DELETE FROM {table_name}")
        to_insert = []
        for i, v in enumerate(scored, 1):
            uid = f"No.{algo_index+1}.{i}"
            row = (
                uid,
                algo_index,
                algo_name,
                int(v["trial_num"]),
                int(v["var_num"]),
                float(v["pp"]),
                float(v["universal_score"]),
                json.dumps(v["params"]),
                json.dumps(param_headers),
                json.dumps(v["stocks"]),
            )
            to_insert.append(row)

        sd_cur.executemany(
            f"INSERT OR REPLACE INTO {table_name} "
            f"(identifier, algo_idx, algo_name, trial_num, var_num, pp, universal_score, params_json, param_headers_json, stocks_json) "
            f"VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            to_insert
        )
        sd_conn.commit()
        print(f"[DB] Wrote {len(to_insert)} rows to {table_name} (IDs No.{algo_index+1}.1 → No.{algo_index+1}.{len(to_insert)})")

        row_cursor = start_cell_row
        variation_counter = 0
        for v in scored:

            stocks = v["stocks"]
            block_height = 2 + 1 + 1 + len(stocks) + row_buffer_between_variations
            variation_counter += 1
            uid = f"No.{algo_index+1}.{variation_counter}"

            # Identifier cells
            id_data = [
                uid,
                block_height,
                v["trial_num"],
                v["var_num"],
                round(v["universal_score"], 2)
            ]
            output_sheet.range((row_cursor, algo_col_start)).value = id_data

            #   PP Profit/Loss Label 
            if variation_counter == 1:
                label_cell = ""
                if stocks:
                    try:
                        result_val = float(stocks[0].get("result", 0))
                        if result_val > 0:
                            label_cell = "Profit"
                        elif result_val < 0:
                            label_cell = "Loss"
                    except:
                        pass
                output_sheet.range((row_cursor - 1, algo_col_start - 1)).value = label_cell

            # Parameters
            output_sheet.range((row_cursor, param_col_start)).value = [param_headers]
            output_sheet.range((row_cursor + 1, param_col_start)).value = [v["params"]]

            # Result headers
            result_headers = [
                "Ticker", "Name", "Final Final Decision", "Result of Buy or Not",
                "Total Profit or Loss from Trade", "Verbal Profit or Loss",
                "Verbal Buy or Not", "symbol selected"
            ]
            output_sheet.range((row_cursor + 3, param_col_start)).value = [result_headers]

            # Stock row data 
            if stocks:
                def is_swapped_format(s):
                    return (
                        isinstance(s.get("result"), str) and s["result"].startswith("A") and
                        isinstance(s.get("verbal_result"), (int, float))
                    )

                if is_swapped_format(stocks[0]):
                    print("[ ] Misaligned verbal/numeric detected — fixing field positions...")
                    for s in stocks:
                        s["result"], s["verbal_result"] = s["verbal_result"], s["result"]
                        s["profit"], s["verbal_decision"] = s["verbal_decision"], s["profit"]

            result_data = []
            for s in stocks:
                def safe(val):
                    return val if val not in [None, "", "null"] else ""
                try:
                    result_val = round(float(s.get("result", 0)), 5)
                except:
                    result_val = ""

                try:
                    profit_val = round(float(s.get("profit", 0)), 2)
                except:
                    profit_val = ""

                result_data.append([
                    safe(s.get("ticker")),
                    safe(s.get("name")),
                    safe(s.get("decision")),
                    result_val,
                    profit_val,
                    safe(s.get("verbal_result")),
                    safe(s.get("verbal_decision")),
                    s.get("symbol_selected") or safe(s.get("ticker"))
                ])

            if result_data:
                output_sheet.range((row_cursor + 4, param_col_start)).value = result_data


            row_cursor += block_height
    try:
        sd_conn.commit()
    finally:
        sd_conn.close()

    wb2.save()
    print("\nMP1 - Part 3 Completed. Variations exported to WB2.")

# dictionary filler
import os
import pandas as pd
import sqlite3
from pathlib import Path
from datetime import datetime

def extract_clean_date(raw):
    try:
        return datetime.strptime(raw.split(" GMT")[0], "%a %b %d %Y %H:%M:%S").strftime("%Y-%m-%d")
    except:
        return raw if pd.notnull(raw) else None

def load_latest_csvs_into_dictionary(db_path="stocks_data.db", count=4):
    downloads_path = str(Path.home() / "Downloads")
    base_name = "All_Stocks"

    all_files = [
        f for f in os.listdir(downloads_path)
        if f.startswith(base_name) and f.endswith(".csv")
    ]
    full_paths = [os.path.join(downloads_path, f) for f in all_files]
    sorted_files = sorted(full_paths, key=os.path.getmtime, reverse=True)[:count]

    print("  Latest CSVs to load:")
    for f in sorted_files:
        print("  -", os.path.basename(f))

    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    #   Clear previous content
    print("\n  Clearing existing DICTIONARY_TABLE...")
    cur.execute("DELETE FROM DICTIONARY_TABLE")
    conn.commit()
    print("  Clearing existing SEEDLING_TABLE...")
    cur.execute("DELETE FROM SEEDLING_TABLE")
    conn.commit()


    total_blocks_inserted = 0

    for file_path in sorted_files:
        try:
            #   Load raw
            raw_df = pd.read_csv(file_path, header=None, dtype=str).iloc[:, :8]
            raw_df.columns = [
                "Ticker_name", "Date", "Open", "High", "Low", "Close", "Volume", "Volume (Again)"
            ]

            #   Convert only date values
            raw_df["Date"] = raw_df["Date"].apply(extract_clean_date)
            skipped_blocks = [] 

            rows_added = 0
            for i in range(0, len(raw_df), 52):
                block = raw_df.iloc[i:i+52].copy()
                if len(block) < 2:
                    continue  
                try:
                    open_val = float(block.iloc[1]["Open"])
                    if block.iloc[1]["Open"].strip().lower() == "open":
                        raise ValueError 
                except:
                    open_val = 0
                if open_val < 2:
                    #   Insert under-$1 block directly into SEEDLING_TABLE with correct formatting
                    empty_row = pd.DataFrame([[""] * 8], columns=block.columns)
                    full_block = pd.concat([empty_row, block], ignore_index=True)
                    full_block.to_sql("SEEDLING_TABLE", conn, if_exists="append", index=False)
                    continue  


                #   Insert block with blank row on top
                empty_row = pd.DataFrame([[""] * 8], columns=block.columns)
                full_block = pd.concat([empty_row, block], ignore_index=True)
                full_block.to_sql("DICTIONARY_TABLE", conn, if_exists="append", index=False)
                rows_added += len(full_block)

            blocks = rows_added // 53
            total_blocks_inserted += blocks

            print(f"\n  File: {os.path.basename(file_path)}")
            print(f"   Columns: {list(raw_df.columns)}")
            print(f"  Inserted: {blocks} ticker blocks ({rows_added} rows) from {os.path.basename(file_path)}")

        except Exception as e:
            print(f"  Error loading {os.path.basename(file_path)}: {e}")

    try:
        cur.execute("PRAGMA table_info(DICTIONARY_TABLE)")
        columns = [row[1] for row in cur.fetchall()]
        print(f"\n  Final columns in DICTIONARY_TABLE: {columns}")
    except:
        print("  Couldn't fetch final column list.")

    conn.close()
    print(f"\n  Finished. Total tickers inserted: {total_blocks_inserted}")

# first dataset filler
import sqlite3
import pandas as pd
import random

def populate_dataset_one_from_dictionary(db_path=r"C:/REDACTED/stocks_data.db", dataset_size=1000):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    #   Load full dictionary
    df = pd.read_sql_query("SELECT * FROM DICTIONARY_TABLE", conn)

    #   Calculate total blocks
    total_rows = len(df)
    total_blocks = total_rows // 53
    start_indices = [i * 53 for i in range(total_blocks)]

    #   Shuffle and pick top amount
    random.shuffle(start_indices)
    selected = start_indices[:dataset_size]

    #   Print top 10 block info for debug
    print("\n Top 10 selected blocks:")
    for idx in selected[:10]:
        ticker = df.iloc[idx + 2]["Ticker_name"] if idx + 2 < total_rows else "OUT OF BOUNDS"
        print(f"  - Row {idx + 1}: {ticker}")

    print("\n  Clearing DATASET_ONE...")
    cur.execute("DELETE FROM DATASET_ONE")
    conn.commit()

    rows_written = 0
    for idx in selected:
        block = df.iloc[idx:idx + 53].copy()
        if len(block) == 53:
            block.to_sql("DATASET_ONE", conn, if_exists="append", index=False)
            rows_written += 53

    conn.close()

    print(f"\n  Done. Wrote {len(selected)} blocks / {rows_written} rows to DATASET_ONE.")

# mp2 p2 start 
import math
import threading
import time
from collections import defaultdict
import sys
sys.stdout.reconfigure(line_buffering=True)
from threading import Event
from collections import defaultdict
class ThreadTimeout(Exception):
    pass

# group-size counters
group_size_counter = defaultdict(int) 
top_size_counter = defaultdict(int)  
# Total time budget set to 0 to disable the limit. For example it could be 86400 for 24 hours
TOTAL_WALL_BUDGET = 0  
_base_per_task_limit = 0.0 
_delta_per_task = 0.0       
_finished_tasks = 0

_limit_lock = threading.RLock()

_global_start = time.time()
_target_end = _global_start + TOTAL_WALL_BUDGET    
_global_end  = _global_start + (TOTAL_WALL_BUDGET * 1.05)

def normalize_stocks(stocks):
    if not stocks:
        return []
    def is_swapped(s):
        return (
            isinstance(s.get("result"), str) and s["result"].startswith("A")
            and isinstance(s.get("verbal_result"), (int, float))
        )
    if stocks and is_swapped(stocks[0]):
        for s in stocks:
            s["result"], s["verbal_result"] = s.get("verbal_result"), s.get("result")
            s["profit"], s["verbal_decision"] = s.get("verbal_decision"), s.get("profit")
    return stocks

exit_event = Event()

import heapq
best_heap = []  
best_index = {} 
temp_all_groups = [] 


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
    global best_heap, best_index, temp_all_groups
    global group_size_counter, top_size_counter

    best_heap.clear()
    best_index.clear()
    temp_all_groups.clear()
    group_size_counter.clear()
    top_size_counter.clear()

    # per‑algo stats box
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

    all_buckets = []

    from threading import Event
    exit_event = Event()

    WATCHDOG_SECONDS = None
    def watchdog():
        if WATCHDOG_SECONDS is None:
            return
        time.sleep(WATCHDOG_SECONDS)
        exit_event.set()
        print("[ Watchdog] Timeout reached. Stopping threads...", flush=True)

    # Start watchdog
    if WATCHDOG_SECONDS:
        threading.Thread(target=watchdog, daemon=True).start()

    from queue import Queue
    db_queue = Queue()
    WB2_FILE = r"C:/REDACTED/WB2.xlsm"
    wb = xw.Book(WB2_FILE)
    sheet = wb.sheets[0]

    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS temp_groups (
            identifier TEXT PRIMARY KEY,
            pp REAL,
            score REAL,
            reason TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS groups (
            identifier TEXT PRIMARY KEY,
            pp REAL,
            score REAL,
            parameters TEXT
        )
    ''')

    cursor.execute("PRAGMA table_info(groups)")
    columns = [row[1] for row in cursor.fetchall()]
    if "parameters" not in columns:
        cursor.execute("ALTER TABLE groups ADD COLUMN parameters TEXT")
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

    best_groups = []
    best_lock = threading.Lock()

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

    # Build a map 
    algo_idx_to_table = {i: t for i, t in enumerate(PRUNED_TABLES)}

    # One bulk read per table
    for algo_idx in range(len(PRUNED_TABLES)):
        table = algo_idx_to_table[algo_idx]
        cursor.execute(f"SELECT identifier, pp, universal_score, params_json, param_headers_json, stocks_json "
                       f"FROM {table}")
        rows = cursor.fetchall()

        if not rows:
            variation_map[algo_idx] = []
            continue

        first_headers = json.loads(rows[0][4]) if rows[0][4] else []
        param_headers_map[algo_idx] = first_headers

        rows.sort(key=lambda r: (float(r[1]), float(r[2])), reverse=True)

        variation_ids = []
        for identifier, pp, uscore, params_json, _headers_json, stocks_json in rows:
            variation_ids.append(identifier)

            # params + stocks
            params = json.loads(params_json) if params_json else []
            raw_stocks = json.loads(stocks_json) if stocks_json else []
            stocks = normalize_stocks(raw_stocks)

            params_map[(algo_idx, identifier)] = params
            full_stocks_map[(algo_idx, identifier)] = stocks

            # true PP from DB
            try:        
                variation_pp_map[(algo_idx, identifier)] = float(pp)
            except Exception:
                variation_pp_map[(algo_idx, identifier)] = None

        variation_map[algo_idx] = variation_ids
      
    ticker_set_map = {}
    for (a, ident), stocks in full_stocks_map.items():
        tset = {s.get("ticker") for s in stocks if s.get("ticker")}
        ticker_set_map[(a, ident)] = tset

    def _is_profit_stock(s):
        vr = str(s.get("verbal_result", "")).upper()
        if "PROFIT" in vr:
            return True
        try:
            return float(s.get("profit", 0) or 0) > 0
        except Exception:
            return False

    profit_ticker_set_map = {}
    for (a, ident), stocks in full_stocks_map.items():
        pset = {s["ticker"] for s in stocks if s.get("ticker") and _is_profit_stock(s)}
        profit_ticker_set_map[(a, ident)] = pset

    variations_by_ticker = defaultdict(set)        
    profit_variations_by_ticker = defaultdict(set)   
    for key, tset in ticker_set_map.items():
        for t in tset:
            variations_by_ticker[t].add(key)
    for key, pset in profit_ticker_set_map.items():
        for t in pset:
            profit_variations_by_ticker[t].add(key)
          
    # setup for threads logic
  
    def extract_params(row_or_identifier, algo_idx):
        return params_map.get((algo_idx, row_or_identifier), [])

    def intersect_full_stock_data(s1, s2):
        t1 = {s["ticker"]: s for s in s1}
        t2 = {s["ticker"]: s for s in s2}
        return [t1[t] for t in t1 if t in t2]

    def calc_pp(stocks):
        if not stocks: return 0
        return round(100 * sum(1 for s in stocks if "PROFIT" in s["verbal_result"]) / len(stocks), 2)

    def group_score(pp, group_size):
        return round((math.log(group_size + 1) * (pp ** 1.15)), 4)

    PAIRSCAN_CAP = 80   
    SHARED2_CAP  = 250 

    def thread_sanity_summary(base_algo_idx, base_stocks, base_key):
        base_profits = list(profit_ticker_set_map.get(base_key, set()))
        if len(base_profits) < 2:
            return "[ Sanity] prof=<2 → no pairs; shared2_vars=0; shared2_algos=0", 0

        def other_profit_support(t):
            return sum(1 for (a, _i) in profit_variations_by_ticker.get(t, set()) if a != base_algo_idx)

        cand_ticks = sorted(base_profits, key=other_profit_support, reverse=True)[:PAIRSCAN_CAP]

        from itertools import combinations
        best_pair = None  
        for t1, t2 in combinations(cand_ticks, 2):
            v1 = profit_variations_by_ticker.get(t1, set())
            v2 = profit_variations_by_ticker.get(t2, set())
            inter = [(a, i) for (a, i) in (v1 & v2) if a != base_algo_idx]
            if not inter:
                continue
            per_algo = defaultdict(int)
            for a_idx, _ in inter:
                per_algo[a_idx] += 1
            total = len(inter)
            if (best_pair is None) or (total > best_pair[2]):
                best_pair = (t1, t2, total, dict(per_algo))

        if best_pair is None:
            pair_line = "best_pair=<none>"
        else:
            t1, t2, total, per_algo = best_pair
            parts = [f"{PRUNED_TABLES[a].replace('pruned_','')}:{c}"
                     for a, c in sorted(per_algo.items(), key=lambda kv: (-kv[1], kv[0]))[:6]]
            pair_line = f"best_pair=({t1},{t2}) in {total} vars (by algo: {', '.join(parts)})"

        from collections import Counter
        ctr = Counter()
        for t in base_profits[:SHARED2_CAP]:
            for (a, i) in profit_variations_by_ticker.get(t, set()):
                if a != base_algo_idx:
                    ctr[(a, i)] += 1

        total_shared2 = 0
        per_algo_shared2 = defaultdict(int)
        for (a, i), k in ctr.items():
            if k >= 2:
                total_shared2 += 1
                per_algo_shared2[a] += 1

        shared2_algo_count = sum(1 for _a, c in per_algo_shared2.items() if c > 0)

        parts2 = [f"{PRUNED_TABLES[a].replace('pruned_','')}:{c}"
                  for a, c in sorted(per_algo_shared2.items(), key=lambda kv: (-kv[1], kv[0]))[:6]]
        shared2_line = f"shared2_vars={total_shared2}; shared2_algos={shared2_algo_count}" + \
                       (f" (by algo: {', '.join(parts2)})" if parts2 else "")

        return f"[Sanity] prof={len(base_profits)} | {pair_line} | {shared2_line}", len(base_profits)

    def score_group(pp, group_size, shared_count):
        return round(pp * math.log1p(group_size) * math.log1p(shared_count), 4)

    final_groups = []  
    final_identifiers = []

    big_candidates = {3: [], 4: [], 5: [], 6: []}
    big_seen_keys = set() 

    def is_duplicate_group(new_group, shared_tickers, pp, score, identifier_map):
        new_algos = set(a for a, _ in new_group)
        new_tickers = {s["ticker"] for s in shared_tickers}
        identifier = "_".join([identifier_map[(a, r)] for a, r in new_group])

        if not best_groups:
            return None
        # prune old single-algo groups
        for i, (old_algos, old_tickers, old_id, old_pp, old_score) in enumerate(best_groups):
            if len(old_algos) == 1:
                print(f" Swap] {old_id} removed — single algorithm → replaced by {identifier}")
                best_groups[i] = (new_algos, new_tickers, identifier, pp, score)
                return i

        lowest_pp = min(bg[3] for bg in best_groups)

        for i, (old_algos, old_tickers, old_id, old_pp, old_score) in enumerate(best_groups):
            reason = f"Evaluating {identifier} vs {old_id}: "

            if old_pp != lowest_pp:
                continue

            if pp > old_pp:
                print(f"{reason}  Replaced due to higher PP ({pp}% > {old_pp}%)")
                best_groups[i] = (new_algos, new_tickers, identifier, pp, score)
                return i

            if len(new_algos) > len(old_algos):
                print(f"{reason}  Replaced due to more algorithms ({len(new_algos)} > {len(old_algos)})")
                best_groups[i] = (new_algos, new_tickers, identifier, pp, score)
                return i

            if len(new_tickers) > len(old_tickers):
                print(f"{reason}  Replaced due to more tickers ({len(new_tickers)} > {len(old_tickers)})")
                best_groups[i] = (new_algos, new_tickers, identifier, pp, score)
                return i

            ticker_counts = {}
            for _, t_set, _, _, _ in best_groups:
                for t in t_set:
                    ticker_counts[t] = ticker_counts.get(t, 0) + 1

            old_unique = sum(1 for t in old_tickers if ticker_counts.get(t, 0) == 1)
            new_unique = sum(1 for t in new_tickers if ticker_counts.get(t, 0) == 1)

            if new_unique > old_unique:
                print(f"{reason}  Replaced due to more unique tickers ({new_unique} > {old_unique})")
                best_groups[i] = (new_algos, new_tickers, identifier, pp, score)
                return i

            print(f"{reason}  Not replaced — all metrics worse or equal")

        return None

    TOP_N_GROUPS = 100

    # keep the single best group for larger sizes
    best_by_size = {}  

    from itertools import combinations
    MAX_GROUP_SIZE = 6  
    all_tested_combos = set()

    def build(group, shared, used_algos, hits, deadline=None):
        if exit_event.is_set() or time.time() >= _global_end:
            return
        # Dynamic time check
        if deadline is not None:
            elapsed = time.time() - deadline
            if elapsed >= current_per_task_limit():
                return
        hits[0] += 1  
        # count attempts for visibility
        if shared:
            profit_flags = []
            for s in shared:
                vr = str(s.get("verbal_result", "")).upper()
                if "PROFIT" in vr:
                    profit_flags.append(1)
                else:
                    try:
                        profit_flags.append(1 if float(s.get("profit", 0) or 0) > 0 else 0)
                    except Exception:
                        profit_flags.append(0)
            profits = sum(profit_flags)
            if profits < 2:
                return 
            id_set_key = None
            if len(group) >= 2:
                group_size_counter[len(group)] += 1
                id_list = [identifier_map[(a, r)] for a, r in group]
                identifier = "_".join(id_list)
                id_set_key = frozenset(id_list) 

                if id_set_key in all_tested_combos:
                    return
                all_tested_combos.add(id_set_key)

            profit_flags = []
            for s in shared:
                vr = str(s.get("verbal_result", "")).upper()
                if "PROFIT" in vr:
                    profit_flags.append(1)
                elif "LOSS" in vr:
                    profit_flags.append(0)
                else:
                    try:
                        profit_flags.append(1 if float(s.get("profit", 0) or 0) > 0 else 0)
                    except Exception:
                        profit_flags.append(0)

            profits = sum(profit_flags)
            pp = round(100 * profits / len(shared), 2)
            score = round((math.log(len(group) + 1) * (pp ** 1.15)), 4)
            _algo_size = len(set(a for a, _ in group))
            if pp == 100.0 and _algo_size in (3, 4, 5, 6) and len(shared) >= 1:
                _ids_fset = frozenset(identifier_map[(a, r)] for a, r in group)
                if _ids_fset not in big_seen_keys:
                    big_seen_keys.add(_ids_fset)
                    big_candidates[_algo_size].append((group.copy(), shared.copy(), pp, score))
            if pp == 100.0:
                with best_lock:
                    algo_count = len(set(a for a, _ in group))
                    ticker_count = len({s["ticker"] for s in shared})

                    def ticker_overlap_ratio(a_shared, b_shared):
                        A = {s["ticker"] for s in a_shared}
                        B = {s["ticker"] for s in b_shared}
                        if not A or not B:
                            return 0.0
                        return len(A & B) / min(len(A), len(B))

                    max_overlap_same_size = 0.0
                    for (_g, _shared, _pp, _sc) in best_index.values():
                        if len(set(ax for ax, _ in _g)) == algo_count:
                            max_overlap_same_size = max(
                                max_overlap_same_size,
                                ticker_overlap_ratio(shared, _shared)
                            )
                    diversity = 1.0 - max_overlap_same_size 

                    rank_key = (algo_count, ticker_count, diversity, pp)

                    # keep full record for temp table
                    temp_all_groups.append((group.copy(), shared.copy(), pp, score))

                    identifier_local = "_".join([identifier_map[(a, r)] for a, r in group])
                    if identifier_local not in best_index:
                        heapq.heappush(best_heap, (rank_key, identifier_local))
                        best_index[identifier_local] = (group.copy(), shared.copy(), pp, score)

                        if len(best_heap) > TOP_N_GROUPS:
                            cur_items = list(best_heap)
                            worst_key = min(k for (k, _) in cur_items)

                            def _overlap_with_shared(candidate_id):
                                _g, _shared, _pp, _sc = best_index[candidate_id]
                                A = {s["ticker"] for s in shared}
                                B = {s["ticker"] for s in _shared}
                                if not A or not B:
                                    return 0.0
                                return len(A & B) / min(len(A), len(B))

                            tied_ids = [i for (k, i) in cur_items if k == worst_key]
                            if len(tied_ids) == 1:
                                _, worst_id = heapq.heappop(best_heap)
                                best_index.pop(worst_id, None)
                            else:
                                worst_id_by_overlap = max(tied_ids, key=_overlap_with_shared)
                                new_heap = [(k, i) for (k, i) in best_heap if i != worst_id_by_overlap]
                                heapq.heapify(new_heap)
                                best_heap[:] = new_heap
                                best_index.pop(worst_id_by_overlap, None)

        # Stop only when max size reached
        if len(group) >= MAX_GROUP_SIZE:
            return
        max_algo_in_group = max(a for a, _ in group)
        for next_algo in range(max_algo_in_group + 1, len(identifier_cols)):
            if exit_event.is_set():
                return
            if time.time() >= _global_end or (time.time() - deadline) >= current_per_task_limit():
                return

            if next_algo in used_algos:
                continue

            for row in variation_map.get(next_algo, []):
                if time.time() >= _global_end or (time.time() - deadline) >= current_per_task_limit():
                    return
                new_stocks = full_stocks_map[(next_algo, row)]
                tset = ticker_set_map[(next_algo, row)]
                new_shared = [s for s in shared if s['ticker'] in tset]

                if not new_shared:
                    continue

                group.append((next_algo, row))
                used_algos.add(next_algo)
                build(group, new_shared, used_algos, hits, deadline=deadline)

                used_algos.remove(next_algo)
                group.pop()

    total_variations = math.prod(len(v) for v in variation_map.values())
    print(f"\n[ ] Total combinations (approx): {total_variations:,}")
    counter = 0
    checkpoint = max(1, total_variations // 10)

    print("\n[] Selecting ALL variations as thread starters (exhaustive coverage)...")

    # Use every variation
    thread_starters = []
    for a in range(len(identifier_cols)):
        for r in variation_map.get(a, []):
            thread_starters.append((a, r))
    identifier_map = {}
    for algo_idx, id_list in variation_map.items():
        for ident in id_list:
            identifier_map[(algo_idx, ident)] = ident

    # Build per-algo starter lists
    starters_by_algo = {ai: [] for ai in range(len(identifier_cols))}
    for a, r in thread_starters:
        starters_by_algo[a].append((a, r))

    def avg_tickers_for_algo(ai):
        lst = starters_by_algo.get(ai, [])
        if not lst:
            return float("inf")
        sizes = [len(full_stocks_map[(a, r)]) for (a, r) in lst]
        return sum(sizes) / max(1, len(sizes))

    heavy_order = [3, 0, 1]

    small_algos = [ai for ai in range(len(identifier_cols)) if ai not in set(heavy_order)]

    # Sort small algos by average ticker count
    small_algos.sort(key=avg_tickers_for_algo)

    def sort_variations_in_algo(ai):
        lst = starters_by_algo.get(ai, [])
        lst.sort(key=lambda ar: len(full_stocks_map[(ar[0], ar[1])]))
        return lst

    ordered = []
    for ai in small_algos:
        ordered.extend(sort_variations_in_algo(ai))
    for ai in heavy_order:
        ordered.extend(sort_variations_in_algo(ai))

    thread_starters = ordered

    # Quick logs to confirm ordering & sizes
    if thread_starters:
        first_len = len(full_stocks_map[(thread_starters[0][0], thread_starters[0][1])])
        last_len  = len(full_stocks_map[(thread_starters[-1][0], thread_starters[-1][1])])
        print(f"[ ] Starter ordering by ticker count within algos applied. First={first_len}, Last={last_len}")
        print(f"[ ] Heavy algos scheduled last in order: 4 → 1 → 2 (0-based: {heavy_order})")

    total_tasks = len(thread_starters)
    max_workers = min(4, max(1, total_tasks))
    batches = math.ceil(total_tasks / max_workers) or 1
    with _limit_lock:

        _base_per_task_limit = float(TOTAL_WALL_BUDGET) / float(batches)
        _delta_per_task = 0.0
        _finished_tasks = 0
    print(f"[init] tasks={total_tasks} | workers={max_workers} | batches={batches} | base_limit={_base_per_task_limit:.3f}s")

    _bank_seconds = 0.0  # total pooled leftover seconds from finished threads

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
            # Always log info
            if seconds <= 0:
                cur = _base_per_task_limit + _delta_per_task
                print(
                    f"[limit] banked 0.000s from {donor_tag or 'thread'} | "
                    f"applied_total 0.000s (per-task +0.000s) | "
                    f"limit {cur:.3f}→{cur:.3f} | bank_left={_bank_seconds:.3f}s"
                )
                return
            _bank_seconds += seconds
            # evenly distribute across remaining tasks
            remaining = max(1, total_tasks - _finished_tasks)
            per_task_bump = _bank_seconds / float(remaining)

            old_limit = _base_per_task_limit + _delta_per_task
            _delta_per_task += per_task_bump
            new_limit = _base_per_task_limit + _delta_per_task

            applied_total = per_task_bump * remaining
            _bank_seconds = max(0.0, _bank_seconds - applied_total)

            print(
                f"[ limit] banked {seconds:.3f}s from {donor_tag or 'thread'} | "
                f"applied_total {applied_total:.3f}s (per-task +{per_task_bump:.3f}s) | "
                f"limit {old_limit:.3f}→{new_limit:.3f} | bank_left={_bank_seconds:.3f}s"
            )

    def mark_finished():
        nonlocal _finished_tasks
        with _limit_lock:
            _finished_tasks += 1

    pre_thread_start_time = time.time()

    print(f"[  Timer] Pre-thread logic done. Time elapsed: {pre_thread_start_time - total_start_time:.2f} seconds")
    print(f"\n[] Starting {len(thread_starters)} starters with a 4-worker pool...\n")

    def run_thread(base_id, algo, row):
        start = time.time()
        limit_at_start = current_per_task_limit()
        var_id = identifier_map[(algo, row)]
        base_stocks = full_stocks_map[(algo, row)]
        base_pp = variation_pp_map.get((algo, var_id), None)
        base_tickers = len(base_stocks)
        pp_str = f"{base_pp:.2f}%" if isinstance(base_pp, (int, float)) else "NA"
        print(f"[ Thread {base_id+1} START] {var_id} | PP={pp_str} | T={base_tickers} | limit_start={limit_at_start:.2f}s")
        sanity_line, _bp = thread_sanity_summary(algo, base_stocks, (algo, var_id))
        print(f"[Thread {base_id+1}] {sanity_line}")

        # Pass the start time;
        deadline = start
        base_stocks = full_stocks_map[(algo, row)]
        hits = [0]
        try:
            build([(algo, row)], base_stocks, {algo}, hits, deadline=deadline)

        except ThreadTimeout as e:
            print(f"[  Thread {base_id+1}] Exited: {e} | Attempts: {hits[0]}", flush=True)
        finally:
            dur = time.time() - start
            limit_at_end = current_per_task_limit()
            timed_out = (
                dur >= (limit_at_end - 0.01)
                or time.time() >= _global_end
                or exit_event.is_set()
            )

            status = "timeout" if timed_out else "done"
            end_pp_str = f"{base_pp:.2f}%" if isinstance(base_pp, (int, float)) else "NA"
            print(
                f"[  Thread {base_id+1} END] {var_id} | {status} | "
                f"dur={dur:.2f}s | attempts={hits[0]} | "
                f"limit_start={limit_at_start:.2f}s | limit_end={limit_at_end:.2f}s | "
                f"PP={end_pp_str} | T={base_tickers}"
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
    from concurrent.futures import ThreadPoolExecutor, as_completed
    try:
        with ThreadPoolExecutor(max_workers=max_workers) as ex:
            futures = [ex.submit(run_thread, i, algo, row) for i, (algo, row) in enumerate(thread_starters)]
            for f in as_completed(futures):
                if exit_event.is_set() or time.time() >= _global_end:
                    exit_event.set()
                    break

    except KeyboardInterrupt:
        print("[] KeyboardInterrupt received — cancelling...", flush=True)
        exit_event.set()
        # Let running tasks see the flag and exit quickly

    thread_end_time = time.time()
    print(f"[  Timer] Threads finished. Time elapsed: {thread_end_time - pre_thread_start_time:.2f} seconds")
    print(f"[  Timer] TOTAL Time (before + during threads): {thread_end_time - total_start_time:.2f} seconds")

    _rows_for_db = []
    print("\n[  Algo Run Stats]")
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

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS algo_run_stats (
            algo_number INTEGER PRIMARY KEY,
            algo_table   TEXT,
            threads      INTEGER,
            variations   INTEGER,
            timeouts     INTEGER,
            total_runtime_sec REAL,
            avg_pp       REAL,
            avg_tickers  REAL
        )
    """)
    cursor.execute("DELETE FROM algo_run_stats")
    cursor.executemany(
        "INSERT INTO algo_run_stats (algo_number, algo_table, threads, variations, timeouts, total_runtime_sec, avg_pp, avg_tickers) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        _rows_for_db
    )
    conn.commit()

    def quality_key(entry):
        group, shared, pp, score = entry
        algo_count = len(set(a for a, _ in group))
        ticker_count = len({s["ticker"] for s in shared})
        # Sort descending
        return (algo_count, ticker_count, pp)

    final_groups = [best_index[i] for _, i in sorted(best_heap)]
    # Sort descending
    final_groups.sort(key=quality_key, reverse=True)
    # Trim to top N 
    final_groups = final_groups[:TOP_N_GROUPS]

    winners_by_size = {}  

    def _size_quality_key(entry):
        _g, _shared, _pp, _score = entry
        _algo_count = len(set(a for a, _ in _g))
        _ticker_count = len({s["ticker"] for s in _shared})
        return (_algo_count, _ticker_count, _pp)

    for _target in (3, 4, 5, 6):
        cand = big_candidates.get(_target, [])
        if not cand:
            continue
        cand.sort(key=_size_quality_key, reverse=True)
        winners_by_size[_target] = cand[:1]
    # Top-100 by size:
    top_size_counter = defaultdict(int)
    for g, shared, pp, score in final_groups:
        top_size_counter[len(set(a for a, _ in g))] += 1

    print("\n[ ] Group Size Distribution (tested → Top-100):")
    for sz in sorted(set(list(group_size_counter.keys()) + list(top_size_counter.keys()))):
        tested = group_size_counter.get(sz, 0)
        kept = top_size_counter.get(sz, 0)
        print(f"  Size {sz}: {tested:,} tested  →  {kept} in Top-100")

    # clear DB
    cursor.execute('DELETE FROM groups')
    cursor.execute('DELETE FROM temp_groups')
    conn.commit()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS extra_big_groups (
            identifier TEXT PRIMARY KEY,
            pp REAL,
            score REAL,
            parameters TEXT
        )
    ''')
    cursor.execute('DELETE FROM extra_big_groups')

    def _ensure_col(tbl, coldef):
        cursor.execute(f"PRAGMA table_info({tbl})")
        have = {r[1] for r in cursor.fetchall()}
        cname = coldef.split()[0]
        if cname not in have:
            cursor.execute(f"ALTER TABLE {tbl} ADD COLUMN {coldef}")

    for _tbl in ("groups", "extra_big_groups"):
        _ensure_col(_tbl, "ticker_count INTEGER")
        _ensure_col(_tbl, "algo_ticker_counts TEXT")

    conn.commit()

    for _sz in (3, 4, 5, 6):
        for _g, _shared, _pp, _sc in winners_by_size.get(_sz, []):
            _ident = "_".join([identifier_map[(a, r)] for a, r in _g])
            _params = {identifier_map[(a, r)]: extract_params(r, a) for a, r in _g}

            _counts = []
            for (a, r) in _g:
                _cnt = len(full_stocks_map.get((a, r), []))
                _counts.append(f"No.{a+1}:{_cnt}")
            _counts_text = "|".join(_counts) if _counts else ""

            # exact group ticker count 
            _ticker_count = len({s.get("ticker") for s in _shared if s.get("ticker")})

            cursor.execute(
                'REPLACE INTO extra_big_groups (identifier, pp, score, parameters, ticker_count, algo_ticker_counts) VALUES (?, ?, ?, ?, ?, ?)',
                (_ident, _pp, _sc, json.dumps(_params), _ticker_count, _counts_text)
            )

    conn.commit()

    final_identifiers.clear()
    best_groups.clear()

    # write Final 100 Groups 
    for group, shared, pp, score in final_groups:
        algos = set(a for a, _ in group)
        tickers = {s["ticker"] for s in shared}
        identifier = "_".join([identifier_map[(a, r)] for a, r in group])

        parameters = {
            identifier_map[(a, r)]: extract_params(r, a)
            for a, r in group
        }

        final_identifiers.append(identifier)
        best_groups.append((algos, tickers, identifier, pp, score))  

        _counts = []
        for (a, r) in group:
            _cnt = len(full_stocks_map.get((a, r), []))
            _counts.append(f"No.{a+1}:{_cnt}")
        _counts_text = "|".join(_counts) if _counts else ""

        _ticker_count = len({s.get("ticker") for s in shared if s.get("ticker")})

        cursor.execute(
            'REPLACE INTO groups (identifier, pp, score, parameters, ticker_count, algo_ticker_counts) VALUES (?, ?, ?, ?, ?, ?)',
            (identifier, pp, score, json.dumps(parameters), _ticker_count, _counts_text)
        )

    conn.commit()

    print(f"\n[TEMP TABLE] Total Groups Generated: {len(temp_all_groups)}")

    top_identifiers = set(
        "_".join([identifier_map[(a, r)] for a, r in g]) for g, _, _, _ in final_groups
    )

    for group, shared, pp, score in temp_all_groups:
        algos = set(a for a, _ in group)
        ticker_count = len({s["ticker"] for s in shared})
        identifier = "_".join([identifier_map[(a, r)] for a, r in group])

        reason = "  Accepted"
        if identifier not in top_identifiers:
            if len(algos) == 1:
                reason = f"  Rejected: Only 1 algorithm"
            else:
                reason = f"  Rejected: Outranked — Not in Top 100 (PP={pp}%, Algos={len(algos)}, Tickers={ticker_count})"

        cursor.execute(
            'REPLACE INTO temp_groups (identifier, pp, score, reason) VALUES (?, ?, ?, ?)',
            (identifier, pp, score, reason)
        )

    conn.commit()


    sheet.range("J3").value = f"wipe_{time.time()}"

    # Export top 100 to WB2
    row_cursor = paste_row
    paste_col = 168  

    print(f"\n[ ] Exporting {len(final_groups)} final groups to Excel...")

    for i, (group, shared, pp, score) in enumerate(final_groups):
        identifier = "_".join([identifier_map[(a, r)] for a, r in group])

        # identifiers
        sheet.range((row_cursor + 1, 165)).value = identifier 
        sheet.range((row_cursor + 1, 166)).value = pp         
        sheet.range((row_cursor + 1, 167)).value = score      

        # parameter headers and values
        total_rows = 0
        for a, r in group:
            headers = param_headers_map[a]
            values = extract_params(r, a)
            for j, h in enumerate(headers):
                sheet.range((row_cursor, paste_col + j)).value = h
            for j, v in enumerate(values):
                sheet.range((row_cursor + 1, paste_col + j)).value = v
            row_cursor += 3
            total_rows += 3

        result_headers = [
            "Ticker", "Name", "Final Final Decision", "Result of Buy or Not",
            "Total Profit or Loss from Trade", "Verbal Profit or Loss",
            "Verbal Buy or Not", "symbol selected"
        ]
        for j, h in enumerate(result_headers):
            sheet.range((row_cursor, paste_col + j)).value = h

        result_values = []
        for stock in shared:
            result_values.append([
                stock.get("ticker", ""),
                stock.get("name", ""),
                stock.get("decision", ""),
                stock.get("result", ""),
                stock.get("profit", ""),
                stock.get("verbal_result", ""),
                stock.get("verbal_decision", ""),
                stock.get("symbol_selected", "")
            ])

        sheet.range((row_cursor + 1, paste_col)).value = result_values

        print(f"  → Exported group {i + 1}/{len(final_groups)} | ID: {identifier} | Stocks: {len(shared)}")

        row_cursor += len(shared) + 3

    side_row = 7
    side_col = 191  
    print("\n[ ] Exporting best size-5 and size-4 groups to GE7...")

    def paste_one_group_at(col, row, group, shared, pp, score):
        ident = "_".join([identifier_map[(a, r)] for a, r in group])

        sheet.range((row + 1, col - 3)).value = ident
        sheet.range((row + 1, col - 2)).value = pp
        sheet.range((row + 1, col - 1)).value = score
        rc = row
        for a, r in group:
            headers = param_headers_map.get(a, [])
            values  = extract_params(r, a)
            if headers:
                sheet.range((rc, col)).value = headers
            if values:
                sheet.range((rc + 1, col)).value = values
            rc += 3 

        result_headers = [
            "Ticker", "Name", "Final Final Decision", "Result of Buy or Not",
            "Total Profit or Loss from Trade", "Verbal Profit or Loss",
            "Verbal Buy or Not", "symbol selected"
        ]
        sheet.range((rc, col)).value = result_headers

        vals = []
        for s in shared:
            vals.append([
                s.get("ticker", ""),
                s.get("name", ""),
                s.get("decision", ""),
                s.get("result", ""),
                s.get("profit", ""),
                s.get("verbal_result", ""),
                s.get("verbal_decision", ""),
                s.get("symbol_selected", "")
            ])
        if vals:
            sheet.range((rc + 1, col)).value = vals

        return rc + 1 + len(vals) 

    for size in (3, 4, 5, 6):
        entries = winners_by_size.get(size, [])
        for (g, sh, pp, sc) in entries:
            end_row = paste_one_group_at(side_col, side_row, g, sh, pp, sc)
            side_row = end_row + 2 

    print(f"\n[ ] Final Best Groups ({len(best_groups)}):")
    printed = set()
    for algo_set, ticker_set, identifier, pp, score in best_groups:
        if identifier not in printed:
            readable_algos = {f"Algo {a+1}" for a in algo_set}
            print((readable_algos, ticker_set, identifier, pp, score))
            printed.add(identifier)

    print("...")

    def export_buckets_to_db():
        cursor.execute('''CREATE TABLE IF NOT EXISTS seen_groups (
            identifier TEXT PRIMARY KEY
        )''')
        for b in buckets:
            for ident in b:
                cursor.execute("INSERT OR IGNORE INTO seen_groups (identifier) VALUES (?)", (ident,))
        conn.commit()

    wb.save()

    total_checked = sum(1 for _ in all_tested_combos)
    print(f"[ ] Unique group identifiers tested: {len(all_tested_combos)}")

    print(f"[  MP2 FINAL] Exported Top {len(final_groups)} groups.")
    print(f"\n[ ] Final Best Groups Array ({len(best_groups)} total):")
    for i, (algo_set, ticker_set, identifier, pp, score) in enumerate(best_groups):
        readable_algos = [f"Algo {a+1}" for a in sorted(algo_set)]
        print(f"{i+1}. ID: {identifier} | Algos: {readable_algos} | Tickers: {len(ticker_set)} | PP: {pp}% | Score: {score}")

    print(f"\n[ ] Total combinations tested: {len(all_tested_combos)}")
    print("\n[ BUCKET SUMMARY] (disabled in exhaustive mode)")

# mp2 p3
import re

EXCEL_FILE = r"C:/REDACTED/WB1.xlsm"
WB2_FILE = r"C:/REDACTED/WB2.xlsm"
wb2 = xw.Book(WB2_FILE)
wb2_sheet = wb2.sheets["ControlSheet"]
DB_FILE = r"C:/REDACTED/stocks_data.db"

CONTROL_SHEET_NAME = "ControlSheet"
WB2_GROUP_START_CELL = "FZ6"
SHIFT_DOWN_MACRO = "ShiftFormulasDown"
SHIFT_UP_MACRO = "ShiftFormulasUp"
TRIGGER_MACRO_CELL = "AA1"
RESET_MACRO_CELL = "AC1"

PARAM_START_COL = 56
PP_COL = 48

def extract_group_parameters(group_id):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    cursor.execute("SELECT identifier, pp, score, parameters FROM groups ORDER BY ROWID ASC")

    all_groups = cursor.fetchall()
    conn.close()

    if group_id >= len(all_groups):
        return None 

    identifier, pp, score, param_str = all_groups[group_id]
    param_map = json.loads(param_str)
    algo_list = []
    param_chunks = []

    for key in param_map:
        match = re.match(r"No\.(\d+)", key)
        if not match:
            continue
        algo_num = int(match.group(1))
        algo_list.append(algo_num)
        param_chunks.append(param_map[key])

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

# macro trigger
def trigger_macro(sheet, macro_cell, delay=1):
    sheet.range(macro_cell).value = 100
    sheet.book.save()
    time.sleep(delay)

# extract buys
def collect_group_buys(wb, algorithms):
    algo_buy_sets = []

    start_data_row = 55
    ticker_col = 9  
    name_col = 10  
    final_decision_col = 11 

    SHEET_MAP = {
        "1": "1 MACD", "2": "2 EMA", "3": "3 RSI", "4": "4 Breakout", "5": "5 ADX",
        "6": "6 Volatility Measure", "7": "7 SMA", "8": "8 Bollinger Bands",
        "9": "9 EMA & MACD", "10": "10 RSI Bollinger"
    }

    print("\n [BUY CHECK] Listing buys for each algorithm:")

    for algo_entry in algorithms:
        algo_num = int(algo_entry)
        sheet_name = SHEET_MAP[str(algo_num)]
        sheet = wb.sheets[sheet_name]

        row = start_data_row
        algo_buys = {}

        print(f"\nAlgorithm {algo_num}: {sheet_name}")
        found = False

        while True:
            ticker = sheet.range((row, ticker_col)).value
            name = sheet.range((row, name_col)).value
            decision = sheet.range((row, final_decision_col)).value
            if not ticker:
                break
            if str(decision).strip().upper() == "BUY":
                print(f"    {ticker} | {name} | {decision}")
                algo_buys[ticker] = (ticker, name, decision)
                found = True
            row += 1

        if not found:
            print("    No buys found.")

        algo_buy_sets.append(algo_buys)

    common_tickers = set(algo_buy_sets[0].keys())
    for buy_dict in algo_buy_sets[1:]:
        common_tickers &= set(buy_dict.keys())

    final_buys = [algo_buy_sets[0][ticker] for ticker in common_tickers]
    return final_buys

def run_mp2_part3_group_executor():
    wb = xw.Book(EXCEL_FILE)
    wb2 = xw.Book(WB2_FILE)
    control_sheet = wb.sheets["ControlSheet"]
    wb2_sheet = wb2.sheets["ControlSheet"]

    trigger_macro(control_sheet, TRIGGER_MACRO_CELL, delay=15)

    rows_per_dataset = 53
    batch_size = int(control_sheet.range("AG4").value)     
    result_limit = int(control_sheet.range("AG3").value)
    batches_per_cycle = result_limit // batch_size           

    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM DICTIONARY_TABLE")
    total_rows = cur.fetchone()[0]
    total_datasets = total_rows // rows_per_dataset

    group_index = 0

    while True:  
        control_sheet.range("Q1").value = BASE_START_ROW
        group_data = extract_group_parameters(group_index)
        if not group_data:
            print("  No more groups.")
            break

        print(f"Group {group_index + 1}: {group_data['identifier']}")
        paste_group_parameters(control_sheet, group_data)
        control_sheet.range("K1").value = batch_size 

        dataset_index = 0
        all_buys = []

        while dataset_index < total_datasets: 
            result_row = BASE_START_ROW
            for batch_num in range(batches_per_cycle): 
                if dataset_index >= total_datasets:
                    break
                offset = dataset_index * rows_per_dataset
                limit_datasets = min(batch_size, total_datasets - dataset_index)
                if limit_datasets != batch_size:
                    control_sheet.range("K1").value = limit_datasets
                limit = limit_datasets * rows_per_dataset
                print(f"  Exporting batch: dataset_index={dataset_index}, offset={offset}, datasets={limit_datasets}")
                cur.execute(f"SELECT * FROM DICTIONARY_TABLE LIMIT {limit} OFFSET {offset}")
                rows = cur.fetchall()
                if not rows:
                    break

                df = pd.DataFrame(rows)
                control_sheet.range("A:I").clear_contents()
                control_sheet.range("A1").value = df.values

                control_sheet.range("Q1").value = result_row
                wb.save()

                original_o1 = control_sheet.range("O1").value
                control_sheet.range("S1").value = 100

                start_time = time.time()
                while True:
                    if control_sheet.range("O1").value != original_o1:
                        break
                    if time.time() - start_time > 180:
                        raise TimeoutError("Macro timed out.")
                    time.sleep(1)

                control_sheet.range("O1").value = 0
                wb.save()

                result_row += batch_size
                control_sheet.range("Q1").value = result_row
                dataset_index += limit_datasets


            # reset q1
            print(f"[  Reset] Hit AG3 limit. Clearing and resetting.")

            algo_sheett_names = [
                "1 MACD", "2 EMA", "3 RSI", "4 Breakout", "5 ADX",
                "6 Volatility Measure", "7 SMA", "8 Bollinger Bands",
                "9 EMA & MACD", "10 RSI Bollinger"
            ]

            buys = collect_group_buys(wb, group_data["algorithms"])
            all_buys.extend(buys)

            trigger_macro(control_sheet, "AE1", delay=2)
            control_sheet.range("Q1").value = BASE_START_ROW
            wb.save()

        if all_buys:
            trigger_macro(control_sheet, RESET_MACRO_CELL, delay=15)
            wb2_sheet.range("M3").value = 100
            time.sleep(2)

            paste_row = 6
            paste_col = 182
            wb2_sheet.range((paste_row, paste_col)).value = group_data["identifier"]
            wb2_sheet.range((paste_row, paste_col + 1)).value = group_data["pp"]
            wb2_sheet.range((paste_row, paste_col + 2)).value = group_data["score"]
            paste_row += 2
            wb2_sheet.range((paste_row, paste_col)).value = ["Ticker", "Name", "Decision"]
            paste_row += 1
            for tck, name, dec in all_buys:
                wb2_sheet.range((paste_row, paste_col)).value = [tck, name, dec]
                paste_row += 1
            wb2.save()
            print("  Group buys pasted into WB2.")
            break  

        else:
            print("  No buys found for group.")

        group_index += 1

# mp2 p3
import re
import time
import json
import math
import sqlite3
import pandas as pd
import xlwings as xw
from datetime import datetime, timezone

EXCEL_FILE = r"C:/REDACTED/WB1.xlsm"
WB2_FILE   = r"C:/REDACTED/WB2.xlsm"
DB_FILE    = r"C:/REDACTED/stocks_data.db"

CONTROL_SHEET_NAME = "ControlSheet"
TRIGGER_MACRO_CELL = "AA1" 
RESET_MACRO_CELL   = "AC1"   
CYCLER_MACRO_CELL = "S1"  

PARAM_START_COL = 56   
BASE_START_ROW  = 55    
ROWS_PER_DATASET = 53  

MAX_WAIT_PER_BATCH_SEC = 900  

START_DATA_ROW = 55
COL_I = 9   
COL_J = 10 

VAR_BLOCK_START = {
    1: 11, 
    2: 17,
    3: 23,  
    4: 29, 
    5: 35, 
    6: 41,  
    7: 47,  
    8: 53, 
}

PRICE_START_COL = 59 
PRICE_WIDTH     = 6  

SHEET_MAP = {
    1: "1 MACD", 2: "2 EMA", 3: "3 RSI", 4: "4 Breakout", 5: "5 ADX",
    6: "6 Volatility Measure", 7: "7 SMA", 8: "8 Bollinger Bands",
    9: "9 EMA & MACD", 10: "10 RSI Bollinger"
}

def ensure_past_buys_table():
    conn = sqlite3.connect(DB_FILE)
    cur  = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS past_buys (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT,
            group_identifier TEXT,
            group_size INTEGER,
            old_pp REAL,
            old_score REAL,
            old_ticker_count INTEGER,
            new_ticker_count INTEGER,
            new_pp REAL,
            algo_slots_json TEXT,
            stocks_json TEXT
        )
    """)
    cur.execute("PRAGMA table_info(past_buys)")
    cols = {row[1] for row in cur.fetchall()}
    if "group_parameters_json" not in cols:
        cur.execute("ALTER TABLE past_buys ADD COLUMN group_parameters_json TEXT")
    conn.commit()
    conn.close()


def _is_profit_from_block(block_vals):
    for v in block_vals:
        if isinstance(v, (int, float)):
            if v > 0:
                return True
        else:
            s = str(v).strip().upper()
            if "PROFIT" in s:
                return True
    return False
MANUAL_BUY_ENABLE  = False       
MANUAL_BUY_GROUPS  = [3,4]          
MANUAL_BUY_ALGOS   = "ALL"           
MANUAL_BUY_ROWS    = [55]        

def _group_is_targeted(gi_1based:int) -> bool:
    if not MANUAL_BUY_ENABLE:
        return False
    if MANUAL_BUY_GROUPS == "ALL":
        return True
    return gi_1based in (MANUAL_BUY_GROUPS or [])

def _algos_for_target(g_algos:list[int]) -> list[int]:
    if MANUAL_BUY_ALGOS == "ALL":
        return list(dict.fromkeys(g_algos)) 
    return MANUAL_BUY_ALGOS or []

def _first_col_for_slot(slot:int) -> int:
    return 11 + 6*(slot-1)

def _sheet_for_algo(algo_num:int):
    return SHEET_MAP.get(algo_num) or SHEET_MAP.get(str(algo_num))
def _sheet_name_for(algo_num:int):
    return _sheet_for_algo(algo_num)

def insert_manual_buy_for_group_algo(wb, group_index_1based:int, groups, group_slot_maps, algo_num:int, rows:list[int]):
    g   = groups[group_index_1based-1]
    sm  = group_slot_maps[group_index_1based-1]
    if algo_num not in sm:
        print(f"   Manual BUY skipped: Algo {algo_num} not used in Group {group_index_1based} ({g['identifier']})")
        return
    slot      = sm[algo_num]
    start_col = _first_col_for_slot(slot)
    sheet_nm  = _sheet_for_algo(algo_num)
    if not sheet_nm:
        print(f"   Manual BUY skipped: sheet not found for Algo {algo_num}")
        return
    sh = wb.sheets[sheet_nm]
    addr_col = sh.range((1, start_col)).get_address(False, False).split('$')[0]
    print(f" Manual BUY → Group {group_index_1based} | Algo {algo_num} | Slot {slot} | {sheet_nm} {addr_col}{rows}")
    for r in rows:
        sh.range((r, start_col)).value = "BUY"

def count_old_tickers_on_sheet(wb, algo_num_for_reference: int) -> int:
    sheet_name = _sheet_name_for(algo_num_for_reference)
    if not sheet_name:
        return 0
    sheet = wb.sheets[sheet_name]
    row = 55
    count = 0
    while row <= 1056:
        t = sheet.range((row, 10)).value  
        if not t:
            break
        count += 1
        row += 1
    return count

def fetch_extra_big_groups(limit=4):
    conn = sqlite3.connect(DB_FILE)
    cur  = conn.cursor()
    cur.execute("SELECT identifier, pp, score, parameters FROM extra_big_groups ORDER BY ROWID ASC LIMIT ?", (limit,))
    rows = cur.fetchall()
    conn.close()
    groups = []
    for ident, pp, score, params_json in rows:
        param_map = json.loads(params_json) if params_json else {}
        algos = []
        params = []
        for key, plist in param_map.items():
            m = re.match(r"No\.(\d+)", str(key))
            if not m:
                continue
            algos.append(int(m.group(1)))
            params.append(plist)
        groups.append({
            "identifier": ident,
            "pp": float(pp) if pp is not None else None,
            "score": float(score) if score is not None else None,
            "algorithms": algos,
            "parameters": params,
            "size": len(set(algos))
        })
    return groups

def get_algo_base_row(algo_num):
    return 4 + (int(algo_num) - 1) * 11

def paste_group_parameters_for_slot(sheet, algo_num, slot, param_list):
    base_row = get_algo_base_row(algo_num) + (slot - 1) 
    for j, val in enumerate(param_list):
        sheet.range((base_row, PARAM_START_COL + j)).value = val

# macro trigger
def trigger_macro(sheet, cell, delay=1.0):
    sheet.range(cell).value = 100
    sheet.book.save()
    time.sleep(delay)
def _normalize_name_ticker(name, ticker):

    name_s = "" if name is None else str(name)
    tck_s  = "" if ticker is None else str(ticker)
    if (" " in tck_s) and (" " not in name_s) and (1 <= len(name_s) <= 6):
        return tck_s, name_s  
    return name_s, tck_s

def read_algo_buys_for_slot(wb, algo_num, slot):
    sheet_name = SHEET_MAP.get(algo_num) or SHEET_MAP.get(str(algo_num))
    if not sheet_name:
        print(f"   read_algo_buys_for_slot: unknown sheet for Algo {algo_num}")
        return {}, 0
    sheet = wb.sheets[sheet_name]

    start_col    = VAR_BLOCK_START[slot]
    block_width  = 6
    decision_col = start_col  

    buys  = {}  
    row   = START_DATA_ROW
    count = 0

    while True:
        name   = sheet.range((row, COL_I)).value
        ticker = sheet.range((row, COL_J)).value
        if not ticker:
            break
        decision = sheet.range((row, decision_col)).value
        if str(decision).strip().upper() == "BUY":
            block_vals = [sheet.range((row, start_col + k)).value for k in range(block_width)]
            price_vals = [sheet.range((row, PRICE_START_COL + k)).value for k in range(PRICE_WIDTH)]
            buys[ticker] = {"name": name, "block": block_vals, "price": price_vals}
            count += 1
        row += 1

    return buys, count

# group testing executor
def run_mp2_part3_group_executor(paper_mode=True):
    print(f"\n==== MP3 START (paper_mode={paper_mode}) ====\n")

    wb  = xw.Book(EXCEL_FILE)
    wb2 = xw.Book(WB2_FILE)
    control_sheet = wb.sheets[CONTROL_SHEET_NAME]
    wb2_sheet     = wb2.sheets["ControlSheet"]

    batch_size      = int(control_sheet.range("AG4").value) 
    result_limit    = int(control_sheet.range("AG3").value)  
    batches_per_cycle = max(1, math.ceil(result_limit / batch_size))

    conn = sqlite3.connect(DB_FILE)
    cur  = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM DICTIONARY_TABLE")
    total_rows = cur.fetchone()[0]
    conn.close()
    total_datasets = total_rows // ROWS_PER_DATASET

    groups = fetch_extra_big_groups(limit=4) if paper_mode else fetch_extra_big_groups(limit=1)

    if not groups:
        print("  No groups found in extra_big_groups.")
        return

    algo_slot_counter = {} 
    group_slot_maps   = []   

    for gi, g in enumerate(groups, start=1):
        slot_map = {}
        for algo_num in g["algorithms"]:
            nxt = algo_slot_counter.get(algo_num, 1)
            if nxt > 8:
                raise RuntimeError(f"Algo {algo_num} needs more than 8 slots.")
            slot_map[algo_num] = nxt
            algo_slot_counter[algo_num] = nxt + 1
        group_slot_maps.append(slot_map)

    print(" Pasting parameters into ControlSheet by algo/slot…")
    for g, slot_map in zip(groups, group_slot_maps):
        for algo_num, plist in zip(g["algorithms"], g["parameters"]):
            slot = slot_map[algo_num]
            paste_group_parameters_for_slot(control_sheet, algo_num, slot, plist)
            print(f"  • Algo {algo_num} → Slot {slot} (row {get_algo_base_row(algo_num)+(slot-1)}) pasted {len(plist)} params")

    wb.save()

    agg_per_group = []
    for g in groups:
        per_algo = {}
        for algo_num in g["algorithms"]:
            per_algo[algo_num] = {}   
        agg_per_group.append(per_algo)

    control_sheet.range("Q1").value = BASE_START_ROW
    dataset_index = 0
    print("\nTriggering compute cycle…")
    window_idx = 1
    while dataset_index < total_datasets:
        remaining_total = total_datasets - dataset_index
        this_window_target = min(result_limit, remaining_total)  
        batches_this_window = math.ceil(this_window_target / batch_size)

        print(f"\n Window {window_idx}: target_rows={this_window_target} "
              f"(AG3={result_limit}), batch_size={batch_size}, batches={batches_this_window}")

        result_row = BASE_START_ROW
        for batch_idx in range(batches_this_window):
            if dataset_index >= total_datasets:
                print(f" No more datasets. Stopping batches in Window {window_idx}.")
                break

            remaining_total = total_datasets - dataset_index
            remaining_in_window = this_window_target - (batch_idx * batch_size)
            limit_datasets = min(batch_size, remaining_total, remaining_in_window)
            offset = dataset_index * ROWS_PER_DATASET
            limit  = limit_datasets * ROWS_PER_DATASET

            print(f"    Window {window_idx} | Batch {batch_idx+1}/{batches_this_window} "
                  f"| dataset_index={dataset_index} | offset={offset} | "
                  f"limit_datasets={limit_datasets} | result_row_start={result_row}")

            conn = sqlite3.connect(DB_FILE)
            cur  = conn.cursor()
            cur.execute(f"SELECT * FROM DICTIONARY_TABLE LIMIT {limit} OFFSET {offset}")
            rows = cur.fetchall()
            conn.close()
            if not rows:
                print("    No rows returned from DB; breaking window.")
                break

            df = pd.DataFrame(rows)
            control_sheet.range("A:I").clear_contents()
            control_sheet.range("A1").value = df.values

            control_sheet.range("Q1").value = result_row
            control_sheet.range("K1").value = limit_datasets
            wb.save()

            original_o1 = control_sheet.range("O1").value
            control_sheet.range(CYCLER_MACRO_CELL).value = 100

            timeout_sec = MAX_WAIT_PER_BATCH_SEC
            start = time.time()
            while True:
                if control_sheet.range("O1").value != original_o1:
                    break
                if time.time() - start > timeout_sec:
                    raise TimeoutError(
                        f"Macro timed out after {timeout_sec}s in Window {window_idx} "
                        f"Batch {batch_idx+1} (limit_datasets={limit_datasets})."
                    )
                time.sleep(0.5)

            control_sheet.range("O1").value = 0
            wb.save()

            result_row    += limit_datasets
            dataset_index += limit_datasets
        
        for gi, g in enumerate(groups, start=1):
            if _group_is_targeted(gi):
                for a in _algos_for_target(g["algorithms"]):
                    insert_manual_buy_for_group_algo(
                        wb,
                        group_index_1based=gi,
                        groups=groups,
                        group_slot_maps=group_slot_maps,
                        algo_num=a,
                        rows=MANUAL_BUY_ROWS
                    )

        window_buy_counts = []
        for gi, (g, slot_map) in enumerate(zip(groups, group_slot_maps), start=1):
            total_for_group = 0
            for algo_num in g["algorithms"]:
                slot = slot_map[algo_num]
                buys_dict, _ = read_algo_buys_for_slot(wb, algo_num, slot)
                agg = agg_per_group[gi-1][algo_num]
                for tck, rec in buys_dict.items():
                    agg[tck] = rec
                total_for_group += len(buys_dict)
            window_buy_counts.append(total_for_group)
        print(f"  Window {window_idx} aggregated buys per group: {window_buy_counts}")

        print(f"  Clear & reset after Window {window_idx}.")
        control_sheet.range("AE1").value = 100  
        control_sheet.range("Q1").value = BASE_START_ROW
        wb.save()
        window_idx += 1
      
    ensure_past_buys_table()
    conn = sqlite3.connect(DB_FILE)
    cur  = conn.cursor()

    paste_row = 6  
    paste_col = 182

    wb2_sheet.range("M3").value = 100
    for gi, (g, slot_map) in enumerate(zip(groups, group_slot_maps), start=1):
        print(f"\n==============================")
        print(f"Group {gi}: {g['identifier']}  (size={g['size']})")
        print(f"Old PP={g['pp']}, Old Score={g['score']}")
        print(f"Algo→Slot map: {slot_map}")

        old_ticker_count = None
        print("  Old ticker count (historical): None")

        per_algo_buys = agg_per_group[gi-1]
        per_algo_counts = {algo_num: len(per_algo_buys.get(algo_num, {})) for algo_num in g["algorithms"]}

        def colname(c):
            name = ""
            x = c
            while x:
                x, r = divmod(x-1, 26)
                name = chr(65+r) + name
            return name

        for algo_num in g["algorithms"]:
            start_col = VAR_BLOCK_START[slot_map[algo_num]]
            end_col   = start_col + 5
            print(f"  • Algo {algo_num} (slot {slot_map[algo_num]}, block {colname(start_col)}:{colname(end_col)}) BUYs: {per_algo_counts[algo_num]}")

        # intersection of tickers
        common = None
        for algo_num in g["algorithms"]:
            tset = set(per_algo_buys[algo_num].keys())
            common = tset if common is None else (common & tset)
        common = sorted(list(common)) if common else []
        print(f"   Intersected tickers ({len(common)}): {', '.join(common) if common else '(none)'}")

        wb2_sheet.range((paste_row, paste_col)).value = ["Group", "Size", "Old_PP", "Old_Score", "Old_Ticker_Count", "New_Ticker_Count", "New_PP"]
        wb2_sheet.range((paste_row+1, paste_col)).value = [g["identifier"], g["size"], g["pp"], g["score"], old_ticker_count, len(common), None]

        _summary_row = paste_row + 1
        paste_row += 3

        headers = ["Ticker","Name","B1","B2","B3","B4","B5","B6","Date","Open","High","Low","Close","Volume"]
        wb2_sheet.range((paste_row, paste_col)).value = headers
        paste_row += 1

        ref_algo = g["algorithms"][0] if g["algorithms"] else None
        stocks_payload = []     
        profit_hits = 0
        new_count   = 0

        for tck in common:
            ref_rec = per_algo_buys[ref_algo][tck]
            norm_name, norm_ticker = _normalize_name_ticker(ref_rec["name"], tck)

            block6  = ref_rec["block"][:6]
            price6  = ref_rec["price"][:6]

            row_vals = [norm_ticker, norm_name] + block6 + price6
            wb2_sheet.range((paste_row, paste_col)).value = row_vals
            paste_row += 1
            new_count += 1

            # profit check with new profit rate
            if _is_profit_from_block(block6):
                profit_hits += 1

            stocks_payload.append({
                "name":   norm_name,
                "ticker": norm_ticker,
                "b1": block6[0],
                "b2": block6[1],
                "b3": block6[2],
                "b4": block6[3],
                "b5": block6[4],
                "b6": block6[5],
                "date":   str(price6[0]) if price6[0] is not None else None,
                "open":   price6[1],
                "high":   price6[2],
                "low":    price6[3],
                "close":  price6[4],
                "volume": price6[5]
            })

        new_pp = round(100.0 * profit_hits / new_count, 2) if new_count > 0 else None
        print(f"Computed new_pp for Group {gi}: {new_pp}% (profits={profit_hits}, total={new_count})")

        wb2_sheet.range((_summary_row, paste_col + 6)).value = new_pp

        group_params_json = json.dumps({f"No.{a}": p for a, p in zip(g["algorithms"], g["parameters"])})
        ts_str = datetime.now(timezone.utc).date().isoformat()

        cur.execute("""
            INSERT INTO past_buys
            (ts, group_identifier, group_size, old_pp, old_score, old_ticker_count, new_ticker_count, new_pp,
             algo_slots_json, stocks_json, group_parameters_json)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            ts_str,
            g["identifier"],
            g["size"],
            g["pp"],
            g["score"],
            old_ticker_count,             
            new_count,
            new_pp,
            json.dumps({str(k): v for k, v in slot_map.items()}),
            json.dumps(stocks_payload),
            group_params_json
        ))
        conn.commit()

        paste_row += 2
        print(f"  Group {gi} pasted to WB2 with {new_count} ticker rows. Logged 1 row to past_buys (full payload).")

    wb2.save()
    print("\n==== MP3 DONE ====\n")

# paper trading helper
import sqlite3

DB_FILE = r"C:/REDACTED/stocks_data.db"

def renumber_past_buys_safe():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    cur.execute("SELECT id FROM past_buys ORDER BY id ASC")
    ids = [row[0] for row in cur.fetchall()]

    for new_id, old_id in enumerate(ids, start=1):
        cur.execute("UPDATE past_buys SET id = ? WHERE id = ?", (new_id, old_id))

    cur.execute("DELETE FROM sqlite_sequence WHERE name='past_buys'")
    cur.execute("INSERT INTO sqlite_sequence (name, seq) VALUES ('past_buys', 20)")

    conn.commit()
    conn.close()
    print(f"  Renumbered {len(ids)} rows (1..{len(ids)}). Next id will be 21.")


# historical buys printer
def reprint_past_buys_compiled_to_excel():
    wb2 = xw.Book(WB2_FILE)
    wb2_sheet = wb2.sheets["ControlSheet"]

    wb2_sheet.range("M3").value = 100

    paste_row = 6
    paste_col = 182

    conn = sqlite3.connect(DB_FILE)
    cur  = conn.cursor()
    cur.execute("""
        SELECT id, ts, group_identifier, group_size,
               old_pp, old_score, old_ticker_count, new_ticker_count, new_pp,
               stocks_json
        FROM past_buys
        ORDER BY ts ASC, id ASC
    """)
    rows = cur.fetchall()
    conn.close()

    if not rows:
        print("  past_buys is empty; nothing to print.")
        return
      
    groups_map = {}
    group_order = []

    for id_pk, ts, gid, gsize, old_pp, old_score, old_count, new_count, new_pp, stocks_json in rows:
        if gid not in groups_map:
            groups_map[gid] = {
                'size': gsize,
                'old_pp': old_pp,
                'old_score': old_score,
                'old_ticker_count': old_count,
                'latest_ts': ts,
                'stocks_by_ticker': {}
            }
            group_order.append(gid)
        gentry = groups_map[gid]

        if ts and (gentry['latest_ts'] is None or str(ts) > str(gentry['latest_ts'])):
            gentry['latest_ts'] = ts
            if old_pp is not None:
                gentry['old_pp'] = old_pp
            if old_score is not None:
                gentry['old_score'] = old_score
            gentry['old_ticker_count'] = old_count

        try:
            slist = json.loads(stocks_json) if stocks_json else []
        except Exception:
            slist = []

        for rec in slist:
            tck = rec.get("ticker")
            if not tck:
                continue
            prev = gentry['stocks_by_ticker'].get(tck)
            if (prev is None) or (str(ts) >= str(prev[0] or "")):
                gentry['stocks_by_ticker'][tck] = (ts, rec)

    # print each group
    for gid in group_order:
        info = groups_map[gid]

        items = sorted(info['stocks_by_ticker'].items(), key=lambda kv: kv[0])
        compiled_count = len(items)


        profit_hits = 0
        for _ts, rec in (v for _, v in items):
            block6 = [rec.get("b1"), rec.get("b2"), rec.get("b3"),
                      rec.get("b4"), rec.get("b5"), rec.get("b6")]
            if _is_profit_from_block(block6):
                profit_hits += 1
        new_pp_compiled = round(100.0 * profit_hits / compiled_count, 2) if compiled_count > 0 else None


        wb2_sheet.range((paste_row, paste_col)).value = [
            "Group", "Size", "Old_PP", "Old_Score", "Old_Ticker_Count", "New_Ticker_Count", "New_PP"
        ]
        wb2_sheet.range((paste_row + 1, paste_col)).value = [
            gid,
            info['size'],
            info['old_pp'],
            info['old_score'],
            info['old_ticker_count'],
            compiled_count,
            None  
        ]
        _summary_row = paste_row + 1
        paste_row += 3

        headers = ["Ticker","Name","B1","B2","B3","B4","B5","B6","Date","Open","High","Low","Close","Volume"]
        wb2_sheet.range((paste_row, paste_col)).value = headers
        paste_row += 1

        for tck, (ts, rec) in items:
            name = rec.get("name")
            ticker = rec.get("ticker")
            if name is not None and ticker is not None:
                name, ticker = _normalize_name_ticker(name, ticker)

            block6 = [rec.get("b1"), rec.get("b2"), rec.get("b3"),
                      rec.get("b4"), rec.get("b5"), rec.get("b6")]
            price6 = [rec.get("date"), rec.get("open"), rec.get("high"),
                      rec.get("low"), rec.get("close"), rec.get("volume")]

            row_vals = [ticker, name] + block6 + price6
            wb2_sheet.range((paste_row, paste_col)).value = row_vals
            paste_row += 1

        wb2_sheet.range((_summary_row, paste_col + 6)).value = new_pp_compiled

        paste_row += 2

    wb2.save()
    print("  Reprinted past_buys to WB2 (compiled by group_identifier).")


# stats of algos ran
import sqlite3
import math

DB_PATH = r"C:/REDACTED/optuna_10A_data.db"
ALGO_NAMES = [
    "MACD", "EMA", "RSI", "Breakout", "ADX",
    "SMA", "Bollinger_Bands", "EMA_MACD_Combo", "RSI_Bollinger", "Volatility"
]

def analyze_pp_vs_tickers():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    thresholds = list(range(20, 105, 5))
    header_row = ["Threshold"] + ALGO_NAMES

    print("=" * 100)
    print("Count of Entries Per Algorithm by PP Threshold (>= Threshold)")
    print("=" * 100)
    print(" | ".join(f"{h:^12}" for h in header_row))
    print("-" * 100)

    threshold_counts = {}

    for threshold in thresholds:
        row = [f"{threshold}%"]
        threshold_counts[threshold] = {} 
        for table in ALGO_NAMES:
            count = 0
            try:
                cursor.execute(f"SELECT profit_percentage FROM {table}")
                for (pp,) in cursor.fetchall():
                    if pp is not None and float(pp) >= threshold:
                        count += 1
            except:
                pass
            row.append(str(count))
            threshold_counts[threshold][table] = count
        print(" | ".join(f"{cell:^12}" for cell in row))

    print("\n" + "=" * 100)
    print("Average Tickers Per Entry by PP Threshold (>= Threshold)")
    print("=" * 100)
    print(" | ".join(f"{h:^12}" for h in header_row))
    print("-" * 100)

    for threshold in thresholds:
        row = [f"{threshold}%"]
        for table in ALGO_NAMES:
            total_tickers = 0
            count = 0
            try:
                cursor.execute(f"SELECT profit_percentage, stocks_bought FROM {table}")
                for pp, stock_json in cursor.fetchall():
                    if pp is not None and float(pp) >= threshold and stock_json:
                        total_tickers += stock_json.count('"ticker"')
                        count += 1
            except:
                pass
            avg = round(total_tickers / count, 2) if count else 0.0
            row.append(f"{avg:.2f}")
        print(" | ".join(f"{cell:^12}" for cell in row))

    print("\n" + "=" * 100)
    print("Average Score Per Algorithm by PP Threshold (>= Threshold)")
    print("=" * 100)
    print(" | ".join(f"{h:^12}" for h in header_row))
    print("-" * 100)

    for threshold in thresholds:
        row = [f"{threshold}%"]
        for table in ALGO_NAMES:
            total_score = 0
            count = 0
            try:
                cursor.execute(f"SELECT profit_percentage, stocks_bought FROM {table}")
                for pp, stock_json in cursor.fetchall():
                    if pp is not None and float(pp) >= threshold and stock_json:
                        ticker_count = stock_json.count('"ticker"')
                        score = float(pp) * math.log(ticker_count + 1)
                        total_score += score
                        count += 1
            except:
                pass
            avg_score = round(total_score / count, 2) if count else 0.0
            row.append(f"{avg_score:.2f}")
        print(" | ".join(f"{cell:^12}" for cell in row))

    conn.close()


# intial seeder for former best groups
import json
import sqlite3
import xlwings as xw

EXCEL_FILE = r"C:/REDACTED/WB1.xlsm"
STOCKS_DB  = r"C:/REDACTED/stocks_data.db"

CONTROL_SHEET_NAME = "ControlSheet"
PARAM_START_COL = 56  
MAX_PARAM_COLS  = 20  
VARIATIONS_PER_ALGO = 8

ALGO_DISPLAY_NAMES = [
    "MACD", "EMA", "RSI", "Breakout", "ADX", "Volatility",
    "SMA", "Bollinger_Bands", "EMA_MACD_Combo", "RSI_Bollinger"
]

def _base_row_for_algo_1based(algo_num: int) -> int:
    return 4 + (algo_num - 1) * 11

def seed_top8_from_pruned():
    wb = xw.Book(EXCEL_FILE)
    sheet = wb.sheets[CONTROL_SHEET_NAME]

    conn = sqlite3.connect(STOCKS_DB)
    cur  = conn.cursor()

    for algo_index_1based, algo_name in enumerate(ALGO_DISPLAY_NAMES, start=1):
        pruned_table = f"pruned_{algo_name}"
        try:
            cur.execute(
                f"""
                SELECT params_json, param_headers_json
                FROM {pruned_table}
                ORDER BY universal_score DESC, pp DESC
                LIMIT ?
                """,
                (VARIATIONS_PER_ALGO,)
            )
            rows = cur.fetchall()
        except sqlite3.OperationalError:
            print(f"[SKIP] Table {pruned_table} not found.")
            continue

        if not rows:
            print(f"[SKIP] Table {pruned_table} is empty.")
            continue

        parsed_rows = []
        max_params_in_table = 0
        for params_json, headers_json in rows:
            try:
                params = json.loads(params_json) if params_json else []
            except Exception:
                params = []
            try:
                headers = json.loads(headers_json) if headers_json else []
            except Exception:
                headers = []
            max_params_in_table = max(max_params_in_table, len(headers) if headers else len(params))
            parsed_rows.append(params)

        clear_cols = min(MAX_PARAM_COLS, max(1, max_params_in_table or 0))
        base_row = _base_row_for_algo_1based(algo_index_1based)

        for i in range(VARIATIONS_PER_ALGO):
            r = base_row + i
            sheet.range((r, PARAM_START_COL)).resize(1, clear_cols).value = None

        for i, params in enumerate(parsed_rows[:VARIATIONS_PER_ALGO]):
            if not params:
                continue
            r = base_row + i
            clean = []
            for v in params:
                if v is None:
                    clean.append("")
                elif isinstance(v, (int, float)):
                    clean.append(v)
                else:
                    clean.append(str(v))
            sheet.range((r, PARAM_START_COL)).value = [clean] 

        print(f"[OK] Seeded {min(len(parsed_rows), VARIATIONS_PER_ALGO)} variations "
              f"for {algo_name} → rows {base_row}-{base_row+VARIATIONS_PER_ALGO-1}")

    wb.save()
    conn.close()
    print("  MP0 Seeder finished — ControlSheet primed from last run.")

# MAIN FUNCTIONS LIST

# MP0 Seeder
# seed_top8_from_pruned()

# MP1 P1
# run_cycling_program()

# MP1 P2
# run_optuna_optimization()

# MP1 P2 helper functions
# save_excel_params_to_temp_table()
# overwrite_trial_params_with_temp_values()
# wipe_optuna_db()
# save_all_variations_to_db()

# MP1 P3
# export_variations_to_excel()

# MP2 P2
# run_mp2_general_group_tester()

# MP2 P3
# run_mp2_part3_group_executor()

# Data provider
# load_latest_csvs_into_dictionary()
# populate_dataset_one_from_dictionary()

# Stats of Algos
# analyze_pp_vs_tickers()

# Reorder of past_buys
# renumber_past_buys_safe()

# MAIN FUNCTIONS LIST

import time

def print_elapsed(label, seconds):
    mins, secs = divmod(int(seconds), 60)
    hours, mins = divmod(mins, 60)
    print(f"[ ] {label} took {hours}h {mins}m {secs}s")

def full_cycle_main():
    NUM_LOOPS = 100
    overall_start = time.perf_counter()

    total_part1_time = 0.0
    total_part2_time = 0.0

    for i in range(NUM_LOOPS):
        print(f"\n[LOOP {i+1}/{NUM_LOOPS}] Starting MP1 - Part 1 (Cycling)...")
        t1 = time.perf_counter()
        try:
            run_cycling_program()
        except Exception as e:
            print(f"[WARN] MP1-Part1 failed on loop {i+1}: {e}")
        part1_duration = time.perf_counter() - t1
        total_part1_time += part1_duration
        print_elapsed("MP1 - Part 1", part1_duration)

        print(f"[LOOP {i+1}/{NUM_LOOPS}] Starting MP1 - Part 2 (Optuna Optimization)...")
        t2 = time.perf_counter()
        try:
            run_optuna_optimization()
        except Exception as e:
            print(f"[WARN] MP1-Part2 failed on loop {i+1}: {e}")
        part2_duration = time.perf_counter() - t2
        total_part2_time += part2_duration
        print_elapsed("MP1 - Part 2", part2_duration)

    print("\n  [SUMMARY] loop section complete.")
    print_elapsed("TOTAL TIME for loop section", total_part1_time + total_part2_time)

    print("\n[MP1] Running Part 3: Exporting variations to WB2...")
    t3 = time.perf_counter()
    try:
        export_variations_to_excel()
    except Exception as e:
        print(f"[WARN] MP1-Part3 failed: {e}")
    part3_duration = time.perf_counter() - t3
    print_elapsed("MP1 - Part 3", part3_duration)

    print("\n[MP2] Running Part 2: Testing all group combinations...")
    t4 = time.perf_counter()
    try:
        run_mp2_general_group_tester()
    except Exception as e:
        print(f"[WARN] MP2-Part2 failed: {e}")
    part4_duration = time.perf_counter() - t4
    print_elapsed("MP2 - Part 2", part4_duration)

    print("\n [FINAL TIMING SUMMARY]")
    print(f"[AVG] MP1 - Part 1 (Cycling): {round(total_part1_time / max(1, NUM_LOOPS), 2)} sec avg")
    print(f"[AVG] MP1 - Part 2 (Optuna): {round(total_part2_time / max(1, NUM_LOOPS), 2)} sec avg")
    print_elapsed("TOTAL TIME for MP1 - Part 3", part3_duration)
    print_elapsed("TOTAL TIME for MP2 - Part 2", part4_duration)
    print_elapsed("OVERALL TIME for entire program", time.perf_counter() - overall_start)

def setup_dictionary_data():
    print("[DATA] Loading CSVs into dictionary...")
    t0 = time.perf_counter()
    load_latest_csvs_into_dictionary()
    print_elapsed("Dictionary Load", time.perf_counter() - t0)

    print("\n[MP2] Running Part 3: Executing and pasting best group result...")
    t1 = time.perf_counter()
    run_mp2_part3_group_executor()
    print_elapsed("MP2 - Part 3", time.perf_counter() - t1)

# Main Runner

if __name__ == "__main__":
    full_cycle_main()
    setup_dictionary_data()
