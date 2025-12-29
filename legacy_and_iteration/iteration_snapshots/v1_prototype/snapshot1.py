import time
import sqlite3
import pandas as pd
import xlwings as xw

import optuna

# Config
EXCEL_FILE = r"<REDACTED_PATH>/WB1.xlsm"
DB_FILE = r"<REDACTED_PATH>/dataset.db"
optuna_db_path = r"<REDACTED_PATH>/results.db"


# Excel Cell Locations
TOTAL_STOCKS_CELL = "AG3"  
STOCKS_PER_CYCLE_CELL = "AG4"  
START_ROW_CELL = "Q1" 
CYCLING_TRIGGER_CELL = "S1"  
CYCLING_COMPLETION_CELL = "O1"  
SHEET_NAME = "ControlSheet"

DATASET_ROWS = 53  
BASE_START_ROW = 55 

ALGO_NAMES = [
    "MACD", "EMA", "RSI", "Breakout", "ADX", "Volatility",
    "SMA", "Bollinger_Bands", "EMA_MACD_Combo", "RSI_Bollinger"
]

conn = sqlite3.connect(optuna_db_path, check_same_thread=False) 
cursor = conn.cursor()

for algo in ALGO_NAMES:
    cursor.execute(f"CREATE TABLE IF NOT EXISTS {algo} (id INTEGER PRIMARY KEY, params TEXT, profit_percentage REAL)")
conn.commit()
conn.close()

# Parameter search ranges are omitted for privacy
PARAM_RANGES: dict = {}

# mp1 p1 start
def run_cycling_program():
    try:
        wb = xw.Book(EXCEL_FILE)
        sheet = wb.sheets[SHEET_NAME]

        # helper cells
        total_stocks = int(sheet.range(TOTAL_STOCKS_CELL).value) 
        batch_size = int(sheet.range(STOCKS_PER_CYCLE_CELL).value)  

        num_cycles = total_stocks // batch_size
        if total_stocks % batch_size != 0:
            num_cycles += 1 

        print(f"[INFO] Running {num_cycles} cycles with {batch_size} datasets per cycle.")

        conn = sqlite3.connect(DB_FILE)

        for cycle in range(num_cycles):
            offset = cycle * batch_size
            total_rows = batch_size * DATASET_ROWS

            print(f"[INFO] Cycle {cycle + 1}/{num_cycles}: Fetching rows {offset} to {offset + total_rows}.")

            query = f'SELECT * FROM "All_Stocks (1)" LIMIT {total_rows} OFFSET {offset * DATASET_ROWS}'
            df = pd.read_sql_query(query, conn)
            sheet.range("A:H").clear_contents()
            sheet.range("A:I").clear_contents()

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

            # trigger excel cycling program
            initial_end_time = sheet.range(CYCLING_COMPLETION_CELL).value 
            sheet.range(CYCLING_TRIGGER_CELL).value = 100 
            wb.save()
          
            estimated_wait_time = (batch_size * 600) / 500  
            estimated_wait_time = max(10, estimated_wait_time) 

            print(f"[INFO] Estimated wait time: {int(estimated_wait_time)} seconds.")

            # wait for completion 
            while True:
                time.sleep(int(estimated_wait_time))  
                current_end_time = sheet.range(CYCLING_COMPLETION_CELL).value
                if current_end_time != initial_end_time:
                    break  
            print(f"[INFO] Cycle {cycle + 1} completed.")
        conn.close()

        wb.save()
        print("Excel Cycling Program Completed Successfully.")

    except Exception as e:
        print(f"[ERROR] Excel Automation or DB Extraction Failed: {e}")

# mp1 p1 end

# mp1 p2 start
def run_optuna_optimization():
    try:
        wb = xw.Book(EXCEL_FILE)
        sheet = wb.sheets[SHEET_NAME]
        row = 4

        for algo_name in ALGO_NAMES:
            used_sets = set()

            study = optuna.create_study(
                study_name=algo_name,
                direction="maximize",
                storage=f"sqlite:///{optuna_db_path}",
                load_if_exists=True
            )

            def objective(trial):
                pp_list = []

                for i in range(8):
                    if algo_name not in PARAM_RANGES or not PARAM_RANGES[algo_name]:
                        raise optuna.exceptions.TrialPruned()
                    ranges = PARAM_RANGES[algo_name]
                    params = []

                    for j, param in enumerate(ranges):
                        name = f"{algo_name}_v{i}_param{j+1}"
                        if isinstance(param, list):
                            val = trial.suggest_categorical(name, param)
                        else:
                            low, high, step = param
                            val = trial.suggest_float(name, low, high, step=step)
                        params.append(val)

                    if algo_name == "Volatility":
                        p1 = trial.suggest_float(f"{algo_name}_v{i}_pmerge1", 1.0, 5.0, step=0.5)
                        p2 = trial.suggest_float(f"{algo_name}_v{i}_pmerge2", 1.0, 5.0, step=0.5)
                        trend = trial.suggest_categorical(f"{algo_name}_v{i}_trend_range", ["low", "high"])
                        t_val = trial.suggest_float(f"{algo_name}_v{i}_trend_val", *(0, 1, 0.1) if trend == "low" else (1, 4, 0.2))
                        params += [f"{p1:.1f}-{p2:.1f}", trend, t_val]

                    if algo_name == "ADX":
                        adx_range = trial.suggest_categorical(f"{algo_name}_v{i}_adx_range", ["low", "high"])
                        adx_val = trial.suggest_float(
                            f"{algo_name}_v{i}_adx_val",
                            *(10, 30, 1) if adx_range == "low" else (60, 100, 1)
                        )
                        params += [adx_range, adx_val]

                    # duplicate checker
                    key = tuple(params)
                    if key in used_sets:
                        raise optuna.exceptions.TrialPruned()
                    used_sets.add(key)

                    # write to Excel
                    for j, val in enumerate(params):
                        sheet.range((row + i, 56 + j)).value = val

                    pp = sheet.range((row + i, 48)).value or 0.0
                    pp_list.append((params, pp))

                conn = sqlite3.connect(optuna_db_path)
                cur = conn.cursor()
                cur.execute(f"""
                    CREATE TABLE IF NOT EXISTS {algo_name} (
                        id INTEGER PRIMARY KEY,
                        params TEXT,
                        profit_percentage REAL
                    )
                """)
                for param_set, pp in pp_list:
                    cur.execute(
                        f"INSERT INTO {algo_name} (params, profit_percentage) VALUES (?, ?)",
                        (str(param_set), pp)
                    )
                conn.commit()
                conn.close()

                return max(pp for _, pp in pp_list)
            try:
                study.optimize(objective, n_trials=1)
            except optuna.exceptions.TrialPruned:
                print(f"[INFO] Duplicate trial skipped for {algo_name}.")
            except Exception as e:
                print(f"[ERROR] Algorithm {algo_name} failed: {e}")

            row += 11

        wb.save()
        print(" MP1 - Part 2 (Optuna Optimization) Completed.")

    except Exception as e:
        print(f"[ERROR] Optuna Optimization Failed: {e}")

# mp1 p1 end

# wipes tables
def wipe_optuna_db():
    conn = sqlite3.connect(optuna_db_path)
    cur = conn.cursor()
    for algo in ALGO_NAMES:
        cur.execute(f"DELETE FROM {algo}")

    optuna_tables = [
        "studies", "study_directions", "study_user_attributes", "trial_params",
        "trial_values", "trials", "trial_system_attributes", "study_system_attributes"
    ]
    for tbl in optuna_tables:
        cur.execute(f"DELETE FROM {tbl}")

    conn.commit()
    conn.close()
    print("[CLEANUP] All Optuna trials and logs deleted.")

# optuna clearer end

# main
if __name__ == "__main__":
    # run_cycling_program()
    run_optuna_optimization()
    # wipe_optuna_db()
    print("MP1 - Part 2 Completed.")
