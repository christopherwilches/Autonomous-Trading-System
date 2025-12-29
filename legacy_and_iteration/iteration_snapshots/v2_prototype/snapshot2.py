import time
import sqlite3
import pandas as pd
import xlwings as xw

import optuna
import threading
import pythoncom

import json

# file paths
EXCEL_FILE = r"<REDACTED_PATH>/WB1.xlsm"
DB_FILE = r"<REDACTED_PATH>/dataset.db"
optuna_db_path = r"<REDACTED_PATH>/results.db"
WB2_FILE = r"<REDACTED_PATH>/WB2.xlsm"

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

# Parameter search ranges omitted for privacy in this snapshot
PARAM_RANGES: dict = {}

# mp1 p1 start
def run_cycling_program():
    """Loops through all sets needed based on AG3, pasting into A1 and cycling through."""
    try:
        wb = xw.Book(EXCEL_FILE)
        sheet = wb.sheets[SHEET_NAME]

        # Read helper cells
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

            # Trigger cycling program
            initial_end_time = sheet.range(CYCLING_COMPLETION_CELL).value
            sheet.range(CYCLING_TRIGGER_CELL).value = 100  
            wb.save()

            estimated_wait_time = (batch_size * 600) / 500 
            estimated_wait_time = max(10, estimated_wait_time) 

            print(f"[INFO] Estimated wait time: {int(estimated_wait_time)} seconds.")

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
        save_all_variations_to_db()

        for algo_name in ALGO_NAMES:
            used_sets = set()

            study = optuna.create_study(
                study_name=algo_name,
                direction="maximize",
                storage=f"sqlite:///{optuna_db_path}",
                load_if_exists=True
            )

            def make_objective(start_row):
                def objective(trial):
                    pp_list = []

                    for i in range(8):
                        params = []

                        if algo_name == "Volatility":
                            val1 = trial.suggest_float(f"{algo_name}_v{i}_param1", 10, 19, step=1) 

                            # Merge two parameter values for the one unique parameter
                            pmerge1 = trial.suggest_float(f"{algo_name}_v{i}_pmerge1", 1.0, 3.0, step=0.1)
                            pmerge2 = trial.suggest_float(f"{algo_name}_v{i}_pmerge2", 3.0, 7.0, step=0.1)
                            merged_param = f"{min(pmerge1, pmerge2):.1f}-{max(pmerge1, pmerge2):.1f}"

                            val3 = trial.suggest_categorical(f"{algo_name}_v{i}_param3", ["Vol", "None"])
                            val4 = trial.suggest_float(f"{algo_name}_v{i}_param4", 1, 30, step=1)
                            val5 = trial.suggest_int(f"{algo_name}_v{i}_param5", 1, 10)
                            trend_range = trial.suggest_categorical(f"{algo_name}_v{i}_trend_range", ["low", "high"])
                            if trend_range == "low":
                                t_val = trial.suggest_float(f"{algo_name}_v{i}_trend_val", 0.0, 1.0, step=0.1)
                            else:
                                t_val = trial.suggest_float(f"{algo_name}_v{i}_trend_val", 1.0, 4.0, step=0.2)

                            params = [val1, merged_param, val3, val4, val5, round(t_val, 1)]

                        elif algo_name == "ADX":
                            val1 = trial.suggest_int(f"{algo_name}_v{i}_param1", 10, 14)

                            # Choose one of two ranges for threshold
                            adx_range = trial.suggest_categorical(f"{algo_name}_v{i}_range", ["low", "high"])
                            if adx_range == "low":
                                val2 = trial.suggest_float(f"{algo_name}_v{i}_param2", 10, 30, step=1)
                            else:
                                val2 = trial.suggest_float(f"{algo_name}_v{i}_param2", 60, 100, step=1)

                            val3 = trial.suggest_categorical(f"{algo_name}_v{i}_param4", ["Vol", "None"])
                            val4 = trial.suggest_int(f"{algo_name}_v{i}_param3", 1, 30)
                            val5 = trial.suggest_int(f"{algo_name}_v{i}_param5", 1, 10)

                            strategy = trial.suggest_categorical(f"{algo_name}_v{i}_param6", ["ADX Only", "ADX + DI+ > DI-"])

                            params = [val1, int(val2), val3, val4, val5, strategy]

                        else:
                            ranges = PARAM_RANGES[algo_name]
                            for j, param in enumerate(ranges):
                                name = f"{algo_name}_v{i}_param{j+1}"
                                if isinstance(param, list):
                                    val = trial.suggest_categorical(name, param)
                                else:
                                    low, high, step = param
                                    val = trial.suggest_float(name, low, high, step=step)
                                params.append(val)

                        # Write parameters to Excel
                        for j, val in enumerate(params):
                            sheet.range((start_row + i, 56 + j)).value = val
                        pp = sheet.range((start_row + i, 48)).value or 0.0
                        pp_list.append((params.copy(), pp))

                    wb.save()
                    return max(pp for _, pp in pp_list)
                return objective
            tested_rows = []

            for i in range(8):
                row_params = []
                for j in range(20): 
                    cell_val = sheet.range((row + i, 56 + j)).value
                    if cell_val is None:
                        break
                    row_params.append(cell_val)
                pp = sheet.range((row + i, 48)).value or 0.0
                tested_rows.append((row_params, pp))
              
            try:
                study.optimize(make_objective(row), n_trials=1)
            except optuna.exceptions.TrialPruned:
                print(f"[INFO] Duplicate trial skipped for {algo_name}.")
            except Exception as e:
                print(f"[ERROR] Algorithm {algo_name} failed: {e}")
            row += 11
        wb.save()
        print(" MP1 - Part 2 (Optuna Optimization) Completed.")

    except Exception as e:
        print(f"[ERROR] Optuna Optimization Failed: {e}")

# mp1 p2 end


def wipe_optuna_db():
    conn = sqlite3.connect(optuna_db_path)
    cur = conn.cursor()

    for algo in ALGO_NAMES:
        cur.execute(f"DELETE FROM {algo}")

    # Delete Optuna tracking tables too 
    optuna_tables = [
        "studies", "study_directions", "study_user_attributes", "trial_params",
        "trial_values", "trials", "trial_system_attributes", "study_system_attributes"
    ]
    for tbl in optuna_tables:
        cur.execute(f"DELETE FROM {tbl}")

    conn.commit()
    conn.close()
    print("[CLEANUP] All Optuna trials and logs deleted.")


# helper function to save variations from excel to db

def save_all_variations_to_db():
    wb = xw.Book(EXCEL_FILE)
    control_sheet = wb.sheets["ControlSheet"]


    start_data_row = 55
    ticker_col = 9  
    name_col = 10   
    result_col_starts = [11, 17, 23, 29, 35, 41, 47, 53] 


    conn = sqlite3.connect(optuna_db_path)
    cur = conn.cursor()

    for algo_index, algo_name in enumerate(ALGO_NAMES):
        sheet_name = SHEET_NAMES[algo_index]
        print(f"\n[INFO] Saving for {algo_name} using sheet '{sheet_name}'")
        try:
            algo_sheet = wb.sheets[sheet_name]
        except:
            print(f"[WARN] Sheet '{sheet_name}' not found. Skipping.")
            continue

        cur.execute(f"""
            CREATE TABLE IF NOT EXISTS {algo_name} (
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

        cur.execute(f"PRAGMA table_info({algo_name})")
        existing_cols = {col[1] for col in cur.fetchall()}

        for col, col_type in required_cols.items():
            if col not in existing_cols:
                cur.execute(f"ALTER TABLE {algo_name} ADD COLUMN {col} {col_type}")
                print(f"[INFO] Column '{col}' added to '{algo_name}' table.")

        cur.execute(f"SELECT COUNT(*) FROM {algo_name}")
        count = cur.fetchone()[0]
        base_trial = (count // 8) + 1

        control_start_row = 4 + (algo_index * 11)

        for var_num in range(1, 9):
            row_offset = var_num - 1
            variation_col = result_col_starts[row_offset]

            params = []
            for j in range(20):
                val = control_sheet.range((control_start_row + row_offset, 56 + j)).value
                if val is None:
                    break
                params.append(val)

            pp = control_sheet.range((control_start_row + row_offset, 48)).value or 0.0

            stocks = []
            row = start_data_row
            while True:
                ticker = algo_sheet.range((row, ticker_col)).value
                name = algo_sheet.range((row, name_col)).value
                if not ticker:
                    break

                decision = algo_sheet.range((row, variation_col)).value
                if str(decision).strip().upper() == "BUY":
                    stocks.append({
                        "ticker": ticker,
                        "name": name,
                        "decision": decision,
                        "result": algo_sheet.range((row, variation_col + 1)).value,
                        "profit": algo_sheet.range((row, variation_col + 2)).value,
                        "verbal_result": algo_sheet.range((row, variation_col + 3)).value,
                        "verbal_decision": algo_sheet.range((row, variation_col + 4)).value,
                        "symbol_selected": algo_sheet.range((row, variation_col + 5)).value
                    })

                row += 1

            cur.execute(f"""
                INSERT INTO {algo_name} (params, profit_percentage, trial_number, variation_number, stocks_bought)
                VALUES (?, ?, ?, ?, ?)
            """, (json.dumps(params), pp, base_trial, var_num, json.dumps(stocks)))

        conn.commit()

    conn.close()
    print("\n All variations saved to DB.")

# filter variations to top one

def export_variations_to_excel():
    wb1 = xw.Book(EXCEL_FILE)
    control_sheet = wb1.sheets["ControlSheet"]

    wb2 = xw.Book(WB2_FILE)
    output_sheet = wb2.sheets[0]

    conn = sqlite3.connect(optuna_db_path)
    cur = conn.cursor()

    start_cell_row = 6
    start_cell_col = 3
    param_start_col = 56  # Column BD
    row_buffer_between_variations = 2
    col_buffer_between_algorithms = 16  # 4+1+10+1

    for algo_index, algo_name in enumerate(ALGO_NAMES):
        cur.execute(f"SELECT * FROM {algo_name}")
        rows = cur.fetchall()

        if not rows:
            continue

        algo_col_start = start_cell_col + (algo_index * col_buffer_between_algorithms)
        param_col_start = algo_col_start + 5

        # Param headers at row 3 + 11*algo_index
        param_headers = []
        header_row = 3 + (algo_index * 11)
        for j in range(10):
            val = control_sheet.range((header_row, param_start_col + j)).value
            if val:
                param_headers.append(val)
            else:
                break

        row_cursor = start_cell_row
        variation_counter = 0

        for row_data in rows:
            _, param_json, pp, trial_num, var_num, stock_json = row_data
            if not pp or float(pp) == 0:
                continue  

            params = json.loads(param_json)
            stocks = json.loads(stock_json)

            # Calculate block height 
            block_height = 2 + 1 + 1 + len(stocks) + row_buffer_between_variations

            # 4 identifier cells
            variation_counter += 1
            uid = f"No.{algo_index+1}.{variation_counter}"

            output_sheet.range((row_cursor, algo_col_start)).value = uid
            output_sheet.range((row_cursor, algo_col_start + 1)).value = block_height
            output_sheet.range((row_cursor, algo_col_start + 2)).value = trial_num
            output_sheet.range((row_cursor, algo_col_start + 3)).value = var_num

            # Write param headers and values
            for j, header in enumerate(param_headers):
                output_sheet.range((row_cursor, param_col_start + j)).value = header
            for j, val in enumerate(params):
                output_sheet.range((row_cursor + 1, param_col_start + j)).value = val

            # Result headers
            result_headers = [
                "Ticker", "Name", "Final Final Decision", "Result of Buy or Not",
                "Total Profit or Loss from Trade", "Verbal Profit or Loss",
                "Verbal Buy or Not", "symbol selected"
            ]
            for j, header in enumerate(result_headers):
                output_sheet.range((row_cursor + 3, param_col_start + j)).value = header

            # Stock rows
            for i, stock in enumerate(stocks):
                result_row = row_cursor + 4 + i
                output_sheet.range((result_row, param_col_start)).value = stock.get("ticker", "")
                output_sheet.range((result_row, param_col_start + 1)).value = stock.get("name", "")
                output_sheet.range((result_row, param_col_start + 2)).value = stock.get("decision", "")
                try:
                    result_val = round(float(stock.get("result", 0)), 2)
                except:
                    result_val = stock.get("result", "")
                try:
                    profit_val = round(float(stock.get("profit", 0)), 2)
                except:
                    profit_val = stock.get("profit", "")
                output_sheet.range((result_row, param_col_start + 3)).value = result_val
                output_sheet.range((result_row, param_col_start + 4)).value = profit_val
                output_sheet.range((result_row, param_col_start + 5)).value = stock.get("verbal_result", "")
                output_sheet.range((result_row, param_col_start + 6)).value = stock.get("verbal_decision", "")
                output_sheet.range((result_row, param_col_start + 7)).value = stock.get("symbol_selected", "")

            row_cursor += block_height

    wb2.save()
    print("\nMP1 - Part 3 Completed. Variations exported to WB2.")

# mp2 p2 start

def run_mp2_general_group_tester():
    WB2_FILE = r"C:/testpy/Combination Grouper MME3.xlsm"
    wb = xw.Book(WB2_FILE)
    sheet = wb.sheets[0]

    identifier_cols = [3, 19, 35, 51, 67, 83, 99, 115, 131, 147]  # C, S, AI, AY, BO, CE, CU, DK, EA, EQ
    start_row = 6
    param_offset = 5
    result_offset = 4
    paste_col = 166  # FK
    paste_row = 6
    threshold_pp = 90.0

    variation_map = {}
    param_headers_map = {}
    full_stocks_map = {}

    for algo_idx, col in enumerate(identifier_cols):
        row = start_row
        variation_map[algo_idx] = []
        blanks = 0

        # Extract parameter headers
        header_row = row
        header_col = col + param_offset
        headers = []
        while True:
            h = sheet.range((header_row, header_col)).value
            if h:
                headers.append(h)
                header_col += 1
            else:
                break
        param_headers_map[algo_idx] = headers

        # Map variations
        while True:
            val = sheet.range((row, col)).value
            if val:
                variation_map[algo_idx].append(row)

                # Cache full stock info for each variation
                r = row + result_offset
                stocks = []
                while True:
                    ticker = sheet.range((r, col + param_offset)).value
                    if not ticker:
                        break
                    stocks.append({
                        "ticker": str(ticker),
                        "name": sheet.range((r, col + param_offset + 1)).value,
                        "decision": sheet.range((r, col + param_offset + 2)).value,
                        "result": sheet.range((r, col + param_offset + 3)).value,
                        "profit": sheet.range((r, col + param_offset + 4)).value,
                        "verbal_result": str(sheet.range((r, col + param_offset + 5)).value).upper(),
                        "verbal_decision": sheet.range((r, col + param_offset + 6)).value,
                        "symbol_selected": sheet.range((r, col + param_offset + 7)).value
                    })
                    r += 1
                full_stocks_map[(algo_idx, row)] = stocks

                row += 1
                blanks = 0
            else:
                blanks += 1
                row += 1
                if blanks >= 3:
                    break

    def extract_params(row, algo_idx):
        col = identifier_cols[algo_idx] + param_offset
        values = []
        for j in range(len(param_headers_map[algo_idx])):
            values.append(sheet.range((row + 1, col + j)).value)
        return values

    def intersect_full_stock_data(s1, s2):
        t1 = {s["ticker"]: s for s in s1}
        t2 = {s["ticker"]: s for s in s2}
        shared = []
        for ticker in t1:
            if ticker in t2:
                # Prefer s1’s version to keep consistency
                shared.append(t1[ticker])
        return shared

    def calc_pp(stocks):
        if not stocks:
            return 0
        profit_count = sum(1 for s in stocks if "PROFIT" in s["verbal_result"])
        return round(100 * profit_count / len(stocks), 2)

    group_results = []

    def build(group, shared_stocks, next_algo):
        if len(group) >= 2:
            pp = calc_pp(shared_stocks)
            if pp >= threshold_pp:
                group_results.append((group.copy(), shared_stocks.copy(), pp))

        for next_algo in range(len(identifier_cols)):
            if any(a == next_algo for a, _ in group):
                continue 

            for row in variation_map[next_algo]:
                new_stocks = full_stocks_map[(next_algo, row)]
                common = intersect_full_stock_data(shared_stocks, new_stocks)
                if not common:
                    continue
                new_pp = calc_pp(common)
                old_pp = calc_pp(shared_stocks)
                if new_pp >= threshold_pp or new_pp > old_pp:
                    group.append((next_algo, row))
                    build(group, common, next_algo + 1)
                    group.pop()

    for algo_idx in range(len(identifier_cols)):
        for row in variation_map[algo_idx]:
            stocks = full_stocks_map[(algo_idx, row)]
            build([(algo_idx, row)], stocks, algo_idx + 1)
          
    # result
    row_cursor = paste_row
    for group, shared_stocks, pp in group_results:
        sheet.range((row_cursor, paste_col)).value = "_".join([sheet.range((r, identifier_cols[a])).value for a, r in group])
        sheet.range((row_cursor, paste_col + 1)).value = pp

        total_rows = 0
        row_cursor += 1
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

        for i, stock in enumerate(shared_stocks):
            rr = row_cursor + 1 + i
            sheet.range((rr, paste_col)).value = stock["ticker"]
            sheet.range((rr, paste_col + 1)).value = stock["name"]
            sheet.range((rr, paste_col + 2)).value = stock["decision"]
            sheet.range((rr, paste_col + 3)).value = stock["result"]
            sheet.range((rr, paste_col + 4)).value = stock["profit"]
            sheet.range((rr, paste_col + 5)).value = stock["verbal_result"]
            sheet.range((rr, paste_col + 6)).value = stock["verbal_decision"]
            sheet.range((rr, paste_col + 7)).value = stock["symbol_selected"]

        group_height = total_rows + len(shared_stocks) + 1
        sheet.range((row_cursor - total_rows, paste_col + 2)).value = group_height

        row_cursor += len(shared_stocks) + 3

    wb.save()
    print(f"[MP2-F2] DONE — {len(group_results)} full groups exported to WB2.")

# mp2 p2 end

if __name__ == "__main__":
run_cycling_program()
run_optuna_optimization()
# wipe_optuna_db()
save_all_variations_to_db()
export_variations_to_excel()
run_mp2_general_group_tester()

print("MP1 - Part 2 Completed.")
