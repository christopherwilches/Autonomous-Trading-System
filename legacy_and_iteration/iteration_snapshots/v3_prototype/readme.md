# Legacy Snapshot — MP1 + MP2 Runner + Real and Paper P3 Testing

## What this version is
This snapshot shows the last major legacy version of the current program, where the project became a real multi-part pipeline, with many working parts and a clearer goal, but was still "manual-run" in some parts.

At this stage:
- **Excel + VBA** is the compute engine, but now has faster batching structures and the cycling macro
- **SQLite** is still the dataset + results store
- **Python** still orchestrates: feeds Excel, triggers macros, reads results, logs performance, and moves the pipeline forward
- MP2 exists as a real concept: **group optimization** by intersecting tickers across strategies/variations, and is now more complex, using more time limits and implementing more changes
- MP2 P3 now exists, and serves the same purpose as it does currently: take the top group produced by MP2 P2 and test it using provided data. Now there is a second paper trading version, capable of testing any date utlizing
  special macros in Excel that allow this

This snapshot is not meant to be fully runnable in a public environment (paths, workbook layout, and private data assumptions are not included here), but it accurately represents the architecture and flow at the time.

---

## What it does

### MP1 Part 1 — Cycling / batch runner (`run_cycling_program`)
This is the core execution loop:
- Pulls batches of tickers from SQLite
- Pastes them into the workbook
- Triggers a VBA “cycling” macro
- Waits on completion signals and advances through result blocks

This was the first stable version of:
**DB → Excel → Macro → Results → Next batch**

---

### MP1 Part 2 — Parameter proposal + logging (`run_optuna_optimization`)
This stage proposes variation parameters and measures performance through Excel:
- Chooses parameters (Optuna-driven at this time)
- Writes them into the Excel parameter block
- Reads back profit-percentage (PP) signals from the sheet
- Stores results into SQLite

It worked, but it was still “early optimization”: minimal metrics, limited stability validation, use of Optuna, and heavy dependence on workbook state.

---

### MP1 Part 3 — Export variations to WB2 (`export_variations_to_excel`)
This stage converts MP1 output into a standardized workbook format (WB2):
- Pastes each variation in a consistent layout (headers, parameter values, result table)
- Enforces spacing so later programs can parse variations reliably

This is what made MP2 scalable: without a strict WB2 layout, group-testing becomes fragile.

---

### MP2 Part 2 — General group tester (`run_mp2_general_group_tester`)
This is the first real “multi-strategy optimizer” layer:
- Reads many variation blocks from WB2
- Intersects tickers across combinations
- Computes group metrics (PP + scoring logic)
- Keeps top-ranked groups and logs them
- Uses more advanced time limit measures

This stage is where the system stops being “pick the best single variation” and becomes:
**find the best compatible combinations**.

---

### MP2 Part 3 — Group executor (`run_mp2_part3_group_executor`)
Takes the best group output and converts it into actionable pasted results:
- Loads external stock data sources (dictionary from CSVs)
- Executes the selected group logic
- Pastes final consolidated outputs back into the workbook for review/trading
- Has a paper trading version (looks like a copy in the code) that can test any date

---

## Utility / helper functions (why they exist)
These are support operations used to keep the pipeline stable:

- `wipe_optuna_db()`  
  Clears Optuna-related state when I want a clean optimization run.

- `load_latest_csvs_into_dictionary()`  
  Loads cached stock data into memory for faster access in later stages.

- `populate_dataset_one_from_dictionary()`  
  Converts dictionary-loaded data into the standardized dataset table format.

- `analyze_pp_vs_tickers()`  
  Quick diagnostic tool to compare PP behavior against ticker counts.

- `renumber_past_buys_safe()` (and other past_buys helpers)  
  Maintenance tools for trade-history blocks during iteration/testing.

---

## How this snapshot was run
A list of the different functions and helper functions are at the bottom for easy access and construction of runner functions

Reason:
- During development, I needed to quickly run individual stages for debugging.
- It also made making larger functions to run the entire pipeline easier.

Then below that, the script includes structured runners like:
- `full_cycle_main()` (loops MP1 P1 + P2 many times, then runs P3 and MP2 P2)
- `setup_dictionary_data()` (loads data and runs MP2 P3)

---

## What was still rough here
This snapshot worked, but it had clear limitations:
- **Single-file orchestration** (not yet split into clean modules)
- **Excel dependency** (in MP2 P2 it still relied on mapping WB2 for its parameters and info, still interacting with it directly)
- **Light metrics** (mostly PP-focused; later versions added stronger gates and stability checks)
- **Early group testing** (functional, but later improved with better pruning, similarity filters, and performance scaling)
- **Lack of trading program** (trades were still done manually, and no custom market strategy was made)

---

## How it evolved into the current system
After this snapshot, the project moved toward:
- stronger logging/telemetry and safety checks,
- more reliable progression through variations and datasets,
- more advanced MP2 group-ranking and filtering logic,
- focus on optimization in how long each part took, and emphasis in automating as many parts as possible

The core idea remained the same:
**Python orchestrates → Excel computes at scale → results get stored → pipeline consumes outputs.**
