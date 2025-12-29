# Legacy Snapshot — Early MP1-MP2 prototype + SQLite Pipeline

## What this snapshot is
This is an early prototype of the system that later became MP1 (variation testing + optimization) and MP2 (group synergy discovery).

## Core Improvements
- Python orchestrates data movement + iteration control.
- Excel performs per-ticker computation through a cycling macro.
- Results are persisted into SQLite so downstream stages can query and recombine them.
- Performs group synergy testing

---

## Components included in this snapshot

### 1) MP1 P1 — Excel Cycling Runner (`run_cycling_program`)
- Pulls batched OHLCV rows from SQLite.
- Pastes into WB1 at `A1`.
- Writes a dynamic `START_ROW` for where Excel should dump results.
- Triggers the Excel cycling macro via a helper cell and blocks until completion.

### 2) MP1 P2 — Early Optuna Parameter Proposer (`run_optuna_optimization`)
- Uses Optuna to propose parameter sets per algorithm/variation.
- Writes parameters into the ControlSheet region.
- Reads back the “profit percentage” output cell(s) for scoring, no other metrics.
- It has a old function of returning the best PP across the 8 variations tested in that trial for debugging

### 3) Variation persistence — Store params + buys (`save_all_variations_to_db`)
- Reads each algorithm sheet in WB1.
- For each variation, collects the tickers marked “BUY” and their associated result/profit fields.
- Saves `(params, PP, trial_number, variation_number, stocks_bought)` into SQLite.
This is the early version of “variation result records” that later enabled pruning + group discovery.

### 4) Export to WB2 — Structured block writer (`export_variations_to_excel`)
- Reads stored variations from SQLite.
- Writes them into WB2 in a block format: identifiers → parameter headers/values → per-stock rows.
- This part is also legacy since it wrote the variations in Excel, still expecting user interaction with it, instead of creating it as a black box of sorts

### 5) MP2 P2 (F2) — General Group Tester (`run_mp2_general_group_tester`)
- Scans WB2 and builds an in-memory map of all variations instead of reading the DB and extracting the information from there.
- Recursively builds multi-algorithm groups by intersecting shared tickers.
- Computes group PP using the “verbal profit/loss” label from stored stock rows.
- Exports qualifying groups back into WB2.

---

## How this differs from the current system
- **Persistence:** this version stores early variation records but lacks the later standardized schemas and pruning tables to cut out the need for WB2, something that the current version does.
- **Optimization quality:** Optuna is used in a basic, single-trial-per-algo way here; later versions add stricter duplicate control, better scoring, and richer metrics.
- **Group discovery:** the recursion/group tester exists here, but later versions implement stronger pruning logic, scoring, deduping, and runtime budgeting.

---

## Privacy note
Paths, dataset identifiers, and workbook names are redacted in this repository snapshot. `WB1` / `WB2` are neutral placeholders for the Excel compute engine and the combination grouper workbook.
