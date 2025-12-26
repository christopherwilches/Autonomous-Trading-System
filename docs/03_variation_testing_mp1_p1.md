# 03_variation_testing_mp1.md

## What mp1 does
mp1 is the weekly “variation testing” stage.

It takes the fixed 5-day dataset from mp0 (`DS_DAY1..DS_DAY5`) and tests 10 algorithms × 8 variations each across all five days, then writes a set of stability and profit metrics into a database so the optimizer can learn what parameter sets are consistently profitable (not just lucky on one day), and learning parameter relationships.

### What the P2 stage is
This stage does not train a predictive model like a neural network.
The “learning” happens by mining a growing SQLite history of parameter trials and using it to propose the next parameter batch.

In practice, mp1 P2 behaves like a history-aware hyperparameter optimizer:
- It scores each parameter set using multi-day stability + confidence metrics (not just raw profit percentage).
- It fits lightweight, per-parameter sampling bias from recent “good” regions.
- It proposes new candidates, then enforces deduplication and diversity before pasting back to Excel.

Key idea
- The Excel workbook computes the trade logic.
- Python orchestrates data loading, triggers the macro, and records results + metrics into SQLite.
- P2 then proposes the next sets of parameters based on those 5-day metrics of the past variations tested.

## Components in this stage
- **P1: 5-day cycling runner** (`run_cycling_program()`)
- **Metrics recorder** (`record_day_results_and_finalize_if_needed()` → `_finalize_5d_metrics()`)
- **P2: rank + recommend (optimizer loop)** (`run_p2_rank_and_recommend()`)

The Excel Cycling Macro handles:
- Running each algorithm sheet logic on the currently loaded batch
- Saving intermediate workbook state and moving through the workbook’s internal flow
- Updating a completion cell (`O1`) so Python can detect completion

Python handles:
- Feeding the next 53-row blocks into ControlSheet
- Telling the macro “run”
- Waiting for completion
- Reading per-variation BUY/HIT outcomes from each algorithm sheet
- Writing per-day and 5-day aggregated metrics into SQLite
- Proposing the next parameter batch (helper cell `P2`) and pasting into ControlSheet (main sheet in Excel that handles dataset and cycling logic)

## Data flow at a glance
Input (from mp0)
- `stocks_data.db`
  - `DS_DAY1..DS_DAY5` (1000 tickers × 53-row blocks)

Execution engine
- Excel workbook: `MakeMoneExcel3 working - Copy.xlsm`
- Macro trigger cell: `S1`
- Macro completion cell: `O1`

Output (for training + optimization)
- `optuna_10A_data.db`
  - One table per algorithm (`MACD`, `EMA`, …, `RSI_Bollinger`)
  - Each row stores:
    - A specific parameter set (`params`)
    - A variation number (1..8)
    - A run_id (one per 5-day anchored run)
    - Day1..Day5 results
    - Aggregated 5-day stability metrics (Wilson LB, MAD, IQR, etc.)
    - Pass/fail gates (Binary metrics designed to speed up filtering in P3)

## P1: 5-day cycling runner (anchored weekly test)
Function
- `run_cycling_program()`

Purpose
- Runs the workbook against `DS_DAY1..DS_DAY5` sequentially so every parameter set is tested across five distinct day-windows.

How it works
1. Create a unique `run_id` for the whole 5-day run (`_new_run_id()`).
2. For each day `day_idx = 1..5`:
   - Read the day’s dataset table: `DS_DAY{day_idx}`
   - Split the 1000 tickers into batches (`AG4` controls size)
   - For each batch:
     - Pull the exact block-range from SQLite by `rowid` window
     - Paste the block into ControlSheet (`A1`)
     - Set `Q1` to the batch’s starting row (this tells Excel where to process)
     - Trigger the macro by writing `S1 = 100`
     - Wait until `O1` changes (macro completion)

3. After all batches for that day:
   - Reset `Q1` to the base row
   - Record day results into the metrics DB
   - Clear the workbook’s staging / cache area (`AE1 = 100`)

The important “anchoring” rule
- Day 1 creates the DB row anchors (`anchor_day1=True`).
- Days 2–5 update the same anchored rows (same `run_id`, same variation_number, same params).

This guarantees that “5-day metrics” truly describe the same parameter set across all five day windows, not five unrelated rows.

What the macro completion mechanism does
- Python reads `initial_o1 = sheet.range("O1").value`
- It triggers the macro: `S1 = 100`
- It waits until `O1 != initial_o1`
- This is a clean handshake that avoids timing guesses.

Why batches exist
- The workbook is designed to process a manageable number of stocks at a time.
- Each stock is a fixed 53-row block; batch size controls workbook load and runtime stability.
- Higher amount of tickers tested at once risks Excel lagging and calculations stalling

## Recording results: per-day buy/hit extraction
Function
- `record_day_results_and_finalize_if_needed(day_idx, run_id, anchor_day1=False)`

What it reads from Excel
For each algorithm and each variation slot (1..8), it reads:
- Ticker / name columns for the processed rows
- A 6-cell variation decision block (BUY/blank)
- The 6-column price block for each row (open/high/low/close/volume)
- It counts:
  - **buys** = number of “BUY” decisions
  - **hits** = number of “profit” outcomes (detected by `_is_profit_cellblock()`)

What gets stored for each day
- `day{d}_buys`
- `day{d}_hits`
- `day{d}_pp` = `100 * hits / buys`
- `day{d}_buys_json` = JSON list of purchased tickers with lightweight price/volume fields

Why store buys_json
- It enables ticker-level stability metrics later:
  - repeat ticker rate
  - top-10 concentration share
  - average buy price / volume
- Without having to re-scrape Excel later

## Metric design used in P1 (computed on day 5)
Function
- `_finalize_5d_metrics(conn, algo_table, row_id, ...)`

Core pooled metrics
- `pooled_buys`, `pooled_hits`, `pooled_pp`

Robust stability stats across 5 days
- `median_pp_5d`
- `pp_mad_5d` (median absolute deviation)
- `pp_iqr_5d` (robust central tendency + stability across 5 days)
- `buycount_med_5d`
- `buycount_mad_5d`
- `buycount_cv_5d` (coefficient of variation)

Confidence / “not luck” metric
- `wilson_lb_5d` = confidence-adjusted hit-rate; penalizes small sample sizes.

Recency-weighted summary (Allows more recent patterns to weigh more)
- `ew_scheme` (default `0.50,0.25,0.15,0.07,0.03`)
- `ew_pp_5d`, `ew_hits_5d`

Concentration / repetition signals (Calculates how many unique tickers it bought throughout the days)
- `repeat_ticker_rate_5d`
- `top_10_ticker_share_5d`

Market regime proxies (from buys_json)
- `avg_buy_price_5d`, `median_buy_price_5d`
- `avg_buy_volume_5d`

Last-day anchors
- `last_day_pp`, `last_day_buys`, `last_day_hits`

Gates
- Consistency gate: `min_daily_hits >= CONSISTENCY_HITS_GATE`
- Export gate: consistency + `wilson_lb_5d >= EXPORT_LB_MIN` + `pooled_buys >= EXPORT_POOLED_BUYS_MIN`

These gates are critical because the optimizer should not chase high pp (profit percentage) with weak evidence (patterns) or unstable day-to-day behavior.

## Database schema (what a row represents)
One DB row represents:
- One algorithm table (e.g., `RSI`)
- One variation slot (1..8)
- One exact parameter list (`params`)
- One anchored 5-day run (`run_id`)

That row accumulates:
- Day 1..5 results
- Finalized 5-day metrics
- Gate results

## P2: rank and recommend (optimizer loop)
Function
- `run_p2_rank_and_recommend()`

Goal
- After P1 produces fresh 5-day metrics, P2 selects the best-performing parameter sets (based on stability + confidence), generates new candidate parameter sets, and pastes them back into Excel so the next P1 run tests them.

P2 is “ML-style” in the sense that it:
- Learns a shape of “good” regions from historical runs
- Proposes new samples biased toward those regions
- Enforces deduplication and diversity so it doesn’t collapse into one near-identical cluster

### Step 1: Pull current params (from DB)
- `_db_current_params_for_algo()` returns 8 parameter lists (variations 1..8) for each algorithm, using the most recent rows.

### Step 2: Score each variation using 5-day metrics
- `_fetch_5d_metrics_for()` retrieves recent rows for the exact (params, variation) pair and aggregates them with a decaying recency weight.
- `_lexi_score()` turns the metric bundle into a single scalar score.

Lexi score design (dominant priorities)
1. Must pass consistency gate (hard fail otherwise)
2. Maximize `wilson_lb_5d` (confidence)
3. Maximize `median_pp_5d`
4. Penalize instability (MAD, IQR, buycount CV)
5. Use evidence + recency as secondary tie-breakers
6. Penalize concentration (repeat tickers / top-10 share)

### Step 3: Generate new candidates (history-aware)
- `_propose_new_params_for_algo()` builds candidates in an internal parameter representation.
- It uses:
  - recent “good” parameters as anchors
  - per-parameter sampling (numeric gaussian around good regions + categorical weighting)
  - ablation-style “change one thing” proposals
  - global exploration samples

Special handling
- ADX and Volatility use internal adapters because the workbook format differs from internal optimization shape (e.g., Volatility merges a/b into `a-b`).
- They both have a unique parameter that is made of two mini parameters. For example, ADX has a parameter where it can exist in one of two different ranges. So one range has to be selected, then a value from that selected range must be selected.

### Step 4: Deduplication and diversity enforcement
This matters because snapping to Excel steps can make different internal proposals collapse into identical JSON after coercion. Sometimes it can round up values, and alter them slightly. 

P2 prevents that by:
- Deduping against historical DB-canonical params (`hist_canonical`)
- Deduping against current on-sheet params (`current_canonical`)
- Dropping “near clones” using a grid-normalized L1 distance threshold
- Forcing diversity across critical axes (`_force_batch_diversity()`)

### Step 5: Paste to Excel in workbook order
- `_write_params_to_excel()` clears old parameter cells and writes the new 8 rows.
- Internal → Excel conversion happens here (`_to_excel_params()`).
- This keeps the workbook UI stable and prevents extra columns from spilling right.

## Why the Excel macro is still the core compute engine
Even though P2 is optimizer-style, the true “model” of profit/loss is embedded in the workbook logic:
- Decision cells
- Profit cell blocks
- Algorithm sheets per variation
- The macro-driven cycling loop

Python treats the workbook as a deterministic black box:
- Provide blocks of data
- Trigger run
- Read standardized outputs
- Save structured results
- Improve parameters using the saved results
- Inputs datasets → Outputs results and metrics

