# 06_pruning_and_stability.md

This part converts the full MP1 history into a compact, stability-first candidate set for MP2.

Input
- **Source DB**: **optuna_10A_data.\*** (contains full 5-day trial rows per algorithm)
- **WB1**: **MakeMoneyExcel3** (used only to read parameter headers)
Output
- **Destination DB**: **stocks_data.db**
  - **pruned_<ALGO>** tables (rich schema, full 5-day metrics + day JSON)
- **WB2**: **Combination Grouper MME3** (ControlSheet receives the pruned variations)

The export focus is consistency-first: it aims to remove unstable or junk parameter sets, deduplicate near-clones, and preserve ticker coverage diversity across selected variations.

---

## Per-algorithm export plan
Algorithms are processed independently, then combined under a global selection policy.

Per-algo cap
- **SPECIAL_ALGOS = {MACD, EMA, Breakout} → cap = 250**
- all others → **cap = 100**
Global bounds
- **GLOBAL_MIN = 700** total exported variations minimum
- **OVERALL_MAX = 1300** total exported variations maximum

The special algos are simply algorithms that have proven under many tests to typically perform better than the rest of the algorithms, so increasing their number maximizes their potential impact. The min and max ensure a healthy amount of varaitions to test combinations with. 

---

## Row extraction
Every DB row is normalized into an internal record:
- **trial_number**
- **variation_number**
- **params** (JSON list)
- **stocks** = Friday-only buys (from **day5_buys_json**)
- **day_pp** list for day1..day5
- **metrics** dictionary (full 5-day pooled and robust stats)

Derived fields computed during extraction
- **pp_min_5d**
- **pp_max_5d**
- **pp_range_5d**

Rows missing required identifiers (`trial_number`, `variation_number`) or other fields are discarded.

---

## Minimal eligibility gate
Goal
- Remove dead / junk runs without over-filtering

Function
- **eligible_basic(m)**

Rules
- **pooled_buys >= 10**
- **min_daily_hits >= 1**
- **median_pp_5d > 0**
- rejects extreme instability only when both are bad:
  - **pp_mad_5d > 25 AND pp_range_5d > 80**

The gate is intentionally lenient: ranking and dedup handle quality after pruning.

---

## Unified export scoring
Goal
- Produce a score value that favors steady, evidence-backed, non-concentrated behavior

Function
- **export_score(m)**

Inputs
- profit stability: **median_pp_5d**, **pp_mad_5d**
- buy stability: **buycount_cv_5d**
- concentration penalties: **repeat_ticker_rate_5d**, **top_10_ticker_share_5d**
- evidence / confidence: **wilson_lb_5d**, **min_daily_hits**

Score formula
- `0.55 * median_pp_5d`
- `+ 0.20 * (100 - pp_mad_5d)`
- `+ 0.10 * (100 - buycount_cv_5d)`
- `+ 0.10 * (100 - repeat_ticker_rate_5d)`
- `+ 0.05 * wilson_lb_5d`
- `- 0.05 * top_10_ticker_share_5d`
- `+ 0.02 * min_daily_hits`
- clamped to **>= 0**

Interpretation
- median-based profit is dominant
- instability and concentration reduce rank even if median profit is high
- Wilson LB provides confidence adjustment for small-sample luck
- min_daily_hits rewards consistency across all five windows

---

## Deduplication logic
Goal
- prevent near-identical parameter sets with overlapping ticker behavior from occupying multiple slots

Function
- **old_dedup_pool(pool, algo_name)**

Mechanisms
- **variation_similarity(params1, params2, algo_name)**
  - tolerance is based on parameter step size and how many discrete options exist
  - Volatility has special handling because it encodes a pair as `"a-b"`
- **ticker_overlap_frac(stocks1, stocks2)**
  - overlap fraction uses intersection / max(|set1|, |set2|)

Duplicate decision
- default: duplicate if **variation_similarity** and **ticker_overlap >= 0.60**
- EMA special case: duplicate if symmetric difference size is <= 1 ticker
Replacement rule
- keep the better candidate by **pooled_pp** and **breadth** (stock count), depending on branch

---

## Coverage-based selection
Goal
- maximize unique ticker coverage across exported variations instead of selecting many that buy the same names

Function
- **coverage_select(dedup_sorted, cap)**

Stage 1
- add candidates that increase union coverage of tickers

Stage 2
- fill remaining slots using a deterministic tie-breaker:
  - prefer low overlap to selected set
  - then higher pooled_pp
  - then larger ticker set

This produces a cap-limited set with broader exposure for MP2 group synergy tests.

---

## Safety net fallback
If all candidates are removed post-dedup or post-gate:
- fallback pool: rows with **min_daily_hits >= 1**
- sort by **pooled_pp desc**, then breadth
- dedup again

This prevents empty exports for an algorithm when the gate is too strict for that sample.

---

## Global backfill and trimming
After per-algorithm selection:
- if total < **GLOBAL_MIN**, fill deficit by pulling from per-algo near-misses
  - respects per-algo caps
  - round-robin across algorithms until minimum is reached or near-miss pools exhaust

If total > **OVERALL_MAX**
- global trim keeps highest **final_order_score** across all algorithms
- removes the rest deterministically

---

## Writing pruned tables into stocks_data.db
For each algorithm, exporter creates:
- **pruned_<ALGO>**

Key properties
- **identifier** is primary key: `No.<algo_idx>.<rank>`
- stores:
  - **final_order_score**
  - **params_json**, **param_headers_json**
  - **stocks_json_day5**
  - full **day1..day5** buys/hits/pp + buys_json
  - pooled and robust metrics
  - recency summary fields
  - concentration fields
  - derived pp range fields
  - gates and min_daily_hits

Proccedure to insert in table:
- schema-driven insertion using **PRAGMA table_info** to guarantee order alignment
- **INSERT OR REPLACE** used for idempotent reruns

---

## Pasting to WB2 (vectorized)
WB2 ControlSheet receives per-algorithm blocks.

Each variation block contains
- identifier row: `[uid, block_height, trial_num, var_num, score]`
- parameter header row
- parameter values row
- result header row (fixed columns)
- one row per Friday buy ticker (BUY rows only)

Writes are vectorized
- one large write for id columns
- one large write for the parameter/results matrix

Only Friday buys are pasted
- WB2 is focused on overlap/synergy grouping using a single consistent anchor day for display
- full 5-day history remains available in **pruned_<ALGO>** tables for deeper analysis

---

## End-state
On completion:
- **stocks_data.db** contains pruned per-algo tables with full 5-day metrics
- **WB2 ControlSheet** displays pruned variations in the standardized layout required by MP2
- totals are constrained to **[GLOBAL_MIN, OVERALL_MAX]** and per-algo caps are enforced
