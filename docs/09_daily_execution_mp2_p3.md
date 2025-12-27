# 09_daily_execution_mp2_p3.md

This section documents the **daily execution pipeline** once MP2 Part 2 has already produced Top Groups, and it is time to generate **today’s tradeable ticker list** for later execution (Real Trader comes after this part).

The daily flow is:

1) MP2 P2 produced a ranked list of groups in DB table: **`extra_big_groups`** (“EBG”).
2) Run the **Daily Dictionary Builder** to refresh the market universe for the **latest fully-closed day**, formatted for Excel and stored in DB.
3) Run **MP2 Part 3**:
   - Select the best group from EBG (between size 3 and 4) by score.
   - Paste that group’s parameters into Excel ControlSheet in the correct slots.
   - Cycle that group through your intraday dictionary blocks, harvesting BUYs.
   - Intersect buys across the group’s algorithms to get “common tickers.”
   - Save candidates to `live_candidates`, print to WB2 output block.
   - Run the historical pruner to rank and reduce to the top 20.
4) Manual qualitative prompts (Copilot / OpenAI) per ticker, isolated per chat, producing a final manually created shortlist.
5) Combine: `P3 rank_score` + qualitative module scores to use to pick top 4 tickers for the next step.

---

## What MP2 P3 reads and writes

### Reads
- `extra_big_groups` (EBG) in `stocks_data.db`:
  - `identifier`, `pp`, `score`, `parameters`, `algo_ticker_counts`
- `DICTIONARY_TABLE` (53-row blocks; 31-day windows for each ticker)
- Alpaca daily bars (in the pruner stage)

### Writes
- `live_candidates` table in DB:
  - initially: candidates from group intersection
  - later overwritten by pruner with `raw_score` and `rank_score`
- WB2 output (ControlSheet FZ block):
  - group header + list of tickers
  - then pruned output labeled `[PRUNED]`

---

## Step 0 — Starting state (after MP2 Part 2)

MP2 Part 2 has already populated:

- **`extra_big_groups`** (EBG): the ranked “best group” database.
  - This is the source for daily P3 group selection.

MP2 Part 3 specifically selects from EBG:
- best-scoring group preference (unless `USE_LOWEST_SCORE_GROUP=True` for testing)

---

## Step 1 — Refresh the Daily Dictionary (DB → Excel-ready format)

Run the Daily Dictionary Builder script:

Function:
- `build_alpaca_intraday_dictionary()`

Output tables:
- `DICTIONARY_TABLE` (normal candidates)
- `SEEDLING_TABLE` (open < `SEEDLING_CUTOFF`, stored separately)

Behavior (as implemented):
- Pulls Alpaca `/v2/assets` and filters aggressively using similar filters from mp0 (`is_clean_common_stock`)
- Fetches **daily bars** (TimeFrame.Day) for each symbol (threaded)
- Drops the newest daily bar if it is “today” during market hours (NY time) so it only has completed days:
  - `drop_incomplete_today_daily_bar(bars)`
- Applies filters on the selected “latest day” (OFFSET):
  - Require:
    - `PRICE_MIN < close < PRICE_MAX`
    - `VOLUME_MIN <= volume <= VOLUME_MAX`
- Writes each passing ticker as an **MP1-style 53-row block**:
  - blank row
  - header row
  - 31 daily rows (latest at top)
  - padding rows to 53

Offset support:
- `OFFSET = 0` means “latest completed day”
- `OFFSET = 1` means “one day back,” etc.

Daily intent:
- It is run once a day, every time before trading so mp2 p3 has the most recent data possible to make its suggestions

---

## Step 2 — MP2 Part 3: pick best group and run the cycle against today’s dictionary

Entry point:
- `run_mp2_part3_group_executor_real()`

High-level procedure:

### 2.1 Select a group from EBG
- Reads EBG:
  - `SELECT ... FROM extra_big_groups ORDER BY score DESC, pp DESC` (default)

- Extracts:
  - `identifier`
  - `pp`
  - `score`
  - per-algo parameter lists (parsed from JSON)

### 2.2 Paste group parameters into Excel
- Opens:
  - MakeMoney Excel workbook (`EXCEL_FILE`)
  - WB2 workbook (`WB2_FILE`) for printing results

- Allocates each algo into a **variation slot 1..8**:
  - `slot_map[algo_num] = next slot`
  - Pasted using:
    - `paste_group_parameters_for_slot(control_sheet, algo_num, slot, param_list)`

This avoids collisions when multiple algos would otherwise paste into the same variation region. 
Since it is organized so each variation has one row under each other, it makes sure none of them overwrites each other just in case

### 2.3 Cycle the intraday dictionary blocks through Excel
- Reads how many tickers exist in `DICTIONARY_TABLE`:
  - `total_datasets = (row_count // 53)`

- Loads dictionary blocks in batches:
  - each “dataset” is a 53-row ticker block
  - each cycle loads `limit_datasets * 53` rows into Excel range A1:I

- Triggers cycler macro in cell:
  - `S1` (same as paper mode)
- Waits on `O1` to change (macro completion signal)
- Clears between windows via `AE1`

### 2.4 Harvest BUYs for each algo at the correct slot
Instead of reading a fixed “variation 1” zone, it reads the exact slot for each algo:

- `read_algo_buys_for_slot(wb, algo_num, slot)`
  - reads Name/Ticker columns
  - reads decisions from the slot’s decision column
  - collects BUY rows
  - dedupes by ticker per algo
  - normalizes name/ticker swap issues
  - enforces ticker sanity (`_looks_like_ticker`)

It aggregates BUYs across all windows so each algo has:
- `per_algo_buys[algo] = {ticker: rec}`
- This is because each of the results are pasted below each algorithm block, left to right, meaning the first variation is the left most set of columns, and the 8th is the farthest at the right.

### 2.5 Intersect across algos to get “common tickers”
After all windows, compute:

- `common = intersection( set(per_algo_buys[algo].keys()) for algo in group_algorithms )`

Guardrails:
- If `len(common) < MIN_COMMON_TICKERS`, skip this group and try next group from EBG.
- Makes sure only a acceptable amount of tickers gets to pass thorough

### 2.6 Save candidates and print to WB2
If group accepted:
- Insert tickers into `live_candidates` with:
  - `ts`, `group_identifier`, `ticker`, `name`
- Print group header and tickers to WB2 ControlSheet at:
  - start col `FZ` (182), row 6

Then the executor exits after the first acceptable group.

### 2.7 Reset macros
At end:
- triggers reset macro cell `AC1`
- This clears the section in Excel under all 10 algorithm blocks that contains the info of tickers bought

---

## Step 3 — Historical Pruner: reduce to top 20 and assign a rank score

Entry point:
- `run_live_candidates_pruner()`

What it does:
1) Pulls the latest batch from `live_candidates` by `MAX(ts)`
2) Fetches daily bars for each ticker (30 trading days target)
3) Classifies each day using predesignated “bucket target” logic from OPEN price:
   - `safe_hit` (target hit + controlled drawdown)
   - `green_close` (closed green + controlled drawdown)
   - `loss`
4) Computes stats:
   - profit_rate, safe_hit_rate, Wilson lower bound, max_loss, avg_range, etc.
5) Applies guardrails:
   - `PROFIT_RATE_MIN`, `SAFE_HIT_RATE_MIN`, `MAX_LOSS_FLOOR`, range bounds
6) Scores each ticker and selects top `MAX_FINAL_TICKERS` (default 20)
7) Assigns a **non-linear 1–20 `rank_score`** based on pruner score distribution
8) Writes the pruned list to WB2 as:
   - `{group_identifier} [PRUNED]`
9) Overwrites `live_candidates` with:
   - `raw_score` (pruner score)
   - `rank_score` (1–20)

Outcome:
- After Step 3, the DB already contains the pruned list and rank scores that you later combine with qualitative module scores.

---
## Step 4 — Manual qualitative prompts (Copilot/OpenAI)

This step exists to **qualitatively stress-test** the quant-selected tickers before trading. 

These 4 prompts are meant to gauge all aspects of a qualitative lens of a stock and quantify its performance to bet gauge which stock is best to choose.

---

### Module A — News, events, governance and structure

**What it evaluates**
- Near-term real-world risks and catalysts
- Corporate governance cleanliness
- Legal, regulatory, and structural hazards

**Why it matters**
Quant systems do not understand:
- Surprise dilution
- Lawsuits
- Governance scandals
- Regulatory or geopolitical shocks

This module answers one question:
**Is there anything in the real world that could abruptly invalidate a clean 1–5 day long?**

**Key outputs**
- news_events_score
- governance_legal_score
- structural_risk_notes
- hard_red_flags

Hard red flags here can **fully disqualify** a ticker regardless of quant rank.

---

### Module B — Fundamentals, balance sheet, dilution, ownership

**What it evaluates**
- Financial survivability
- Capital structure stress
- Dilution likelihood
- Insider and institutional alignment

**Why it matters**
Short-term trades still fail when:
- Companies are forced to raise capital
- ATMs or shelves activate unexpectedly
- Insiders are aggressively exiting

This module answers:
**Is the company structurally stable enough to not sabotage a short-term trade?**

**Key outputs**
- fundamentals_score
- dilution_risk_score
- ownership_score
- hard_red_flags

Repeated dilution or going-concern language can override technical strength.

---

### Module C — Liquidity and microstructure

**What it evaluates**
- Tradability
- Slippage risk
- Halt and gap behavior
- Spread behavior and float dynamics

**Why it matters**
Even correct direction fails when:
- Liquidity disappears
- Spreads explode
- Halts trap positions
- Meme dynamics override rational execution

This module answers:
**Can this stock be traded cleanly with tight execution and controlled exits?**

**Key outputs**
- liquidity_microstructure_score
- microstructure_notes
- hard_red_flags

Extremely illiquid or halt-prone stocks are rejected here.

---

### Module D — Social and public sentiment narrative

**What it evaluates**
- Crowd psychology
- Narrative stability
- Meme or scandal risk
- Sentiment volatility

**Why it matters**
Narratives can overpower fundamentals and technicals in the short term.
This module detects:
- Toxic sentiment
- Unstable hype
- Coordinated narrative risk

It answers:
**Is the public story around this ticker stable enough for rational trading?**

**Key outputs**
- social_sentiment_score
- narrative_notes
- hard_red_flags

Severe sentiment risk can disqualify a ticker even if all other modules pass.

---

## Output of step 4

For each ticker:
- Four independent module scores
- Explicit hard red flags
- A defensible qualitative profile

These outputs are **not averaged blindly**.
Hard red flags can override numeric scores.

---

## Role in the system

Step 4 acts as a **qualitative firewall** between:
- Quant signal generation
- Live trade execution

Only tickers that pass both:
- Quant stability and performance
- Qualitative risk inspection

are allowed into Real Trader/chosen to trade.

This is intentionally manual right now. However, there is already a plan to automate this part, using APIs and programmatic extraction of the scores. This explained more in `11_limitations_and_future.md`


## Step 5 — Final daily selection (your combined scoring)

Final score per ticker (conceptually):

- `final_score = (P3 rank_score) + (A_score + B_score + C_score + D_score)`

Then:
- pick top ~4 tickers
- pass those to Real Trader (explained in `10_trade_management_tsh.md`)

---

## Minimal daily checklist (checklist to do every day before day trading)

1) Run **Daily Dictionary Builder (Offset=0)**  
   - fills `DICTIONARY_TABLE`

2) Run **MP2 Part 3**  
   - chooses best group from `extra_big_groups`
   - cycles dictionary, intersects common buys
   - saves `live_candidates`
   - prunes to top 20 and assigns `rank_score`

3) For each ticker in the pruned top list:
   - run prompts A/B/C/D in an isolated ticker chat

4) Combine:
   - `rank_score` + module scores give the top 4 tickers for Real Trader
