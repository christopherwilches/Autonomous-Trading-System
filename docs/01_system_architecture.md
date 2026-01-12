# System Architecture

This doc explains how the system is organized: the major components, what each one consumes and produces, and how the pipeline runs on a weekly and daily schedule.

The system has two schedules:

- a weekend training + group-selection run
- a weekday run that generates tickers and executes trades

Each weekend, it trains on the most recent 5 trading days and selects a group. That group is then used for the next 5 trading days.

---

## High-Level Layers

The architecture is organized in layers. Each layer has a clear purpose and a clean input/output boundary.

### Data Layer
Responsible for building the weekly datasets and resetting weekly state.

- builds a 5-day dataset from Alpaca
- filters the universe down to tradable tickers
- ensures complete 5-day OHLCV coverage
- samples a single 1,000-ticker universe used across all 5 days
- wipes old weekly results tables before new training begins

### Training Layer
Responsible for generating variation performance data.

- runs the mp1 loop for N iterations
- mp1_p1 executes variation tests in Excel across all 5 days
- mp1_p2 reads results and generates the next parameter sets

### Decision Layer
Responsible for selecting what will actually be used next week.

This is no longer training. This is where the system chooses what moves on and builds groups.

- mp1_p3 prunes variations into pruned tables using score + hard gates
- mp2_p2 searches for multi-algorithm groups using recursion + threading
- mp2_p2 selects final candidate groups for next week using scoring + anti-overfit selection logic

### Daily Decision Layer
Responsible for generating the daily ticker list from the chosen group.

- pulls the latest daily snapshot data
- runs the chosen group in Excel to generate “tomorrow” candidates
- ranks and trims to a shortlist

### Qual Layer
Responsible for scoring the daily shortlist using news/narrative prompts.

- uses the mp2_p3 shortlist
- produces a final smaller set to trade (currently: top 4)

### Execution Layer
Responsible for placing trades and managing exits.

- allocates total capital across the final tickers
- executes market open entries
- exits using custom TSH rules (conservative profit capture and trailing stops)

---

## Core Tables and Storage Databases

There are two primary SQLite databases, plus the Excel workbook which functions as a compute engine.

### stocks_data_db
This database holds the weekly datasets, daily dictionary snapshots, and pruned tables used for group discovery.

Table families:

- dict_day1..dict_day5  
  raw weekly OHLCV tables for the 5 training days

- ds_day1..ds_day5  
  the 1,000-ticker dataset tables used by training, one per day  
  important: the ticker list is the same 1,000 across all five days

- dictionary_table  
  the daily snapshot table used for weekday execution

- pruned_adx, pruned_ema, pruned_macd, ...  
  the output of mp1_p3  
  these are the only variations mp2_p2 is allowed to use

- final_groups
  the output of mp2_p2
  gives the final two groups that will be used for testing
  
### variation_results_db
This database holds the variation-level performance results produced by mp1_p1 across all cycles.

In earlier versions of the system, this database appears under a legacy name tied to an abandoned Optuna-based approach. Functionally, it serves as the weekly results store used by the mp1 loop.

Typical contents per algorithm:

- parameter sets for each tested variation
- daily PP values and derived stability metrics
- buy counts and overlap metrics
- any derived score fields used downstream

### Excel Compute Engine
The Excel workbook is not storage. It is the compute engine for the variation tester and the daily execution run.

It contains:

- algorithm blocks (10 algos, 8 variations each)
- a cycling macro that runs datasets and records outputs
- helper macros used for shifting/clearing/control
- helper cells used by Python to coordinate the run state

Python commands Excel. Excel performs the per-ticker computation.

---

## Module Map (Inputs and Outputs)

This section is the “wiring diagram” in text form. Each module is described as input -> process -> output.

### weekly_5d_data
input:
- Alpaca asset universe
- 5 trading days of OHLCV

process:
- filter out non-common-stock style assets (etfs, unsupported symbols, punctuation tickers, etc.)
- filter by liquidity and price constraints using the earliest day as the initial screen
- verify complete OHLCV coverage for all 5 days
- randomly sample 1,000 tickers from the surviving pool
- write dict_day1..dict_day5 and ds_day1..ds_day5 into stocks_data_db

output:
- stocks_data_db tables: dict_day1..dict_day5, ds_day1..ds_day5

runtime:
- ~110 seconds typical

### wipe_weekly_results
input:
- optuna_10a_data_db

process:
- wipes weekly performance tables and group tables before training begins
- keeps the system “fresh” each week by not carrying forward last week’s outcomes

output:
- empty weekly results tables ready for new writes

### mp1_p1_variation_tester
input:
- ds_day1..ds_day5 from stocks_data_db
- current parameter sets for 10 algos * 8 variations
- Excel compute engine (macros + sheets)

process:
- for each day in day1..day5:
  - split 1,000 tickers into 4 batches of 250
  - for each batch:
    - paste data into Excel
    - trigger cycling macro
    - wait on helper-cell completion signal
    - advance helper cells to control where results are written
  - after all 1,000 tickers are processed for that day:
    - extract results from Excel
    - write results into optuna_10a_data_db
    - wipe/clear Excel result region
- repeats for all 5 days

output:
- per-algorithm performance data in optuna_10a_data_db

### mp1_p2_optimizer
input:
- optuna_10a_data_db (all accumulated results for the week so far)

process:
- reads historical variation outcomes and their metrics
- uses a deterministic ML-inspired search process to generate 8 new parameter sets per algorithm
- avoids duplicate parameter proposals
- writes the next parameters back into the Excel blocks for the next mp1_p1 run

output:
- updated parameter sets in Excel for the next cycle

### mp1_loop
input:
- mp1_p1_variation_tester
- mp1_p2_optimizer
- cycle_count (current count: 200)

process:
- repeats:
  - run mp1_p1
  - run mp1_p2
- accumulates large weekly performance history per algorithm

output:
- a large weekly results dataset across all algorithms and variations

### mp1_p3_pruner
input:
- optuna_10a_data_db (all weekly results)

process:
- scores variations using performance + stability + confidence metrics
- applies hard gates to remove unstable or low-signal variations
- keeps different counts per algorithm based on algorithm role (example: 100-250)
- writes surviving variations into pruned tables inside stocks_data_db

output:
- stocks_data_db tables: pruned_algo_name

### mp2_p2_group_discovery
input:
- pruned tables in stocks_data_db

process:
- loads all pruned variations into RAM to avoid DB bottlenecks
- spawns worker threads (limited by CPU cores, example: 4 concurrent threads)
- each thread anchors on one variation and expands combinations via recursion:
  - intersects tickers across variations
  - recalculates group metrics on shared tickers across 5 days
  - prunes branches early when guardrails fail
- uses a 24-hour global runtime budget enforced by a watchdog system:
  - budget is distributed per thread
  - early threads that finish fast donate time to later threads
  - total runtime converges to the global budget

output:
- a ranked set of candidate groups stored in the final_groups table
- the selected group(s) saved for weekday execution

### daily_snapshot
input:
- Alpaca daily bars for the current trading day close

process:
- fetches daily OHLCV snapshot for the eligible universe
- writes results into dictionary_table in stocks_data_db

output:
- stocks_data_db: dictionary_table

### mp2_p3_executor
input:
- selected group parameters (from the final_groups table)
- dictionary_table (today’s snapshot)
- Excel compute engine

process:
- pastes group parameters into the workbook
- runs an execution version of the Excel cycle process on today’s snapshot
- corrects for the “first row confirmation” behavior by shifting macros so “today’s data” generates “tomorrow’s decision”
- compiles candidate tickers
- ranks candidates using a short lookback sanity check (example: 1 month data fetch per candidate)
- trims to a shortlist (currently: top 20) with a rank-based score

output:
- daily candidate tickers + ranking signals

### qual_scoring
input:
- mp2_p3 shortlist tickers

process:
- runs prompt-based news/narrative scoring
- combines prompt scores with rank score
- selects final tickers to trade (top 4)

output:
- final trade list for the next session

### real_trader
input:
- final trade list
- account equity and constraints

process:
- waits until market open
- splits capital across tickers (currently: 1/4 each)
- places marketable entry orders

output:
- open positions at/near market open

### tsh_trade_management
input:
- entry prices
- real-time price updates

process:
- applies bracket logic by price range
- defines a conservative target change per bracket
- tracks price movement and sells on:
  - hitting target behavior
  - trailing stop conditions after reaching a local high
  - end-of-day exit if still open

output:
- exit orders + realized intraday profits/losses

---

## Weekly and Daily Control Flow

### Weekend Cycle
1. weekly_5d_data writes dict_day1..dict_day5 and ds_day1..ds_day5
2. wipe_weekly_results resets weekly tables
3. mp1_loop runs for N cycles (example: 200)
4. mp1_p3_pruner writes pruned tables
5. mp2_p2_group_discovery runs under a 24-hour watchdog budget
6. final group selection is saved for the upcoming week

### Weekday Cycle
1. after close, daily_snapshot writes dictionary_table
2. mp2_p3_executor generates and ranks candidates for tomorrow
3. qual_scoring selects the final trade list
4. real_trader runs at market open and enters positions
5. tsh_trade_management manages exits intraday (during stock market hours or at closing)

This weekday cycle repeats for 5 trading days using the group selected on the weekend.

---

## Design Constraints that Shaped the Architecture

- runtime constraint: full weekly retrain + group search requires weekend time, about 40 hours to train
- batching constraint: Excel processing is done in 250-ticker batches to stay stable and controllable
- stability constraint: multi-day testing is required to avoid one-day overfit behavior
- compute constraint: group discovery explores a tiny part of a massive amount of possible search space and must be bounded by time budgets and guardrails
- iteration constraint: early development was done without IDE version control, so explicit iteration versions are limited but included separately in `legacy_and_iteration`
