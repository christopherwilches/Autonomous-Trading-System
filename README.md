# Autonomous-Trading-System
Experimental end-to-end algorithmic optimization framework integrating Python, SQL, and Excel.
Built to study how multi-day performance data can refine algorithmic trading strategies through adaptive parameter tuning and group optimization. This system automates the full research cycle — from dataset collection to multi-day testing, ML-style optimization, and recursive group formation.
Each module builds on the previous one, forming a continuous pipeline that self-improves over time through iteration, scoring, and pruning logic.

## MP0 - Initial Set-up

- Uses Google Sheets + Apps Script macros to collect NASDAQ and NYSE tickers automatically.
- Due to size restrictions, distributes tickers across six Data Documents (DDs), each using custom GOOGLEFINANCE() calls to fetch 31 days of OHLCV data per ticker.
- An Apps Script macro consolidates the six DDs into one master export and generates clean CSVs of price and volume data.
- CSVs are downloaded to a remote server, then processed by a Python data assimilation program that: Filters tickers by price, liquidity, and data completeness, organizes them into dictionary tables and SQLite database entries, and randomizes 1,000-ticker datasets per day (Mon–Fri) for consistent 5-day testing.
- These are the databases that all later programs use to reference stock data.


  ## MP1 P1 - Variation Tester (Excel Cycler)

- Tests 10 algorithms (MACD, EMA, RSI, ADX, SMA, Bollinger Bands, Breakout, etc.) each under 8 different parameter variations.
- Uses Excel VBA macros to automatically replace datasets under fixed algorithm blocks, saving outcomes directly below each block of variations.
- Each dataset represents one day in the 5-day training window; the macro cycles all five days.
- The Python script:

  Triggers macros remotely using helper cells.

  Monitors dataset progress.

  Collects results and appends them to SQLite result tables for each algorithm.

- The system logs detailed performance data for each variation, including per-day profit rate (PP), buy counts, ticker overlap, and variance metrics all in the database.
- Past versions of this program used single day analysis on Friday for all testing, but after measuring the volatility of performance the next week using only Friday to test my algorithms, I changed the program to do 5 day testing on my groups to account for consistent performance, along with many new metrics to increase reliability

  ## MP1 P2 - Sci-kit ML Optimizer

- Reads all historical variation results from SQLite tables, including 5-day metrics and parameter sets.
- Converts Excel-style parameters into internal JSON schemas for cross-session consistency (later reptitions can easily interpert the data).
- Each variation is scored with a lexicographic scoring function balancing:

  Profitability (PP)

  Stability (MAD/IQR)

  Reliability (Wilson Lower Bound)

  Buy Count CV and Ticker Concentration penalties

  Recency weighting for recent-week improvements

- Uses a custom ML-inspired search loop:
Runs an initial range scan to find optimal value regions for each parameter.
Performs ablation testing around top performers.
Uses model-based sampling (Gaussian around good regions, uniform elsewhere) for controlled exploration.
Enforces diversity and eliminates potential duplicate suggestions by comparing parameter sets.
- Replaces Optuna entirely — past versions of this part used Optuna (hyperparameter optimization software) but due to mroe metrics and introduction of 5-day testing, it was removed for the current system. Now it learns faster, remains deterministic, and has more predictable and tunable parameter decisions.
- MP1-P1 and MP1-P2 run as a continuous adaptive loop:

  Test → Rank → Generate → Retest

  Each cycle improves variation quality and algorithm performance.

## MP1 P3 - Filter of Variations

- Filters and ranks all variations by combined metrics across the 5 days to find stable, repeatable performers.
- Applies hard thresholds (minimum hits, max dispersion) and score-based ranking.
- Exports top variations per algorithm into “pruned tables” — ready for group-level testing.
- These pruned variations represent the best individual performers for each algorithm, forming the base for synergy testing in MP2-P2.

## MP2 - P2 - Threading Recursion Grouper

- Uses multi-threading and recursion to identify which algorithm combinations produce the most consistent and profitable shared performance.
- Reads pruned tables and constructs in-memory maps of every algorithm’s variation, tickers, and buy/profit data for easy access in RAM instead of continously accessing the DB.
- Each thread starts from one “anchor variation”, building up group combinations recursively:
Intersects shared tickers and profit data for each combination.
Computes new consensus buys across all 5 days and same metrics calcluated in MP1-P1
Prunes weak groups early if daily hit requirements or minimum PP thresholds are not met.
- For each valid group:

  Calculates pooled 5-day PP, median/second-worst PP, MAD/IQR, CV of buys, and Wilson lower bound.

  Uses those metrics for a new composite score, which it uses to compare groups and save the best ones for each size (groups with 3-6 algorithms) into another table in my SQLite file.

- Features:

  Adaptive time budgeting: earlier faster threads donate unused time to slower ones later on.
  
  Heartbeat system: writes live memory and performance stats to CMD logs.

  Thread watchdogs: enforce timeouts and prevent infinite recursion.

  Per-algorithm runtime stats tracked and saved to another table with weekly timestamps for future reference. It stores total time taken, number of threads, and number of timeouts for the threads for each algorithm.

- The result is dynamic exploration of synergy in top groups, by finding which combinations of variations produce overlapping profitable tickers with minimal volatility to give the best chance of high performance next week.

## MP2 - P3 - Executor of Top Group

- Takes the top-ranked group from MP2-P2 and re-tests it live in Excel with the same algorithms.
- Puts the chosen parameters into the correct algorithm slots, then runs the Cycler VBA macro using the most recent historical ticker data to provide the predicted buys.
- Saves shared buys and prints results to CMD logs and into another table in the DB. 

## Cycler VBA Macro - Excel Cycler program

- Cycles through all datasets already loaded into the ControlSheet and updates the active data block used by every algorithm
- Replaces the top dataset each cycle, forcing a full recalculation across all algorithm sheets to simulate how each variation reacts to new market data
- Records each algorithm’s decisions, metrics, and price information into result tables below their respective sections in Excel, one row per dataset
- Uses helper cells to manage batches and track progress, signaling completion using helper cells that Python can use to know when each cycle is done
- Automatically restores the original dataset at the end of the process to keep Excel ready for the next run

## Scoring Logic and Metrics

- Each variation and group is evaluated using multiple 5-day performance metrics, stored in SQLite databases. Some of these metrics include:

  1. PP (Profit Percentage): Primary profit metric for each day
  2. Median PP/Wilson Lower Bound: Emphasizes reliable returns with statistical confidence to ensure lower volatility
  3. MAD/IQR (Dispersion Metrics): Penalizes volatility if the values skews too much. One of my top metrics in my ML logic and score formulas
  4. Buy/Count CV (Coefficent of Variation): Analyzes daily buy count consistency
  5. Repeat Ticker Rate (RTR): Measures concentration of buys and overlap of same tickers bought across multiple days
  6. Recency Weighting/EWPP: Favors improvments in later dates, meaning better performance in Friday is favored over Monday

- These metrics combine in a lexicographic scoring function:
  1. Hard gate by minimum daily buys/hits and Wilson LB score
  2. Rank by median PP
  3. Penalize instability (MAD/IQR) and concentration (overlapping of same tickers bought multiple times)
  4. Reward recent performance and evidence (pooled buys)

## System Flow Chart (Visualizer)

  1. (MP0) Fetch all tickers using macro for selected date in Google Sheets (GS)
  2. (MP0) Send parts of ticker list to the DDs and copy the GoogleFinance formula for the data
  3. (MP0) Run macro that collects all the data and compiles it into CSV files all formatted
  4. (MP0) Upload tickers to dictionaries and datasets in SQLite tables after filtering
  5. (MP1 - P1 LOOP START) Pastes batches of the different datasets into Excel, and triggers the cycling macro
  6. (MP1 - P1) Cycling macro cycles the datasets and records data, using helper cells to keep track of progress and communicate with python
  7. (MP1 - P1) Saves data for 5 days and calculates metrics and saves them in the database
  8. (MP1 - P2) Processes all 5 day metrics and parameters in the tables, runs ML inspired logic to choose the next best combinations of values for each algorithm
  9. (MP1 - P2 LOOP END) Pastes the new 8 variation parameters for each of the 10 algorithms in Excel
  10. (MP1 - P3) Filters all variations tested after the loop, and leaves only a few hundred of the top variations per algo using complex scoring formulas that use various metrics
  11. (MP2 - P2) Uses pruned tables to then start testing possible group combinations with every variation used as a base anchor in their own thread
  12. (MP2 - P2) Compares the new shared ticker list, and recalculates new 5 day metrics, and scores the current group to save it or not. Enforces hard gates to end recursion branches early if shows early undesirable results
  13. (MP2 - P2) Saves top groups for each size (3-6) including all metrics into the database along with their parameters
  14. (MP2 - P3) Tests the top performing group by exporting its parameters to Excel and running the new datasets for the latest day to collect shared buys for tomorrow

## Design Notes

- This entirely built on Notepad, and executed via CMD by running the file (python mp1.py), without IDEs, auto-save, or version control.
- The original concept of this project came from building simple algorithms on Google Sheets, and exploded into what it is today due to my principles of optimization, automation, and perfection
- This is my first large-scale coding project, designed and maintained independently.
- I developed it iteratively through trial, debugging, and optimization. One small change at a time, refining it little by little. 
