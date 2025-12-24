# Autonomous-Trading-System

An experimental, end-to-end algorithmic optimization framework integrating **Python, SQL, and Excel**.  
This project studies how **multi-day performance data** can be used to refine algorithmic trading strategies through **adaptive parameter tuning, stability-based scoring, and group-level optimization**.

Rather than optimizing individual algorithms in isolation, the system is designed as a **continuous research pipeline** that tests, ranks, refines, and recombines strategies over time, prioritizing **consistency, reliability, and risk control** over overfitted predictions.

The system is used in a recurring weekly cycle. Each weekend, it processes the previous five trading days of market data, tests and optimizes algorithm variations, and generates stable algorithm groups based on shared performance. Those selected groups are then executed during the following trading week to generate daily buy decisions. This structure allows the system to continuously retrain on recent market behavior while applying its results in live conditions.

---

## How to Read This Repository

This repository is intentionally structured so that it can be understood at multiple depths:

- **If you want a high-level understanding:**  
  Start with this README, then read  
  `docs/00_overview.md` → `docs/01_system_architecture.md`

- **If you want technical depth in the various concepts used:**  
  Read the remaining files in `/docs/` (they correspond directly to system modules)

- **If you want to inspect implementation:**  
  Browse the `/src/` directory, which contains the live Python code for each module/part

- **If you want to see system behavior and scaling evidence:**  
  See `/diagrams/` and `/logs_examples/`

- **If you want to understand iteration and early design constraints and non-python programs used:**  
  See `/iteration_history/`

Each folder and document exists to explain *why* the system works the way it does, not just *what* it outputs.

---

## Project Origin and Evolution

This project began as a **simple exploratory experiment**.

The earliest version was built in **Google Sheets**, where I manually copied historical price data from Yahoo Finance and experimented with basic logic using open and close prices. I had to screen for tickers using the trending and top gainers market, making a list of them every day and analyzing them. I would paste their data and use formulas to see how many days of profit or loss it had up to that point, and manually tried to find a pattern. At that stage, the system consisted of simple cell references, conditional logic, and required full manual control at every step.

As the project evolved:

- I transitioned from manual data entry to **automated ticker collection**
- I went from manually screening potential tickers to screening the **entire stock market's tickers**
- I implemented **industry-standard indicators** (MACD, RSI, EMA, Bollinger Bands, ADX, Breakout, etc.)
- I introduced **multiple parameter variations per algorithm** instead of single fixed strategies
- I automated dataset cycling instead of repeatedly pasting data or pasting algorithm blocks
- I moved from Google Sheets to **Excel VBA + Python orchestration**
- I introduced **multi-day testing**, stability metrics, and statistical confidence measures
- I experimented with third-party hyperparameter tools to replace manual pattern analysis
- I replaced third-party hyperparameter tools with a **custom deterministic optimizer**
- I expanded from single-algorithm optimization to **group-level synergy system**

What began as a small exploratory spreadsheet testing simple 2 day profit or loss patterns became a **large-scale, modular research system** focused on identifying strategies that stay reliable under **near-term market changes**, not just ones that look good in a single window.

---

## System Overview

At a high level, the system operates as a multi-stage pipeline:

Data Collection
→ Algorithm Variation Testing
→ Stability & Metric Evaluation
→ Parameter Optimization
→ Pruning & Filtering
→ Group Synergy Discovery
→ Execution on New Data and Trade Management

Each stage feeds directly into the next, forming a **closed-loop optimization cycle**. After a group produces daily candidates, a separate execution step applies risk filters and manages positions using my custom trade management logic (TSH, Trailing Stop High), designed to keep exits systematic and conservative for guarenteed profits.

---

## Repository Structure

### `README.md`
This file.  
Provides an introduction and guide on how to explore the project.

---

### `docs/`
Detailed technical explanations for each major subsystem.

- `00_overview.md` — Conceptual overview of the full system
- `01_system_architecture.md` — Structural design and module interactions
- `02_data_pipeline_mp0.md` — Ticker collection, dataset construction, and filtering
- `03_variation_testing_mp1.md` — Multi-variation algorithm testing framework
- `04_metric_design.md` — Performance, stability, and reliability metrics
- `05_optimizer_logic_mp1_p2.md` — Custom ML-inspired parameter optimizer
- `06_pruning_and_stability.md` — Filtering logic for repeatable performance
- `07_group_synergy_mp2.md` — Recursive group combination testing
- `08_threading_and_scaling.md` — Multi-threading, time budgeting, and safeguards
- `09_execution_and_outputs.md` — Final execution and output generation
- `10_limitations_and_future.md` — Known limitations and future directions

These documents explain design decisions, tradeoffs, and failure modes for each subsystem, not just implementation details.

---

### `src/`
Live Python source code for the system.

Organized by module:

- `mp0_data_pipeline/` — Dataset creation and filtering
- `mp1_variation_testing/` — Excel-driven algorithm testing
- `mp1_optimizer/` — Parameter optimization logic
- `mp1_pruning/` — Stability and consistency filters
- `mp2_grouping/` — Recursive group synergy discovery
- `mp2_execution/` — Final group execution
- `trade_execution/` — Trade Management and Strategy
- `utils/` — Shared utilities and helpers

---

### `diagrams/`
System visualizations used to explain flow and architecture.

Examples:
- Overall system flow
- Data pipeline
- Optimizer feedback loop
- Recursive group construction

---

### `logs_examples/`
Representative logs showing:
- Threading behavior
- Optimization progress
- Runtime constraints and safeguards

These demonstrate how the system behaves under real workloads.

---

### `iteration_history/`
Documentation and artifacts from earlier versions.

Includes:
- Early design notes
- Google Sheets macros
- Excel VBA automation

This folder exists to show **iteration, constraint-driven design, and evolution over time**.

---

### `attribution.md`
Details external tools, libraries, and how AI assistance was incorporated under direct human control, along with clarification of authorship.

---

### `LICENSE`
License information for the project.

---
