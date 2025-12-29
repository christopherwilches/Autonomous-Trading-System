# Legacy Snapshot — Early MP1 (Excel Cycle + Optuna Parameter Proposals)

## What this version is
This snapshot is an early attempt at automating weekly variation testing using:
- **Excel as the compute engine** (VBA cycling macro + helper cells)
- **SQLite** as the dataset/results store
- **Python** as the orchestrator
- **Optuna** as a parameter proposal mechanism (one trial per algorithm per run)

This file combines “MP1 Part 1” (cycling / batch execution) and “MP1 Part 2” (parameter proposal + logging).

Note: paths, dataset table names, and parameter ranges are redacted in this version.

## What it does
### MP1 Part 1 — Cycling / batch runner (run_cycling_program)
- Reads **total stocks** and **batch size** from helper cells in `ControlSheet`.
- Pulls a batch from SQLite, pastes it into Excel, and triggers the cycling macro.
- Waits until a completion cell changes, then advances the result start row and repeats.

This was the first stable “DB → Excel → macro → results block” loop.

### MP1 Part 2 — Early optimizer loop (run_optuna_optimization)
- Creates/loads an Optuna study per algorithm.
- Proposes parameters for 8 variations, writes them into the Excel parameter block,
  reads back the profit percentage cell, and logs `(params, pp)` into SQLite.
- Includes a simple duplicate-checker to avoid repeating identical parameter sets.

In this snapshot, parameter ranges are omitted for privacy, so the optimizer portion is intentionally non-runnable without the private configuration and environment in my server.

## How this differs from the current system
Compared to the current architecture, this version was rougher in a few key ways:
- **Single-file design**: All parts existed in same script, instead of being seperated
- **Minimal metrics**: stored mostly `(params, profit_percentage)` rather than multi-day stability metrics, buy-counts, overlap, confidence gates, etc.
- **Optuna use was shallow**: one trial at a time, and was overall weaker than the current optimization engine.
- **Group-optimization** didn't exist yet, and was simply finding the best single algorithm variation to use for stock-picking
- Many of the parts that currently exist weren't implemented yet, including the automated trader program or data fetcher
- In this iteration, data was collected using Google Sheets and use of Google Finance formulas, and duplicating formulas to fetch data

The core idea however, is the same:  
**Python orchestrates --> Excel computes at scale --> results get stored --> next parameters are generated.**
