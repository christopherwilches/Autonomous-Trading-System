# Logs Examples

## Purpose
This folder contains **real console output** from long-running pipeline executions.  
It is here to show that the code runs, because the programs cannot be executed directly from GitHub.

Execution is PC-bound because the runtime depends on:
- local Excel workbooks and datasets (not included),
- installed dependencies and drivers,
- and external integrations (Alpaca API keys and URLs).

These logs demonstrate:
- the pipeline’s runtime behavior,
- progress tracking during long durations,
- and the final summary metrics produced by each stage like outcomes or time it ran for.

---

## Excerpts

### 1) MP1 Loop (Part 1 + Part 2) — `mp1_p1_p2_loop_excerpt.txt`

This file shows an example execution of the MP1 control loop configured for **200 cycles**.

Each cycle consists of two Python-controlled stages:

#### MP1 Part 1 — Dataset execution (Python → Excel)
- Python selects a standardized 5-day dataset window and iterates through it day by day.
- For each day, Python:
  - sets run configuration values,
  - triggers the Excel compute macro,
  - blocks until completion,
  - and verifies completion through checkpoint signals.
- Each day is processed in multiple batches, with debug confirmations after each batch to ensure state reset and result capture before advancing.
- The loop continues deterministically across all days before marking the cycle complete.

Excel acts purely as a compute backend for the algorithms. Python controls execution order, coordination, and validation.

#### MP1 Part 2 — Parameter update (Python-only)
- After Part 1 completes, Python reads the scored results from storage.
- A new parameter set is selected for each variation slot and written back for the next cycle.

The log excerpt shows:
- the beginning of the run (initial cycles),
- one full representative cycle,
- and the end of the run (final cycles and timing summary).

---

### 2) MP1 Part 3 Export — `mp1_p3_excerpt.txt`
  
It shows a complete log from MP1 Part 3.

**What it shows**
- Reads variation records from the results database.
- Applies gating / selection diagnostics per strategy (stability and consistency-style filters).
- Caps or selects the final set of variations per algorithm.
- Exports the selected variations into the combination workbook format used downstream.

**Significance**
- Demonstrates that the system is not only generating results, but also applying structured selection logic before group testing.
- The output includes per-algorithm counts and an end-to-end export summary.

---

### 3) MP2 Part 2 Group Tester — `mp2_p2_threaded_weekrun.txt`
**What it is**  
A representative log from the threaded group synergy search.

**What it shows**
- Startup configuration with baseline profit percentage weighing for future scoring.
- Thread scheduling across an exhaustive starter set.
- Per-thread summaries including:
  - timeouts,
  - attempt counts,
  - and adaptive runtime limit adjustments.
- Heartbeat checkpoints (`HB`) to track:
  - progress through the starter list,
  - runtime since start,
  - and system health (RAM usage) during multi-hour execution.
- Final selection and summary statistics:
  - best group selection for size 3 and size 4,
  - algorithm usage distributions,
  - per-algorithm timing and average performance metrics,
  - and final runtime totals.

**Meaning of `Total groups evaluated (counter)`**
- This counter represents the number of candidate group evaluations performed by the search process (candidate expansions / intersections), not the number of exported “final groups.”
- The final exported set is intentionally bounded (best-by-size selection) and final groups selected use bell-curve-like logic to choose.

---

## Notes on redaction and formatting

- File paths, timestamps, and environment-specific identifiers are replaced with placeholders such as `<REDACTED_PATH>` and `<TIMESTAMP_REDACTED>`.
- `[...]` indicates omitted sections that repeat the same execution pattern. It replaces the hundreds or thousands of the same logs to show the most important parts.
- Start and end segments are retained to show initialization, steady-state execution, and final timing summaries.
