# Excel Cycling Compute Macro — (VBA)

## What this component is
This is the Excel-based compute engine used to execute all strategy logic at scale.  
Although located under `legacy_and_iteration` for organizational reasons, this macro is **still actively used** in the current system.

Its purpose is to let a single algorithm block per sheet evaluate thousands of datasets sequentially, while snapshotting each variation's output into structured result tables.

This was a key discovery early on that made large-scale testing feasible inside Excel without having to duplicate the algorithms and leave them static.

---

## Why this exists
Early versions duplicated entire strategy blocks for each ticker or dataset, which:
- Hit Excel size limits quickly
- Became fragile and slow
- Made iteration impractical

The realization was that **only the output matters**, not static copies of formulas.

This macro executes that idea by using:
- One live input region
- One recalculation per dataset
- Deterministic extraction of results
- Repeatable coordination with Python

---

## What it does

### 1) Dataset cycling
- Assumes a pre-staged vertical stack of datasets in the ControlSheet, each following the same fixed row format.
- Uses helper cells to determine:
  - how many datasets exist,
  - how many to process per run,
  - and where results should be appended.
- For each iteration:
  - The next dataset block is copied into a fixed live input region.
  - All strategy sheets reference this live region, so recalculation updates every algorithm simultaneously.
  - A single full workbook recalculation is forced.

### 2) Per-strategy output capture
- Each strategy sheet has a fixed internal layout for its signals and metrics.
- After recalculation, the macro:
  - Copies specific output fields from each strategy sheet
  - Pastes them as **values** into a standardized result table
  - Advances the destination row so results are appended, not overwritten
- Supports variations with different internal structures by using explicit, per-sheet copy mappings.

### 3) External orchestration handshake
- Reads run configuration (batch size, result start location, iteration count) from helper cells.
- Writes start and completion timestamps to helper cells.
- This allows an external controller (in this case Python) to:
  - Trigger execution
  - Block until completion
  - Safely advance to the next stage

### 4) State safety and recovery
- Temporarily disables screen updates, events, and automatic calculation for speed.
- Restores all Excel state on exit, including error paths.
- Restores the original dataset after cycling completes so the workbook remains usable.
- Disabling screen updates and event handling reduced full-cycle runtime by over an order of magnitude,
  making large-scale multi-variation testing practical instead of prohibitively slow.

## Significance in the overall system
This macro is the **execution backbone** for MP1 Part 1:
- Python controls **when** and **how much** to run
- Excel computes **what happens per dataset**
- SQLite saves **what was observed**

This separation made it possible to:
- Scale to thousands of tickers
- Keep Excel stable
- Treat the workbook as a deterministic compute module instead of a UI tool that constantly crashed and took a long time to finish

---

## Additional Excel orchestration (context only)

The workbook contains additional helper macros used for orchestration and maintenance
(for example: execution triggers, state resets, and forward-execution alignment during testing).
These macros support the pipeline but do not implement crucial logic like the cycling macro.

The algorithm formulas themselves are intentionally excluded
This section details execution flow, result capture, and coordination with external controllers of important macros
