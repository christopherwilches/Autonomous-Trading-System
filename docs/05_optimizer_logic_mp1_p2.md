# 05_optimizer_logic_mp1_p2.md

This stage explains how mp1 P2 selects, ranks, and proposes new parameter sets for each algorithm based on historical performance.

P2 is not a predictive model.
It does not predict prices or learn a parametric function.

Instead, it functions as a **history-aware hyperparameter optimizer** that adapts search behavior using stored trial results from previous weekly runs.

---

## Conceptual framing
P2 has both classical grid search and full Bayesian optimization elements.

It borrows ideas from:
- Sequential model-based optimization
- Multi-objective ranking
- Exploration vs exploitation control

But it remains intentionally lightweight, interpretable, and fully deterministic.

The core idea:
- Past parameter sets that demonstrated **stable, confident performance** define “good regions”
- Future proposals are biased toward those regions, while preserving diversity and exploration

---

## What P2 optimizes (and what it does not)

P2 optimizes:
- Parameter stability across multiple days
- Confidence-adjusted performance
- Robustness to small-sample noise

P2 does not optimize:
- Maximum raw profit
- Single-day outcomes
- Any learned prediction of future prices

---

## Inputs to P2
P2 takes in metrics produced by mp1 P1 and stored in SQLite.

Primary inputs:
- `optuna_10A_data.db`
- One table per algorithm (e.g. `MACD`, `EMA`, `RSI`, etc.)

Each row represents:
- One algorithm
- One variation slot (1..8)
- One exact parameter set
- One anchored 5-day run
- A full set of finalized metrics (described in 04)

---

## Core components

### Parameter history loader
Function:
- `_db_current_params_for_algo()`

Purpose:
- Fetch the most recent parameter sets from the SQLite DB which correspond to the current variations in Excel
- Maintain alignment between:
  - Excel sheet state
  - Database state
  - Internal optimizer representation

This avoids desynchronization between UI-visible parameters and DB records. In case they don't match, this will fix the parameters in the DB.

---

### Metric aggregation and scoring
Function:
- `_fetch_5d_metrics_for()`
- `_lexi_score()`

P2 does not collapse metrics into a naive weighted sum.

Instead, it applies **lexicographic dominance**, where priorities are enforced in order:

1. Hard gates (consistency, minimum evidence)
2. Confidence (`wilson_lb_5d`)
3. Central tendency (`median_pp_5d`)
4. Stability penalties (MAD, IQR, buycount CV)
5. Recency-weighted performance
6. Concentration penalties

This prevents high-variance or low-evidence parameter sets from ranking highly, even if raw profit is strong.

---

## Candidate generation strategy

Function:
- `_propose_new_params_for_algo()`

This step is the core of P2’s adaptive behavior.

It generates new candidates using a mixture of:

### 1. Anchor-based sampling
- Selects recent high-ranking parameter sets as anchors
- Samples numeric parameters using small Gaussian perturbations
- Preserves categorical choices with weighted probability

This biases search toward regions that already demonstrated stability and success.

---

### 2. Ablation-style proposals
- Modifies exactly one parameter at a time
- Keeps all others fixed

Purpose:
- Isolate sensitivity
- Detect whether a parameter actually caused change, or is incidental

---

### 3. Global exploration
- Samples from the full allowed parameter space
- Injects randomness to prevent local minima collapse

This ensures the optimizer does not converge prematurely.

---

## Special parameter handling
Some algorithms use internal representations that do not map 1:1 to Excel cells.

Examples:
- Volatility parameters that merge multiple sub-values
- ADX parameters that switch between distinct numeric regimes

P2 handles this through:
- Internal canonical representations
- Explicit conversion layers:
  - `_to_excel_params()`
  - `_from_excel_params()`

This allows optimization to operate in a clean internal space while respecting workbook constraints.

---

## Deduplication and diversity enforcement

This stage is important, because Excel enforces rounding and discrete steps, different internal samples can collapse into identical workbook parameters.

P2 prevents this via:

### Canonical deduplication
- Dedupes against:
  - Historical DB parameters
  - Current on-sheet parameters

### Near-clone rejection
- Computes normalized L1 distance in parameter space
- Drops candidates below a similarity threshold

### Batch-level diversity
- `_force_batch_diversity()` ensures:
  - Variation slots are not minor perturbations of the same anchor
  - Coverage across key parameter axes

Without this, the optimizer would converge to visually different but functionally identical rows (variations).

---

## Output to Excel

Function:
- `_write_params_to_excel()`

Responsibilities:
- Clear old parameter cells
- Write exactly 8 new rows in workbook order
- Maintain column alignment and UI stability
- Avoid accidental formula shifts or spillover

Excel remains the authoritative execution engine.
P2 only controls what parameters get tested next.

---

## Why this is not classical ML
This stage intentionally avoids:
- Training differentiable models
- Fitting probabilistic surrogate models
- Gradient-based optimization

Reasons:
- The execution engine (Excel) is treated as a black box
- The objective is multi-modal and non-smooth
- Stability and confidence dominate raw optimization

This makes lightweight, rule-driven optimization more appropriate than heavy ML.

---

## Why this is still “learning”
P2 adapts because:
- Search behavior changes based on historical outcomes
- Parameter regions gain or lose probability mass over time
- Poor regions are implicitly abandoned
- Good regions are explored more deeply

This achieves the practical definition of learning without using a predictive model for now.

---

## Relationship to later stages
- mp1 P3 prunes based on these same metrics
- mp2 builds group-level behavior on top of optimized variations
- Daily execution relies on P2 having suggested stable, improving parameter sets
- 
