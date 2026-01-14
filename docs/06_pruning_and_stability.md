# Pruning and Stability

This part converts the full MP1 history into a compact, stability-first candidate set for MP2.

Input
- **optuna_10A_data** (contains full 5-day trial rows per algorithm)
- **WB1** (used only to read parameter headers)
Output
- **stocks_data.db**
  - **pruned_<ALGO>** tables (full 5-day metrics)
- **WB2** (ControlSheet receives the pruned variations)

The export focus is consistency first: it aims to remove unstable or junk parameter sets, deduplicate near-clones, and preserve ticker coverage diversity across selected variations.

---
## Per-algorithm export plan
Algorithms are pruned independently to preserve their distinct behavior, then combined under a global selection policy.

Some algorithms are allowed greater representation based on consistent historical robustness, while global limits ensure the final candidate pool remains diverse and tractable for group discovery.

---

## Row normalization
Each trial result is normalized into a compact internal record containing:
- parameter identity
- five-day performance summaries
- stability and dispersion metrics
- ticker-level behavior from the anchor day

Incomplete or malformed records are discarded early to keep the pruning stage reliable.

---

## Minimal eligibility gate
Before scoring, each candidate must meet basic activity and consistency requirements.

This gate removes inactive, degenerate, or extremely unstable runs while remaining intentionally permissive. Fine-grained quality control is handled later through scoring and deduplication.

---

## Unified export scoring
Candidates that pass the eligibility gate are ranked using a stability-first score.

The score prioritizes:
- steady multi-day performance
- consistent trading behavior
- confidence-adjusted outcomes
- low ticker concentration

This ranking is used only to order candidates within an algorithm. Final selection emphasizes diversity and coverage rather than absolute score.

---

## Deduplication logic
After scoring, candidates are filtered to remove near-identical variations.

Deduplication considers:
- similarity of parameter configurations
- overlap in purchased tickers on the anchor day

When two candidates are deemed equivalent, only one is retained. The selection favors broader ticker coverage and stronger overall performance to avoid concentrating the export pool around redundant behavior.

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

This produces a final set with broader exposure for MP2 group synergy tests.

---

## End-state
On completion:
- **stocks_data.db** contains pruned per-algo tables with full 5-day metrics
- **WB2 ControlSheet** displays pruned variations in the standardized layout required by MP2
- the final candidate pool is size-controlled to remain tractable for MP2
