# 07 — Group Synergy (MP2 Part 2)

This part explains the **group-building + synergy scoring** layer of MP2 Part 2.

Scope:
- Recursive group construction (how algorithms are added)
- Consensus intersection logic (what a “shared ticker” means)
- Group evaluation metrics (PP, stability, conservative bounds, coverage, overfit control)
- The **core guardrails** that control repeat factor and minimum amount of hits (profits) daily

---

## Definitions

### Variation
A single algorithm configuration exported by MP1 Part 3 into `pruned_*` tables.
Each variation includes 5 day buy lists (`day1_buys_json..day5_buys_json`), where each buy row contains:
- ticker, name, open, close
- profit flag is normalized as `is_profit = 1 if close > open else 0`

### Group
A set of **distinct algorithms** (no duplicates), each contributing exactly one variation:
- `group = [(algo_idx, variation_identifier), ...]`

Group size is capped by `MAX_GROUP_SIZE` (in this program: 4).

### Consensus tickers by day
For each day (1-5), the group maintains:
- `consensus_by_day[d] = set of tickers that ALL current members bought on day d`

This is the “shared ticker” definition used for selection and metrics.

---

## Inputs and Precompute Maps in RAM

MP2 loads variations from the 10 `pruned_*` tables, then precomputes fast primitives:

### 1) Day buys set
`day_buys_map[(algo_idx, ident, day)] -> set[ticker]`

Meaning: “tickers bought by this variation on this day”.

### 2) Day profit lookup
`day_profit_map[(algo_idx, ident, day)][ticker] -> 0/1`

Meaning: “if this variation bought ticker T on day D, did it win (close>open)?”.

These maps are the performance-critical foundation of recursion:
- recursion uses set intersections (`&`) and O(1) dict lookups for wins.

---

## Recursive Group Builder (Core Engine)

Entry point per starter variation:
- seed group with one member
- seed consensus sets from that member’s day buys

Recursive function:
`build(group, consensus_by_day, used_algos, hits, deadline=None)`

### Recursion Invariants
- `group` contains unique algorithm indices
- `consensus_by_day[d]` is always the intersection of all member day buy sets for day d
- Adding a member can only shrink (or keep) consensus sets
- Example: if a new algo buys only 4 tickers, the entire group is limited by the tickers they share with the new algo

### Base Termination (Structural)
A branch terminates if there is no consensus left to evaluate:

- If all days are empty consensus:
  - return immediately

This is the primary natural pruning mechanism:
- intersections collapse quickly as group size grows.

### Expansion Rule (Adding Members)
To expand the current group, pick a next algorithm index not already in the group, then iterate its candidate variations.

For each candidate `(next_algo, ident)`:
- For each day `d`, compute:
  - `inter = consensus_by_day[d] ∩ day_buys_map[(next_algo, ident, d)]`
- If every day becomes empty (`has_any == False`), skip this candidate.
- Otherwise:
  - append member
  - recurse with the updated `new_consensus`
  - backtrack

Ordering constraint:
- only add algorithms with index greater than current max in the group
  - guarantees deterministic enumeration
  - avoids permutations of the same combination AKA wasted resources calcualting the same groups

---

## Group Evaluation (Size ≥ 2)

A group is evaluated once it has at least 2 members and at least one non-empty consensus day.

### Daily buys / hits
For each day `d`:

- `buys_d = |consensus_by_day[d]|`

- `hits_d` is computed as:
  - for each ticker `t` in consensus tickers that day:
    - count a win if **ANY member** has profit flag 1 for `(member, day, ticker)`

This choice makes the group:
- strict for selection (must be bought by all)
- flexible for win credit (any member’s win counts)

Outputs:
- `day_buys = [b1..b5]`
- `day_hits = [h1..h5]`
- `daily_pp[d] = 100 * hits_d / buys_d` (0 if buys_d = 0)

### Weekly pooled PP
Let:
- `pooled_buys = sum(day_buys)`
- `pooled_hits = sum(day_hits)`

Then:
- `pooled_pp = 100 * pooled_hits / max(1, pooled_buys)`

---

## Context Weighting: Day Difficulty

The system builds `DAY_DIFFICULTY_WEIGHTS[1..5]` from baseline win rates in `DS_DAY1..DS_DAY5`.
It is meant to give priority to days that had lower overall profit rate, meaning the group that can perform better those days is more favorable

Mechanism:
- compute baseline day win rate (close>open) from the dataset’s “first OHLC row per ticker”
- compare each day’s baseline to the average baseline
- derive weight: `w = 1 + k*(avg_wr - day_wr)`
- clamp weights to a sane band (0.5 .. 1.5)

Used metric:
- `weighted_pp = avg( daily_pp[d] * weight[d] )`

Purpose:
- prevent score inflation on “easy” days
- reward groups that perform on “harder” days

---

## Stability and Robustness Metrics

From `daily_pp` (5 values):
- `median_pp` (middle of sorted daily PP)
- `second_worst_pp` (2nd lowest daily PP)
- `mad_pp = avg(|pp_d - median_pp|)` across 5 days
- `iqr_pp = sorted_pp[3] - sorted_pp[1]`

From daily buy counts:
- `mean_buys = avg(day_buys)`
- `cv_buys = stdev(nonzero_buys) / mean_buys` (0 if insufficient data)

Conservative bound:
- `wilson_lb` = Wilson lower bound of pooled hit-rate (z = 1.96), expressed in percent
  - punishes small-sample “perfect” overlaps
  - boosts large-sample consistent performance

Trend:
- `pp_slope` = linear regression slope of daily PP across days 1..5
  - negative slope penalized (decay)
  - positive slope lightly rewarded

---

## Overfitting Control: Repeat Factor

Let:
- `weekly_set = union of all consensus tickers across days`
- `unique_consensus = |weekly_set|`
- `total_consensus_buys = sum(day_buys)`

Define:
- `repeat_factor = total_consensus_buys / max(1, unique_consensus)`

Interpretation:
- ~1.0 means consensus picks are mostly unique across the week
- higher values mean the same names repeat across days (over-concentration)
- the score penalizes repeat_factor above a threshold (quadratic)

---

## Coverage Control: “Enough picks, not too many”

Coverage term is designed to:
- avoid tiny overlaps that look great but trade nothing
- avoid spray overlaps that stop being “consensus”

Components:
- `coverage_core = 6 * log1p(mean_buys)` (with mild capping)
- `sweet_spot_bonus(mean_buys)` shaped to prefer roughly ~14/day
- `hard_over_penalty` quadratic if any single day exceeds 20 consensus buys

This enforces a usable consensus list size.

---

## Group Score (Synergy Score)

The score combines performance + conservative floor + usability and subtracts instability + overfit.

### Profit base
- `profit_base = 0.60*weighted_pp + 0.25*median_pp + 0.15*second_worst_pp`

### Penalties / boosts
- `trend_penalty = 6 * max(0, -pp_slope)`
- `stability_penalty = 0.25*mad_pp + 0.18*iqr_pp + 0.10*cv_buys*100 + trend_penalty`
- `repeat_penalty = 25 * max(0, repeat_factor - 1.20)^2`
- `trend_boost = 3 * max(0, pp_slope)`

### Final score
- `score = profit_base + 0.80*wilson_lb + coverage_term + trend_boost - stability_penalty - repeat_penalty`

This score is the primary ranking used for Top-K and later final selection.

---

## Guardrails That Matter (Branch Killers)

These are the *meaningful* pruning rules that affect recursion coverage:

1) **No-consensus termination**
- if all 5 consensus sets are empty, the branch ends

2) **Daily hit floor** (size ≥ 2)
- requires `hits_d >= 8` for each day (for d=1..5)
- eliminates weak or luck-based overlaps early

3) **Max group size**
- recursion stops expanding when `size_now >= MAX_GROUP_SIZE`

Everything else is either:
- a score penalty (does not prune),
- or an eligibility filter for Top-K selection (does not prune recursion).
- This is because even small rules that affect recursion (`build()`) can completely change the combinations tested, and can potentially stop exploration of groups with potential.

---

## Candidate Retention (Top-K + Final Pick)

Groups are only *retained* for final consideration when size is 3 or 4.

Retention pipeline:
1) Build the `entry` struct (pp, score, daily_stats, buy_stats, stability metrics, consensus sets)
2) Feed into Top-K accumulator for that size
3) After the run, choose final winners from Top-K using a bell-curve + stability filter

Key property:
- retention rules do not affect recursion coverage
- recursion explores, selection filters

---

## Persistence of Results

During the run:
- The current best-by-size (SIZE3 / SIZE4) can be continuously upserted into `extra_big_groups`

At the end:
- final selected winners are persisted and exported into the workbook output area
- consensus tickers are stored explicitly per day, including profit flags

---

## Summary

MP2 Group Synergy is:
- a recursive intersection engine that enforces “all members/algos agree” on selection
- a scoring system that rewards:
  - hard-day performance,
  - conservative pooled confidence,
  - week-level stability,
  - usable coverage,
- and penalizes:
  - volatility, decay, buy-count instability, and repeated-name overfitting.
