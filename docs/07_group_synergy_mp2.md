# Group Synergy

This part explains the **group-building + synergy scoring** layer of MP2 Part 2.

Scope:
- Recursive group construction (how algorithms are added)
- Consensus intersection logic (what a “shared ticker” means)
- Group evaluation metrics (PP, stability, conservative bounds, coverage, overfit control)
- The core guardrails that control repetition, stability, and minimum usable signal

---

## Definitions

### Variation
A single algorithm configuration exported by MP1 Part 3 into `pruned_*` tables.
Each variation includes five days of historical buy decisions with associated profit outcomes.

### Group
A group is a combination of distinct algorithms, where each algorithm contributes exactly one variation.

Group size is intentionally limited to keep combinations interpretable and the search space manageable.

### Consensus tickers by day
For each day, the group tracks the set of tickers that all current members bought independently.

This is the “shared ticker” definition used for selection and metrics.

---

## Group Evaluation (Size ≥ 2)

A group is evaluated once it has at least 2 members and at least one non-empty consensus day.

### Daily buys / hits
For each day, the group records how many shared tickers were selected and how many of those resulted in a profit.
A shared pick is considered successful if any member achieved a profit on that ticker.

This choice makes the group:
- strict for selection (must be bought by all)
- flexible for win credit (any member’s win counts)

Daily results are then aggregated into a week-level performance view, combining all shared picks across the five days.

This pooled view provides a more reliable signal than any single day, especially when daily overlap sizes vary.

---

## Context Weighting: Day Difficulty

The system builds `DAY_DIFFICULTY_WEIGHTS[1..5]` from baseline win rates in `DS_DAY1..DS_DAY5`.
It is meant to give priority to days that had lower overall profit rate, meaning the group that can perform better those days is more favorable

Not all trading days are equally difficult.  
Some days naturally produce higher win rates across the market, while others are more challenging.

To account for this, group performance is evaluated with awareness of day difficulty:
- stronger performance on harder days is weighted more favorably
- performance on unusually easy days is naturally tempered

This prevents score inflation caused by favorable market conditions and rewards robustness across regimes.

Purpose:
- prevent score inflation on “easy” days
- reward groups that perform on “harder” days

---

## Stability and Robustness Metrics

Groups are evaluated not just on average performance, but on how consistently that performance holds across the week.

Key stability signals include:
- how tightly daily results cluster around a central tendency
- resistance to sharp swings between strong and weak days
- whether performance decays or improves over time

In addition, results are adjusted conservatively to account for sample size.
Groups with very few shared trades are treated with caution, even if their raw performance looks strong.

---

## Overfitting Control: Repetition

A key failure mode in group construction is over-reliance on a small set of recurring names.

Groups are therefore penalized if the same tickers dominate consensus selections across multiple days.
Healthy groups demonstrate breadth over time, rather than repeatedly exploiting a narrow subset of symbols.

- 1.0 means consensus picks are mostly unique across the week
- higher values mean the same names repeat across days (over-concentration)
- the score penalizes repeat_factor above a threshold (quadratic)

---

## Coverage Control

Coverage is used to ensure that consensus selections are actually tradable.

The system discourages:
- extremely small overlaps that produce unreliable signals
- overly large overlaps that dilute the idea of consensus

Instead, it favors a consistent, manageable number of shared picks that can realistically be acted upon.

---

## Guardrails that shape the search

Only a small number of hard constraints affect group construction:
- groups must maintain shared selections
- group size is limited to preserve interpretability and tractability

All other signals influence ranking rather than exploration, ensuring promising combinations are not prematurely excluded.

---

### Retention logic

Groups are retained only after exploration is complete.

Rather than pruning during construction, the system:
- explores broadly
- records high-quality candidates
- then applies stability-first selection

Final candidates are chosen from the strongest-performing groups of each size, with preference given to those that combine:
- consistent week-level performance
- conservative confidence bounds
- stable behavior across days

This separation ensures promising combinations are not prematurely excluded.

Key property:
- retention rules do not affect recursion coverage
- recursion explores, selection filters

---

## Summary

Group Synergy focuses on the core logic of MP2 P2, including:
- a recursive intersection engine that enforces “all members/algos agree” on selection
- a scoring system that rewards:
  - hard-day performance,
  - conservative pooled confidence,
  - week-level stability,
  - usable coverage,
- and penalizes:
  - volatility, decay, buy-count instability, and repeated-name overfitting.
