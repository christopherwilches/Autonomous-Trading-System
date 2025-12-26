# 04_metric_design.md

This section explains **why each metric exists**, what it measures, and which failure mode it is designed to prevent.

mp1 does not optimize for raw profit percentage alone.
Instead, it evaluates **evidence strength**, **stability**, and **generalization risk** across five independent day windows.

All metrics described here are computed after day 5 in mp1 P1 and stored in SQLite.
They are later used by the optimizer (mp1 P2) and by the pruning stage (mp1 P3) to make score formula to rank the variations. Then those same metrics are recalculated in mp2 P2 and used in another score formula.

---

## Core Idea
Markets are volatile, and a strategy that works on one day may fail the next even if nothing “breaks.”

This metric system is therefore designed to:
- Reward **consistency across days**
- Penalize **small-sample luck**
- Penalize **unstable or fragile behavior**
- Detect **overfitting via ticker concentration**
- Preserve **recency awareness** without overprioritizing

No single metric is trusted by it self.

---

## Metric families

### 1. Volume of evidence (how much data exists)
These metrics account for the question: **Is there enough evidence to trust this result?**

#### pooled_buys
Total number of BUY decisions across all five days.

Why it is important:
- A strategy that trades twice and hits (profit) both is not comparable to one that trades 80 times. 
- Low trade counts inflate apparent performance.

Usage:
- Acts as a hard floor for export and pruning gates.
- Used implicitly by Wilson lower bound.

---

### 2. Raw performance summaries
These describe central tendency but are **never trusted alone**.

#### pooled_pp
Overall hit rate across all days: pooled_pp = 100 * pooled_hits / pooled_buys

Why it exists:
- Simple, easy to use performance signal.
- Useful for sanity checks and ranking after confidence is established.

Limitation:
- Extremely sensitive to small sample sizes.
- Never used as a primary decision metric.

---

### 3. Robust daily stability metrics
These answer: **Does this strategy behave similarly day to day?**

#### median_pp_5d
Median of daily profit percentages across five days.

Why median:
- Robust to outlier days.
- Prevents a single extreme win or loss from dominating.

#### pp_mad_5d (median absolute deviation)
Measures how far daily results deviate from the median.

Why it exists:
- Penalizes volatility in outcomes.
- A strategy that alternates between +90% and −10% is unstable.
- Ideally one that sticks to the same profit percentage (pp) throughout the training week is more stable

#### pp_iqr_5d
Interquartile range of daily profit percentages.

Why it exists:
- Another robust dispersion measure.
- Catches asymmetric instability that MAD may miss.

Usage:
- These metrics jointly penalize strategies that “work sometimes” or "up and down".

---

### 4. Buy behavior stability metrics
These answer: **Does the algorithm behave consistently in how often it trades?**

#### buycount_med_5d
Median number of buys per day.

#### buycount_mad_5d
Day-to-day variation in buy counts.

#### buycount_cv_5d
Coefficient of variation of daily buy counts.

Why these exist:
- Some strategies only look good because they “skip” bad days entirely.
- High variability in trade frequency often signals regime fragility.

Usage:
- Penalized during scoring.
- Used as pruning gates in mp1 P3.

---

### 5. Confidence-adjusted performance
These answer: **Is this performance statistically believable?**

#### wilson_lb_5d
Wilson score lower bound of hit probability across five days.

Why it exists:
- Adjusts hit rate downward when sample size is small.
- Prevents 100% hit-rate illusions from 3–5 trades.

Why Wilson instead of normal CI:
- Works well for binomial proportions.
- Behaves sensibly near 0% and 100%.

Usage:
- Primary confidence metric.
- Dominant term in lexicographic scoring.
- Hard gate for export and pruning.

---

### 6. Recency-weighted summaries
These answer: **Is the strategy adapting to current conditions?**

#### ew_scheme
Default weights:
[0.50, 0.25, 0.15, 0.07, 0.03]
(Most recent day → oldest)

#### ew_pp_5d
Exponentially weighted profit percentage.

#### ew_hits_5d
Exponentially weighted hit count.

Why these exist:
- Markets drift.
- A strategy that worked four weeks ago but not last week should decay naturally.

Usage:
- Secondary ranking signal.
- Tie-breaker among similar variations.

---

### 7. Ticker concentration and repetition metrics
These answer: **Is this strategy just exploiting a small set of tickers?**

#### repeat_ticker_rate_5d
Fraction of buys that reoccur across multiple days.

#### top_10_ticker_share_5d
Share of total buys accounted for by the top 10 tickers.

Why these exist:
- Overfitting often manifests as repeated exploitation of a few symbols.
- This fails catastrophically when those symbols change behavior.

Usage:
- Penalized during scoring.
- Used as pruning filters in mp1 P3.

---

### 8. Market regime proxies
These describe **what kind of trades** the strategy prefers.

- avg_buy_price_5d
- median_buy_price_5d
- avg_buy_volume_5d

Why they exist:
- Provide context for later grouping (mp2).
- Help detect strategies that silently drift into illiquid or microcap behavior.

They are descriptive, meant for later use.

---

### 9. Last-day anchors
These capture the most recent behavior explicitly.

- last_day_pp
- last_day_buys
- last_day_hits

Why they exist:
- Prevents stale strategies from surviving purely on historical strength.
- Used as sanity checks and weak recency signals.

---

## Gates vs scores
The system deliberately separates:
- **Gates** (hard accept/reject conditions)
- **Scores** (relative ordering among survivors)

Examples of gates:
- Minimum daily hits: So it doesn't go from buying 3 times to 40 in the same week
- Minimum pooled buys: So the overall buy count is high enough to be useful. A group that doesn't buy anything is just as bad as one that buys losses
- Wilson lower bound threshold

Only strategies that pass gates are ever ranked.

---

## Relationship to later stages
- mp1 P2 consumes these metrics to propose new parameters.
- mp1 P3 uses them to prune aggressively.
- mp2 later uses similar principles at the group level.

These metrics are the foundation of the entire system.
