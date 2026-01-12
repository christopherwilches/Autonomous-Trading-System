# System Overview

This project is a testing and evaluation system designed to study how algorithmic trading algorithms perform across repeated short-term market conditions, rather than in a single isolated backtest.

Instead of attempting broad market predictions, the system evaluates different algorithms under repeated testing on recent historical market data, applying the results to live conditions in the following period.

---

## Core Idea

Most algorithmic strategies fail not because they never work, but because they are overfit to a narrow window of data. A strategy that performs well on one day or one dataset often breaks down when market behavior shifts.

This system is built around three core principles:

- **Multi-day evaluation** over short-term market snapshots  
- **Stability and consistency** over maximum return  
- **Iterative filtering and refinement** rather than one-time optimization  

Rather than selecting a single “best” algorithm, the system tests many variations, measures their behavior across multiple days, and gradually filters out approaches that are unstable, inconsistent, underperforming, or overly sensitive to noise.

---

## Pipeline Structure

The system operates as a structured evaluation pipeline rather than a static model.

At each stage:
- Strategies are tested under the same market conditions
- Performance is evaluated using both return and stability metrics
- Weak candidates are removed early
- Strong candidates are refined or combined

Instead of treating algorithms independently, the system explores **group-level behavior**, testing whether combinations of algorithms reinforce or cancel each other out. This allows it to identify sets of algorithms that perform more consistently together than any single strategy alone.

Results from each evaluation pass determine which algorithms are kepy, modified, or removed in later parts. Each cycle feeds into the next, allowing changes to be introduced gradually on a weekly basis

---

## Algorithm families

The system does not rely on a single strategy type.  
Instead, it evaluates and combines multiple algorithm families, each designed to capture a different market behavior to make a combined, quantiative decision on which tickers to buy.

The current implementation includes the follwing 10 algorithms:

- **Momentum and trend-following algorithms**  
  Designed to capture continuation when price strength persists under sufficient liquidity  
  (e.g., EMA, MACD, EMA–MACD).

- **Mean-reversion and range-based algorithms**  
  Target temporary price dislocations that statistically revert over short horizons  
  (e.g., RSI, Bollinger Bands, RSI–Bollinger combinations).

- **Breakout and volatility-sensitive algorithms**  
  Detect regime shifts, expansions, or compressions where price behavior changes rapidly  
  (e.g., Breakout, Volatility-based strategies).

- **Smoothing and baseline filters**  
  Used to reduce noise and stabilize signals across different market conditions  
  (e.g., SMA-based logic).

- **Trend strength and regime-filtering algorithms**  
  Measure whether the market is trending or range-bound, allowing other strategies to activate or stand down accordingly  
  (e.g., ADX).

Each algorithm is intentionally simple in isolation, focusing on one aspect of a stock. Each algorithm is made up hundreds of individual formulas that contribute to a final buy or no-buy decision.

All algorithms are implemented in Excel and tuned through custom parameter tables.

---

## Practical Usage

In practice, the system runs on a weekly evaluation schedule.

Using the most recent five trading days, it:
- Collects and filters market data
- Tests algorithm variations
- Optimizes parameters
- Identifies stable strategy groups

Those selected groups are then applied during the following trading week to generate daily buy decisions. Execution is handled by a separate trade-management layer (TSH) that controls entries, exits, and position sizing using conservative rules.

This structure allows the system to adapt to recent market behavior while avoiding overreaction to short-term noise.

---

## What This System Is Not

This project is not designed to:
- Predict exact price movements
- Chase maximum short-term profit
- Optimize on hindsight alone
- Eliminate risk entirely

Instead, it is a system that evaluates its own outputs and improves through controlled iteration.

The goal is not certainty, but repeatable and predicatble performance in a short periods of time.
