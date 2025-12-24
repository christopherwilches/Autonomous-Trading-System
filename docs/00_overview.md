# System Overview

This project is an experimental research system designed to study how algorithmic trading strategies behave over time, rather than how they perform in a single, isolated test.

Instead of attempting to “predict” markets, the system evaluates strategies under repeated testing on recent historical market data, applying the results to live conditions afterward.

---

## Core Idea

Most algorithmic strategies fail not because they never work, but because they are overfit to a narrow window of data. A strategy that performs well on one day or one dataset often breaks down when market behavior shifts.

This system is built around three core principles:

- **Multi-day evaluation** over short-term snapshots  
- **Stability and consistency** over maximum return  
- **Iterative refinement** rather than one-time optimization  

Rather than selecting a single “best” algorithm, the system tests many variations, observes their behavior across multiple days, and gradually filters out approaches that are unstable, inconsistent, underperforming, or overly sensitive to noise.

---

## System Operation

The system operates as a research pipeline rather than a static model.

At each stage:
- Strategies are tested under the same market conditions
- Performance is evaluated using both return and stability metrics
- Weak candidates are removed early
- Strong candidates are refined or combined

Instead of treating algorithms independently, the system explores **group-level behavior**, testing whether combinations of strategies reinforce or cancel each other out. This allows it to identify sets of algorithms that perform more consistently together than any single strategy alone.

The system continuously learns from its own results. Each cycle feeds into the next, allowing improvements to compound gradually during each training week rather than relying on large, risky changes.

---

## Practical Usage

In practice, the system runs on a weekly cycle.

Using the most recent five trading days, it:
- Collects and filters market data
- Tests algorithm variations
- Optimizes parameters
- Identifies stable strategy groups

Those selected groups are then applied during the following trading week to generate daily buy decisions. Execution is handled by a separate trade-management layer built around a custom strategy I designed (TSH), which manages entries, exits, and position sizing conservatively.

This structure allows the system to adapt to recent market behavior while avoiding overreaction to short-term noise.

---

## What This System Is Not

This project is not designed to:
- Predict exact price movements
- Chase maximum short-term profit
- Optimize on hindsight alone
- Eliminate risk entirely

Instead, it is a system that evaluates itself holistically, navigates uncertainty, and improves through controlled iteration.

The goal is not certainty, but repeatable and predicatble performance in a short span of time.
