# 10 Trade management with strict TSH

## Purpose
This stage handles live position management after entries are placed using the tickers from the daily execution logic.
Its goal is to let trades develop naturally, then lock in a profit using a strict,
deterministic trailing-stop-high model that avoids premature exits and prevents
potential winners from falling back down and losing money.

This part documents the **behavioral model**, not a predictive system.

---

## High-level flow
- Positions are opened at market open using notional market orders
- Each symbol is assigned a fixed capital allocation
- A strict TSH model manages exits for the rest of the session
- No discretionary exits exist
- A final end-of-day flatten acts as a safety backstop

---

## Target selection
Each position receives a target percentage based on its entry price bucket.

Lower-priced stocks receive wider targets  
Higher-priced stocks receive tighter targets  

This normalizes expected intraday movement across price ranges, accounting for smaller predicted movements for larger stocks.

The selected value is referred to as **x**.

---

## Core variables
- **E** = entry price (actual filled price)
- **x** = target percentage
- **T** = target price = `E * (1 + x)`
- **trailing_pct** = `x / 2`
- **TA** = trailing amount (fixed dollars) = `E * trailing_pct`
- **H** = highest observed price after arming
- **F** = enforced trailing floor

---

## Phase 1: not armed
- If price < **T**
- No exits are allowed
- No stop-loss
- No trailing logic
- Position is allowed to fluctuate freely

Rationale:
This prevents early exits caused by noise and ensures only real momentum
activates trade management.

---

## Arming condition
The model arms the first time: 
price >= T

On arming:
- `triggered = True`
- `H = price`

From this point forward, the position is actively managed.

---

## Phase 2: armed trailing logic
After arming, the model updates continuously.

### High watermark
If: price > H

Then: H = price

### Floor computation
The trailing floor is computed as:

raw_floor = H - TA
F = max(T, raw_floor)

Properties:
- The floor never drops below the target price
- The floor only moves upward
- Gains are locked progressively as new highs form

---

## Exit rule
An exit is triggered immediately when:

price <= F

Action:
- Submit a market sell for the full filled quantity
- Record exit price, exit time, and exit floor
- Mark the trade as closed

---

## Why the trailing amount is fixed dollars
Trailing distance is computed from the **entry price**, not the current price.

This prevents the trailing distance from expanding during strong rallies and
keeps giveback proportional to the original open price.

---

## Price feed behavior
- When SIP NBBO data is used:
  - The model tracks the **bid** price
  - This reflects the realistic sellable price
- If bid is unavailable:
  - The ask is used as a fallback
- When IEX trade data is used:
  - The last trade price drives the model

---

## One-entry constraint
- Each symbol is bought once at market open
- All trailing logic assumes a single entry and a single exit
- Re-entries are not allowed during the session

---

## End-of-day safety flatten
- Near market close, all remaining positions are flattened once
- This prevents unintended overnight exposure
- Any position not exited by TSH is closed here

---

## Logged fields per trade
- Entry time and entry price
- Exit time and exit price
- Target percentage and trailing percentage
- Target price and trailing floor at exit
- Whether TSH was armed
- Intraday high, low, and final price
- Profit percentage

---

## Design intent
- The entire goal of this project to generate guarenteed profits. This custom market strategy
  allows for profits far higher than its predicted percentage, but is built to quickly sell once it starts
  dropping from its predicted high to guarentee profits, not risk and try to maximize profit amount
- Fully deterministic behavior
- No predictive logic inside the execution layer
```
