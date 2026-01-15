# Trade management with TSH

## Purpose
This stage handles live position management after entries are placed using the tickers from the daily execution logic.
Its goal is to let trades develop naturally, then lock in a profit using a strict,
deterministic trailing-stop-high model that avoids premature exits and prevents
potential winners from falling back down and losing money.

---

## High-level flow
- Positions are opened near market open using marketable orders sized by notional allocation
- Each symbol is assigned a fixed capital allocation
- A strict TSH model manages exits for the rest of the session
- No discretionary exits exist
- A final end-of-day flatten acts as a safety backstop

---

## Target selection
Each position receives a target percentage based on its entry price bucket.

Lower-priced stocks receive wider targets  
Higher-priced stocks receive tighter targets  

This normalizes typical intraday movement across price ranges. Lower-priced names tend to move in wider percentages; higher-priced names tend to move in tighter percentages.

The selected value is referred to as **x**.

---

## Core terms
- **E**: entry price (actual fill)
- **x**: target percentage (chosen from an entry-price bucket)
- **T**: target price (entry price plus x percent)
- **trailing_pct**: trailing percentage (set to half of x)
- **TA**: trailing amount in dollars (entry price times trailing_pct)
- **H**: highest observed price after arming
- **F**: active trailing floor used for exit decisions

---

## Before arming
While price remains below the target price **T**:
- the system does not exit
- there is no stop-loss or trailing behavior
- the position is allowed to fluctuate freely

Rationale:
This avoids selling on noise. Trade management only activates after the move reaches the target threshold.

---

## Arming condition
The model arms the first time the price reaches or exceeds the target price **T**.

On arming:
- the high watermark **H** is initialized to the current price
- the position transitions into active trailing management

From this point forward, the position is actively managed.

---

## After arming (trailing management)

### High watermark
After arming, the system tracks a running high **H**. If a new high is reached, **H** is updated immediately.

### Trailing floor
The trailing floor **F** is computed from the high watermark and the fixed trailing amount **TA**:

- start with a raw floor of **H minus TA**
- enforce a minimum floor at the target price **T**
- use the higher of the two as **F**

Properties:
- **F** never drops below **T**
- **F** only moves upward as **H** rises
- gains are locked progressively as new highs form

---

## Exit rule
An exit is triggered immediately when the current price falls to the trailing floor **F** (or below).

Action:
- Submit a market sell for the full filled quantity
- Record exit price, exit time, and exit floor
- Mark the trade as closed

---

## Why the trailing amount is fixed dollars
Trailing distance is computed from the **entry price**, not the current price.

This prevents the trailing distance from expanding during strong rallies and keeps allowable giveback proportional to the entry price.

---

## Price feed behavior
- With SIP NBBO:
  - the system tracks the **bid** as the sellable reference price
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
- Lock in gains after a target move is achieved, rather than letting winners round-trip
- Deterministic behavior with no discretionary exits
- No predictive logic inside the execution layer
- Favor repeatable intraday outcomes over maximizing any single trade
