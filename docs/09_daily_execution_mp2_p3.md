# Daily Execution

This section explains the daily pipeline used to turn the weekend-selected groups into a short list of tickers for the next session. Trade execution happens in the next stage (Real Trader).

Daily execution has two parts: a fully automated candidate generator (MP2 Part 3) and a manual qualitative screen.
1) selects a top-ranked group produced by MP2 Part 2
2) refreshes the daily market snapshot used for evaluation
3) runs MP2 Part 3 to generate and prune a candidate list
4) applies a manual qualitative screen to catch real-world risks
5) combines quantitative rank with qualitative scores to choose the final tickers

---

## Step 0 — Starting state (after MP2 Part 2)

MP2 Part 2 produces a ranked list of candidate groups stored in `extra_big_groups` (EBG).  
MP2 Part 3 uses this table as its starting point and typically selects the highest-scoring eligible group (a test mode can force selection of weaker groups for validation).

---

## Step 1 — Refresh the Daily Dictionary

Before group execution, the system rebuilds a daily snapshot of the tradable universe using the most recent fully closed trading day.

What it does:
- pulls the Alpaca asset universe and filters to clean, tradable common stocks
- fetches recent daily bars and discards any partial “today” bar during market hours
- applies basic price and liquidity screens to remove extreme illiquids and outliers
- formats each ticker into the fixed 53-row block structure the Excel workbook expects

Outputs:
- `DICTIONARY_TABLE` as the main daily universe used by MP2 Part 3
- `SEEDLING_TABLE` as a separate bucket for low-priced tickers kept for reference

---

## Step 2 — MP2 Part 3: pick best group and run the cycle against today’s dictionary

MP2 Part 3 runs the selected group against the latest daily dictionary using the Excel workbook as the compute engine, then extracts consensus tickers across the group.

High-level procedure:

### 2.1 Select a group
MP2 Part 3 reads the ranked group list from EBG and selects the best-scoring eligible group (typically size 3 or 4). It extracts the group identifier and the per-algorithm parameter sets needed for execution.

### 2.2 Load group parameters into Excel
Each algorithm’s parameters are placed into separate variation slots so multiple algorithms can be evaluated in one workbook pass without collisions. This is effectively a temporary mapping layer that keeps the workbook layout stable while testing a multi-algorithm group.

### 2.3 Cycle the daily dictionary through Excel
The daily universe is processed in batches of ticker blocks. For each batch, the workbook is triggered to evaluate the group and output buy decisions per algorithm. A completion signal is used so Python never relies on timing guesses.

### 2.4 Collect buy decisions per algorithm
After each batch, the executor reads the buy outputs for each algorithm’s assigned slot, normalizes ticker/name issues, and deduplicates tickers within each algorithm. Results are aggregated across all batches so each algorithm ends with a clean set of tickers it would buy for the next session.

### 2.5 Compute consensus tickers
The candidate list is formed by intersecting the tickers across all algorithms in the group. If the overlap is too small to be useful, MP2 Part 3 falls back to the next group in the ranked list.

### 2.6 Save and display candidates
Once a usable consensus set is found, candidates are stored in `live_candidates` and printed into the WB2 output area for inspection and downstream pruning. The executor exits after selecting the first acceptable group.

### 2.7 Reset workbook state
At the end of the run, the workbook is reset so the next execution starts clean and does not leak results from the previous cycle.

---

## Step 3 — Historical pruner (rank to a short list)

After MP2 Part 3 generates consensus candidates, a historical pruner reduces the list to a manageable shortlist (typically ~20).

What it evaluates (using recent daily bars):
- consistency of positive closes under conservative definitions
- downside behavior (loss and drawdown sensitivity)
- evidence strength (penalizing tiny samples)
- basic volatility and range sanity checks

The output is a ranked list with a `rank_score` that is later combined with qualitative module scores.

---
## Step 4 — Manual qualitative prompts (Copilot/OpenAI)

This step exists to **qualitatively stress-test** the quant-selected tickers before trading. 

These 4 prompts are meant to gauge all aspects of a qualitative lens of a stock and quantify its performance to bet gauge which stock is best to choose.

---

### Module A — News, events, governance and structure

**What it evaluates**
- Near-term real-world risks and catalysts
- Corporate governance cleanliness
- Legal, regulatory, and structural hazards

**Why it matters**
Quant systems do not understand:
- Surprise dilution
- Lawsuits
- Governance scandals
- Regulatory or geopolitical shocks

It checks for near-term real-world risks (news, legal, governance) that could invalidate a short-horizon trade.

**Key outputs**
- news_events_score
- governance_legal_score
- structural_risk_notes
- hard_red_flags

Hard red flags here can **fully disqualify** a ticker regardless of quant rank.

---

### Module B — Fundamentals, balance sheet, dilution, ownership

**What it evaluates**
- Financial survivability
- Capital structure stress
- Dilution likelihood
- Insider and institutional alignment

**Why it matters**
Short-term trades still fail when:
- Companies are forced to raise capital
- ATMs or shelves activate unexpectedly
- Insiders are aggressively exiting

It checks whether the company’s financial and ownership situation creates near-term risks for a short-horizon trade (especially dilution or funding pressure).

**Key outputs**
- fundamentals_score
- dilution_risk_score
- ownership_score
- hard_red_flags

Repeated dilution or going-concern language can override technical strength.

---

### Module C — Liquidity and microstructure

**What it evaluates**
- Tradability
- Slippage risk
- Halt and gap behavior
- Spread behavior and float dynamics

**Why it matters**
Even correct direction fails when:
- Liquidity disappears
- Spreads explode
- Halts trap positions
- Meme dynamics override rational execution

It checks whether the stock is realistically tradable without excessive slippage, spread risk, halts, or gap behavior that can break execution.

**Key outputs**
- liquidity_microstructure_score
- microstructure_notes
- hard_red_flags

Extremely illiquid or halt-prone stocks are rejected here.

---

### Module D — Social and public sentiment narrative

**What it evaluates**
- Crowd psychology
- Narrative stability
- Meme or scandal risk
- Sentiment volatility

**Why it matters**
Narratives can overpower fundamentals and technicals in the short term.
This module detects:
- Toxic sentiment
- Unstable hype
- Coordinated narrative risk

It checks whether the public narrative is stable enough to avoid sudden sentiment-driven moves that can override technical execution.

**Key outputs**
- social_sentiment_score
- narrative_notes
- hard_red_flags

Severe sentiment risk can disqualify a ticker even if all other modules pass.

---

## Output of step 4

For each ticker:
- Four independent module scores
- Explicit hard red flags
- A defensible qualitative profile

These outputs are **not averaged blindly**.
Hard red flags can override numeric scores.

---

## Role in the system

Step 4 acts as a **qualitative firewall** between:
- Quant signal generation
- Live trade execution

Only tickers that pass both:
- Quant stability and performance
- Qualitative risk inspection

are allowed into Real Trader/chosen to trade. 

This process is currently manual, but with plans to turn automated. More on planned automation is in `11_limitations_and_future.md`.

## Step 5 — Final daily selection

Final selection combines:
- the pruner’s `rank_score`
- module scores from A–D
- any disqualifying red flags (which can override rank)

The output is a small final set (typically ~4 tickers) passed into Real Trader.

---

## Daily routine (in practice)

Each day follows the same pattern:
1) refresh the daily dictionary snapshot
2) run MP2 Part 3 to generate consensus candidates and prune them to a shortlist
3) apply the qualitative screen to filter risks and choose the final tickers
