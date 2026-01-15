# Limitations and Future

## Purpose

This file explains **current limitations** and **planned upgrades** for the project.
The goal is to show what constrained design choices, early limitations, what risks remain, and what the next engineering steps are for this project.

---

## Limitations

### 1) Starting knowledge and lack of mentorship

I began this project in **9th grade** without prior background in
- algorithm design
- trading strategy mechanics
- combinatorics and probability
- statistics and optimization
- ML workflow concepts
- Python engineering practices
- structured project development

I also had **no formal mentorship** while building the system.
Most learning came from self-study, experimentation, and iteration.
Most progress came from repeated trial, failure, refactoring, and rebuilding.

**What this led to**
- Some parts of the system are more primitive than they would be with earlier exposure to professional tooling and software architecture
- Some redundancy exists because major components were rebuilt multiple times as my understanding deepened, and stability was prioritized over elegance
- Some edge-case risks may still exist because there was no senior engineer or professional reviewing decisions and implementation early

---

### 2) Tooling and workflow constraints

Early development used simple tools, including coding in basic editors like Notepad, and running it as a python file.
Version control and structured testing were introduced later, after the core system had already grown in complexity.

**Impact on the system**
- Fewer reliable checkpoints for “known good” versions early on
- More time spent rebuilding when something broke
- Slower debugging progress compared to a full IDE plus structured tests from day one

---

### 3) Legacy platform cost

A major portion of early development used Google Sheets as a core environment before shifting toward a more programmatic pipeline using Excel and Notepad python scripts.
This shift was due to realizing Google Sheets' limited capabilities for a early version of the system I had at that point, forcing me to move.

**Practical impact**
- A large chunk of time produced learning and prototypes in Google Sheets, but not reusable long-term infrastructure that is still being used today
- Some earlier progress does not translate directly into today’s architecture, even though it shaped later design decisions

---

### 4) Compute and hardware constraints

During development, hardware limitations affected design choices. I am using legacy hardware with the PC I currently use to run my programs, requiring me to use batching
and smaller number of threads, and RAM considerations when moving large maps of parameters and group testing in mp2_p2. 

**Impact on the system**
- Forced batching strategies and tighter memory usage
- Limited practical parallelism during heavy runs
- Increased runtime made full exploration of larger parameter and group spaces unrealistic in earlier stages

---

### 5) Excel dependency and interpretability tradeoffs

Excel provided rapid iteration and visibility, but it also created limits.

**Why this matters**
- Excel is not a high-performance execution engine for large-scale experimentation
- Recalculation overhead and workbook structure constraints affect scaling
- Some complexity exists mainly to work around spreadsheet constraints rather than because the trading logic requires it
- Managing hundreds of interdependent formulas across replicated algorithm blocks caused frequent reference errors and significant iteration slowdowns, amplified by hardware limits
- Any fixes or changes to any part of my algorithm would result in me having to manually change each formula affected in all 8 of the variations in the algorithm
  block (since my algorithm block has 8 identical copies of the same algorithm, with different references to a different set of parameters, but each seperate). 
- **No reliable AutoSave / version history for workbooks:** I still don’t have consistent Excel AutoSave because enabling it would require storing the workbook in an externally managed OneDrive account (not under my control). As a result, workbook changes rely on manual saves and local backups by making copies of the workbook constantly, which due to a lack of experience, I often didn't have a backup, and multiple times lost progress and slowed iteration down.

---

### 6) Data coverage limits

The system relies on public market data and a defined dataset window.

**Why this matters**
- Results are only as good as the representativeness of the tested period
- Market conditions change, so older results can decay
- Some tickers appear or disappear over time, which complicates perfectly uniform historical testing when making my ticker universe
- Some issues with NBBO (National Best Bid and Offer) and buying shares right at market open resulted in different prices

---

### 7) Qualitative layer is manual today

The qualitative screen uses manual prompts for Modules A B C D.

**Why this matters**
- Manual review adds judgment and reduces automation risk
- Manual review introduces the possibility of ticker input errors or mismatch of prompts, or handling the process of compiling the scores.
- Manual review does not scale well to hundreds of tickers without automation

This layer is intentionally manual at this stage to preserve judgment and accountability. Automation is planned once reliable data sources and scoring standards are finalized.

---

### 8) Real world execution risk

Live trading introduces risks not visible in backtests.

Examples
- slippage and spread changes
- liquidity drop-offs near open and close
- halts and news shocks
- API and infrastructure failures
- NBBO open price is different than recorded in standard market data providers

**Why this matters**
- A strategy that looks strong on paper can still fail without robust execution controls
- The project includes explicit safety logic and run logging, but execution is always a separate issue and can be volatile

---

### 9) Use of assistive tools (including AI)

This project began in 9th grade, before I had much experience in software engineering, optimization, or large-scale system design.
I had ideas about what I wanted to build as I experimented and learned, but limited technical background to implement them efficiently.

During later stages of development, assistive tools (including AI-based tools) were used selectively to:
- help reason through implementation approaches and learn new Python libraries and concepts,
- sanity-check logic while debugging,
- and speed up refactoring, formatting, or repetitive coding tasks.

All system design, algorithm choices, architectural structure, and execution flow were decided by the author.
Assistive outputs were used as reference material and were rewritten, tested, or discarded during implementation.

Even with assistive tools, development remained slow and difficult due to lack of proper training.
Debugging across Excel, Python, SQLite, and multi-threaded execution required manual iteration, testing, and long-run validation.

Assistive tools accelerated learning and iteration, but did not replace learning or testing; all core logic was implemented, tested, and revised by hand.

---

## Future upgrades

### 1) Full rebuild into a black box architecture in Python

Planned change
- replace spreadsheet-coupled logic with **pure Python black box functions**
- each algorithm becomes a single callable unit

Definition
- input is a dataset window such as 31 days of OHLCV and a parameter vector
- output is a decision and a structured result summary such as buy list plus metrics

**Why this matters**
- removes Excel overhead
- enables faster experimentation
- improves testability and reproducibility
- makes it easier to add new algorithms and compare them under one interface

---

### 2) Professional engineering workflow

Planned change
- migrate development into a full IDE workflow
- enforce version control
- add unit tests and regression tests for critical components
- add experiment logging standards and health checks

**Why this matters**
- faster iteration without breaking known-good behavior
- safer refactors
- easier collaboration and review

---

### 3) Parameter research engine using surrogate modeling

Goal
- explore far more of the parameter space than brute force allows

Core idea
- run a large number of real parameter trials
- train models that approximate performance as a function of parameter settings and environment features
- use those models to evaluate **millions of candidate parameter sets virtually**
- backtest only the top predicted candidates for confirmation

**Why this matters**
- brute force scales exponentially
- surrogate models convert expensive exploration into cheap ranking
- the system can refresh and adapt as market behavior shifts

---

### 4) Group synergy engine at scale

Planned change
- store each variation’s ticker outputs and metrics as reusable objects
- compute group intersections and group metrics without rerunning the full algorithm logic
- use sampling and scoring to avoid combinatorial explosion
- train a group-level surrogate to rank candidate groups before confirmation

**Why this matters**
- group search grows too fast to brute force
- this keeps the search focused on promising regions
- it turns group selection from guesswork into a measured engineering process

---

### 5) Automated qualitative data ingestion and scoring

Planned change
- programmatically gather structured public signals used by Modules A B C D
- store per-ticker metadata and red-flag fields
- extract module scores automatically once scoring reliability is validated

Guardrails
- keep hard red flags as first-class veto conditions
- record sources and timestamps
- preserve an audit trail of why a ticker was accepted or rejected

**Why this matters**
- scales qualitative screening without losing traceability
- improves consistency across days
- reduces the chance of narrative contamination between tickers

---

### 6) Rolling-window updating and decay control

Planned change
- maintain a rolling history window such as 30 to 60 trading days
- add new day results
- drop older days
- retrain and refresh summaries on a regular cadence

**Why this matters**
- prevents stale regimes from dominating
- keeps the system aligned with current market structure
- supports daily adaptation without retesting everything from scratch

---

### 7) Cleaner daily workflow

Target end state
- Program 1 collects market data and relevant metadata
- Program 2 updates parameter knowledge and outputs best per-algorithm candidates
- Program 3 updates group knowledge and outputs best group candidates
- Program 4 selects the final group for the next session using stored knowledge plus the newest day update

**Why this matters**
- daily selection becomes a fast decision step
- heavy testing becomes a background research cycle run on a schedule
- the live system benefits from accumulated evidence instead of repeating the same tests
