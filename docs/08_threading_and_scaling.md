# 08 — Threading & Scaling (MP2 Part 2)

This part explains the **execution layer** of MP2 Part 2: how the recursion engine is run at scale with a 4-worker pool, how per-starter time is budgeted, and how the run is kept bounded and observable with logs.

Scope:
- Thread starter design (what is parallelized)
- 4-worker scheduling model (bounded in-flight, backfill)
- Adaptive per-task time budgeting + leftover “donations”
- Global stop conditions (hard end)
- What actually matters for scaling

---

## What is parallelized (Starter-level parallelism)

MP2 runs the recursion builder starting from many “starter” variations.

A **starter** is:
- `(algo_idx, variation_identifier)`

Each starter launches:
- `build([(algo, ident)], seed_consensus, {algo}, hits, deadline=start_time)`

Important: Parallelism is **NOT** inside the recursion.
- Recursion stays single-threaded within each starter.
- Parallelism is across starters, not recursion logic

Why this is the right cut:
- Starters are independent from one another.
- They share read-only maps (`day_buys_map`, `day_profit_map`, `variation_map`, etc.).
- It avoids locks on tight loops inside recursion.
- It makes runtime predictable: each starter has a time budget.

---

## The 4-worker execution model

MP2 uses a `ThreadPoolExecutor` with:
- `max_workers = min(4, total_tasks)` (hard cap at 4 due to number of cores)

Scheduling is **bounded + backfilled**:
1) Submit up to `max_workers` starters initially.
2) Wait for the first completed (`FIRST_COMPLETED`).
3) Submit the next starter to keep the pool full.
4) Repeat until all starters are processed or global stop triggers.

Key property:
- At most 4 starters are running at any moment.
- This prevents “submit floods” (thousands of combos tested) and keeps RAM stable.

---

## Per-starter time budgeting (the only thing that matters)

### Goals
- Each starter gets a time slice.
- If a starter finishes early, unused time can be reallocated to remaining starters.
- The per-starter limit should increase monotonically (no thrashing).
- There is still a hard global stop.

### Core variables
- `TOTAL_WALL_BUDGET`: target wall budget for the run (seconds)
- `_global_start`: wall start time
- `_target_end = _global_start + TOTAL_WALL_BUDGET`
- `_global_end`: hard stop (absolute kill boundary)

This system keeps a per-starter time limit:
- `current_per_task_limit() = _base_per_task_limit + _delta_per_task`

Where:
- `_base_per_task_limit` is computed once from the target budget and the pool geometry
- `_delta_per_task` is raised over time by “donations” (unused time)

### How base_limit is computed
Let:
- `total_tasks = number of starters`
- `max_workers = 4`
- `batches = ceil(total_tasks / max_workers)`

Then:
- `_base_per_task_limit = TOTAL_WALL_BUDGET / batches`

This means:
- If every running starter consumes roughly `base_limit`,
  total wall time stays near the target because 4 run concurrently.

This is the critical insight:
- **Per-task budget scales with the number of batches**, not the number of tasks.
- This allows for consistent times equaling the total limit no matter how many threads there are, or how long a individual takes. 

---

## The donation model (banking leftover time)

When a starter finishes:
- `dur = time.time() - start`
- `limit_at_end = effective_thread_limit()`
- If the thread is considered “done” (not timed out), it donates:
  - `leftover = max(0, limit_at_end - dur)`

That leftover is fed into:
- `donate_leftover(leftover, donor_tag=...)`

Donation behavior in this implementation:
- leftover seconds go into a pooled bank
- the bank is converted into a **per-task bump** spread across remaining tasks
- `_delta_per_task` increases monotonically

Result:
- as “easy” starters finish early, harder starters later get more time
- the system self-tunes without manual per-algo tuning
- The starters are organized by amount of tickers and special algos. Special algos go last, and in both pools it is in order of ticker size
- The logic being that starters with more tickers have the potential for more combinations, and should have more time

### Why it works
- Recursion depth/width varies wildly by starter.
- Some starters collapse quickly (little consensus).
- Others explore deep (persistent consensus overlap).
- Banking converts that variance into better overall coverage without blowing wall time.

---

## Enforcing the budget inside recursion

The recursion function accepts:
- `deadline` (starter start time)

At entry and at critical points:
- `_thread_time_up(deadline)` checks elapsed time against the current per-task limit.

Two key helpers:
- `current_per_task_limit()` — dynamic (base + donations)
- `SOFT_THREAD_CAP_SEC` — clamps the per-task limit so it doesn’t explode

Then:
- if time is up, recursion returns (branch terminates)
- there is no exception-driven unwind required; it’s a cooperative stop

Key property:
- This keeps the recursion from “running away” on a single starter.

---

## Global stop conditions

Even with per-task budgeting, MP2 has an absolute stop boundary.

Global stop triggers:
- `time.time() >= _global_end`
- `exit_event.is_set()`

When triggered:
- recursion functions return quickly
- the executor is asked to stop scheduling further work
- in-flight tasks are not started again; the run winds down

This is the “hard wall” safety net:
- it prevents any logic bug from keeping the run alive forever.

---

## Starter ordering and why it matters

Starters are ordered before execution. In this build:
- starters are sorted by ticker count (ascending)
- some “special algos” are placed at the end (still sorted by ticker count)

Why ordering matters:
- With budgeting + donations, early completion creates “bank” for later tasks.
- Running smaller/easier starters earlier tends to:
  - complete fast
  - donate time
  - improve coverage for later/harder starters

This is a runtime strategy:
- **front-load cheap tasks** to donate time to more expensive tasks (starters with more tickers).

---

## Shared state and thread safety (what’s actually needed)

MP2 is intentionally designed so that:
- the hot-path structures are read-only during threading:
  - `day_buys_map`, `day_profit_map`, `variation_map`, `full_stocks_map`, `params_map`
- the recursion itself does not require locks

Locks are only needed for:
1) **Counters / bookkeeping**
   - tested group counters
   - per-thread Top-K contribution counts
2) **DB writes**
   - streaming upsert of best groups
   - final stats writes

This is important for scale:
- locking inside the recursion loop would destroy throughput
- keeping recursion lock-free is a core design decision
- Managing the locks is important when multiple threads attempt to change or access a shared list or variable

---

## Scaling characteristics (what changes and what doesn’t)

### What increases runtime to limit
- More starters (more variations across algos)
- Starters with larger day buy sets (consensus shrinks slower)
- Higher `MAX_GROUP_SIZE` (more levels of intersection expansion)
- Looser recursion guardrails (07) that prevent early termination
- Higher total run time increase

### What does NOT change asymptotic risk much
- The thread count (4 vs 6) only scales throughput linearly until CPU becomes the limit.
- The core combinatorial explosion is controlled primarily by:
  - consensus intersections shrinking
  - group size cap
  - daily hit floor
  - time budgeting

### Practical expectation
With 4 workers:
- wall time is bounded by `_global_end`
- coverage improves as donations accumulate
- the run degrades gracefully: late-stage starters may be partial, but the run still completes

---

## Why 4 workers is the current setting

In this design:
- each worker is doing heavy Python set ops + recursion
- heavy shared-memory access patterns
- Excel and SQLite exist in the process (even if lightly used in the begininng mainly)

A fixed cap of 4:
- avoids memory pressure spikes
- keeps CPU utilization high without thrashing
- reduces contention for the GIL compared to higher thread counts
- stays stable across most Windows desktop environments

If you increase workers:
- you usually get diminishing returns due to the GIL and memory bandwidth
- and you can worsen variance (more “hard” starters running simultaneously)
- Without having a server with more cores, the threads will simply queue, not making it faster

So 4 is the best choice until the server is upgraded.

---

## Summary

MP2’s scaling model is:

- **Parallelize across starter variations**, not inside recursion
- Run a **bounded 4-worker pool** with backfill scheduling
- Use a **per-starter time slice** derived from the target wall budget and batches
- **Bank leftover time** from fast starters to fund hard starters later
- Enforce both **per-task** and **global hard-stop** time boundaries
- Keep recursion hot-path lock-free by making core maps read-only during execution
