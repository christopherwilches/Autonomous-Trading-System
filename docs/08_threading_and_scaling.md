# Threading & Scaling

This section explains the **execution layer** of MP2 Part 2: how group search is run at scale using a bounded 4-worker pool, how time is allocated across starters, and how the overall run is kept predictable and bounded.

Contents:
- Thread starter design (what is parallelized)
- 4-worker scheduling model (bounded in-flight, backfill)
- Adaptive per-task time budgeting + leftover “donations”
- Global stop conditions (hard end)

---

## How parallelism is applied

MP2 runs the recursion builder starting from many “starter” variations.

A **starter** is:
- `(algo_idx, variation_identifier)`

Each starter launches an independent group-building search beginning from a single algorithm variation.

Parallelism is applied **across starters**, not inside the recursion itself.
Each recursive search remains single-threaded and self-contained.

Why this is the right cut:
- Starters are independent from one another.
- They share read-only maps (`day_buys_map`, `day_profit_map`, `variation_map`, etc.).
- It avoids locks on tight loops inside recursion.

---

## The 4-worker execution model

MP2 runs a fixed pool of four concurrent workers.
This cap reflects practical CPU, memory, and shared-state limits on the current execution environment.

Scheduling is bounded and continuously backfilled:
- the pool starts with a small number of starters
- completed starters immediately free a worker slot
- new starters are submitted incrementally to keep the pool full
- the system never submits all work at once


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
Each starter is given a time budget derived from the overall wall-time target and the number of execution waves required.

The key idea:
- starters are processed in batches of four
- each batch receives a proportional slice of the total runtime
- this guarantees the run stays close to the intended wall-time regardless of how many starters exist

This keeps execution predictable while still allowing uneven workloads.
The initial time slice is derived from:
- the total wall-time budget
- the number of execution waves required when running four starters at a time

This effectively divides the run into batches, ensuring each batch receives a proportional share of the total runtime.

This means:
- If every running starter consumes roughly `base_limit`,
  total wall time stays near the target because 4 run concurrently.

The key point is that time is allocated per execution wave, not per individual task.
This keeps total runtime close to the target even when individual starters vary widely in complexity.

---

## The donation model (banking leftover time)

When a starter finishes early, its unused time is not wasted.
Instead, that time is collected and redistributed to remaining starters.

Unused time is added to a shared pool and redistributed to remaining starters.
Time allowances only increase over the course of the run and are never reduced.

Result:
- As “easy” starters finish early, harder starters later get more time
- The system self-tunes without manual per-algo tuning
- Starters are ordered by expected difficulty.
- Those with smaller consensus potential run first, allowing them to finish quickly and donate time to more complex starters later.
- Starters with larger potential search space are intentionally scheduled later so they can benefit from accumulated donated time.

### Why it works
- Recursion depth/width varies wildly by starter.
- Some starters collapse quickly (little consensus).
- Others explore deep (persistent consensus overlap).
- Banking converts that variance into better overall coverage without blowing wall time.

---

## Enforcing the budget inside recursion

Each recursive search periodically checks whether it has exceeded its allotted time.
If so, the search exits cleanly without forcing exceptions or unwinding state.

This cooperative cutoff ensures:
- no single starter can monopolize runtime
- partial progress is preserved
- the overall run remains stable and predictable

---

## Global stop conditions

Even with per-task budgeting, MP2 has an absolute stop boundary.

Global stop triggers:
- the total wall-time budget is reached
- an external termination signal is issued

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

---

## Scaling behavior and worker choice

Runtime pressure increases primarily with:
- the number of starters
- how slowly consensus sets shrink as groups expand
- higher group size limits

Thread count affects throughput but not combinatorial risk.
Beyond four workers, gains diminish due to shared-memory pressure, Python execution limits, and higher variance from running multiple complex starters simultaneously.

A four-worker design provides:
- stable memory usage
- high CPU utilization without thrashing
- predictable behavior on desktop-class hardware

This balance allows the donation system to be effective without overwhelming the runtime.

---

## Summary

MP2 scales by parallelizing independent group searches rather than recursive logic.
A bounded four-worker pool, adaptive per-starter time limits, and a donation-based reallocation system ensure:

- predictable wall-time
- graceful handling of uneven workloads
- deeper exploration where it is most valuable
