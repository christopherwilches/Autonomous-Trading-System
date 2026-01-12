# Data Pipeline

## What this stage does
This stage builds the weekly dataset used by mp1.

It produces in `stocks_data.db`
- `DICT_DAY1..DICT_DAY5`
- `DATASET_TICKERS` (the 1000 chosen tickers + order)
- `DS_DAY1..DS_DAY5` (same 1000 tickers, same order, all 5 days)

Goal
- Same 1000 tickers across all 5 days
- Formatted exactly how the Excel workbook expects

Component name in this doc
- `weekly_5d_data`

## Source universe
Starts from Alpaca `/v2/assets` and applies various filters after.

Asset must be
- `class == us_equity`
- `status == active`
- `exchange != OTC`
- Ticker symbol is letters only (`A–Z`), no dots, hyphens, or punctuation
- Name filter removes wrappers like funds / etfs / reits / notes / units / warrants
- Also removes common ETF families by name prefix (spdr, ishares, vanguard, invesco, proshares, global x, direxion)

After this, the universe is usually ~5k symbols.

## Price + volume screen
For each symbol, the script fetches daily bars and screens using the most recent completed day.

Filters
- `10 < close < 90`
- `250,000 <= volume <= 100,000,000`

Completed day rule
- If the newest daily bar is “today” and it is before 4pm NY time, that bar is dropped
- This avoids using an incomplete daily candle during market hours

## Minimum history requirement
This builder needs enough bars to create 5 shifted windows.

Requirement
- enough history to support five shifted 31-day windows
- for a 5-day setup, this requires 35 completed trading days

The script uses a long enough calendar lookback (around ~65 days) so 35 trading days is guaranteed.

## Why the 53-row format exists
Each ticker is stored as a fixed 53-row block because the Excel algorithms are built around it.

Block layout
- Row 0: blank
- Row 1: header row
- Rows 2–32: 31 daily bars (newest at the top)
- Rows 33–52: blank padding

Details that matter
- The ticker symbol appears only on the first data row, then remains blank for the rest of the block
- This matches how the workbook slices blocks and stays stable when batching

## DICT_DAY tables
Tickers that pass the screens get written to `DICT_DAY1..DICT_DAY5`.

Each table holds a different 31-day window for the same ticker set, shifted by 1 day each time.

Offset rule
- `DICT_DAY5` uses the most recent window (offset 0)
- `DICT_DAY4` offset 1
- `DICT_DAY3` offset 2
- `DICT_DAY2` offset 3
- `DICT_DAY1` offset 4

This gives each ticker five overlapping windows, one for each day.


## Building the weekly dataset (DS tables)
`DS_DAY` tables are the final weekly dataset mp1 trains on.

How it gets built
- Extract ticker sets from `DICT_DAY1..DICT_DAY5`
- Take the intersection (tickers that exist in all five)
- Randomly shuffle, then pick 1000
- Save the 1000 list into `DATASET_TICKERS` with `order_idx`
- Copy the same 53-row blocks into `DS_DAY1..DS_DAY5` in that exact order

Important
- `DS_DAY` does not add extra filtering
- It is intersection + sampling only, to be realistic and allow a representative sample of the remaining universe

## Runtime and limits
- Daily bar fetching is multi-threaded (`N_WORKERS = 4`)
- Typical runtime is about 100–120 seconds depending on API speed
