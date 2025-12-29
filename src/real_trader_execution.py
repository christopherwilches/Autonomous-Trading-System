"""
RealTrader — Live execution + strict Trailing-Stop-High (TSH)

Purpose:
- Place one entry per symbol at/near the open (no re-entries).
- Manage exits with a strict, deterministic “TSH” model:
  - No exits allowed until the target is reached (“arming”).
  - After arming, track the session high and enforce a rising floor.
  - Exit immediately when price falls to the floor (market sell).
  - End-of-day flatten as a safety backstop.

Privacy / portfolio notes:
- This version is intentionally redacted:
  - API keys, endpoints, local paths, and symbols are removed.
  - All explicit dollar amounts and capital figures are removed.
  - Exact bucket thresholds / target percentages may be simplified or omitted.
- The implementation is shown to demonstrate system design and the custom,
  deterministic trade-management abstraction. I originally derived this “TSH”
  behavior independently (similar to building abstractions before learning the
  formal names), then refined it into a clean, testable model.
"""

import time
import sqlite3 
from datetime import datetime, timezone
from typing import Dict, Optional

from alpaca_trade_api.rest import REST, APIError
from alpaca_trade_api.stream import Stream
from alpaca_trade_api.common import URL
import threading

# Config

USE_PAPER = False           # Paper vs live keys for paper or live trading
USE_TEST_CAP = True
TEST_CAP = 0.0 # explicit capital limits removed        

LEVERAGE_FACTOR = 1.0  

API_KEY_PAPER = ""
API_SECRET_PAPER = ""
ENDPOINT_PAPER = ""

API_KEY_LIVE = ""
API_SECRET_LIVE = ""
ENDPOINT_LIVE = ""

DB_FILE = r"<REDACTED_PATH>/stocks_data.db"
LOG_FILE = "alpaca_realtrader_log.txt"

SYMBOLS: list[str] = [] # symbols removed for privacy

DATA_FEED = "sip"
USE_NBBO = DATA_FEED.lower() == "sip" 

PERCENT_TARGETS = [
    # exact bucket thresholds + target percentages removed for privacy
    (0.0, 0.0, 0.0),
]

def trailing_from_target(t: float) -> float:
    """Trailing percent is half the target percent."""
    return t / 2.0

def pick_target_pct(price: float) -> float:
    """Choose the target % based on entry price bucket."""
    for lo, hi, p in PERCENT_TARGETS:
        if lo <= price < hi:
            return p
    return 0.0 # redacted for privacy 

def log(msg: str) -> None:
    """Log to both console and file."""
    ts = datetime.now().isoformat()
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def get_api() -> REST:
    """Return a REST client for paper or live."""
    if USE_PAPER:
        return REST(API_KEY_PAPER, API_SECRET_PAPER, ENDPOINT_PAPER)
    else:
        return REST(API_KEY_LIVE, API_SECRET_LIVE, ENDPOINT_LIVE)

# Trailing Model (Custom TSH (Trailing Stop High))

class LiveTrade:
    def __init__(self, symbol: str, entry_price: float, shares: float, target_pct: float):
        self.symbol = symbol
        self.entry = entry_price
        self.shares = shares

        self.target_pct = target_pct
        self.trailing_pct = trailing_from_target(target_pct)

        # Core TSH parameters
        self.target_price = self.entry * (1 + self.target_pct)
        # Fixed dollar trailing amount based on ENTRY price
        self.trailing_amount = self.entry * self.trailing_pct

        # State
        self.highest = self.entry
        self.triggered = False
        self.sold = False

        # Logging / DB fields
        self.entry_time: Optional[str] = datetime.now(timezone.utc).isoformat()
        self.exit_time: Optional[str] = None
        self.exit_price: Optional[float] = None
        self.exit_floor: Optional[float] = None

        log(
            f"[INIT] {symbol}: entry={self.entry}, "
            f"target={self.target_price}, trail%={self.trailing_pct}, "
            f"trail_amt={self.trailing_amount}"
        )

    def update(self, price: float) -> bool:
        """
        Update the model with a new price.

        Returns True if a trailing market sell should be triggered.
        """
        if self.sold:
            return False

        # Phase 1: not yet triggered
        if not self.triggered:
            if price >= self.target_price:
                self.triggered = True
                self.highest = price
                log(
                    f"{self.symbol}: TSH ARMED at price={price}, "
                    f"target={self.target_price}"
                )
            return False

        # Phase 2: triggered 
        if price > self.highest:
            self.highest = price
            log(f"{self.symbol}: new HIGH = {price}")

        # Fixed-dollar trailing
        raw_floor = self.highest - self.trailing_amount
        floor = max(self.target_price, raw_floor)
        self.exit_floor = floor 

        log(f"{self.symbol}: price={price} floor={floor}")

        # Once the price is at or below the floor, sell
        if price <= floor:
            log(f"TRIGGER SELL: {self.symbol} @ {price} (floor={floor})")
            return True

        return False

# ================================================================
# STREAMING STATE
# ================================================================
api: Optional[REST] = None
entry_prices: Dict[str, float] = {}
alloc_per_symbol: Dict[str, float] = {}
live_trades: Dict[str, LiveTrade] = {}
bought_symbols: Dict[str, bool] = {sym: False for sym in SYMBOLS}
price_stats: Dict[str, Dict[str, Optional[float]]] = {
    sym: {"high": None, "low": None, "last": None}
    for sym in SYMBOLS
}

def update_price_stats(symbol: str, price: float) -> None:
    """Track session high/low/last price for each symbol."""
    stats = price_stats.setdefault(
        symbol, {"high": None, "low": None, "last": None}
    )
    if stats["high"] is None or price > stats["high"]:
        stats["high"] = price
    if stats["low"] is None or price < stats["low"]:
        stats["low"] = price
    stats["last"] = price

def wait_for_market_open(rest_api: REST) -> None:
    # Wait for market to open
    from zoneinfo import ZoneInfo
    NY = ZoneInfo("America/New_York")

    while True:
        now_ny = datetime.now(NY)
        today_929 = now_ny.replace(hour=9, minute=29, second=0, microsecond=0)

        if now_ny < today_929:
            seconds = (today_929 - now_ny).total_seconds()
            log("Pre-open: sleeping until the start-of-session polling window.")
            time.sleep(seconds)
            continue

        clock = rest_api.get_clock()
        if clock.is_open:
            log("Market OPEN.")
            return

        log("Waiting for market to open... checking again in 1s.")
        time.sleep(1)

def flatten_existing_positions(rest_api: REST) -> None:
    try:
        positions = rest_api.list_positions()
    except APIError as e:
        log(f"Error listing positions for flatten: {e}")
        return

    if not positions:
        log("No existing positions to flatten.")
        return

    log(f"Flattening {len(positions)} existing positions...")
    for p in positions:
        sym = p.symbol
        qty_raw = float(p.qty)   
        if qty_raw == 0:
            continue

        side = "sell" if qty_raw > 0 else "buy" 
        qty = abs(qty_raw)

        try:
            rest_api.submit_order(
                symbol=sym,
                qty=qty,         
                side=side,
                type="market",
                time_in_force="day",
            )
            log(f"Submitted flatten {side} for {sym}, qty={qty}")

            model = live_trades.get(sym)
            if model and not model.sold:
                last_price = price_stats.get(sym, {}).get("last")
                if last_price is not None:
                    model.sold = True
                    model.exit_price = last_price
                    model.exit_time = datetime.now(timezone.utc).isoformat()
                    raw_floor = model.highest - model.trailing_amount
                    model.exit_floor = max(model.target_price, raw_floor)
        except APIError as e:
            log(f"Error flattening {sym}: {e}")

def wait_for_fill(
    rest_api: REST,
    order_id: str,
    symbol: str,
    timeout: float = 20.0,
) -> Optional[float]:

    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            o = rest_api.get_order(order_id)
        except APIError as e:
            log(f"{symbol}: error fetching order {order_id} for fill check: {e}")
            return None

        status = o.status
        filled_avg_price = o.filled_avg_price

        if status == "filled" and filled_avg_price is not None:
            price = float(filled_avg_price)
            log(
                f"{symbol}: order {order_id} filled, "
                f"filled_avg_price={price}"
            )
            return price

        if status in ("canceled", "rejected"):
            log(f"{symbol}: order {order_id} {status}, no fill.")
            return None
        time.sleep(0.5)
    log(
        f"{symbol}: order {order_id} not filled within timeout."
    )
    return None
  
# ================================================================
# TRADE LOG DB HELPERS
# ================================================================

def init_trade_log_table() -> None:
    try:
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS REAL_TRADER_LOG (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                run_date TEXT,
                is_paper INTEGER,
                symbol TEXT,
                alloc_dollars REAL,
                entry_time TEXT,
                exit_time TEXT,
                entry_price REAL,
                exit_price REAL,
                target_pct REAL,
                trailing_pct REAL,
                target_price REAL,
                tsh_triggered INTEGER,
                trailing_floor_exit REAL,
                day_high REAL,
                day_low REAL,
                close_price REAL,
                profit_pct REAL
            )
            """
        )
        conn.commit()
        conn.close()
        log("[DB] REAL_TRADER_LOG table ready.")
    except Exception as e:
        log(f"[DB] init_trade_log_table error: {e}")

def clear_trade_log_table() -> None:
    try:
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("DELETE FROM REAL_TRADER_LOG")
        conn.commit()
        conn.close()
        log("[DB] REAL_TRADER_LOG cleared.")
    except Exception as e:
        log(f"[DB] clear_trade_log_table error: {e}")

def write_trade_log(is_paper: bool) -> None:
    try:
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        run_date = datetime.now(timezone.utc).date().isoformat()

        for sym in SYMBOLS:
            model = live_trades.get(sym)
            stats = price_stats.get(sym, {})
            day_high = stats.get("high")
            day_low = stats.get("low")
            close_price = stats.get("last")

            if model is None:
                continue

            entry_price = model.entry
            exit_price = model.exit_price or close_price
            entry_time = model.entry_time
            exit_time = model.exit_time
            target_pct = model.target_pct
            trailing_pct = model.trailing_pct
            target_price = model.target_price
            tsh_triggered = 1 if model.triggered else 0
            trailing_floor_exit = model.exit_floor
            alloc = alloc_per_symbol.get(sym, 0.0)

            profit_pct = None
            if entry_price and exit_price:
                profit_pct = (exit_price - entry_price) / entry_price * 100.0

            cur.execute(
                """
                INSERT INTO REAL_TRADER_LOG (
                    run_date, is_paper, symbol, alloc_dollars,
                    entry_time, exit_time, entry_price, exit_price,
                    target_pct, trailing_pct, target_price,
                    tsh_triggered, trailing_floor_exit,
                    day_high, day_low, close_price, profit_pct
                )
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    run_date,
                    1 if is_paper else 0,
                    sym,
                    alloc,
                    entry_time,
                    exit_time,
                    entry_price,
                    exit_price,
                    target_pct,
                    trailing_pct,
                    target_price,
                    tsh_triggered,
                    trailing_floor_exit,
                    day_high,
                    day_low,
                    close_price,
                    profit_pct,
                ),
            )

        conn.commit()
        conn.close()
        log("[DB] REAL_TRADER_LOG updated for this run.")
    except Exception as e:
        log(f"[DB] write_trade_log error: {e}")
def place_initial_buys(rest_api: REST) -> None:
    for sym in SYMBOLS:
        alloc = alloc_per_symbol.get(sym, 0.0)
        if alloc <= 0:
            log(f"{sym}: no capital allocation, skipping BUY.")
            continue

        try:
            order = rest_api.submit_order(
                symbol=sym,
                notional=alloc,     
                side="buy",
                type="market",
                time_in_force="day",
            )
            log(
                f"Submitted NOTIONAL BUY for {sym}, "
                f"notional={alloc:.2f}, order_id={order.id}"
            )

            filled_price = wait_for_fill(rest_api, order.id, sym, timeout=20.0)
            if filled_price is None:
                log(f"{sym}: no fill price available, skipping LiveTrade model.")
                continue

            try:
                o = rest_api.get_order(order.id)
                shares = float(o.filled_qty or 0)
            except APIError as e:
                log(f"{sym}: error re-fetching order for filled_qty: {e}")
                shares = 0.0

            if shares <= 0:
                log(f"{sym}: filled_qty is 0, skipping LiveTrade model.")
                continue

            update_price_stats(sym, filled_price)

            tgt_pct = pick_target_pct(filled_price)
            live_trades[sym] = LiveTrade(sym, filled_price, shares, tgt_pct)

        except APIError as e:
            log(f"Error placing NOTIONAL BUY for {sym}: {e}")

# ================================================================
# TRADE-DRIVEN CALLBACK (IEX or non-NBBO mode)
# ================================================================

async def on_trade(trade) -> None:
    global api
    symbol = trade.symbol
    if symbol not in SYMBOLS:
        return

    price = float(trade.price)
    ts = getattr(trade, "timestamp", None)
    update_price_stats(symbol, price)
    log(f"TRADE {symbol} price={price} t={ts}")

    if api is None:
        log("API reference missing in on_trade; ignoring trade.")
        return

    model = live_trades.get(symbol)
    if model is None:
        return

    # Trailing-stop-high logic
    if model.update(price):
        qty = model.shares
        try:
            api.submit_order(symbol, qty, "sell", "market", "day")
            log(f"Submitted trailing market SELL for {symbol}, qty={qty}")
            model.sold = True
            model.exit_price = price
            model.exit_time = datetime.now(timezone.utc).isoformat()

            entry = model.entry
            gain = price - entry
            pct_gain = (gain / entry) * 100 if entry else 0
            log(
                f"[EXIT DIAG] {symbol}: entry={entry:.4f}, exit={price:.4f}, "
                f"profit={gain:.4f} ({pct_gain:.4f}%)"
            )

        except APIError as e:
            log(f"Error submitting trailing SELL for {symbol}: {e}")

# ================================================================
# QUOTE-DRIVEN CALLBACK (SIP NBBO mode)
# ================================================================

async def on_quote(quote) -> None:
    global api
    symbol = quote.symbol
    if symbol not in SYMBOLS:
        return

    # NBBO bid / ask
    bid = float(quote.bid_price) if quote.bid_price is not None else 0.0
    ask = float(quote.ask_price) if quote.ask_price is not None else 0.0
    ts = getattr(quote, "timestamp", None)

    log(f"QUOTE {symbol} bid={bid} ask={ask} t={ts}")
    if api is None:
        log("API reference missing in on_quote; ignoring quote.")
        return
    model = live_trades.get(symbol)
    if model is None:
        return
    price_for_trailing = bid if bid > 0 else ask
    if price_for_trailing <= 0:
        return
    update_price_stats(symbol, price_for_trailing)
    if model.update(price_for_trailing):
        qty = model.shares
        try:
            api.submit_order(symbol, qty, "sell", "market", "day")
            log(f"Submitted trailing market SELL for {symbol}, qty={qty}")
            model.sold = True
            model.exit_price = price_for_trailing
            model.exit_time = datetime.now(timezone.utc).isoformat()
            entry = model.entry
            gain = price_for_trailing - entry
            pct_gain = (gain / entry) * 100 if entry else 0
            log(
                f"[EXIT DIAG] {symbol}: entry={entry:.4f}, "
                f"exit={price_for_trailing:.4f}, "
                f"profit={gain:.4f} ({pct_gain:.4f}%)"
            )

        except APIError as e:
            log(f"Error submitting trailing SELL for {symbol}: {e}")

def main() -> None:
    global api
    open(LOG_FILE, "w", encoding="utf-8").close()

    api = get_api()
    log("=== RealTrader starting ===")
    log(f"Using tickers: {SYMBOLS}")
    log(f"Data feed: {DATA_FEED} (USE_NBBO={USE_NBBO})")
    init_trade_log_table()

    flatten_existing_positions(api)
    wait_for_market_open(api)

    acct = api.get_account()
    bp = float(acct.buying_power)
    equity = float(acct.equity)

    if USE_TEST_CAP:
        effective_bp = min(bp, TEST_CAP)
        mode_desc = f"TEST_CAP={TEST_CAP}"
    else:
        effective_bp = min(equity * LEVERAGE_FACTOR, bp)
        mode_desc = (
            f"equity={equity:.2f}, bp={bp:.2f}, LEVERAGE_FACTOR={LEVERAGE_FACTOR}"
        )

    if effective_bp <= 0:
        log(f"No effective buying power available ({mode_desc}). Exiting.")
        return

    n = len(SYMBOLS)
    effective_bp_rounded = round(float(effective_bp), 2)

    base = round(effective_bp_rounded / n, 2)
    allocs = [base] * n

    total = round(sum(allocs), 2)
    leftover_cents = int(round((effective_bp_rounded - total) * 100))

    for i in range(max(0, leftover_cents)):
        allocs[i % n] = round(allocs[i % n] + 0.01, 2)

    alloc_per_symbol.clear()
    for sym, a in zip(SYMBOLS, allocs):
        alloc_per_symbol[sym] = a

    log(
        f"Allocations (cents-safe): total={sum(allocs):.2f} "
        f"-> { {s: alloc_per_symbol[s] for s in SYMBOLS} }"
    )

    avg_alloc = round(sum(allocs) / len(allocs), 2)

    log(
        f"Capital config: {mode_desc}, effective_bp={effective_bp:.2f}, "
        f"avg per-symbol allocation~{avg_alloc:.2f}"
    )

    place_initial_buys(api)

    log(
        "Price polling loop starting. Buys and trailing stops are now driven by "
        "REST latest quote/trade data."
    )
    eod_flatten_done = False

    try:
        while True:
            clock = api.get_clock()
            if not clock.is_open:
                log("Market closed according to Alpaca clock. Ending price loop.")
                break

            now = datetime.now(timezone.utc)
            seconds_to_close = (clock.next_close - now).total_seconds()
            if seconds_to_close <= 120 and not eod_flatten_done:
                log(
                    f"Within {seconds_to_close:.0f}s of close; "
                    "flattening all positions as EOD safety."
                )
                flatten_existing_positions(api)
                eod_flatten_done = True

            for sym in SYMBOLS:
                model = live_trades.get(sym)
                if model is None or model.sold:
                    continue

                try:
                    if USE_NBBO:
                        # SIP NBBO: use latest quote (bid/ask)
                        q = api.get_latest_quote(sym)
                        bid = float(q.bid_price) if q.bid_price is not None else 0.0
                        ask = float(q.ask_price) if q.ask_price is not None else 0.0
                        ts = getattr(q, "timestamp", None)

                        price_for_trailing = bid if bid > 0 else ask
                        if price_for_trailing <= 0:
                            continue

                        log(f"QUOTE {sym} bid={bid} ask={ask} t={ts}")
                        update_price_stats(sym, price_for_trailing)

                        if model.update(price_for_trailing):
                            qty = model.shares
                            api.submit_order(sym, qty, "sell", "market", "day")
                            log(f"Submitted trailing market SELL for {sym}, qty={qty}")
                            model.sold = True
                            model.exit_price = price_for_trailing
                            model.exit_time = datetime.now(timezone.utc).isoformat()

                            entry = model.entry
                            gain = price_for_trailing - entry
                            pct_gain = (gain / entry) * 100 if entry else 0
                            log(
                                f"[EXIT DIAG] {sym}: entry={entry:.4f}, "
                                f"exit={price_for_trailing:.4f}, "
                                f"profit={gain:.4f} ({pct_gain:.4f}%)"
                            )
                    else:
                        # IEX / non-NBBO: use latest trade price
                        t = api.get_latest_trade(sym)
                        price = float(t.price)
                        ts = getattr(t, "timestamp", None)

                        log(f"TRADE {sym} price={price} t={ts}")
                        update_price_stats(sym, price)

                        if model.update(price):
                            qty = model.shares
                            api.submit_order(sym, qty, "sell", "market", "day")
                            log(f"Submitted trailing market SELL for {sym}, qty={qty}")
                            model.sold = True
                            model.exit_price = price
                            model.exit_time = datetime.now(timezone.utc).isoformat()

                            entry = model.entry
                            gain = price - entry
                            pct_gain = (gain / entry) * 100 if entry else 0
                            log(
                                f"[EXIT DIAG] {sym}: entry={entry:.4f}, "
                                f"exit={price:.4f}, "
                                f"profit={gain:.4f} ({pct_gain:.4f}%)"
                            )

                except APIError as e:
                    log(f"{sym}: error during price polling: {e}")
                except Exception as e:
                    log(f"{sym}: unexpected error during price polling: {e}")

            time.sleep(5)

    finally:
        # flatten anything left and log trades
        flatten_existing_positions(api)
        write_trade_log(USE_PAPER)
        log("RealTrader finished.")

if __name__ == "__main__":
    main()
