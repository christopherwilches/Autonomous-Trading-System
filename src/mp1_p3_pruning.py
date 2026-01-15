"""
MP1 — Part 3: Export pruned variations (DB -> WB2)

Selects and deduplicates the best variations per algorithm from the Optuna results DB,
ensures minimum coverage / global minimum counts, writes pruned_* tables into stocks_data.db,
and pastes a Friday-only view into WB2 (ControlSheet) for MP2.

Input: optuna_10A_data.* (source results DB), ControlSheet headers (WB1)
Output: pruned_<ALGO> tables (stocks_data.db) + WB2 pasted blocks
"""

def export_variations_to_excel():
    import json, math, sqlite3, time
    import xlwings as xw
    EXCEL_FILE = r"<REDACTED_PATH>/WB1"
    optuna_db_path = r"<REDACTED_PATH>/optuna_data.db"
    # --- Open books/sheets
    wb1 = xw.Book(EXCEL_FILE)
    control_sheet = wb1.sheets["ControlSheet"]

    WB2_FILE = r"<REDACTED_PATH>/WB2"
    wb2 = xw.Book(WB2_FILE)
    output_sheet = wb2.sheets["ControlSheet"]

    app = wb2.app
    app.screen_updating = False
    app.display_alerts  = False
    try:
        app.api.Calculation = -4135 
    except Exception:
        pass
    try:
        for __sh in wb2.sheets:
            try:
                __sh.api.ScrollArea = ""
            except Exception:
                pass
    except Exception:
        pass
    try:
        app.enable_events = True
        output_sheet.range("P3").value = "wipe_trigger"
        try:
            app.api.CalculateFullRebuild()
        except Exception:
            wb2.app.calculate()
        try:
            app.api.DoEvents()
        except Exception:
            pass
        time.sleep(0.3)
    finally:
        app.enable_events = False
    import sqlite3, os, glob

    ALGO_TABLES = [
        "MACD", "EMA", "RSI", "Breakout", "ADX", "Volatility",
        "SMA", "Bollinger_Bands", "EMA_MACD_Combo", "RSI_Bollinger"
    ]

    def _db_has_algo_tables(_conn):
        try:
            _cur = _conn.cursor()
            qmarks = ",".join(["?"] * len(ALGO_TABLES))
            _cur.execute(
                f'SELECT COUNT(*) FROM sqlite_master WHERE type="table" AND name IN ({qmarks})',
                ALGO_TABLES
            )
            return int(_cur.fetchone()[0] or 0)
        except Exception:
            return 0

    # Build candidate list
    _candidates = []
    try:
        if optuna_db_path:
            _candidates.append(optuna_db_path)
    except Exception:
        pass

    _base = r"<REDACTED_PATH>
    _names = [
        "optuna_10A_data",          
        "optuna_10A_data.db",
        "optuna_10A_data.sqlite",
        "optuna_10A_data.sqlite3"
    ]
    for nm in _names:
        _candidates.append(os.path.join(_base, nm))

    _candidates.extend(glob.glob(os.path.join(_base, "optuna_10A_data*.*")))

    _seen = set(); _candidates = [c for c in _candidates if not (c in _seen or _seen.add(c))]

    conn = None
    cur = None
    SOURCE_DB_FILE = None
    _best = (None, -1) 

    for _path in _candidates:
        if not _path or not isinstance(_path, str):
            continue
        try:
            _conn = sqlite3.connect(_path)
        except Exception:
            continue
        try:
            cnt = _db_has_algo_tables(_conn)
            if cnt > _best[1]:
                if _best[0]:
                    try: _old = sqlite3.connect(_best[0]); _old.close()
                    except Exception: pass
                _best = (_path, cnt)
            _conn.close()
        except Exception:
            try: _conn.close()
            except Exception: pass

    if _best[0] is None or _best[1] <= 0:
        # Hard fail with clear message
        raise RuntimeError(
            "P3 source DB not found: none of the expected algo tables exist in the searched locations. "
            "Ensure optuna_10A_data.* is present under r"<REDACTED_PATH>" or set optuna_db_path correctly."
        )

    SOURCE_DB_FILE = _best[0]
    conn = sqlite3.connect(SOURCE_DB_FILE)
    cur = conn.cursor()
    print(f"[P3] Source DB chosen: {SOURCE_DB_FILE} (algo tables found={_best[1]})")

    STOCKS_DB_FILE = r"<REDACTED_PATH>/stocks_data.db"
    sd_conn = sqlite3.connect(STOCKS_DB_FILE)
    sd_cur = sd_conn.cursor()

    # Layout
    start_cell_row = 6
    start_cell_col = 3
    param_start_col = 56
    row_buffer_between_variations = 2
    col_buffer_between_algorithms = 16

    # Config
    ALGO_DISPLAY_NAMES = [
        "MACD", "EMA", "RSI", "Breakout", "ADX", "Volatility",
        "SMA", "Bollinger_Bands", "EMA_MACD_Combo", "RSI_Bollinger"
    ]
    SPECIAL_ALGOS = {"MACD", "EMA", "Breakout"}
    OVERALL_MAX = 1300
    GLOBAL_MIN  = 700

    # Per-algo caps
    def algo_cap(name: str) -> int:
        return 250 if name in SPECIAL_ALGOS else 100
    def get_colmap(table):
        cur.execute(f'PRAGMA table_info("{table}")')
        rows = cur.fetchall()
        if not rows:
            raise RuntimeError(f'Source table "{table}" not found in {SOURCE_DB_FILE}')
        return {r[1]: r[0] for r in rows}

    def g(row, cmap, key, default=None, cast=float):
        idx = cmap.get(key)
        if idx is None:
            return default
        try:
            v = row[idx]
            if v is None:
                return default
            return cast(v) if cast else v
        except Exception:
            return default
    def variation_similarity(p1, p2, algo_name):
        try:
            total = min(len(p1), len(p2), len(PARAM_RANGES[algo_name]))
        except Exception:
            total = min(len(p1), len(p2))
        same = 0
        for i in range(total):
            try:
                base = PARAM_RANGES[algo_name][i]
            except Exception:
                base = None
            try:
                # Volatility special formatting
                if algo_name == "Volatility" and i == 1 and base is not None:
                    a_min, a_max = map(float, str(p1[1]).split("-"))
                    b_min, b_max = map(float, str(p2[1]).split("-"))
                    base_min = PARAM_RANGES["Volatility"][1]
                    base_max = PARAM_RANGES["Volatility"][2]
                    low, high, step = base_min
                    tol_min = step
                    if abs(a_min - b_min) <= tol_min:
                        same += 1
                    low, high, step = base_max
                    tol_max = step
                    if abs(a_max - b_max) <= tol_max:
                        same += 1
                    continue
                if algo_name == "Volatility" and i == 2:
                    continue
                if isinstance(base, list):
                    if str(p1[i]) == str(p2[i]):
                        same += 1
                    continue
                if base is None:
                    if p1[i] == p2[i]:
                        same += 1
                    continue
                low, high, step = base
                options = round((high - low) / step) + 1
                tol = 0 if options <= 10 else step if options <= 40 else 2*step if options <= 150 else 3*step
                if abs(p1[i] - p2[i]) <= tol:
                    same += 1
            except Exception:
                continue
        return same > (total // 2)

    def ticker_set(stocks):
        return {s.get("ticker") for s in stocks if s and s.get("ticker")}

    def ticker_overlap_frac(s1, s2):
        t1 = ticker_set(s1); t2 = ticker_set(s2)
        if not t1 or not t2:
            return 0.0
        inter = len(t1 & t2)
        denom = max(len(t1), len(t2))
        return (inter / denom) if denom else 0.0
    def old_dedup_pool(pool, algo_name):
        deduped = []
        for v in pool:
            is_dup = False
            for d in list(deduped):
                v_pp = float((v.get("metrics") or {}).get("pooled_pp", 0.0) or 0.0)
                d_pp = float((d.get("metrics") or {}).get("pooled_pp", 0.0) or 0.0)
                if algo_name == "EMA":
                    t1 = ticker_set(v["stocks"]); t2 = ticker_set(d["stocks"])
                    if abs(len(t1.symmetric_difference(t2))) <= 1:
                        if (len(t1), v_pp) > (len(t2), d_pp):
                            deduped.remove(d); deduped.append(v)
                        is_dup = True; break
                else:
                    if variation_similarity(v["params"], d["params"], algo_name) and \
                       ticker_overlap_frac(v["stocks"], d["stocks"]) >= 0.60:
                        if (v_pp, len(v["stocks"])) > (d_pp, len(d["stocks"])):
                            deduped.remove(d); deduped.append(v)
                        is_dup = True; break
            if not is_dup:
                deduped.append(v)
        return deduped

    def eligible_basic(m):
        """Minimal gate — filters out junk or dead runs only."""
        pb = m.get("pooled_buys", 0) or 0
        mdh = m.get("min_daily_hits", 0) or 0
        med = m.get("median_pp_5d", 0) or 0
        mad = m.get("pp_mad_5d", 999)
        rng = m.get("pp_range_5d", 999)
        if pb < 10: return False
        if mdh < 1: return False
        if med <= 0: return False
        if mad > 25.0 and rng > 80.0: return False
        return True

    def export_score(m):
        """Unified continuous scoring formula — favors steady & diverse."""
        avg_pp = float(m.get("median_pp_5d", 0.0) or 0.0)
        mad    = float(m.get("pp_mad_5d", 0.0) or 0.0)
        cv     = float(m.get("buycount_cv_5d", 0.0) or 0.0)        
        rpt    = float(m.get("repeat_ticker_rate_5d", 0.0) or 0.0)
        top    = float(m.get("top_10_ticker_share_5d", 0.0) or 0.0)
        hits   = float(m.get("min_daily_hits", 0.0) or 0.0)
        wl     = float(m.get("wilson_lb_5d", 0.0) or 0.0)

        score = (
              0.55 * avg_pp
            + 0.20 * (100.0 - mad)
            + 0.10 * (100.0 - cv)    
            + 0.10 * (100.0 - rpt)
            + 0.05 * wl
            - 0.05 * top
            + 0.02 * hits
        )
        return max(score, 0.0)

    def apply_gate(pool):
        """Applies minimal eligibility gate, discards only useless rows."""
        return [v for v in pool if eligible_basic(v["metrics"])]

    # Diagnostics
    def _snapshot_quants(vals):
        if not vals: return (0, 0, 0)
        a = sorted(vals)
        def q(p):
            idx = int((len(a)-1) * p / 100)
            return round(float(a[idx]), 2)
        return (q(50), q(75), q(90))

    def print_gate_diagnostics(algo_name, pool, thr):
        if not pool:
            print(f"[P3][diag] {algo_name}: no rows for diagnostics")
            return
        keys = [
            "wilson_lb_5d","median_pp_5d","pp_mad_5d","pp_iqr_5d","pp_range_5d",
            "buycount_cv_5d","min_daily_hits","buycount_med_5d",
            "repeat_ticker_rate_5d","top_10_ticker_share_5d","passed_consistency_gate"
        ]
        metrics = {k: [] for k in keys}
        for v in pool:
            m = v["metrics"]
            for k in keys:
                try:
                    metrics[k].append(float(m.get(k, 0) or 0))
                except Exception:
                    metrics[k].append(0.0)

        # Quantiles for a few key metrics
        wl_q = _snapshot_quants(metrics["wilson_lb_5d"])
        mp_q = _snapshot_quants(metrics["median_pp_5d"])
        mad_q = _snapshot_quants(metrics["pp_mad_5d"])
        cv_q = _snapshot_quants(metrics["buycount_cv_5d"])
        print(f"[P3][diag] {algo_name}: wilson(m/p90)={wl_q[0]}/{wl_q[2]}  "
              f"median_pp(m/p90)={mp_q[0]}/{mp_q[2]}  "
              f"pp_mad(m/p90)={mad_q[0]}/{mad_q[2]}  "
              f"buy_cv(m/p90)={cv_q[0]}/{cv_q[2]}")

        # Component pass counts
        comps = [
            ("passed_consistency_gate", lambda m: int(m.get("passed_consistency_gate", 0)) >= thr.get("passed_consistency_gate", 1)),
            ("min_daily_hits",         lambda m: (m.get("min_daily_hits", 0) or 0) >= thr.get("min_daily_hits", 4)),
            ("buycount_med_5d",        lambda m: (m.get("buycount_med_5d", 0) or 0) >= thr.get("buycount_med_5d", 4)),
            ("wilson_lb_5d",           lambda m: (m.get("wilson_lb_5d", 0) or 0) >= thr.get("wilson_lb_5d", 50.0)),
            ("median_pp_5d",           lambda m: (m.get("median_pp_5d", 0) or 0) >= thr.get("median_pp_5d", 55.0)),
            ("pp_mad_5d",              lambda m: (m.get("pp_mad_5d", 1e9)) <= thr.get("pp_mad_5d", 10.0)),
            ("pp_iqr_5d",              lambda m: (m.get("pp_iqr_5d", 1e9)) <= thr.get("pp_iqr_5d", 20.0)),
            ("pp_range_5d",            lambda m: (m.get("pp_range_5d", 1e9)) <= thr.get("pp_range_5d", 30.0)),
            ("buycount_cv_5d",         lambda m: (m.get("buycount_cv_5d", 1e9)) <= thr.get("buycount_cv_5d", 0.60)),
            ("repeat_ticker_rate_5d",  lambda m: (m.get("repeat_ticker_rate_5d", 1e9)) <= thr.get("repeat_ticker_rate_5d", 85.0)),
            ("top_10_ticker_share_5d", lambda m: (m.get("top_10_ticker_share_5d", 1e9)) <= thr.get("top_10_ticker_share_5d", 60.0)),
        ]
        parts = []
        for name, fn in comps:
            cnt = sum(1 for v in pool if fn(v["metrics"]))
            parts.append(f"{name}={cnt}")
        print(f"[P3][diag] {algo_name}: gate component passes -> " + ", ".join(parts))

    DAY_KEYS = ["day1","day2","day3","day4","day5"]

    def extract_row(row, cmap):
        params = json.loads(g(row, cmap, "params", "[]", cast=str))
        trial  = g(row, cmap, "trial_number", None, cast=int)
        varno  = g(row, cmap, "variation_number", None, cast=int)
        if trial is None or varno is None:
            return None

        # Day-wise
        day_pp  = []
        for d in DAY_KEYS:
            day_pp.append(float(g(row, cmap, f"{d}_pp", 0.0)))

        pp_min  = float(min(day_pp)) if day_pp else 0.0
        pp_max  = float(max(day_pp)) if day_pp else 0.0
        pp_rng  = float(pp_max - pp_min)

        # Metrics
        m = dict(
            pooled_buys = g(row, cmap, "pooled_buys", 0, cast=int),
            pooled_hits = g(row, cmap, "pooled_hits", 0, cast=int),
            pooled_pp   = g(row, cmap, "pooled_pp", 0.0),

            median_pp_5d = g(row, cmap, "median_pp_5d", 0.0),
            pp_mad_5d    = g(row, cmap, "pp_mad_5d", 0.0),
            pp_iqr_5d    = g(row, cmap, "pp_iqr_5d", 0.0),
            buycount_med_5d = g(row, cmap, "buycount_med_5d", 0.0),
            buycount_mad_5d = g(row, cmap, "buycount_mad_5d", 0.0),
            buycount_cv_5d  = g(row, cmap, "buycount_cv_5d", 0.0),

            wilson_lb_5d = g(row, cmap, "wilson_lb_5d", 0.0),

            ew_scheme = g(row, cmap, "ew_scheme", "", cast=str),
            ew_pp_5d  = g(row, cmap, "ew_pp_5d", 0.0),
            ew_hits_5d= g(row, cmap, "ew_hits_5d", 0.0),

            repeat_ticker_rate_5d = g(row, cmap, "repeat_ticker_rate_5d", 0.0),
            top_10_ticker_share_5d= g(row, cmap, "top_10_ticker_share_5d", 0.0),

            avg_buy_price_5d   = g(row, cmap, "avg_buy_price_5d", 0.0),
            median_buy_price_5d= g(row, cmap, "median_buy_price_5d", 0.0),
            avg_buy_volume_5d  = g(row, cmap, "avg_buy_volume_5d", 0.0),

            last_day_pp   = g(row, cmap, "last_day_pp", 0.0),
            last_day_buys = g(row, cmap, "last_day_buys", 0, cast=int),
            last_day_hits = g(row, cmap, "last_day_hits", 0, cast=int),

            min_daily_hits = g(row, cmap, "min_daily_hits", 0, cast=int),
            passed_consistency_gate = g(row, cmap, "passed_consistency_gate", 0, cast=int),
            passed_export_gate = g(row, cmap, "passed_export_gate", 0, cast=int),
        )
        m["pp_min_5d"]   = pp_min
        m["pp_max_5d"]   = pp_max
        m["pp_range_5d"] = pp_rng

        day5_json = g(row, cmap, "day5_buys_json", "[]", cast=str)
        try:
            friday_buys = json.loads(day5_json) if day5_json else []
        except Exception:
            friday_buys = []
        pack = dict(
            trial_number=trial,
            variation_number=varno,
            params=params,
            stocks=friday_buys, 
            day_pp=day_pp,
            metrics=m
        )
        return pack

    def marginal_gain(tset, current_sets):
        if not current_sets: return len(tset)
        union_curr = set()
        for s in current_sets: union_curr |= s
        return len(tset - union_curr)
    def coverage_select(dedup_sorted, cap):
        selected = []
        selected_sets = []

        for v in dedup_sorted:
            ts = ticker_set(v["stocks"])
            if marginal_gain(ts, selected_sets) > 0:
                selected.append(v); selected_sets.append(ts)
                if len(selected) >= cap:
                    break

        if len(selected) < cap:
            already = {id(v) for v in selected}
            fillers = [v for v in dedup_sorted if id(v) not in already]

            def overlap_to_selected(tset):
                if not selected_sets: return 0.0
                best = 1.0
                for s in selected_sets:
                    u = len(tset | s)
                    if u: best = min(best, len(tset & s) / u)
                return best

            scored_fillers = []
            for v in fillers:
                ts = ticker_set(v["stocks"])
                ov = overlap_to_selected(ts)
                pooled_pp = float((v.get("metrics") or {}).get("pooled_pp", 0.0) or 0.0)
                tno = int(v.get("trial_number") or 0)
                var = int(v.get("variation_number") or 0)
                scored_fillers.append((ov, -pooled_pp, -len(ts), tno, var, id(v)))

            scored_fillers.sort()
            id_map = {id(x): x for x in fillers}
            for _, _, _, _, _, vid in scored_fillers:
                v = id_map[vid]
                selected.append(v); selected_sets.append(ticker_set(v["stocks"]))
                if len(selected) >= cap:
                    break
        return selected

    # Storage for global backfill
    per_algo_selected = {}
    per_algo_nearmiss = {}

    for algo_index, algo_name in enumerate(ALGO_DISPLAY_NAMES):
        # Param headers
        header_row = 3 + (algo_index * 11)
        param_headers = []
        for j in range(10):
            val = control_sheet.range((header_row, param_start_col + j)).value
            if val:
                param_headers.append(val)
            else:
                break
        try:
            cmap = get_colmap(algo_name)
        except Exception as e:
            if "not found" in str(e).lower():
                if not hasattr(export_variations_to_excel, "_listed_tables_once"):
                    export_variations_to_excel._listed_tables_once = True
                    try:
                        cur.execute('SELECT name FROM sqlite_master WHERE type="table" ORDER BY name')
                        _tbls = [r[0] for r in cur.fetchall()]
                        print(f"[P3] Source DB available tables in {SOURCE_DB_FILE}: {_tbls}")
                    except Exception:
                        pass
            print(f"[{algo_name}] Table not found; skip.")
            continue

        cur.execute(f'SELECT * FROM "{algo_name}"')

        rows = cur.fetchall()
        if not rows:
            print(f"[{algo_name}] No rows; skip.")
            continue
        # Build pool
        all_vars = []
        for row in rows:
            pack = extract_row(row, cmap)
            if not pack:
                continue
            all_vars.append(pack)
        print(f"[P3] {algo_name}: fetched={len(rows)} usable_pre_gate={len(all_vars)}")
        print_gate_diagnostics(algo_name, all_vars, {
            "passed_consistency_gate": 1,
            "min_daily_hits": 1,
            "buycount_med_5d": 1,
            "wilson_lb_5d": 0.0,
            "median_pp_5d": 0.0,
            "pp_mad_5d": 25.0,
            "pp_iqr_5d": 80.0,
            "pp_range_5d": 80.0,
            "buycount_cv_5d": 1.0,
            "repeat_ticker_rate_5d": 100.0,
            "top_10_ticker_share_5d": 100.0
        })

        if not all_vars:
            print(f"[{algo_name}] No usable rows; skip.")
            continue
        cap = algo_cap(algo_name)

        base_pool = [v for v in all_vars if eligible_basic(v["metrics"])]

        for v in base_pool:
            v["final_order_score"] = export_score(v["metrics"])

        base_pool.sort(
            key=lambda v: (-v["final_order_score"], -len(v["stocks"]), -(v["metrics"].get("pooled_pp", 0.0) or 0.0))
        )

        pool_dedup = old_dedup_pool(base_pool, algo_name)
        print(f"[P3] {algo_name}: gate_pass={len(base_pool)} after_dedup={len(pool_dedup)} cap={cap}")

        if not pool_dedup:
            safe = [v for v in all_vars if (v["metrics"].get("min_daily_hits", 0) or 0) >= 1]
            safe.sort(key=lambda v: (-(v["metrics"].get("pooled_pp", 0.0) or 0.0), -len(v["stocks"])))
            pool_dedup = old_dedup_pool(safe[:max(50, cap)], algo_name)
            print(f"[P3] {algo_name}: SAFETY_NET used -> {len(pool_dedup)}")

        selected = coverage_select(pool_dedup, cap)

        dedup_sorted = sorted(
            pool_dedup,
            key=lambda v: (-(v["metrics"].get("pooled_pp", 0.0) or 0.0), -len(v["stocks"]))
        )

        selected = coverage_select(dedup_sorted, cap)

        # Score ordering
        for v in selected:
            v["final_order_score"] = export_score(v["metrics"])

        # Secondary tie-breaks
        def tie_key(v):
            m = v["metrics"]
            return (
                -v["final_order_score"],
                -m.get("pp_min_5d", 0.0),
                -(m.get("wilson_lb_5d", 0.0) or 0.0),
                m.get("pp_range_5d", 0.0) or 0.0,
                -(m.get("pooled_hits", 0) or 0),
            )
        selected.sort(key=tie_key)

        picked_ids = { (v["trial_number"], v["variation_number"], json.dumps(v["params"])) for v in selected }
        remain = [v for v in dedup_sorted
                  if (v["trial_number"], v["variation_number"], json.dumps(v["params"])) not in picked_ids]
        for v in remain:
            v["final_order_score"] = export_score(v["metrics"])

        nearmiss = sorted(remain, key=lambda v: (-v["final_order_score"]))

        per_algo_selected[algo_name] = dict(selected=selected, param_headers=param_headers)
        per_algo_nearmiss[algo_name] = nearmiss

        print(f"[P3] {algo_name} | selected={len(selected)} (cap={cap})")

    total_selected = sum(len(v["selected"]) for v in per_algo_selected.values())
    if total_selected < GLOBAL_MIN:
        deficit = min(GLOBAL_MIN - total_selected, OVERALL_MAX - total_selected)
        while deficit > 0:
            progressed = False
            for algo_name in ALGO_DISPLAY_NAMES:
                if deficit <= 0:
                    break
                cap = algo_cap(algo_name)
                pack = per_algo_selected.get(algo_name, {})
                sel = pack.get("selected", [])
                nm  = per_algo_nearmiss.get(algo_name, [])
                if len(sel) >= cap or not nm:
                    continue
                v = nm.pop(0)
                sel.append(v)
                per_algo_selected[algo_name]["selected"] = sel
                progressed = True
                deficit -= 1
            if not progressed:
                break

    # Trim if over limit
    total_selected = sum(len(v["selected"]) for v in per_algo_selected.values())
    if total_selected > OVERALL_MAX:
        all_items = []
        for a in ALGO_DISPLAY_NAMES:
            for v in per_algo_selected[a]["selected"]:
                all_items.append((v["final_order_score"], a, v))
        all_items.sort(reverse=True)
        keep_ids = set(id(v) for _, _, v in all_items[:OVERALL_MAX])
        for a in ALGO_DISPLAY_NAMES:
            per_algo_selected[a]["selected"] = [v for v in per_algo_selected[a]["selected"] if id(v) in keep_ids]

    for algo_index, algo_name in enumerate(ALGO_DISPLAY_NAMES):
        selected = per_algo_selected.get(algo_name, {}).get("selected", [])
        param_headers = per_algo_selected.get(algo_name, {}).get("param_headers", [])

        # Excel anchors per algo
        algo_col_start = start_cell_col + (algo_index * col_buffer_between_algorithms)
        param_col_start = algo_col_start + 5

        table_name = f"pruned_{algo_name}"
        sd_cur.execute(f"DROP TABLE IF EXISTS {table_name}")
        sd_cur.execute(f"""
            CREATE TABLE {table_name} (
                identifier TEXT PRIMARY KEY,
                algo_idx INTEGER,
                algo_name TEXT,
                trial_num INTEGER,
                var_num INTEGER,

                -- ranking
                final_order_score REAL,

                -- params & headers
                params_json TEXT,
                param_headers_json TEXT,

                -- Friday-only (for traceability)
                stocks_json_day5 TEXT,

                -- day-level
                day1_buys INTEGER, day1_hits INTEGER, day1_pp REAL, day1_buys_json TEXT,
                day2_buys INTEGER, day2_hits INTEGER, day2_pp REAL, day2_buys_json TEXT,
                day3_buys INTEGER, day3_hits INTEGER, day3_pp REAL, day3_buys_json TEXT,
                day4_buys INTEGER, day4_hits INTEGER, day4_pp REAL, day4_buys_json TEXT,
                day5_buys INTEGER, day5_hits INTEGER, day5_pp REAL, day5_buys_json TEXT,

                -- pooled & robust
                pooled_buys INTEGER, pooled_hits INTEGER, pooled_pp REAL,
                median_pp_5d REAL, pp_mad_5d REAL, pp_iqr_5d REAL,
                buycount_med_5d REAL, buycount_mad_5d REAL, buycount_cv_5d REAL,

                -- anchors & recency
                wilson_lb_5d REAL, ew_scheme TEXT, ew_pp_5d REAL, ew_hits_5d REAL,
                last_day_pp REAL, last_day_buys INTEGER, last_day_hits INTEGER,

                -- concentration
                repeat_ticker_rate_5d REAL, top_10_ticker_share_5d REAL,

                -- price/vol (kept for reference)
                avg_buy_price_5d REAL, median_buy_price_5d REAL, avg_buy_volume_5d REAL,

                -- derived extras
                pp_min_5d REAL, pp_max_5d REAL, pp_range_5d REAL,

                -- gates
                min_daily_hits INTEGER, passed_consistency_gate INTEGER, passed_export_gate INTEGER
            )
        """)
        sd_cur.execute(f'PRAGMA table_info("{table_name}")')
        _cols_info = sd_cur.fetchall()
        dest_cols = [r[1] for r in _cols_info] 

        def _build_record(algo_idx, algo_name, uid, v, param_headers, dr_tuple):
            m = v["metrics"]
            rec = {c: None for c in dest_cols}

            # identifiers
            rec["identifier"]  = uid
            rec["algo_idx"]    = algo_idx
            rec["algo_name"]   = algo_name
            rec["trial_num"]   = int(v["trial_number"])
            rec["var_num"]     = int(v["variation_number"])

            # score + params/headers
            rec["final_order_score"]  = float(v["final_order_score"])
            rec["params_json"]        = json.dumps(v["params"])
            rec["param_headers_json"] = json.dumps(param_headers)
            rec["stocks_json_day5"]   = json.dumps(v["stocks"] or [])

            d = list(dr_tuple) if dr_tuple else [0,0,0.0,"[]"]*5
            rec["day1_buys"], rec["day1_hits"], rec["day1_pp"], rec["day1_buys_json"] = int(d[0] or 0), int(d[1] or 0), float(d[2] or 0.0), d[3] or "[]"
            rec["day2_buys"], rec["day2_hits"], rec["day2_pp"], rec["day2_buys_json"] = int(d[4] or 0), int(d[5] or 0), float(d[6] or 0.0), d[7] or "[]"
            rec["day3_buys"], rec["day3_hits"], rec["day3_pp"], rec["day3_buys_json"] = int(d[8] or 0), int(d[9] or 0), float(d[10] or 0.0), d[11] or "[]"
            rec["day4_buys"], rec["day4_hits"], rec["day4_pp"], rec["day4_buys_json"] = int(d[12] or 0), int(d[13] or 0), float(d[14] or 0.0), d[15] or "[]"
            rec["day5_buys"], rec["day5_hits"], rec["day5_pp"], rec["day5_buys_json"] = int(d[16] or 0), int(d[17] or 0), float(d[18] or 0.0), d[19] or "[]"

            rec["pooled_buys"]      = int(m.get("pooled_buys",0)  or 0)
            rec["pooled_hits"]      = int(m.get("pooled_hits",0)  or 0)
            rec["pooled_pp"]        = float(m.get("pooled_pp",0.0) or 0.0)

            rec["median_pp_5d"]     = float(m.get("median_pp_5d",0.0) or 0.0)
            rec["pp_mad_5d"]        = float(m.get("pp_mad_5d",0.0)    or 0.0)
            rec["pp_iqr_5d"]        = float(m.get("pp_iqr_5d",0.0)    or 0.0)

            rec["buycount_med_5d"]  = float(m.get("buycount_med_5d",0.0) or 0.0)
            rec["buycount_mad_5d"]  = float(m.get("buycount_mad_5d",0.0) or 0.0)
            rec["buycount_cv_5d"]   = float(m.get("buycount_cv_5d",0.0)  or 0.0)

            rec["wilson_lb_5d"]     = float(m.get("wilson_lb_5d",0.0)  or 0.0)
            rec["ew_scheme"]        = str(m.get("ew_scheme","") or "")
            rec["ew_pp_5d"]         = float(m.get("ew_pp_5d",0.0)     or 0.0)
            rec["ew_hits_5d"]       = float(m.get("ew_hits_5d",0.0)   or 0.0)

            rec["last_day_pp"]      = float(m.get("last_day_pp",0.0)  or 0.0)
            rec["last_day_buys"]    = int(m.get("last_day_buys",0)    or 0)
            rec["last_day_hits"]    = int(m.get("last_day_hits",0)    or 0)

            rec["repeat_ticker_rate_5d"]  = float(m.get("repeat_ticker_rate_5d",0.0)  or 0.0)
            rec["top_10_ticker_share_5d"] = float(m.get("top_10_ticker_share_5d",0.0) or 0.0)

            rec["avg_buy_price_5d"]     = float(m.get("avg_buy_price_5d",0.0)     or 0.0)
            rec["median_buy_price_5d"]  = float(m.get("median_buy_price_5d",0.0)  or 0.0)
            rec["avg_buy_volume_5d"]    = float(m.get("avg_buy_volume_5d",0.0)    or 0.0)

            rec["pp_min_5d"]  = float(m.get("pp_min_5d",0.0)  or 0.0)
            rec["pp_max_5d"]  = float(m.get("pp_max_5d",0.0)  or 0.0)
            rec["pp_range_5d"]= float(m.get("pp_range_5d",0.0)or 0.0)

            rec["min_daily_hits"]          = int(m.get("min_daily_hits",0)          or 0)
            rec["passed_consistency_gate"] = int(m.get("passed_consistency_gate",0) or 0)
            rec["passed_export_gate"]      = int(m.get("passed_export_gate",0)      or 0)

            return rec

        rows_for_insert = []
        for i, v in enumerate(selected, 1):
            uid = f"No.{algo_index+1}.{i}"
            sql = (
                "SELECT "
                "day1_buys, day1_hits, day1_pp, day1_buys_json, "
                "day2_buys, day2_hits, day2_pp, day2_buys_json, "
                "day3_buys, day3_hits, day3_pp, day3_buys_json, "
                "day4_buys, day4_hits, day4_pp, day4_buys_json, "
                "day5_buys, day5_hits, day5_pp, day5_buys_json "
                f'FROM "{algo_name}" '
                "WHERE variation_number=? AND params=? "
                "ORDER BY id DESC LIMIT 1"
            )
            cur.execute(sql, (int(v["variation_number"]), json.dumps(v["params"])))
            dr = cur.fetchone()
            if dr is None:
                dr = (0,0,0.0,"[]", 0,0,0.0,"[]", 0,0,0.0,"[]", 0,0,0.0,"[]", 0,0,0.0,"[]")

            rec = _build_record(algo_index, algo_name, uid, v, param_headers, dr)
            rows_for_insert.append(tuple(rec.get(c) for c in dest_cols))

        if rows_for_insert:
            placeholders = ",".join(["?"] * len(dest_cols))
            col_list     = ", ".join(dest_cols)
            sd_cur.executemany(
                f'INSERT OR REPLACE INTO "{table_name}" ({col_list}) VALUES ({placeholders})',
                rows_for_insert
            )
            sd_conn.commit()

        row_cursor = start_cell_row
        positions = []
        for v in selected:
            stocks = v["stocks"] or [] 
            block_height = 2 + 1 + 1 + len(stocks) + row_buffer_between_variations
            positions.append({
                "start": row_cursor,
                "trial_num": int(v["trial_number"]),
                "var_num":   int(v["variation_number"]),
                "score":     round(v["final_order_score"], 2),
                "params":    v["params"],
                "stocks":    stocks,
            })
            row_cursor += block_height

        total_rows = row_cursor - start_cell_row
        if total_rows > 0:
            id_cols = 5
            id_matrix = [[None]*id_cols for _ in range(total_rows)]
            pr_cols = max(len(param_headers), 8)
            pr_matrix = [[None]*pr_cols for _ in range(total_rows)]

            for idx, pos in enumerate(positions, start=1):
                r0 = pos["start"] - start_cell_row
                stocks = pos["stocks"]
                block_height = 2 + 1 + 1 + len(stocks) + row_buffer_between_variations

                uid = f"No.{algo_index+1}.{idx}"
                id_matrix[r0] = [uid, block_height, pos["trial_num"], pos["var_num"], pos["score"]]

                pr_matrix[r0] = list(param_headers) + [None]*(pr_cols - len(param_headers))
                pvals = list(pos["params"])
                if len(pvals) > pr_cols: pvals = pvals[:pr_cols]
                pr_matrix[r0 + 1] = pvals + [None]*(pr_cols - len(pvals))

                rh = [
                    "Ticker","Name","Final Final Decision","Result of Buy or Not",
                    "Total Profit or Loss from Trade","Verbal Profit or Loss",
                    "Verbal Buy or Not","symbol selected"
                ]
                pr_matrix[r0 + 2] = rh + [None]*(pr_cols - len(rh))

                for j, s in enumerate(stocks):
                    rr = r0 + 3 + j
                    def safe(val): return val if val not in [None,"","null"] else ""
                    rd = [
                        safe(s.get("ticker")),
                        safe(s.get("name")),
                        "BUY",                         # day5 buys are buys
                        "",                            # no result here
                        "",                            # no P/L here
                        "",                            # no verbal result
                        "BUY",
                        s.get("symbol_selected") or safe(s.get("ticker"))
                    ]
                    pr_matrix[rr] = rd + [None]*(pr_cols - len(rd))

            output_sheet.range((start_cell_row, algo_col_start)).value = id_matrix
            output_sheet.range((start_cell_row, param_col_start)).value = pr_matrix

        print(f"[EXPORT] {algo_name}: {len(selected)} variations pasted.")

    # Close/Save
    sd_conn.commit(); sd_conn.close()
    conn.close()

    try:
        wb2.app.calculate()
    except Exception:
        pass
    wb2.save()
    try:
        app.api.Calculation = -4105 
    except Exception:
        pass

    app.display_alerts  = True
    app.screen_updating = True

    try:
        app.Interactive = True
        app.EnableEvents = True
        app.DisplayAlerts = True
        app.DisplayStatusBar = True
        app.ScreenUpdating = True
        app.Calculation = -4105
        try:
            for _sh in wb2.sheets:
                try:
                    _sh.api.ScrollArea = ""
                except Exception:
                    pass
        except Exception:
            pass
        app.api.DoEvents()
    except Exception:
        pass
    try:
        wb_mm = xw.Book(EXCEL_FILE)
        mm_control = wb_mm.sheets["ControlSheet"]
        mm_control.range("A:H").clear_contents()
        try:
            wb_mm.macro("AE1")()
        except Exception:
            try:
                wb_mm.app.macro("AE1")()
            except Exception:
                pass
        wb_mm.save()
    except Exception:
        pass

    total_selected = sum(len(v["selected"]) for v in per_algo_selected.values())
    print(f"\n[P3 COMPLETE] Exported total variations: {total_selected} (min {GLOBAL_MIN}, max {OVERALL_MAX})")
    
# End of MP1 Part 3
if __name__ == "__main__":
    export_variations_to_excel()
