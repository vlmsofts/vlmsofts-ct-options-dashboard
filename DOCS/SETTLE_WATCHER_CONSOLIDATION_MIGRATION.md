# Settle-Watcher → Single EOD-Bundle Consolidation — Migration & Coordination Plan

**Status: PLANNING / not started. settle_watcher stays LIVE and authoritative until parallel-validation passes.**
**Owners: dashboard side = ct-options-dashboard (this repo). Producer side = options-flow-analyzer.**
**Produced 2026-06-29 from a 3-agent cross-repo blast-radius investigation + live EOD-bundle inspection.**

---

## 0. THE GOAL (Lou)

Collapse the **7 fragile, intraday-raced, Excel-fed outputs** that `settle_watcher.py` produces into **ONE cadence-based EOD data bundle**, built from what the analyzer already captures at end of day:
- the **close time & sales** (options + futures blotter, ~2:25 ET CT / ~1:35 ET softs),
- the **full options blotter**,
- the **4:30 PM settlement poll** (settled surface: settles, Greeks, IV, OI per strike).

One source of truth, no Excel, no 3-minute polling race. Then **retire settle_watcher**.

**This is a MIGRATION, not a delete.** settle_watcher is load-bearing (see §2). It is retired ONLY after the consolidated bundle reproduces all 7 outputs byte-for-byte for several live sessions.

---

## 1. WHY (the trigger)

`settle_watcher.py` reads the **Excel ICE RTD workbook** (`read_ice_workbook('CT')` at settle_watcher.py:470/488/591) — the same COM-wedge / corrupt-strike fragility that took the dashboard down twice on 2026-06-29. Its intraday 3-min polling (SNAPSHOT 14:16, poll 14:25→16:00) is the fragile part. The EOD bundle is icepython-API-based (no Excel) and cadence-based (no race).

---

## 2. WHAT settle_watcher PRODUCES (the 7 outputs) — verified file:line

Written to **ct-options-dashboard/** unless noted. Schemas verified against live file headers 2026-06-29.

| # | Output | Live schema (verified header) | Timing |
|---|---|---|---|
| 1 | `local_futures_history.csv` | `date,commodity,contract,settle,yest_settle,change,high,low,volume,efp_vol,efs_vol,block_vol,open_int,oi_chg,first_notice,last_trade` | futures settle ~14:30 |
| 2 | `local_futures_spreads_history.csv` | `date,commodity,contract,settle,yest_settle,change,high,low,volume,efp_vol,efs_vol,block_vol` | with futures |
| 3 | `local_options_history.csv` | `date,commodity,security_des,contract_month,put_call,strike_px,open_int,oi_chg,px_settle,px_volume,vlm_key` | options settle ~14:45 |
| 4 | `settle_status.json` | `{date,futures_settled,futures_time,options_settled,options_time}` | both settle events + startup |
| 5 | `data/<date>/flow_rtd.json` (analyzer dir) | `{snapshot_time,mode,contracts:{ICE:{futures{bid,offer,last,settle,oi},options[{strike,call_vol,call_block,call_settle,call_oi,put_*}]}}}` | futures settle |
| 6 | `data/<date>/rtd_snap.json` (analyzer dir) | raw read_ice_workbook dict | 14:16 baseline |
| 7 | `settle_watcher.lock` | PID text | startup/exit |

KC sibling `settle_watcher_kc.py` writes the `local_kc_*` equivalents + `settle_status_kc.json` (KC options CSV is an 8-col compact schema; no flow_rtd/rtd_snap for KC). Timing: settle ~12:25, freeze ping 13:20, close refresh 13:35, stop 15:30.

---

## 3. CONSUMERS — the blast radius (why deletion breaks things)

**15+ consumers across BOTH repos.** Retiring settle_watcher without a replacement writer breaks:

**Analyzer-side (HARD):**
- `weekly_brief_runner.py` — **HARD ABORT** (no timeout) if `settle_status.json` options_settled≠true → Friday brief dies every week.
- `gex_calculator.py`, `pipeline/iv_snapshot.py`, `flow_watcher.py`, `outlook_watcher.py` — gate on `settle_status.json` (some have `--force`/timeout, some don't).
- `gateway_settle.py`, `forward_reconciliation.py`, `weekly_aggregator.py`, `eod_options_brief/run.py` — read `local_*_history.csv` / `flow_rtd.json`.

**Dashboard-side:**
- Straddle FREEZE-LIFT (app.py:1639), EOD-snapshot 409 gate (app.py:4699), write-guards (app.py:2440/2548), `/api/settle-status` (60s poll → settled banners + auto-reload), `/api/watcher-status` (90s poll → watcher-down banner), `_in_ct_settle_window` COM-yield (app.py:697/1621/2835/4094).
- Retire with no signal → **permanent watcher-down banner 95 min/day, EOD 409s all day, freeze stuck at 14:16, COM-yield blinds the dashboard 95 min/day.**

**Conclusion: the replacement bundle MUST reproduce all 7 outputs (esp. `settle_status.json`) or these break.**

---

## 4. WHAT THE EOD BUNDLE HAS TODAY vs THE GAP (verified from live files)

EOD bundle dir: `C:\Ice eod records\<COMMODITY>\<date>\`. Verified headers 2026-06-29:
- `settled_surface_<tenor>_<date>.csv`: `Date,Symbol,Right,Strike,Settle,ImpVol_decimal,Delta,Gamma,Theta,Vega,ThDelta,ThGamma,ThTheta,ThVega,Black76Vol,TheoPrice,CVol,Volume,OpenInt,Expiry,Forward Settle,Forward Contract` → **covers options** settle/Greeks/IV/OI/vol.
- `options_blotter_<tenor>_<date>.csv`: `Symbol,Right,Strike,Time,Exchange Time,Price,Size,Conditions,Seq Num` → options T&S.
- `enriched_options_ALL_<date>.csv`, `traded_iv_<date>.csv` → derived.
- **KC ONLY** has `futures_ohlc_<date>.csv`: `Date,Contract,Open,High,Low,Last,Settle,RecSet,Volume,OpenInt,Expiry`. **CT/SB/CC have NO futures file.**

### Gap table — to reach parity with the 7 outputs

| Need | In bundle? | Action |
|---|---|---|
| Options settle/IV/Greeks/OI/vol | ✅ settled_surface | reformat Symbol→security_des, Right→put_call, add contract_month + vlm_key |
| Options `oi_chg` | ❌ | compute = today OI − prior-day OI (prior CSV row) |
| **Futures settle/high/low/volume/OI** | ⚠️ KC only (futures_ohlc) | **port the futures pass to CT**; SB/CC too |
| **Futures EFP/EFS/block vol** | ❌ not even in KC's futures_ohlc | **ADD** — native `get_quotes` fields ("EFP Volume","EFS Volume","Block Volume") |
| Futures `yest_settle`/`change`/`oi_chg` | ❌ | compute from prior-day row (trivial diff) |
| Futures first_notice/last_trade | ❌ (Expiry present) | derive FND from expiry / contract spec |
| **Calendar spreads** | ❌ | **ADD** a spreads pass (consecutive-month pairs `CT m1:CTm2`) |
| **`settle_status.json`** | ❌ | **WRITE at end of the 4:30 pull** — the signal 5 pipelines gate on |
| flow_rtd.json | ❌ (different schema) | decide: keep producing, or migrate its consumers to the bundle |
| rtd_snap.json (14:16 freeze) | ❌ (pre-settle, dashboard freeze) | unrelated to 4:30; keep dashboard-side freeze mechanism |

**Verdict: viable. ~85% of futures already proven (KC futures_ohlc works). Remaining build = futures pass for CT + EFP/EFS/block fields + spreads pass + settle_status write + options reformat + prior-day diffs.** This is largely the analyzer's own `PLAN_eod_full_migration` (futures/spreads/HV passes) plus the settle_status signal.

---

## 5. DIVISION OF LABOR

### Analyzer owns (producer side — their files, their 4:30 pull)
1. Add **futures pass** to CT EOD capture (port from softs `futures_ohlc`); extend SB/CC.
2. Add **EFP/EFS/block volume** to the futures field list (native get_quotes fields).
3. Add **spreads pass** (consecutive-month pairs, NOT `**` autolist — it hangs).
4. Compute **yest_settle/change/oi_chg** on write (prior-day diff).
5. **Write `settle_status.json`** (+ `settle_status_kc.json`) at the end of the EOD pull, schema-identical (§2 #4).
6. **Reformat** the settled_surface → `local_options_history.csv` schema (Symbol→security_des, Right→put_call, +contract_month, +vlm_key, +oi_chg). KC → its 8-col compact schema.
7. Decide flow_rtd.json fate (keep, or migrate consumers).
8. Keep the 14:16 `rtd_snap.json` freeze write (dashboard depends on it) OR hand the freeze to the dashboard's own backup writer (app.py:1647 already writes it if a load occurs ≥14:16).

### Dashboard owns (this repo — consumer side)
1. Nothing changes until parity is proven. The dashboard already reads all 7 outputs; if the bundle writes them byte-identically, the dashboard is untouched.
2. Once retired: remove the now-vestigial `_in_ct_settle_window` COM-yield guards (app.py:697/1621/2835/4094) and the watcher-down banner logic — ONLY after settle_watcher is gone.
3. Order-of-write contract: the bundle MUST write `settle_status.json` (both flags true) **before/with** the CSVs, to avoid the dashboard's ≥16:30 auto-persist race (app.py:2030-2051).

---

## 6. PARALLEL-VALIDATION GATE (the "few sessions" Lou wants)

settle_watcher is NOT retired until ALL of these pass for **≥3 consecutive trading sessions** (CT + KC):

1. **Byte-diff**: bundle-written `local_futures_history.csv` / `local_options_history.csv` / `local_futures_spreads_history.csv` rows == settle_watcher's rows for the same date (settle, OI, vol, EFP/EFS/block, oi_chg all match).
2. **settle_status.json**: bundle writes it with correct date + both flags + times; the 5 analyzer gates + dashboard freeze-lift fire correctly.
3. **Dashboard render**: straddles, skew history, HV, IV-percentile, EOD email identical to settle_watcher-fed output.
4. **Analyzer pipelines**: GEX, IV snapshot, weekly brief, EOD brief all run green off the bundle.
5. **No race**: settle_status written before CSVs; no auto-persist overwrite.

Run BOTH systems in parallel (settle_watcher live + bundle writing to a shadow path) and diff, before cutting the dashboard/analyzer reads over to the bundle and stopping settle_watcher.

---

## 7. SEQUENCING

1. **Now**: dashboard live ATM:15 options pull (independent — does NOT touch settle_watcher or this migration). Validate it across sessions.
2. **Analyzer**: build §5 producer additions; write to a SHADOW path first.
3. **Parallel-validate** §6 for ≥3 sessions.
4. **Cut over**: point dashboard CSV reads + analyzer gates at the bundle.
5. **Retire** settle_watcher (stop the .bat / Task Scheduler task); remove vestigial COM-yield + banner.
6. Keep settle_watcher's code in git history; do not delete the file until cutover is proven stable for a further period.

---

## 8. CONSTRAINTS (carry into every step)
- QUOTE_BATCH ≤ 150 (get_quotes hangs above ~200). Main-thread COM. 32-bit Python for icepython.
- NEVER `get_autolist("**...")` (spreads) — it HANGS. Enumerate consecutive-month pairs.
- settle=RecSet (Settle is None intraday). yest_settle=PrevPrice. EFP/EFS/Block native. ImpVol decimal.
- October (V) excluded for CT. SB no Dec. KC 2.5c strike grid, 8-col compact options schema.
- **Additive + parallel-validate. Never break a working feature. settle_status.json schema is FROZEN — 5 consumers depend on it.**

---

## 9. OPEN QUESTIONS FOR ANALYZER
1. Will the EOD pull write `settle_status.json` itself, or should a thin signal-shim do it? (Timing: it must fire ~14:30/14:45 for the gates, but the full surface is 4:30 — does the signal need to lead the surface?)
2. flow_rtd.json: keep producing it, or migrate `run.py`/`residual.py`/`iv_snapshot`/`parse_rtd` to read the bundle?
3. The 14:16 rtd_snap.json freeze: analyzer keeps writing it, or hand to the dashboard's own backup writer?
4. Is the EOD-bundle close T&S (2:25/1:35) the right cadence to satisfy the intraday gates, or do the gates need an earlier settle signal?
