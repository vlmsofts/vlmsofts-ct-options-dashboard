# CT Options Dashboard — Change Catalog

---

## EOD snapshot date-shift fix + contract tickers + auto-backfill on options-settle (2026-06-12)

### [fix] eod_snapshot.json paired yesterday's options-date with today's futures settles

**Root cause:** `_assemble_eod_data()` returns `'date': d.get('last_date')`, where `last_date` is the max of the **options** history dates (`load_data`), but `ct1/2/3_settle` come from the **futures** data. Futures settle ~1 hr before options (observed 2026-06-11: futures 14:34 ET, options 15:37 ET). A snapshot written in that window stamps yesterday's options-date onto today's futures prices. Evidence: on-disk snapshot read `{date:2026-06-10, ct1_settle:72.49}` while CTN6 was 71.10 on 06-10 and 72.49 on 06-11. Separately, the snapshot carried no contract identity, so the downstream carry calc could not compute the true calendar-month gap.

**Fix (design):** Make the settle-watcher the canonical EOD writer — it POSTs to `/api/save-eod-snapshot` only once options-settlement is confirmed, so options-date == futures-date == today and the date-shift is structurally eliminated. Manual "Save Snapshot" and the Push-to-Site side-effect remain as fallbacks. Three surgical edits to `app.py`:

1. `_assemble_eod_data()` futures loop (~4157): added `'ticker': tkr,` to each `futures_rows` entry (raw ICE ticker, e.g. CTN6).
2. `_write_eod_snapshot()` `extracted` dict (~4388–4390): added `ct1_ticker`/`ct2_ticker`/`ct3_ticker` from `std_futs[0..2].get('ticker')` (None-guarded on length).
3. `/api/save-eod-snapshot` (~4467): after a successful `_write_eod_snapshot(data)`, call `_spawn_append_backfill()`. This endpoint uses `skip_if_same_date=False` so the write always succeeds; the unconditional spawn is safe because `append_backfill.py` has its own duplicate-date guard (no-ops on an already-present date).

Coordinated multi-repo change: the settle-watcher repo adds the options-settle POST; market-intelligence's `append_backfill.py` is updated separately (not touched here). Shared contract: snapshot gains `ct{1,2,3}_ticker`; watcher calls `http://127.0.0.1:5050/api/save-eod-snapshot`.

---

## Serial month straddle ATM vol identical to standard month (2026-06-05)

### [fix] CTU6 showing same ATM vol as CTZ6 post-settlement

**Root cause:** When no live bid/ask is present for a serial option chain post-settle, the B76 fallback in the straddle loop used `std_iv` (the standard month's ATM IV — e.g. CTZ6 = 22.01%) to price the serial's straddle. Back-solving that B76 value recovers approximately the same input vol regardless of T, so CTU6 and CTZ6 both displayed ~22.01%. The serial's own ATM IV (`atm_iv.get(ticker)`) was already computed from yesterday's settled CTU6 option prices (22.86%) but was never used in this path.

**Fix:** `app.py` serial B76 fallback now uses `atm_iv.get(ticker) or std_iv` as the vol input. CTU6 uses its own CSV-derived settled IV; falls back to the standard month IV only if the serial has no CSV history. Confirmed with flow analyzer: iv_snapshot.py independently computed CTU6 ATM IV = 22.86% from CTU6's own settled prices — consistent with the fix.

---

## Spreads H/L/V showing dashes — CSV fallback for rtd_spreads (2026-06-04)

### [fix] H/L/V columns always blank in spreads table

**Root cause:** The frontend builds spread rows from `D.rtd_spreads` (`spr = D.rtd_spreads`). When `read_spreads` in `ice_rtd_reader.py` returns empty (product name check or strip format mismatch on the ICE workbook), `rtd_spreads = {}` and the frontend's `addSpread` function always falls to the computed fallback (`_computed_spread`) which hardcodes `high:null, low:null, volume:null`. The settle_watcher writes H/L/V to `local_futures_spreads_history.csv` at ~14:31 but `load_data` never read that file.

**Fix:** Added CSV fallback block in `load_data` (`app.py` lines 1345–1385), mirroring the existing futures CSV fallback at lines 1303–1343. When `fut_last == today`, loads the most-recent row per contract from `LOCAL_SPR_HISTORY`. For keys absent from `rtd_spreads`, inserts a full record (display name, settle, yest_settle, change, pct_chg, H/L/V). For keys already present, fills null H/L/V from CSV. Frontend `spr[key]` now finds the record and renders H/L/V. October (CTV*) contracts are excluded by `_has_excluded_leg` in the frontend loop.

---

## Straddle strikes locked to yesterday's ATM — price tape fallback (2026-06-04)

### [fix] ATM strike used yesterday's futures settle when ICE workbook COM read failed

**Root cause:** The ICE workbook COM interface is shared with the price tape recorder (20s poll).
When the dashboard's `read_ice_workbook` call lands during a price tape poll cycle, COM returns
"Call was rejected by callee" and `_ice_raw` is set to `None`. With `_ice_raw = None`,
`ice_fut_row = {}` → futures bid/offer/last all `None` → `fwd` fell to `futures.get(ticker)`
(yesterday's CSV settle). Strike was then `atm_strike.get(ticker)` (computed from yesterday's
options CSV). On a day with a ~2-point move, CTH7 showed Strike=82 (yesterday settle=81.72)
instead of Strike=80 (live mid=79.80). Settlement column was correct (always from CSV), but
Value and % CHG Vol used the wrong ATM.

**Fix — four changes to app.py:**

1. **Price tape loaded before straddle loop:** Reads `ct_price_tape.csv` (updated every 20s by the
   price tape recorder) once and builds `_tape_live = {ticker: mid}` with the most recent row per
   contract.

2. **`fwd` fallback chain extended:** `bid/offer → last → price tape mid → settle(yesterday)`.
   The tape is used when ICE workbook COM misses, ensuring `fwd` reflects live market prices.

3. **`atm` from live `fwd` when `ice_chain` empty:** Replaced `atm_strike.get(ticker)` (yesterday's
   CSV ATM) with `ceil/floor(fwd)` using the standard rounding rule — same logic as serial months.

4. **`prev_atm = atm` — Settlement uses today's ATM strike:** Settlement column now looks up
   yesterday's straddle price AT TODAY's ATM strike (not yesterday's ATM). Change and % CHG Vol
   are same-strike comparisons. `fwd_settle` is retained only for `settle_iv_pct` back-solve
   (derives yesterday's implied vol at today's ATM strike using yesterday's forward).

5. **Serial-month `std_fwd` also gets tape fallback:** Same bid/offer → last → tape → settle chain
   applied to the standard-month forward used for serial ATM computation.

---

## Straddle snapshot moved to 14:16 + retry + dashboard backup (2026-06-03)

### [fix] rtd_snap.json was never written — COM race condition between price tape and settle_watcher

**Root cause:** settle_watcher's 14:18 snapshot read competed with the price tape recorder (running
every 20s on the same Excel COM object). At 14:16:26 and 14:22:53 the price tape logged
`RTD_READ_ERR: Call was rejected by callee`. settle_watcher's single read at 14:18:00 was silently
rejected. `rtd_snap.json` was never created, so the straddle freeze never activated. The dashboard
continued using live RTD in `prior_settle` mode — all mode-transition bugs (Bugs 4–6) remained.

**Fix — three layers:**

1. **settle_watcher.py — snapshot moved to 14:16 and retried 3×:** Target changed from 14:18 to
   14:16 (earlier than the typical COM error window). Added retry loop: up to 3 attempts, 20 seconds
   apart. On success writes `rtd_snap.json` and breaks. File renamed from `rtd_1418.json` to
   `rtd_snap.json` (time-agnostic).

2. **app.py — dashboard backup save:** If the dashboard reads live RTD successfully AND the current
   ET time is ≥ 14:16 AND `rtd_snap.json` doesn't exist yet → dashboard writes the file immediately.
   This guarantees the freeze activates even if settle_watcher misses the window entirely.

3. **Freeze load updated:** `_1418_path` renamed `_snap_path` pointing to `rtd_snap.json`.

**Files:**
- `Options_flow_analyzer/settle_watcher.py` lines 466–527: 14:16 target, 3× retry loop
- `ct-options-dashboard/app.py` lines 1490–1522: `rtd_snap.json` path, backup save block, freeze load

---

## Straddle 14:18 snapshot freeze (2026-06-03)

### [arch] Straddles freeze at 14:18 RTD snapshot until options settle

**Problem:** Repeated mode-transition bugs (Bugs 4–6 above) caused Value, Settlement, Change, and
% CHG VOL to be wrong after futures settled at ~14:31. Root cause: the RTD workbook transitions
from 'live' to 'prior_settle' mode at futures settlement, breaking all IV, prev_atm, and
fwd_settle calculations in the straddle loop.

**Solution:** settle_watcher already reads the RTD workbook at 14:18 in clean 'live' mode. Now
saves the full raw RTD snapshot to `data/YYYY-MM-DD/rtd_1418.json`. The CT dashboard loads this
file as `_ice_raw` for the straddle computation from 14:18 onward (regardless of live RTD state).
The freeze lifts when `settle_status.json → options_settled: true`.

**Scope:** Straddle tab and EOD email only. Live futures header prices, vol smile, skew history,
and all other dashboard sections use separate data paths — completely unaffected.

**The ~2-minute window (14:18–14:20 market close):** acceptable; market barely moves then.

**Files:**
- `Options_flow_analyzer/settle_watcher.py` lines 507–517: write `rtd_1418.json` at 14:18
- `ct-options-dashboard/app.py` lines 1486–1508: load snapshot, override `_ice_raw` if not settled

---

## Straddle Value: use live bid/ask in prior_settle mode, not call_settle (2026-06-03)

### [fix] Value column showed B76/call_settle after futures settled — options still trading

**Bug:** In `prior_settle` RTD mode (after 14:31 ET futures settlement), `today_c/today_p` were
read from `atm_row.get('call_settle')` = yesterday's option settle prices. The B76 override then
used yesterday's IV, giving a theoretical value (e.g. 9.88) that didn't match the live market
(Bloomberg showed 9.54 bid/ask mid). Root cause: the code assumed `prior_settle` = "market closed"
but options trade independently until ~14:50 ET and have live bid/ask throughout.

**Fix:** Always read `call_bid`/`call_offer`/`put_bid`/`put_offer` first, regardless of RTD mode.
Set `_val_from_live_bid_ask = True` when both sides available. Fall back to `call_last`, then
`call_settle` (only in `prior_settle` mode) if bid/ask absent. B76 override at lines 1638–1649
now guarded by `not _val_from_live_bid_ask` — skipped entirely when live prices were used.

**Files:** `app.py` lines 1546–1564 (today_c/today_p read), line 1638 (B76 override guard).

---

## Straddle Settlement column: use prev_atm (yesterday's strike) not live ATM (2026-06-03)

### [fix] Settlement column now stable when intraday move shifts ATM strike

**Bug (3rd occurrence):** Settlement column and % CHG Vol changed during the session when futures
moved enough to push the ATM strike to the next whole number (e.g. Sep 26: 80→81 intraday).
Root cause: `prev_c`/`prev_p` were read from `atm_row` — selected using the live forward `fwd`.
When ATM shifted, the settlement lookup row changed, producing a different Settlement value and
a wrong Change/% CHG Vol.

**Fix:** Compute `prev_atm` and `prev_atm_row` from `fwd_settle` (ICE RTD Settle column =
yesterday's published futures settle, static all day). All settlement straddle lookups now use
`prev_atm` throughout: RTD call_settle/put_settle read, flow_rtd.json key, CSV fallback strike
match, serial-month B76 fallback (new `_prev_atm_s2`), and settle_iv_pct implied_vol call.
Live ATM (`atm`) unchanged — still drives Value, Strike display, breakeven, and live IV.

**Files:** `app.py` lines 1674–1748 (straddle loop settlement section).

### [fix] fwd_settle post-settlement: use csv_prev_settle not RTD settle (2026-06-03)

**Bug:** Post-futures-settlement, `fwd_settle` was computed from `ice_fut_row.get('settle')` or
`futures.get(ticker)`. Both flip to today's settled price once the ICE RTD updates (~14:31 ET).
`prev_atm` was then derived from today's price, not yesterday's — same ATM-shift bug as above
but in the post-settlement phase. Result: Settlement column and DOD change wrong after 14:31.

**Fix:** When `post_settle=True`, use `csv_prev_settle.get((month_num, yr))` (prev_date CSV rows,
already loaded) as the primary source for `fwd_settle`. Serial months use standard-month key
`(std_m, std_y)`. Serial B76 fallback `_fwd_s2` also updated to `csv_prev_settle` first.
Pre-settlement path unchanged — RTD settle is still static/correct before 14:31.

**Files:** `app.py` lines 1683–1690 (`fwd_settle` derivation), line 1737 (`_fwd_s2` fallback).

---

## SINGLE-PROTOCOL SETTLEMENT — full cross-system implementation complete (2026-06-01)

### [arch] settle_watcher.py is sole authority on all settlements — both systems confirmed

**Rule:** `settle_watcher.py` (Options_flow_analyzer/) detects, writes, and signals all settlements.
No other process touches COM during 14:25–16:00 ET. No other process writes to shared CSVs
during the trading session. All consumers wait for `settle_status.json → options_settled: true`.

**Dashboard side (app.py):**
- `_in_ct_settle_window()` blocks all four `read_ice_workbook()` call sites during 14:25–16:00 ET
- `_auto_persist()` restricted to cold-start after 16:30 ET when settle_watcher did not run
- `LOCAL_OPT_HIST` NameError fixed → `LOCAL_OPT_HISTORY` (line 668, `_csv_opt_settle` now loads correctly)

**Flow analyzer side:**
- `flow_watcher.py`: five COM functions deleted (`_read_rtd_current`, `_bootstrap_baseline`,
  `_load_prior`, `_detect_settlement`, `_try_snapshot`). Settlement detection replaced with
  `_read_options_settled()` polling `settle_status.json` every 5 minutes. Zero COM access.
- `price_tape.py`: `_in_settle_window()` COM pause 14:25–16:00 ET live (was already written)
- `iv_snapshot.py`: `build_enriched_iv()` gates on `options_settled: true` before reading
  `local_options_history.csv` — prevents partial read during settle_watcher's delete-then-append
- `weekly_brief_runner.py`: settle check upgraded from warning to hard block (CompletenessCheckError);
  checks `options_settled` not just `futures_settled`

Both scheduled tasks restarted and confirmed Ready. First live test: 2026-06-02 settlement window.

---

## COM COLLISION FIX — all processes yield ICE COM to settle_watcher 14:25–16:00 ET (2026-06-01)

### [fix] price_tape.py pauses COM reads during settlement window

**File:** `Options_flow_analyzer/price_tape.py`

`price_tape.py` runs from 05:30 AM to 15:30 ET, polling `read_ice_workbook()` every 20
seconds. During 14:25–15:30, it fired ~65 COM reads straight through the settlement
detection window — the primary source of `Call was rejected by callee` errors that
bounced settle_watcher off the workbook and caused `options=no` to persist for hours.

**Fix:** Added `_in_settle_window()` helper and a pause block in the main loop. During
14:25–16:00 ET the loop sleeps instead of polling, logs `SETTLE_WINDOW: pausing` once
on entry and `SETTLE_WINDOW: ended - resuming` on exit. settle_watcher has zero
competition for COM during its detection window.

## COM COLLISION FIX — dashboard yields ICE COM to settle_watcher during settlement window (2026-06-01)

### [fix] Dashboard stops calling read_ice_workbook() during 14:25–16:00 ET

**File:** `app.py` (new `_in_ct_settle_window()` helper + guards at lines ~580 and ~1452)

**Root cause:** Both the dashboard (`load_data()` on every `/api/data` request, ~3 min cycle)
and `settle_watcher.py` called `ice_rtd_reader.read_ice_workbook('CT')` through Excel's
single-threaded COM interface. When they collided, Excel rejected one caller with
`Call was rejected by callee` or `NoneType has no attribute 'UsedRange'`. This forced
settle_watcher to retry after 180s — causing repeated missed settlement detection windows
and the observed `futures=YES options=no` state persisting far past expected settlement time.

**Fix:** Added `_in_ct_settle_window()` — returns True during 14:25–16:00 ET on CT trading
days. Both COM call sites in `load_data()` now check this flag and skip the workbook read
entirely during that window. The dashboard falls back to `_flow_rtd_opts` (flow_rtd.json)
and B76 theoretical for straddle values. settle_watcher has exclusive COM access during
the settlement detection period.

**One settlement process:** settle_watcher.py is the sole settlement detection + persistence
process. The dashboard is read-only during the window.

---

## CONFLICT AUDIT — app.py CSV write isolation (2026-06-01)

### [fix] _auto_persist() restricted to cold-start bootstrap only

**File:** `app.py` (~line 1714)

`_auto_persist()` previously fired on every `/api/data` request during trading hours,
writing to `local_options_history.csv` and `local_futures_history.csv` via
`_persist_ct_options_ice()` and `_persist_futures_ice()`. This violated the rule that
`settle_watcher.py` is the sole writer to those files.

**Fix:** Added cold-start guard. Auto-persist now runs only when ALL are true:
1. Current ET time is after 16:30 (settle_watcher hard stop + buffer)
2. `settle_status.json` does not exist or shows a prior date (settle_watcher did not run)
3. ICE workbook has contracts not in the CSV

During trading hours the guard always suppresses the thread launch.

### [fix] _csv_opt_settle parser handles both security_des formats

**File:** `app.py` (~line 642)

The parser previously required the embedded `security_des` format (`CTN6P    75`) written
by `settle_watcher`. Rows written by `_persist_ct_options_ice` use a short format
(`CTN6P`) with strike in the separate `strike_px` column — these were silently skipped,
making the CSV settle fallback unreliable for cold-start rows.

**Fix:** Parser now uses `strike_px` column as primary strike source; falls back to
parsing embedded format only when `strike_px` is empty. Both formats now populate
`_csv_opt_settle` correctly.

---

## SETTLEMENT PIPELINE REBUILD + EOD EMAIL COMPLETE (2026-05-28)

### [arch] New master settlement detection process — settle_watcher.py

**Files changed:**
- `Options_flow_analyzer/settle_watcher.py` — complete rewrite (was unused old version)
- `Options_flow_analyzer/settle_watcher.bat` — new Task Scheduler trigger
- `Options_flow_analyzer/SETTLEMENT_PIPELINE.md` — full handoff document
- `ct-options-dashboard/app.py` — removed `_persist_ice_all()`, scheduled fetch timer; added `/api/settle-status`
- `ct-options-dashboard/templates/index.html` — auto-poll + dismissible settlement banner
- `ct-options-dashboard/local_futures_history.csv` — schema expanded, bbg_ticker removed
- `ct-options-dashboard/local_futures_spreads_history.csv` — new file

**Architecture:**
`settle_watcher.py` is the single owner of settlement persistence. It runs daily
from 14:25 ET, polls ICE RTD FEED CT.xlsx every 3 minutes, and detects settlement
when the `Settle` column changes by ≥ 0.01 from the prior CSV row.

**The Settle column rule:** It is static during the session (= yesterday's settlement).
It updates once ICE publishes today's settlement (~14:30 ET futures, ~14:45 ET options).
This is the detection signal.

**On futures settlement:** writes `local_futures_history.csv`, `local_futures_spreads_history.csv`, `flow_rtd.json`, updates `settle_status.json`.
**On options settlement:** writes `local_options_history.csv`, updates `settle_status.json`.
**Hard stop:** 16:00 ET.

**Two-watcher distinction:**
- `settle_watcher.py` → settlement detection + CSV writes + dashboard notification
- `outlook_watcher.py` → Gmail flow email watcher + options flow analysis pipeline (separate, independent)

**CSV schema (futures):**
`date, commodity, contract, settle, yest_settle, change, high, low, volume, efp_vol, efs_vol, block_vol, open_int, oi_chg, first_notice, last_trade`

**CSV schema (spreads — new file):**
`date, commodity, contract, settle, yest_settle, change, high, low, volume, efp_vol, efs_vol, block_vol`

**Computed at write time:** `yest_settle` = prior CSV row; `change` = settle − yest_settle; `oi_chg` = oi − prev_oi.

**Dashboard integration:**
- `/api/settle-status` endpoint reads `settle_status.json`
- Frontend polls every 60s; auto-reloads data + shows dismissible gold banner on settlement
- `_ld_cache` invalidates on CSV file mtime change → fresh data on next refresh

**EOD email chain (complete):**
settle_watcher writes CSVs → `_ld_cache` invalidates → `D` reloads fresh →
user clicks EOD button → `_buildEodData()` assembles from `D` →
all fields populated (straddles, futures, spreads, HV) → canvas PNG → push/email.

**Removed from app.py:**
`_persist_ice_all()`, `_schedule_settle_fetch()`, `_run_scheduled_fetch()`,
`SETTLE_FETCH_HOUR/MINUTE`, `/api/fetch-settles` route.

**Kept:** `_persist_today()` GitHub backup fetch, `_schedule_preclose_flush()`.

**Known data issue (2026-05-28):**
Corrupt row (Bloomberg ticker in date col) blocked CSV updates for months.
May 27 futures data was missing — manually inserted correct ICE values.
Resolved going forward by settle_watcher.

---

## ICE RTD CHANGE/SETTLE/YEST_SETTLE FIX + EOD EMAIL BRANDING (2026-05-28)

### [⚠️] Futures `yest_settle` and `change` were wrong — root cause: RTD `Change` col is last-trade based
**Files:** `app.py` (live_futures + rtd_spreads override), `templates/index.html` (_buildEodData)

**Root cause — three distinct layers:**

**Layer 1 — ICE RTD `Change` column (col S) is last-trade based, not settle-to-settle.**  
`Change = last_trade − S_{t−1}` where `S_{t−1}` is ICE's internal reference (approximately
yesterday's settle but NOT the official published settlement). The formula
`yest_settle = settle − change` therefore gives a wrong `yest_settle`.  
Example: CTZ6 settle=76.16, Change=−1.09 (last-trade based) → computed yest_settle=77.25,
but correct yest_settle=77.37.

**Layer 2 — ICE RTD `settle` column bug (ICE's bug, fixed by ICE support 2026-05-27).**  
Before ICE's fix the `settle` column showed *yesterday's* settlement price, not today's.
After their fix `settle` = today's published settlement. This was a server-side ICE bug;
no code change needed once they fixed it.

**Layer 3 — Pre-settlement vs post-settlement distinction.**  
Before today's CT session settles (~3 PM ET), RTD `settle` still holds yesterday's
published settlement. Using `change = settle − csv_yest_settle` therefore gives `0` because
both values are yesterday's settle. The correct intraday change is `last − csv_yest_settle`.

**Fixes applied:**

1. **`app.py` — build `csv_prev_settle`** after `fut_lookup`. Reads `ct_fut` rows for
   `prev_date` (the trading day before `last_date`) to get yesterday's official settlement for
   each contract, keyed by `(month_num, year)` — same key scheme as `fut_lookup`.

2. **`app.py` — `yest_src` selection:** if `last_date == today` (settlement already fetched),
   use `csv_prev_settle`; else use `fut_lookup` values (which already hold yesterday's settle).

3. **`app.py` — `live_futures` override (outrights):** after RTD read, override `yest_settle`
   from `yest_src`; compute `change = last − yest_settle` (NOT `settle − yest_settle`).
   Using `last` is correct for intraday display; `settle` = yesterday's value pre-settlement.

4. **`app.py` — `rtd_spreads` override:** same pattern — override `yest_settle` from near/far
   outright yest_settle difference; recompute `change` and `pct_chg`.

5. **`index.html` — EOD canvas (`_buildEodData`):** computes `change = settle − yest_settle`
   directly for both futures rows and RTD spread rows (settle-to-settle for EOD email).
   Does NOT use the `change` field from `live_futures`, which is intraday last-trade based.

**Rule to remember:**
- Live ticker / tab display → `change = last − yest_settle` (intraday move from yesterday close)
- EOD email → `change = settle − yest_settle` (settle-to-settle, computed client-side)
- Never use RTD `Change` col as-is for either; always recompute from `csv_yest_settle`.

---

### [⚠️] Straddle `Settlement` column wrong — CSV `px_settle` has stale/wrong values
**File:** `app.py` — CT straddle loop, `prev_c`/`prev_p` lookup (~line 1361)

**Root cause:** The straddle settlement straddle was computed from `ct_opts` CSV rows
(`px_settle` field) at `settle_ref = last_date`. The CSV can be stale or hold prices that
don't match ICE's official published option settlements. The RTD options sheet carries
`call_settle`/`put_settle` fields that are always current and confirmed correct.

**Fix:** Swap priority — try RTD `atm_row.call_settle` / `atm_row.put_settle` first;
fall back to CSV `px_settle` only when RTD has no data for that contract.

**Rule:** RTD option settle fields are authoritative. CSV is backup only.

---

### [fix] Spread year labels showing `Mar07/May07` instead of `Mar27/May27`
**File:** `templates/index.html` — `parseTkr()` inside `_buildEodData()`

**Root cause:** ICE tickers use single-digit years (`CTH7` = Mar 2027). `parseInt('7') = 7`.
The old pivot `yy >= 30 ? 1900+yy : 2000+yy` gave `2000+7 = 2007`. Python's
`parse_ct_ticker` already uses `2020 + year_digit` — JS was inconsistent.

**Fix:** `yy < 10 → 2020 + yy` added as the first branch in the year resolution chain,
matching `parse_ct_ticker` exactly.

---

### [feat] EOD Email PNG — VLM brand palette applied to all exports
**Files:** `templates/index.html` — `exportStraddlePanel()`, `_buildEodCanvas()`,
and all surface/smile/skew export functions.

All PNG exports now share a single palette:
- Header background: `#1a1a2e` (dark navy, matches COT report)
- Title accent: `#E8C547` (VLM gold, matches COT report)
- Column bar: `#2c3e50`
- Positive: `#15803d` | Negative: `#b91c1c`
- Gold accent strip (4 px) separates header from body
- VLM footer bar on all exports

Previously each export had its own ad-hoc colours (`#0f1923`, `#c9a227`, `#1a3a6e`, `#2b63b8`).

---

## PUSH TO SITE — FLASK PROXY FIX (2026-05-26)

### [fix] "Failed to fetch" CORS error on PUSH TO SITE button
**Files:** `app.py` (new route), `templates/index.html` (JS URL + header)  
**Root cause:** Browser at `http://127.0.0.1:15050` cannot POST directly to `https://vlmdata.com`
— browsers block cross-origin requests unless CORS headers are present on the remote server.

**Fix:**
- Added Flask route `POST /push-to-vlm` in `app.py` — receives multipart form data from
  the browser and forwards it to `https://vlmdata.com/api/analysis/push` server-side using
  `requests`, injecting the `x-push-secret` header. Secret never leaves the server.
- JS `_PUSH_URL` changed from the external URL to `'/push-to-vlm'` (same-origin).
- `x-push-secret` header removed from JS `fetch()` call entirely.

---

## SETTLEMENT IV T_SETTLE ROOT-CAUSE FIX (2026-05-26)

### [⚠️] % CHG ON DAY showed wrong sign — e.g. CTN6 −1.26 when vol was actually UP
**File:** `app.py` — straddle loop, `_actual_settle_str` block (~line 1155)
**Root cause (confirmed by exact reproduction):**
The GitHub options CSV is published **one business day late** — the rows stored under
`last_date` in `local_options_history.csv` contain the **previous trading session's**
settlement prices, but the date label is TODAY (the publication date).
For example, on May 26 the CSV has rows labeled `2026-05-26` that hold May 22 settlement
prices (call=1.81, put=2.39 at K=78 for CTN6).  
`last_date = '2026-05-26'` → `T_settle = (Jun 12 − May 26)/365 = 17/365`.  
But the settlement was actually priced when DTE = 21 (May 22 → Jun 12).  
Using 17/365 instead of 21/365 inflates `settle_IV` from 28.15% to 31.27%,
giving `% CHG = 30.01 − 31.27 = −1.26` — wrong sign, wrong magnitude.

**Exact parameter reproduction confirming the bug:**
| fwd_settle | T_settle | settle_IV | % CHG |
|---|---|---|---|
| 77.42 (ICE RTD ✓) | 17/365 (bug) | 31.27% | **−1.26** |
| 77.42 (ICE RTD ✓) | 21/365 (fix) | 28.15% | **+1.86** |

**Fix:** Before the straddle loop, compute `_actual_settle_str` = the last CT trading day
before today by walking back from yesterday using `_is_ct_trading_day()`:
```python
_prev_day = datetime.now() - timedelta(days=1)
while not _is_ct_trading_day(_prev_day.strftime('%Y-%m-%d')):
    _prev_day -= timedelta(days=1)
_actual_settle_str = _prev_day.strftime('%Y-%m-%d')
```
`T_settle` now uses `_actual_settle_str` instead of `last_date`. On May 26 this gives
`_actual_settle_str = '2026-05-22'` → T_settle = 21/365 → CHG = +1.86 ✓.

**Why it cannot revert:** The computation is dynamic — it always finds the correct prior
trading day regardless of what date labels the GitHub CSV uses, including future holidays.
The `_is_ct_trading_day()` function already holds the full 2026 ICE holiday calendar.

**Supersedes:** The earlier fix in "STRADDLE VOL-CHANGE T-MISMATCH FIX" that used
`T_settle from last_date` was incomplete — it only worked correctly when `last_date`
happened to equal the actual settlement date (which it never does intraday).

---

## STRADDLE TWO-VOL + SERIAL ATM FIXES (2026-05-26)

### [⚠️] Vol Smile overlay T-mismatch → two different IVs for same contract
**File:** `app.py` — live smile overlay loop (~line 1002)
**Root cause:** The Vol Smile live overlay used `get_dte(ticker, last_date)` for its T, giving
21/365 (May 22 → Jun 12). The straddle loop (fixed earlier) uses today's actual T = 17/365.
Same straddle price back-calculated at two different T values → two different IVs on screen
(27.4% header vs 30.39% table for CTN6).
**Fix:** Changed to `get_dte(ticker, datetime.now().strftime('%Y-%m-%d'))` so the live overlay
T always reflects today's actual DTE, matching the straddle table.

### [⚠️] Serial month ATM strike wrong (CTU6 showing 79, should be 80)
**File:** `app.py` — CT straddle loop, serial month branch (~line 1228)
**Root cause:** `atm` is selected before the serial month branch runs, using the stale
`futures.get('CTU6')` forward (~79.X from May 22 CSV). `std_fwd` (live CTZ6 = 80.28) is
computed inside the branch but `atm` was never updated from the stale value.
**Fix:** After establishing `std_fwd`, re-derive `atm` using the official rounding rule
(fractional ≥ 0.50 → ceil/upper, < 0.50 → floor/lower). Also refreshes `atm_row` so the
settlement fallback uses the correct strike row. CTZ6 at 80.28 → .28 < .50 → ATM = 80 ✓.
**Rule:** This rounding rule is applied only in the serial month branch because standard months
use nearest-chain-strike which gives the same result for $1-interval chains.

---

## STRADDLE VOL-CHANGE T-MISMATCH FIX (2026-05-26)

### [⚠️] Straddle % CHG on Day — Wrong T in Settlement IV Back-Calculation
**File:** `app.py` — CT straddle loop, `_persist_ice_all()`
**Root cause (two-part):**
1. `_persist_ice_all()` ran on Memorial Day (May 25, ICE closed) and wrote 478 option rows +
   10 futures rows dated `2026-05-25` that contained unchanged May 22 Friday settlement prices.
   This made `last_date = '2026-05-25'`, so DTE was computed as 18 (Jun 12 − May 25) instead
   of the correct 21 (Jun 12 − May 22 when the prices were actually set).
2. The straddle loop used the same `T = (lt − last_date).days / 365` for BOTH the live IV and
   the settlement IV back-calculation, so neither was computed with the correct DTE.
**Impact on CTN6 (Jul 26, 18 DTE):** settle IV inflated by √(21/18) = 1.08 → showed −0.57
vol pts when vol actually went UP from Friday. Short-dated contracts most affected.
**Fixes applied:**
- Added `_ICE_CT_CLOSED` (2026 ICE Cotton holiday set) and `_is_ct_trading_day()`.
- `_persist_ice_all()` now skips entirely on weekends and ICE holidays — prevents bad date rows.
- Straddle loop now computes `T` from `datetime.now()` (correct DTE for live IV) and
  `T_settle` from `last_date` (correct DTE for settlement IV back-calculation) independently.
- Deleted the 488 stale May 25 rows from both local CSVs.
**DO NOT** revert to a single `T` for both live and settlement IV — that will re-introduce the
multi-day-gap error on any day before the 3:45 PM fetch updates `last_date`.

---

## STRADDLE & PNG FIXES (2026-05-22)

### [⚠️] CT Serial Month Straddles — Live B76 Value (CTU6, CTX6, CTF7, etc.)
**File:** `app.py` — CT straddle loop (~line 1188)
**Root cause:** Serial CT months (Aug/Sep/Nov/Jan/Feb/Apr/Jun) have no dedicated futures contract.
Their ICE RTD option strips contain frozen `call_settle`/`put_settle` values; the old code used
these as `val`, so the displayed straddle value never moved intraday (Change = 0 all day).
**Fix:** Added an `if month_num not in CT_STANDARD_MONTHS:` branch that runs BEFORE the
`today_c/today_p` check. This path always uses B76 with the underlying standard month's live
IV (`atm_iv[std_tkr]`) and live forward (bid/ask mid → last → settle fallback from ICE RTD
futures row). `val` therefore moves with the market during the session.
Last-resort fallback: if `std_iv` or `std_fwd` are unavailable, uses frozen settle prices.
**Affected months:** CTU6 (Sep26→Dec26), CTX6 (Nov26→Dec26), CTF7 (Jan27→Mar27), etc.
**DO NOT** move the `CT_STANDARD_MONTHS` check below the `today_c/today_p` check — serial months
will silently revert to frozen settlement values.

### [✅] atm_row UnboundLocalError — Fixed
**File:** `app.py` — CT straddle loop (~line 1158)
**Root cause:** `atm_row` was only assigned inside `if fwd and ice_chain:`, not in the `elif fwd:` branch.
Accessing it at line 1241 (`if prev_c is None and atm_row:`) crashed the server.
**Fix:** Added `atm_row = None` unconditionally before the `if fwd and ice_chain:` block.

### [✅] CTU7 / CTX7 (Sep27, Nov27) Removed from CT Straddle Display
**File:** `app.py` — CT straddle loop (~line 1132)
**Change:** `straddle_tickers = [t for t in straddle_tickers if t not in {'CTU7', 'CTX7'}]`
**Why:** Illiquid; including them added rows with unreliable values. Re-add when liquidity warrants.

### [✅] `_ice_to_rtd_shape` — call_settle/put_settle Passed Through
**File:** `app.py` — `_ice_to_rtd_shape()` (~line 483)
**Root cause:** The adapter that converts the raw ICE workbook dict to the generic RTD shape
was only forwarding `bid`, `ask`, `last` — not settle prices. Deferred contracts with no live
bid/offer were getting `mid=None` with no settle fallback.
**Fix:** Added `sk` (settle key) to the inner loop: `('Call','call_bid','call_offer','call_last','call_settle')`.
Settle is now forwarded to each strike entry and available for the fallback logic downstream.

### [✅] Generic Straddle Loop — Settle Fallback for Deferred Contracts (KC/SB/CC)
**File:** `app.py` — generic straddle loop (~line 2440)
**Change:**
1. `today_c/today_p` mid calculation now falls back to `s.get('settle')` when bid/offer absent.
2. `prev_c/prev_p` lookups now fall back to RTD `call_settle`/`put_settle` when the contract
   is not yet in the local CSV history (avoids missing Settlement/Change columns for back months).

### [⚠️] PNG Exports — WhatsApp Readability (All Commodities)
**File:** `templates/index.html` — `_drawSkewPng()`, `exportSurfacePanel()`
**Problem:** At WhatsApp display width (~400–800px), fonts below ~12px logical were unreadable
and low-contrast colours (`MUTED='#6b7280'`, `BORDER='#d1d5db'`) were invisible on phones.
**Fix — `_drawSkewPng` (Skew History PNG):**
- Canvas: W 840→1100, hdrH 64→80, chartH increased to 390, stats/footer rows enlarged
- All fonts: axis labels NF(14), legend B(15), latest values B(19), delta labels B(16)
- Line weight: 2.5→4.0
- MUTED: '#6b7280'→'#374151', BORDER: '#9ca3af', GRID: '#cbd5e1'
**Fix — `exportSurfacePanel` (Vol Surface PNG):**
- Canvas: W 840→1100, hdrH 64→80, rowH 22→30, tblHdrH 28→36, skewH 130→180, footH 30→44
- All fonts: headers B(22), section titles B(13), table data B(12-13), HV cells N(12)
- IV column width: 57px→82px (no more clipping)
- MUTED: '#6b7280'→'#374151', BORDER: '#9ca3af'
- Footer shaded background '#e8ecf0'
Both exports are commodity-agnostic — improvements apply to CT/KC/SB/CC automatically.
**Awaiting user verification.**

### [✅] `_drawStradStrip` — Label Fonts Increased
**File:** `templates/index.html` — `_drawStradStrip()`
**Change:** N(9)→N(11) for `ATM IMPLIED VOL` / `LIVE` / `SETTLE` / `ATM STRADDLE` labels;
B(13)→B(14) for the contract ticker. Improves strip readability on WhatsApp exports.

### [✅] generatePng / exportTradePanel — Hardcoded 'CT' Label Fixed
**File:** `templates/index.html` — `generatePng()` line ~2683, `exportTradePanel()` line ~2245
**Fix:** Replaced hardcoded `'CT Options Analytics'` with `(D.commodity_name||'CT Options Analytics')`
so the header reflects the active commodity (KC/SB/CC) instead of always showing CT.

---

## PHASE 1: ICE RTD AS PRIMARY LIVE SOURCE + DAILY SETTLE PERSISTENCE (2026-05-21)

### [✅] ICE RTD wired as primary; Bloomberg RTD inactive fallback

**Files:** `app.py`, `ice_rtd_reader.py`

**What changed:**
- `ice_rtd_reader.py` — added COM thread-safety wrapper (`pythoncom.CoInitialize/CoUninitialize`)
  around the public entry point so xlwings calls are safe from Flask worker threads.
- `app.py` imports `ice_rtd_reader` at startup. `_ice_to_rtd_shape()` adapter converts the
  `{'mode', 'futures', 'options'}` dict to the same shape `rtd_reader.read_live()` returned,
  so all downstream IV/greeks logic is unchanged.
- `load_data('CT')` and `/api/debug` now try ICE RTD first; `rtd_reader` (Bloomberg) is the
  inactive fallback only called when ICE workbook is unavailable.

**Daily settle persistence (Bloomberg-free going forward):**
- `_persist_ct_options_ice(ice_data, today_str)` — writes CT option rows using the *new* CSV
  format (`security_des='CTN6C'`, `strike_px=82.5`) to `local_options_history.csv`.
- `_persist_generic_options_ice(...)` — same for KC/SB/CC to their respective CSVs.
- `_persist_futures_ice(...)` — writes ordinal contract rows (`CTJUL1`) derived from ICE contract
  codes using the exact same formula as `get_hist_fwd`. Copies `last_trade`/`first_notice` from
  existing CSV rows via `_build_lt_fn_lookup`. All functions are idempotent (skip if date exists).
- `_persist_ice_all()` — orchestrates all 4 commodities, clears skew-history cache after write.
- Auto-scheduler: `threading.Timer` fires at **20:45 local** (= 15:45 ET / BST-adjusted) daily,
  reschedules itself. Daemon thread — does not prevent server shutdown.
- `POST /api/fetch-settles` — manual trigger; returns `{'status': 'ok', 'results': {...}}` keyed
  by commodity with row counts or `'unavailable'`/`'error: ...'` strings.

---

## ICE RTD PIPELINE + BLOOMBERG HISTORICAL NORMALISATION (2026-05-21)

### [✅] ICE RTD reader for CT (parallel pipeline, existing CT unchanged)

**Files:** `ice_rtd_reader.py` (new), `test_ice_pipeline.py` (new), `normalize_bbg_softs.py` (new)

**Architecture:**
- `ice_rtd_reader.py` — standalone xlwings reader for `ICE RTD FEED CT.xlsx` (and KC/SB/CC when ready).
  Returns `{'mode': 'live'|'today_settle'|'prior_settle'|'unavailable', 'futures': {...}, 'options': {...}}`.
  Three-mode detection: `live` (Market State == 'Open'), `today_settle` (new ICE settle differs from stored),
  `prior_settle` (market closed, using yesterday's data).
- `test_ice_pipeline.py` — 8-section cross-check: workbook open, sheet discovery, futures parsing,
  mode detection, per-strip option chains, ATM IV term structure, Bloomberg CSV comparison, OI sanity.
  Bloomberg comparison confirmed **0.0 bps difference** on CTN6/CTU6 ATM strikes vs local CSV.
- `normalize_bbg_softs.py` — one-time conversion of Bloomberg softs options/futures CSVs to the
  column format app.py expects. Input: `coffee/sugar/cocoa_*_history.csv` from Downloads.
  Output: `local_kc/sb/cc_*_history.csv` in this folder.

**Normalisation detail:**
- Options: `KCN6C 257.5 Comdty` → `security_des='KCN6C'`, `strike_px=257.5`, `px_settle=PX_LAST`.
  PX_LAST is last-traded price (not official settle); close proxy for liquid near-ATM strikes.
- Futures: `KCH6 Comdty` → ordinal contract `KCMAR3` where ordinal matches `get_hist_fwd`'s
  formula exactly (`ordinal = delivery_year - first_year_from_row_date + 1`). Year derived from
  `last_trade` (always populated by Bloomberg pull — fallback ordinal parse never triggers).

**Result:** KC/SB/CC skew history extended from 2026-05-01 back to **2024-01-02** (585/600/463 dates).

**ICE RTD NOT yet wired into app.py** — tested in isolation only. Existing CT pipeline untouched.

---

## MULTI-COMMODITY EXPANSION — KC / SB / CC (2026-05-20)

### [✅] KC Coffee, SB Sugar, CC Cocoa added alongside CT

**Files:** `app.py`, `templates/index.html`, `bootstrap_other_softs.py` (new)

**Architecture:**
- `COMMODITY_CONFIG` dict centralises per-commodity paths, standard months, serial maps, expiry overrides.
- `/api/data?commodity=KC` and `/api/skew-history?commodity=KC` routes added.
- `_load_generic_data(commodity)` handles KC/SB/CC (settle-only, no RTD/live smile).
- `_persist_today_generic(commodity)` auto-appends new GitHub rows daily for each commodity.
- `compute_skew_history(commodity='CT')` parameterised; per-commodity cache in `_skew_hist_cache` dict.

**Expiry logic (critical — do not change):**
- **KC options:** `_kc_opt_expiry(month, year)` — 8 biz days before first biz day of delivery month.
  *NOT* `last_trade` from futures CSV (KC futures last_trade is in delivery month = ~32 days wrong).
- **SB options:** `fut_lookup[(month, year)]['last_trade']` — SB July futures last_trade = June 30.
- **CC options:** `fut_lookup[(month, year)]['last_trade']` — same pattern.
- Futures key uses `first_notice`-derived delivery month (not `last_trade.month`) because SB/CC
  futures expire in the month before delivery — `last_trade.month` ≠ delivery month.

**Frontend:** commodity dropdown in topbar; `loadAndRender(commodity)` destroys charts and re-fetches;
CT default on fresh load; last selection persisted in `localStorage`.

**Data:** Local CSVs seeded from OI dashboard GitHub via `bootstrap_other_softs.py`.
Options history from 2026-05-01 (Bloomberg pre-history pending); futures from 2008.

**CT unchanged** — all CT routing, RTD, live smile, expiry dates identical to before.

---

Running log of fixes, what was broken, how it was fixed, and current status.
**Before touching any of these areas, read this file first.**

---

## VOL SURFACE — LIVE DATA + BAR/TABLE CONSISTENCY (2026-05-20)

### [⚠️] Vol Surface skew bar now uses live data when Bloomberg connected
**File:** `templates/index.html` — `loadSurface()`, `exportSurfacePanel()`
**Issue 1:** Skew bar in Vol Surface panel used `D.skew_value` (settle-based backend value) while the
delta-bucket table rows used `getIVatDelta()` which prefers live smile when Bloomberg is connected.
This caused a visible discrepancy (e.g. bar = −6.7, table 25D P−C implied −5.6).
**Issue 2:** Export PNG always labelled "SETTLE" and always computed IVs from settle prices even
when live Bloomberg data was available.
**Fix:**
- `loadSurface()`: after computing `cells` for each ticker, derive `_surfSkew[ticker]` = cells[2]−cells[6]
  (25D Put − 25D Call) using the same live-preferring `getIVatDelta()` source.  Skew bar now uses
  `_surfSkew` instead of `D.skew_value`. Bar and table are always identical.
- `loadSurface()`: card title updates to "Vol Surface — Live" / "Vol Surface — Settle" based on
  whether any ticker has live smile data present.
- `exportSurfacePanel()`: `ivAtDelta()` now interpolates from `D.live_smile` first (same logic as
  `getIVatDelta()`), falls back to settle.  Uses `D.live_smile_fwd` for forward when available.
  Section title label changes from "SETTLE" to "LIVE" when live data used.
  Skew bar in PNG also derived from cells (consistent with table).
**DO NOT** revert to `D.skew_value` for the skew bar — it will re-introduce the discrepancy.

---

## SKEW PNG — Δ1d / Δ1w CHANGE ROW (2026-05-20)

### [⚠️] Day-over-day and week-over-week changes added to skew export PNGs
**File:** `templates/index.html` — `_drawSkewPng()`
**Change:** Added a second stats row below the "Latest:" line showing Δ1d and Δ1w changes.
- Δ1d = last non-null value − second-to-last non-null value (1 trading day)
- Δ1w = last non-null value − 6th-to-last non-null value (~5 trading days)
- 2-series charts (10d/25d/35d): shows Call Δ, Put Δ, and C−P Spread Δ for both periods
- ATM chart: shows Δ1d and Δ1w for ATM IV
- Values colour-coded: green = positive, red = negative
- `footH` increased from 28 to 52 to accommodate the extra row

---

## STATUS LEGEND
- ✅ VERIFIED WORKING — confirmed by user or test
- ⚠️ APPLIED, NOT YET CONFIRMED — fix implemented, awaiting user verification
- ❌ BROKEN — known issue, not yet fixed
- 🔁 REVERTED — undone or superseded

---

## CORE DATA PIPELINE

### [✅] RTD / Bloomberg Live Feed
**File:** `rtd_reader.py`
**Issue:** Flask worker threads do not auto-initialise a COM apartment.
`win32com.client.GetActiveObject()` raised `CO_E_NOTINITIALIZED` silently,
falling through to openpyxl (XLSX mode) on every request even with Excel open.
**Fix:** `_read_via_com()` now calls `pythoncom.CoInitialize()` / `CoUninitialize()`
around the inner COM logic (extracted to `_read_via_com_inner()`).
**Verified:** Thread-test from worker thread confirmed `live_rtd` source.
**DO NOT** remove the `CoInitialize` call — it will immediately revert to XLSX mode.

### [✅] Live Smile — High-IV Spikes (far-dated, deep OTM)
**File:** `app.py` — live smile section (~line 799)
**Issue:** Far-dated options (e.g. CTZ7, 543 DTE) have delta > 0.03 even at deep
OTM strikes, so stale no-market quotes passed the delta filter and spiked the smile.
**Fix (two layers):**
1. Relative IV bound: reject any strike where IV < 50% or > 250% of settle ATM IV.
2. Minimum price floor: `px < 0.05` rejected.
**DO NOT** remove `iv_lo`/`iv_hi` or lower the 50% floor — spikes will return.

### [✅] Live Smile — Suppressed for Illiquid Contracts
**File:** `app.py` — live smile section (~line 802 and ~line 836)
**Root cause:** Far-dated options (CTZ7, 543 DTE) have too few BQL quotes
(7 strikes) and stale/inconsistent prices. Filters on individual points cannot
fix a dataset where the overall shape contradicts the market structure (e.g.
drawing a valley on a contract the dashboard itself labels "CALLS BID").
**Fix (three layers in order):**
1. Require raw bid AND ask both >= 0.02, ask <= bid * 8 (no last-trade fallback)
2. Two-tier IV bounds: delta > 0.15 → [75%, 130%] of ATM IV; delta <= 0.15 → [50%, 250%]
3. **Minimum 8 clean strike points** — if fewer survive filters, live smile is
   suppressed entirely for that ticker. Liquid contracts (CTN6, CTZ6) have 15+.
**Result:** CTZ7 shows settle smile only (reliable, exchange-computed).
**DO NOT** lower the minimum to < 8 — the wrong U-shape will return.
**DO NOT** remove the two-tier IV bounds — near-ATM stale quotes will slip through.
**Verified by user 2026-05-19.**

### [✅] ATM IV — Straddle Method
**File:** `app.py` — `atm_iv_for_date()` (~line 539)
**Issue:** Single-option IV is noisier than straddle when F ≠ K.
**Fix:** Uses call+put straddle: `call_eq = (strad + (F-K)*e^{-rT}) / 2`, then
solves IV on the call-equivalent. Falls back to single option if one side missing.
**DO NOT** revert to single-option only.

### [✅] Forward Price — Put-Call Parity
**File:** `app.py` — expiry loop (~line 465)
**Issue:** Futures CSV settle can lag intraday moves; options data was
inconsistent with stale futures price.
**Fix:** Median implied forward from call-put pairs: `F = K + (C-P)*e^{rT}`.
Applied for today, prev_date, and week_date separately.
**DO NOT** replace with raw CSV settle as the forward.

### [✅] DTE — ICE Cotton Expiry Dates
**File:** `app.py` — `ICE_CT_EXPIRY` dict + `get_dte()`
**Issue:** Computed "last Friday of preceding month" formula is ~14 days wrong
for Cotton (ICE uses a different rule).
**Fix:** Hardcoded Bloomberg-sourced dates in `ICE_CT_EXPIRY` dict. RTD 'all
options' sheet can override for new contracts not yet in the dict.
**DO NOT** replace these with the computed `option_expiry_date()` formula.

---

## FUTURES BAR

### [✅] Non-Standard Month Filtering
**File:** `templates/index.html` — `buildFuturesBar()` (~line 741)
**Issue:** Serial option months (CTU6, CTX6, CTF7, etc.) were appearing in the
ticker bar. ICE Cotton No. 2 standard futures only trade H/K/N/V/Z.
**Fix:** `CT_FUT_MONTHS = new Set(['H','K','N','V','Z'])` filter on expiry list.
**DO NOT** remove this filter — non-standard months will reappear.

### [✅] OI / Change Data in Futures Bar
**File:** `templates/index.html` — `buildFuturesBar()`
**Issue:** OI and day-change data only populate when RTD is live (Bloomberg
connected). They are empty in XLSX/CSV fallback mode — this is expected,
not a bug. Was mistakenly attributed to a code regression.

---

## HEADER STRIP

### [✅] ATM IV Header — RTD Straddles Sheet Override
**Files:** `app.py` (~line 911); `templates/index.html` — `updateHeader()` (~line 962)
**Root cause:** RTD straddles sheet is keyed "Dec26"/"Jul26" format. Prior merge code used `'CT' + contract_code` producing "CTDec26" — never matched any ticker in `expiry_list`. `atm_vol_rtd` was therefore never populated and `D.atm_iv[e]` always showed settlement value (24.8%) instead of live Bloomberg (24.55%).
**Fix (backend):** After the live smile for-loop (~line 911), added a standalone `if rtd:` block that reads `rtd.get('straddles')` directly. Uses `_STRAD_MONTH` dict to map "Dec" → "Z", "Jul" → "N" etc., constructs "CTZ6"/"CTN6" etc., and overwrites `atm_iv[tkr]` with `sd['atm_vol'] * 100`. This block runs AFTER the live smile computation so it wins over stale BQL data.
**Fix (frontend):** `_isLive = !!_lsm` (Bloomberg connected when live smile populated). `liveAtmIV = D.atm_iv[e]` — trusts the backend value, which is now the live Bloomberg straddle vol when RTD is connected.
**Priority chain:** CSV settle → live bid/ask straddle IV → RTD straddles sheet (Bloomberg live, wins last).
**DO NOT** revert `liveAtmIV` to `_lsm[atmK]` — the BQL all-options sheet is a static snapshot; its per-strike IVs equal settlement values and are NOT live streaming.
**DO NOT** remove the `_STRAD_MONTH` block or the `if rtd:` wrapper around it.
**DO NOT** change the key format assumption — Bloomberg RTD straddles sheet Contract column uses "Dec26"/"Jul26" not "Z6"/"N6".

### [⚠️] ATM Straddle Price Tile — Correct K + Live IV
**File:** `templates/index.html` — `intel-strip` HTML + `updateHeader()`
**What it does:** Tile between "ATM Implied Vol" and "IV Change" showing the straddle price
in cents per pound. Sub-label shows the ATM strike used.
**Data flow:** Backend (`app.py` live smile section ~line 869) computes `strad_mid = call_bid/ask_mid + put_bid/ask_mid` at `live_atm` (nearest strike to live forward), solves IV → stores as `D.atm_iv[e]`. This is the raw bid/ask straddle the user wants.
**Critical K rule:** ATM strike K for repricing must be nearest strike in `D.live_smile[e]` to the live forward. `D.atm_strike[e]` is NOT used — it can be overridden by Bloomberg RTD to a different strike (e.g. K=84 when forward=83.03), causing a wrong `(F−K)` displacement of ~1¢ and incorrect IV lookup.
**JS implementation:** K = `liveSmileKeys.reduce(nearest to stradF)`. σ = `D.atm_iv[e] / 100`. F = `D.live_smile_fwd[e]` || `D.futures[e]`. Formula: `straddle = 2 × Black76.call(F,K,T,r,σ) − (F−K)×e^{−rT}` — recovers `call_mid + put_mid` exactly.
**DO NOT** use `D.atm_strike[e]` as K — Bloomberg override can be 1 strike off the forward, adding ~1¢ error.
**DO NOT** use `D.live_smile[e][K]` as σ — that is an OTM-selected IV per strike, not the straddle IV. `D.atm_iv[e]` is the correct straddle IV.
**DO NOT** use raw `D.options_today.px` sums — those are CSV settle units and produce wrong values.
**Awaiting user verification.**

---

## VOL SMILE PANEL

### [✅] IV Changes Table — Uses Live IV When Available
**File:** `templates/index.html` — top-strikes table
**Issue:** Table used settle IV for `ivNow` but header used live BBG mid,
making changes look wrong.
**Fix:** `ivNow` now checks `liveSmileMap[k]` first before falling back to
settle IV, matching the header's live ATM display.

### [✅] Smile Expiry Selector Syncs With Surface Panel
**File:** `templates/index.html` — `smileExpChange()` / `surfaceExpChange()`
**Issue:** Vol Surface panel was locked to CTN6 regardless of selected expiry.
**Fix:** Bidirectional sync: smile selector change pushes to `#surface-expiry`
and vice versa. Tab switch also syncs.

---

## PNG EXPORT — ALL TABS

### [⚠️] Straddle Info Strip on All Export PNGs
**File:** `templates/index.html` — `_drawStradStrip()` + all 4 export functions (~line 2045)
**What it does:** A 44px dark strip drawn immediately below the navy header on every export PNG (Vol Smile, Vol Surface, Trade Analyzer, Skew History). Shows: ATM IV % (green when live, white when settle) + LIVE/SETTLE tag on the left; ATM Straddle ¢ (gold) on the right; K and F values centered.
**Shared helpers:** `_stradInfo(ticker)` → `{iv, strad, K, F, isLive}` (reuses same Black76 + live smile logic as `updateHeader()`). `_drawStradStrip(cx, ticker, x, y, w)` renders the strip.
**Ticker resolution per tab:** Smile = `#exp-expiry`; Trade = `#smile-expiry`; Surface = `#surface-expiry`; Skew = `_skewActive` (rolling → first expiry with ATM IV data).
**Canvas heights updated:** All 4 functions add `_STRIP_H` (44) to their canvas height to accommodate the strip.
**DO NOT** remove `_STRIP_H` from the canvas height calculations — the strip will overflow into content below.
**Awaiting user verification.**

---

## PNG EXPORT — VOL SMILE

### [✅] Export Defaulted to CTN6 Regardless of Selected Expiry
**File:** `templates/index.html` — `openExport()`
**Fix:** `openExport()` now reads from `smile-expiry` selector before opening modal.

### [⚠️] Export Overlays Respect Toggle State
**File:** `templates/index.html` — `generatePng()` (~line 2290)
**Issue:** `prevData` and week data were hardcoded — always rendered regardless
of `smileOverlays.prev` / `smileOverlays.week` toggle state.
**Fix:** Both datasets gated on overlay state. Week line added (red `#dc2626`).
**Awaiting user confirmation.**

---

## PNG EXPORT — VOL SURFACE

### [⚠️] Serial Option Months Included in Surface Export
**File:** `templates/index.html` — `exportSurfacePanel()` (~line 2036)
**Change:** Removed `CT_SURF_MONTHS = new Set(['H','K','N','V','Z'])` filter. Serial option months (CTU6/Sep, CTX6/Nov, CTF7/Jan) now appear in the delta-bucket table alongside standard months.
**HV columns:** Show `—` for serial months — correct, since there is no dedicated futures HV for these contracts. Delta-bucket IVs still render using the underlying futures price from `CT_SERIAL_FUTURES` mapping.
**DO NOT** re-add the `CT_SURF_MONTHS` filter — serial months have active option trading and should appear.
**Awaiting user verification.**

### [✅] Skew Title Clipped — textAlign Bug
**File:** `templates/index.html` — `exportSurfacePanel()` skew section
**Issue:** `surfRows.forEach` ended with `cx.textAlign='center'`. The skew
section title at `PAD+12` was being center-anchored, putting most of the text
off the left edge of the canvas.
**Fix:** `cx.textAlign='left'` reset before drawing the skew title.

### [✅] Skew Bars Overflowing Bottom
**File:** `templates/index.html` — `exportSurfacePanel()` skew geometry
**Issue:** `zeroY` was near the bottom of the chart; negative skew (calls bid)
bars grew downward past the canvas boundary.
**Fix:** Symmetric layout: `zeroY` centred with `skewTopPad=10` / `skewBotPad=24`.
Scale: `(skewH - skewTopPad - skewBotPad) / 2 / skewMax`.

### [✅] Unicode Characters Clipping Section Titles
**File:** `templates/index.html` — `exportSurfacePanel()`
**Issue:** `◆` (U+25C6) and `Δ` (U+0394) have no glyph in Arial on Windows
canvas, causing text to be mispositioned.
**Fix:** Replaced with ASCII: `>>` for section marker, `D` for delta symbol.

### [✅] DTE Sub-Label Overflowing Row
**File:** `templates/index.html` — `exportSurfacePanel()` row loop
**Issue:** DTE label drawn at `y+rowH+0` — below the row boundary, clipping
into the next row.
**Fix:** Sub-label removed entirely (DTE is readable from expiry + date header).

### [✅] HV Columns Empty in Export
**Note:** HV data only populates when RTD is live (Bloomberg connected).
Empty HV in the surface export is EXPECTED in XLSX/CSV fallback mode.
Now resolved since RTD COM fix restores live feed.

---

## SKEW HISTORY TAB

### [✅] Full Local Data Migration — No GitHub Dependency
**Files:** `app.py` — `load_data()`, `compute_skew_history()`, `/api/debug`, `_persist_today()`, `read_local_csv()`; `bootstrap_local.py` (one-time); `local_options_history.csv`; `local_futures_history.csv`
**What changed:** All `fetch_csv(OPT_CSV_URL)` and `fetch_csv(OI_CSV_URL)` calls replaced with `read_local_csv(LOCAL_OPT_HISTORY)` / `read_local_csv(LOCAL_FUT_HISTORY)`. `_persist_today()` now writes raw GitHub-format rows (full column set) rather than the old simplified format.
**Bootstrap:** `bootstrap_local.py` run 2026-05-19 — downloaded full history from GitHub. Options: 37,397 CT rows, 354 dates (2025-01-03 → 2026-05-19). Futures: 46,905 CT rows, 4,700 dates (2008-01-02 → 2026-05-19). Spot-checks MATCHED.
**Local file formats:** Options matches `options_oi.csv` (`date, commodity, security_des, contract_month, put_call, strike_px, open_int, oi_chg, px_settle, px_volume`). Futures matches `oi_data.csv` (`date, commodity, contract, bbg_ticker, settle, open_int, oi_chg, first_notice, last_trade`).
**Verified:** Dashboard starts, `/api/data` returns correct `last_date`, expiries, ATM IVs, `data_source: live_rtd`.
**Rollback:** Change 4 lines in `load_data()`, `compute_skew_history()`, `/api/debug` back to `fetch_csv(OPT_CSV_URL)` / `fetch_csv(OI_CSV_URL)`. `fetch_csv()` is still in place and untouched.
**Daily data flow:** `_persist_today(last_date)` is called on every `load_data()`. It checks if `last_date` is already in local files (fast, no network). If it's a new trading day, it fetches from GitHub (one HTTP call per CSV), appends only today's CT rows, and logs the count. Page loads are always local-only; GitHub is only hit once per new day.
**DO NOT** delete `local_options_history.csv` or `local_futures_history.csv` — these are the accumulated history.
**DO NOT** delete `fetch_csv()` — still used by `_persist_today()` for daily updates and by `/api/debug-rtd`.

### [⚠️] Skew History Endpoint + Tab
**Files:** `app.py` — `compute_skew_history()` + `/api/skew-history`; `templates/index.html` — `panel-skew`
**What it does:**
- API returns `{rolling: {...}, tickers: {CTN6: {...}, ...}}` — rolling + per-contract series
- Rolling front-month = nearest H/K/N/Z with DTE ≥ 30; per-ticker includes all months ≥ 10 dates
- Ticker selector pills: ROLLING + individual months (N6, U6, X6, Z6, H7, K7 etc.)
- 4 separate charts in 2×2 grid: ATM | 10Δ C/P | 25Δ C/P | 35Δ C/P
- Hover shows ⛶ expand icon; click any chart to open fullscreen overlay — title includes contract (e.g. "25Δ Call vs Put — CTZ6")
- PNG button: export modal with 4 selectable tiles; click = single select, Ctrl+click = multi-select
- Each selected chart exports as a separate WhatsApp-ready PNG (navy header, white background, VLM branding)
- Export PNGs include "Latest" stats bar: call vol (green), put vol (red), C−P differential (gold)
- 1-hour server-side cache; lazy fetch on first tab open
**Awaiting user verification.**

### [✅] Skew History — Wrong Forward Price (IV Accuracy)
**File:** `app.py` — `compute_skew_history()` (~line 1111)
**Issue:** `compute_skew_history()` used raw CSV futures settle as the forward for delta
calculations on every historical date. The Vol Surface panel (which user confirmed correct)
uses a put-call parity corrected forward via `load_data()`. The discrepancy caused skew
history delta-IV values to diverge significantly from the Vol Surface and Bloomberg OVDV.
**Fix:** After reading the CSV settle, the function now overrides `fwd` with the median
implied forward from call-put pairs on that date: `F = K + (C-P)*e^{rT}` — identical
logic to `load_data()`. Falls back to CSV settle if fewer than 3 pairs exist.
**DO NOT** remove the parity override — raw CSV settle lags intraday and diverges from
option-implied forward, producing wrong deltas and systematically biased skew values.
**Verified by user 2026-05-19.**

### [✅] CT_SERIAL_FUTURES Mapping — October Excluded Correctly
**File:** `app.py` — `CT_SERIAL_FUTURES` dict (~line 261)
**Rule:** Aug (8) and Sep (9) serial option months anchor to **December futures (12)**.
October futures (CTV) exist on ICE but anchor to NO option months — October is excluded
from all metrics entirely. The mapping `{8: 12, 9: 12}` is correct and must not be changed.
**DO NOT** map Aug or Sep to October (10) — CTU6/CTX6 options price off Dec futures, not Oct.

### [✅] Skew History Tooltip — C−P Differential
**File:** `templates/index.html` — `_skewOpts()` tooltip callbacks (~line 2791)
**Fix:** Added `afterBody` callback to the shared `_skewOpts()` tooltip config. For 10Δ/25Δ/35Δ
charts (two datasets), the tooltip now shows a `Spread (C−P): ±X.X vol pts` line after the
individual call/put values. ATM chart (one dataset) is unaffected.
**Applies to:** both the 2×2 grid charts and the fullscreen expand overlay (both use `_skewOpts`).

### [✅] PNG Export Button — Routes to Wrong Panel
**File:** `templates/index.html` — `openExport()` (~line 1973)
**Issue:** The global EXPORT button and the smile-panel PNG button both call `openExport()`.
That function only checked for `panel-trade` and `panel-surface` — when the Skew History tab
was active it fell through to the smile export modal.
**Fix:** Added `if(activePanel.id === 'panel-skew') { openSkewExport(); return; }` check.
**DO NOT** remove the `panel-skew` routing or add any other default — the smile modal must
only open when the smile panel is actually active.

### [⚠️] Vol Surface + Skew History — Live Delta IVs from Live Smile
**File:** `templates/index.html` — `getIVatDelta()` in `loadSurface()` (~line 1947); `_liveSkewPoint()` (~line 2856)
**Issue:** Both tabs used CSV settle option prices (`getChainPx` / `D.options_today`) for delta strike IVs (10/15/25/35Δ). ATM was live (from `D.atm_iv`) but delta strikes were stale — producing a mixed live/settle picture.
**Fix:** Both functions now prefer `D.live_smile[ticker]` (live bid/ask mid IVs) for delta interpolation:
1. Build (delta, iv) pairs from live smile strikes using Black76.delta
2. Sort and interpolate to target delta
3. Fall back to settle (`getChainPx` / `D.options_today`) when live smile is absent (illiquid back months like CTZ7)
**Result:** Front months (CTN6, CTZ6, CTH7, CTK7) get fully live delta IVs. CTZ7+ keep settle. `_liveSkewPoint` also uses `D.live_smile_fwd` (parity forward) for consistency.
**DO NOT** remove the live smile preference — without it, delta IVs in the surface and skew history are from yesterday's settle while ATM is live.

### [⚠️] Skew History — Live Sync for Front Contracts (Today's Last Data Point)
**File:** `templates/index.html` — `_liveSkewPoint()` + `_renderAllSkewCharts()`
**Issue:** The Skew History has a 1-hour server cache and never uses RTD data. Settlement
data arrives next day at 9:30am via OI CSV. During the open session, the "today" data point
was always yesterday's settle, diverging from the Vol Surface (which is live/60s refresh).
**Fix:** When the last date in the series matches `D.last_date`, `_liveSkewPoint()` recomputes
all 7 delta-IV values (ATM, 10/25/35Δ call+put) using `D.options_today` + `D.futures` —
identical formula to `compute_skew_history()`. The last array element for each series is patched
in a copy (original `_skewData` not mutated). Front contracts (in `D.atm_iv`) get live values;
K7+ are not in `D.atm_iv` so `_liveSkewPoint()` returns null and they keep yesterday's settle.
**Data flow:** OI CSV settlements daily before 3pm → CSV updated at 9:30am next day → all
historical points in the series use settle. Only the "today" point is live-patched.
**DO NOT** patch historical dates — only the last point when it equals `D.last_date`.
**Outlier guard:** if any delta IV in the live point exceeds 150% of ATM IV, the entire
live point is discarded and yesterday's settle is kept. Threshold is intentionally extreme —
normal skew never approaches 150% of ATM; only stale/bad option prices would trigger it.
**DO NOT** lower the 150% cap — legitimate cotton skew (calls bid on supply squeeze) can
push wing IVs well above ATM without being bad data.
**Awaiting user verification.**

### [✅] Expand Chart Title Shows Contract
**File:** `templates/index.html` — `expandSkewChart()` (~line 2836)
**Fix:** Title now reads `"<metric> — <contract>"` e.g. `"25Δ Call vs Put — CTZ6"` using
`_skewActive`. Rolling shows `"25Δ Call vs Put — Rolling"`.
**DO NOT** hardcode the title — it must reflect whatever ticker pill is currently active.

---

## INFRASTRUCTURE

### [✅] Flask Auto-Reload on OneDrive Path
**File:** `app.py` — server startup
**Fix:** `reloader_type='stat'` (polling-based) instead of default watchdog,
which fails on OneDrive network paths on Windows.

### [⚠️] pywin32 Added to requirements.txt
**Fix:** `pywin32>=306` added. Was missing, making the dependency implicit.

---

## PROCESS RULES — NEVER VIOLATE

- **Never claim a fix is done without verifying the exact code change was applied.** Read the relevant lines after every edit. "I've made the change" is only valid after the tool confirms the edit and the code is correct.
- Before any edit: read the target lines. After any edit: verify the new code is correct. Do not rely on memory of what was written.

---

## SETTLEMENT PIPELINE HARDENING + EOD EMAIL DATA INTEGRITY (2026-05-31)

### [fix] Multiple settle_watcher instances corrupting settle_status.json
**Files:** `Options_flow_analyzer/settle_watcher.py`

**Root cause:** Task Scheduler launched 3 separate instances on 2026-05-29 (13:00, 14:45, 14:59 ET).
The 3rd instance hit the `else` branch of the idempotent-restart check before the 2nd instance had
finished, calling `_write_status(today, False, None, False, None)` — resetting both flags to False.
Result: `settle_status.json` ended the day as `options_settled: false` despite two successful writes.

**Fix:** Added PID lock file (`settle_watcher.lock`) at startup. If another instance is alive, logs
error and exits immediately. Lock removed via `atexit` on clean exit; stale locks auto-overwritten.
See ERRORS.md Rule 10.

---

### [fix] `_persist_ct_options_ice` and `_persist_futures_ice` overwriting settled CSV data
**File:** `ct-options-dashboard/app.py`

**Root cause:** When dashboard was loaded on Sunday evening (2026-05-31), app.py's auto-persist
detected RTD tickers and called `_persist_ct_options_ice`, which deleted then rewrote Friday's
1036 correctly-settled option rows with 483 Sunday-night live prices. Similarly `_persist_futures_ice`
had no protection against overwriting settled futures data.

**Fix:** Both functions now read `settle_status.json` at entry. If `futures_settled` / `options_settled`
is `true` for `today_str`, they return immediately without touching the CSV.
`settle_watcher.py` is now the **sole writer** to both CSVs. See ERRORS.md Rule 11.

---

### [fix] Friday 2026-05-29 option settle data recovered
**File:** `ct-options-dashboard/local_options_history.csv`

The 1036 correct ICE-settled option rows written by settle_watcher at 14:45 ET were overwritten
by app.py's auto-persist with 483 Sunday-night live prices. Recovered by running a one-shot script
(`patch_options_may29.py`, since deleted) that re-read Friday's settled prices directly from the
ICE RTD Settle column (static until Monday settlement) and wrote 1036 correct rows.
`settle_status.json` manually corrected to `options_settled: true`.

---

### [fix] Nov 26 (CTX6) straddle showing — for Settlement, Change, % CHG Vol
**File:** `ct-options-dashboard/app.py` — CT straddle loop, serial month fallback

**Root cause:** ATM strike for CTX6 was 81 (derived from CTZ6 forward ~81.14), but
`local_options_history.csv` had no strike-81 rows for CTX6 (jumps 80→82 for puts, 80→83 for calls).
CSV fallback at `abs(r['strike'] - atm) >= 0.01` found no match → `prev_val = None` → all three
fields showed `—`.

**Fix:** After the CSV fallback, if `prev_c`/`prev_p` still None for a serial month, compute
settlement straddle via B76 using the standard month's prior settlement IV
(`atm_iv_for_date(std_tkr, settle_ref)`). Same method used for the live value — always available.
See ERRORS.md Rule 12.

---

### [feat] Watcher-down alert banner on dashboard
**Files:** `ct-options-dashboard/app.py`, `ct-options-dashboard/templates/index.html`

Added `/api/watcher-status` endpoint that checks the lock file PID. Frontend polls every 90s and
shows a red banner "settle_watcher.py is NOT running" during the settlement window (14:25–16:00 ET
on trading days). Banner disappears automatically when watcher comes back up.

---

### [feat] Silent 3-minute cache warmer
**File:** `ct-options-dashboard/templates/index.html`

Added `setInterval` that silently fetches `/api/data` every 3 minutes without re-rendering.
Ensures the server-side `_ld_cache` always has values from within the last 3 minutes when RTD
goes offline at market close (~14:20 ET). Dashboard continues to display correct straddle values
through the post-close window and until the next session opens (9pm ET next trading day).
Manual refresh button unaffected.

---

### [fix] EOD email missing futures high/low/volume/OI when RTD workbook closed
**File:** `ct-options-dashboard/app.py` — CSV fallback block after RTD section

**Root cause:** `live_futures` was built exclusively from the RTD. If the workbook was closed,
`live_futures = {}` — all futures fields (high, low, volume, efp, efs, block, OI, OI_chg) would
be blank in the EOD email. `settle_watcher` already writes all these fields to the CSV at ~14:45.

**Fix:** Added CSV fallback block after the RTD section. When `post_settle = True` (futures CSV
has today's date), any missing `live_futures` fields are filled from the CSV row written by
settle_watcher. RTD values take priority; CSV fills gaps. EOD email now has complete data
regardless of workbook state.

---

### [confirmed] EOD email data pipeline — all fields verified post-settlement

| Field | Source | Status |
|---|---|---|
| Futures settle, yest_settle, change | settle_watcher CSV | ✓ Always |
| Futures high, low, volume, efp, efs, block | RTD → CSV fallback | ✓ Always |
| Futures OI, OI_chg | RTD → CSV fallback | ✓ Always |
| Straddle value (pre-settle) | Stale cache guard + cache warmer | ✓ Always |
| Straddle value (post-settle) | local_options_history.csv | ✓ Always |
| Straddle Settlement (prev) | flow_rtd.json → CSV prev_date | ✓ Always |
| Straddle change, % CHG Vol | Computed from above | ✓ Always |
| HV10/30/60/90 | local_futures_history.csv (_compute_hv) | ✓ Always |

Post-close behaviour confirmed: last live bid/ask straddle values preserved via cache warmer +
stale straddle guard from market close (~14:20 ET) through to next session open (9pm ET next
trading day). After settlement (~14:45 ET), values automatically update to ICE official settled
prices. No manual action required.

---

## SECURITY CONSTRAINTS — NEVER VIOLATE

- `RTD SIte.xlsx` / `RTD SIte.xlsm` at `Site Sync/` path: **READ-ONLY, NEVER MODIFY**
- `.ARBCTVV F Index`: **COMPLETELY IGNORED — never use**
- Bloomberg data: **local only, never transmitted to cloud**
