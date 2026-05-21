# CT Options Dashboard — Change Catalog

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

## SECURITY CONSTRAINTS — NEVER VIOLATE

- `RTD SIte.xlsx` / `RTD SIte.xlsm` at `Site Sync/` path: **READ-ONLY, NEVER MODIFY**
- `.ARBCTVV F Index`: **COMPLETELY IGNORED — never use**
- Bloomberg data: **local only, never transmitted to cloud**
