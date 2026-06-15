# Coffee (KC) EOD Snapshot → Backfill — BUILD HANDOFF

**Status:** NOT STARTED. This document is the complete brief for building the coffee
EOD signal-backfill pipeline as a mirror of the working cotton one. Read it end to
end before writing any code. Build it in a **fresh chat**, in **steps**, with each
step verified before the next.

**Date written:** 2026-06-15.
**Author context:** written immediately after the cotton EOD auto-fire pipeline was
fixed and verified firing end-to-end (see §2). Every file/line reference below was
read and verified on 2026-06-15; re-verify before relying on a line number, because
files change.

---

## 0. ONE-LINE SUMMARY

Cotton has a fully automatic chain: settle-watcher → dashboard snapshot endpoint →
`eod_snapshot.json` → `append_backfill.py` → cotton signal CSV in the `oi-dashboard`
repo → served by the gateway. Coffee has **none** of the snapshot/backfill half. The
job is to build the coffee equivalent, reusing the cotton pattern, **without touching
cotton**.

---

## 1. THE COTTON PIPELINE (REFERENCE — this is what "same as cotton" means)

The pipeline spans **four separate git repos**. Each coffee piece goes where its
cotton counterpart already lives.

| # | Stage | Cotton file (verified) | Repo |
|---|-------|------------------------|------|
| 1 | Detect settlement, then POST | `Options_flow_analyzer/settle_watcher.py` (POST at lines 645–661, fired only when `fut_done and opt_done`) | `options-flow-analyzer` |
| 2 | Assemble EOD data + write compact snapshot | `ct-options-dashboard/app.py` — `_assemble_eod_data()` (~L4106), `_write_eod_snapshot()` (~L4344), endpoint `api_save_eod_snapshot()` (~L4453) | `vlmsofts-ct-options-dashboard` |
| 3 | Read snapshot, compute columns, push row | `market-intelligence/append_backfill.py` | `market-intelligence` (pushes to oi-dashboard) |
| 4 | Physical storage of the backfill CSV | `Open interest dashboard/data/signals/vlm_signal_backfill.csv` | `oi-dashboard` |
| 5 | Serve the CSV over the API | `vlm-data-gateway/routes/signals.py`, `routes/catalog.py` | `vlm-data-gateway` |

**Consumers of the cotton backfill CSV (verified via grep):**
`market-intelligence/cot_score.py`, `market-intelligence/repair_backfill.py`,
`vlm-data-gateway/routes/signals.py`, `vlm-data-gateway/routes/catalog.py`.

### 1a. The snapshot file (`market-intelligence/data/eod_snapshot.json`)

Compact JSON the dashboard writes and the appender reads. Cotton schema:
```json
{
  "date": "2026-06-15",
  "ct1_settle": 76.81, "ct2_settle": 78.12, "ct3_settle": 79.06,
  "ct1_ticker": "CTZ6", "ct2_ticker": "CTH7", "ct3_ticker": "CTK7",
  "atm_iv_30d": 20.87, "hv30": 27.06, "hv60": 23.03
}
```
`ct1/2/3` = the first three **standard-month** futures (settle + ticker).
`atm_iv_30d` = first standard straddle with `dte >= 30`. `hv30/hv60` = historical vol.

### 1b. The backfill row (`append_backfill.py`)

16 columns, in order (see `append_backfill.py` `COLUMNS`):
```
date, ct1_close, ct2_close, ct3_close, atm_iv_30d,
ct1_ct2_spread, ct2_ct3_spread, si_carry_approx, pct_si_approx,
hv30, hv60, iv_hv30_ratio, atm_iv_zscore_1yr,
iv_hv30_ratio_zscore, pct_si_zscore_1yr, ct1_ct2_zscore
```
Key derivations (cotton):
- `gap_months` = true calendar-month gap from `ct1_ticker`→`ct2_ticker`. Cotton valid
  gaps = `{2, 3, 5}` (MAR→MAY=2, MAY→JUL=2, DEC→MAR=3, JUL→DEC=5). **Loud-fail** if
  the gap isn't in the set.
- `si_carry_approx = (STORAGE_CPM + ct1_close * FINANCING_RATE/12) * gap_months`, with
  `STORAGE_CPM = 0.50` (cents/lb/month), `FINANCING_RATE = 0.0515` (5.15%/yr).
- `pct_si_approx = (ct2_close - ct1_close) / si_carry_approx * 100`.
- z-scores use a **252-row** window; null until 252 prior rows exist.
- Duplicate-date guard: `sys.exit` if `date` already in the CSV (safe no-op re-run).
- Pushes to `oi-dashboard` `main` with `git pull --ff-only` first, then commit + push.

---

## 2. WHAT WAS FIXED IN COTTON ON 2026-06-15 (already shipped — context)

Three production bugs surfaced and were fixed (commit alongside this doc):

1. **Snapshot date was the lagged options-CSV `last_date`** (one business day stale),
   so a snapshot taken on Mon 06-15 was stamped Fri 06-12 → backfill aborted on a
   false duplicate. **Fix:** `_assemble_eod_data()` now returns `today_et`
   (ET session date); `_write_eod_snapshot()` stamps `today_et`, not `last_date`.
2. **Save Snapshot button was commodity-blind** — `_assemble_eod_data()` is hardcoded
   `load_data('CT')`, so pressing Save on the *coffee* tab silently wrote a *cotton*
   snapshot. **Fix:** the button now sends `?commodity=`; the endpoint **refuses
   non-CT** (HTTP 400).
3. **No options-settlement gate** — `atm_iv_30d` comes from the live straddle and is
   non-final until options settle, but nothing stopped a pre-settle save from
   producing a backfill-eligible row. **Fix:** endpoint returns **HTTP 409** unless
   `settle_status.json` shows `options_settled == true` for `today_et`. The
   settle-watcher writes that flag *before* it POSTs, so the auto path passes; only a
   premature manual press is blocked.

Also: the manual **Save Snapshot button is now disabled/grayed** in
`templates/index.html` (label `💾 Auto (on settle)`), because the chain fires
automatically. Remove `disabled` to restore it as a manual fallback.

**Verified working 2026-06-15:** options settled 15:28 ET → watcher POSTed → snapshot
written dated 06-15 with settled IV → `append_backfill.py` appended row
`2026-06-15,76.81,78.12,79.06,20.87,...` and pushed to oi-dashboard. Zero manual steps.

**These cotton fixes are the template for coffee. Do not regress them.**

---

## 3. WHAT COFFEE ALREADY HAS vs NEEDS

### Already exists (verified):
- **`COMMODITY_CONFIG['KC']`** in `app.py` (~L144–163):
  - `std_months = {3, 5, 7, 9, 12}` — **FIVE** standard months (cotton has four).
  - `excl_months = frozenset()` — **no excluded month** (cotton excludes OCT; coffee
    has no such rule — do NOT copy cotton's OCT exclusion into coffee).
  - `strike_increment = 2.5` (coffee trades a 2.5-cent strike grid).
  - `straddle_tickers = {KCN6, KCU6, KCZ6, KCH7, KCK7}` (offline fallback list).
  - separate CSV paths: `LOCAL_KC_OPT_HISTORY`, `LOCAL_KC_FUT_HISTORY`,
    `LOCAL_KC_SPR_HISTORY`.
- **`load_data('KC')`** works (routes to `_load_generic_data`).
- **`_assemble_eod_data()`** — currently hardcoded `load_data('CT')`. Needs to accept a
  commodity so it can assemble KC.
- **`settle_watcher_kc.py`** in `Options_flow_analyzer` — a full KC sibling of the
  cotton watcher (own lock/paths/workbook; KC settles ~12:25 ET, closes 1:30 ET; has
  `settle_status_kc.json`). **It does NOT currently POST to `/api/save-eod-snapshot`**
  (verified: only `app.py` and the cotton `settle_watcher.py` reference that endpoint).
- **`settle_status_kc.json`** exists in `ct-options-dashboard/` (KC settle flags).

### Does NOT exist (must be built):
- No coffee snapshot file (no `eod_snapshot_kc.json` or equivalent).
- No coffee appender (no `append_backfill_kc.py`).
- No coffee backfill CSV (no `kc_signal_backfill.csv` in oi-dashboard).
- No coffee gateway route.
- **No coffee composite-signal model** that consumes a coffee backfill. (The
  `CTA MONITOR/data/kc_*` files are a SEPARATE coffee model fed by its own inputs —
  NOT this EOD pipeline. Do not assume they connect.)

---

## 4. OPEN DECISIONS — RESOLVE THESE *BEFORE* WRITING CODE

These are not implementation details; they change what gets built.

1. **What consumes the coffee backfill?** Cotton's exists to feed the cotton composite
   signal engine. If there is no coffee composite engine, this build is
   "collect-now-consume-later" data plumbing (fine, but say so). This decides whether
   stage 5 (gateway route + a coffee signal model) is in scope.
2. **Do the cotton carry columns even apply to coffee?** `si_carry_approx` /
   `pct_si_approx` encode cotton storage-and-interest carry. Decide whether coffee
   needs the same columns, different ones, or the same formula with coffee params.
3. **Coffee storage / financing params.** Cotton uses `STORAGE_CPM = 0.50` cents/lb/mo
   and `FINANCING_RATE = 0.0515`. Coffee C is also cents/lb but has different storage
   economics. Get the correct coffee storage rate (and confirm the financing rate).
4. **Coffee month-gap set.** From `{3,5,7,9,12}` the consecutive-month gaps are
   H→K=2, K→N=2, N→U=2, U→Z=3, Z→H=3 → candidate valid set `{2, 3}` (NOT cotton's
   `{2,3,5}`). Confirm against the real front-pair roll behavior before locking the
   loud-fail guard.
5. **Backfill CSV location + name.** Recommended: `oi-dashboard`
   `data/signals/kc_signal_backfill.csv` (same repo as cotton, so the gateway serves
   it the same way). Confirm.
6. **z-score window.** Coffee has little/no history yet — z-scores will be null until
   252 rows accrue. Confirm that's acceptable, or define a shorter window for coffee.

---

## 5. STEP-BY-STEP BUILD PLAN (each step ends with a STOP + verify)

> Build in this order. Do not proceed to the next step until the current one is
> verified. Nothing pushes to a live repo until Step 6 is explicitly approved.

**Step 0 — Resolve §4 decisions.** Write the answers down. Do not code until done.

**Step 1 — Make `_assemble_eod_data()` commodity-aware.**
Add a `commodity='CT'` parameter; replace the hardcoded `load_data('CT')`. Cotton
behavior must be byte-identical when called with no arg / `'CT'`. Verify the cotton
snapshot output is unchanged.

**Step 2 — Make `_write_eod_snapshot()` + the endpoint commodity-aware.**
- Snapshot path per commodity (e.g. `eod_snapshot.json` for CT,
  `eod_snapshot_kc.json` for KC) — do NOT overwrite the cotton file.
- Standard-month selection from `COMMODITY_CONFIG[commodity]['std_months']`, not the
  cotton-hardcoded `('Mar','May','Jul','Dec')`.
- Options-settled gate reads the right status file (`settle_status.json` for CT,
  `settle_status_kc.json` for KC).
- Endpoint accepts `?commodity=KC` and routes accordingly; keep the non-supported
  refuse for anything without a pipeline.

**Step 3 — Build `append_backfill_kc.py` in `market-intelligence`.**
Mirror `append_backfill.py` with: KC snapshot path, KC month-code map, KC valid-gap
set (§4.4), KC storage/financing params (§4.3), KC backfill CSV target (§4.5),
KC commit message. Keep the duplicate-date guard and 252-row z-score logic. **Do NOT
copy cotton's OCT exclusion** (coffee has no excluded month).

**Step 4 — Wire `settle_watcher_kc.py` to POST.**
Add the POST to `/api/save-eod-snapshot?commodity=KC` at the watcher's "both settled"
point (mirror cotton `settle_watcher.py` L645–661). Best-effort, never crash the
watcher. It must write `options_settled=true` to `settle_status_kc.json` BEFORE the
POST, so the endpoint's 409 gate passes.

**Step 5 — Create the empty KC backfill CSV + (if in scope) the gateway route.**
Header-only `kc_signal_backfill.csv` in oi-dashboard; add a gateway route if §4.1 says
a consumer needs it.

**Step 6 — End-to-end dry run, then live.**
`append_backfill_kc.py --dry-run` first. Then a real KC settlement cycle:
watcher detects KC options settle → POST → KC snapshot written (correct date, settled
IV) → KC row appended + pushed. Verify the row in the CSV and `local == remote`.

---

## 6. VERIFICATION CHECKLIST (per the project's hard rules)

- Cotton is **untouched** and still auto-fires (re-run a cotton check; the cotton
  snapshot, endpoint, and `append_backfill.py` must be unchanged in behavior).
- KC snapshot writes to a **separate** file; never overwrites `eod_snapshot.json`.
- KC date stamp = KC ET session date (the `today_et` fix), never a lagged CSV date.
- KC backfill aborts safely on duplicate date; gap guard loud-fails off-set gaps.
- ASCII-only console output (Windows cp1252 chokes on Δ/Σ/em-dashes).
- Nothing committed/pushed without explicit go.

---

## 7. HARD RULES (do not violate)

- **Verify before stating** anything about file contents / data / system state — read
  the source in the same step. Memory is not verification.
- **Read the whole enclosing function** before editing inside it; grep for a variable
  name in scope before declaring it.
- **Surgical changes** — touch only what the step requires; match existing style.
- **Do NOT touch the cotton pipeline** while building coffee. Separate files, separate
  paths, separate status, separate CSV.
- **Coffee ≠ cotton** on: number of standard months (5 vs 4), OCT exclusion (none vs
  hard rule), month-gap set, settle timing (~12:25 ET vs cotton's ~14:1x), and
  storage/financing params.
