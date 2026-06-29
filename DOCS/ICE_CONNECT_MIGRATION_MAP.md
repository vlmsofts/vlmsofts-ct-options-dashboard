# ICE Connect Migration Map — ct-options-dashboard

**Produced 2026-06-29 by a 5-agent read-only trace + live Connect verification. No code changed.**
This is the build blueprint for replacing the Excel RTD feed with the icepython ICE Connect API.

## VERDICT
Clean migration. ONE new function `read_ice_api(commodity)` returning the EXACT dict
`_ice_to_rtd_shape` emits (app.py:604-609) makes the swap render-invisible: all 12 rtd consumers
and every template key unchanged. Standalone 32-bit producer process (icepython is main-thread-only
+ 32-bit; dashboard is 64-bit Flask threaded=True). Additive, flag-gated, parallel-validate vs Excel.

## THE CONTRACT read_ice_api MUST EMIT (verbatim, Agent 1)
```
{
  'outrights': { 'CTZ6': {last, settle, yest_settle, change, pct_chg, oi, oi_chg(=None),
                          volume, high, low, block_vol, efs_vol, efp_vol,
                          hv10/30/60/90(=None)}, ... },
  'spreads':   { 'CTN6/CTZ6': {display, settle, last, change, pct_chg, yest_settle,
                               high, low, volume, block_vol, efs_vol, efp_vol}, ... },
  'live_options': { 'CTZ6': {expiry(=None), strikes:[{strike, pc('Call'|'Put'), bid, ask,
                             mid(=(bid+ask)/2 or last), last, settle, vol(=None)}, ...]}, ... },
  'source': 'ice_rtd_live' | 'ice_rtd_today_settle' | 'ice_rtd_prior_settle',
}
# OR None when unavailable. Always-None today: oi_chg, hv10/30/60/90, live_options.expiry, strikes[].vol
```

## FIELD MAP (Agent 2, LIVE-VERIFIED 2026-06-29 — three flags resolved by live probe)
| dashboard field | Connect source | type |
|---|---|---|
| last / high / low | Last / High / Low | NATIVE |
| settle | **RecSet** (NOT Settle — None intraday) | NATIVE |
| yest_settle | **PrevPrice** (drop the prior-day CSV diff) | NATIVE |
| change | **Change** (net change) | NATIVE |
| pct_chg | Change/PrevPrice*100 | COMPUTE |
| volume | Cumulative Volume | NATIVE |
| oi | OpenInt | NATIVE |
| block/efs/efp_vol | Block/EFS/EFP Volume | NATIVE |
| **futures bid/offer** | Bid/Offer | **NATIVE (live-verified: CT Z26 Bid 76.93)** |
| **option bid/offer per strike** | Bid/Offer | **NATIVE (live-verified: CT Z26C76 Bid 4.28, IV 0.209, Delta 0.556)** |
| **spread yest_settle** | spread's own PrevPrice | **NATIVE (live-verified: CT Z26:CTH27 PrevPrice -1.36)** |
| oi_chg | OpenInt − yesterday CSV OI | COMPUTE+CSV (unavoidable) |
| HV10/30/60/90 | get_timeseries Settle hist → local log-return stdev*sqrt(252), exclude rolls(|lr|>0.15) | COMPUTE |
| ATM IV / straddle / breakeven / smile / skew | B76 from Connect mids+IV | COMPUTE (formulas verified vs app.py:1961-1971) |

**Symbology:** futures `CT Z26`; options `CT Z26C76`/`CT H27C82.5`; spreads `CT Z26:CTH27` (colon, 2nd
leg keeps CT, no space — NEVER `get_autolist("**...")`, it HANGS). chain via `get_autolist('***CT Z26')`.

**Connect ImpVol is DECIMAL** (0.209) vs ×100 in timeseries. **DECISION:** keep dashboard B76+SOFR
self-solve for per-strike IV (continuity); use Connect ImpVol as cross-check, not source.

## COM REDESIGN (Agent 3) — the one architectural requirement
`_read_ice_workbook_safe` (app.py:27-37) wraps reads in a ThreadPoolExecutor for an 8s timeout.
icepython COM is **main-thread only** (threaded → com_error -2147221008) AND needs **32-bit Python**
(`py -3.13-32`; dashboard is 64-bit Py3.14). → producer MUST be a separate 32-bit process that writes
JSON; the dashboard reads the JSON (no in-process icepython). See [[reference_icepython_32bit_requirement]].

## SETTLE WINDOW + CROSS-REPO (Agent 3) — what stays
- Keep `_in_ct_settle_window()` 14:25-16:00 gate (app.py:614-627) — producer honors it.
- KEEP reads: flow_rtd.json (app.py:738-754, yest settle), rtd_snap.json (14:16 freeze, app.py:1615-1645),
  ct_price_tape.csv (app.py:1686-1690, live-mid fallback). settle_status.json write-guards (app.py:2418-2435).
- These are fallbacks/freezes — unaffected by the source swap.

## MULTI-COMMODITY (Agent 4)
`_ice_to_rtd_shape` is commodity-agnostic; symbology generalizes. Honor per-commodity:
KC 2.5c strike grid; SB std_months {3,5,7,10} NO Dec; CT excludes Oct(V); KC settle ~12:25 / window
13:28-15:00 (_WATCHER_META app.py:4717); KC freeze 13:20 vs CT 14:18.
**Pre-existing bug:** `_load_generic_data` gates KC/SB/CC on CT's window (app.py:2816) — KC's own settle unguarded.

## RENDER CONTRACT (Agent 5) — ZERO template changes
Both `/` and `/api/data` call load_data/_load_generic_data → serialize the same _result dict
(emitted app.py:2053-2086 CT / 3658-3689 generic). Template reads via `const D = {{data|tojson}}`.
Only cosmetic edit: feed_down reason string (app.py:2045) "ICE RTD workbook" → Connect-appropriate.
Freshness polls unaffected: /api/settle-status (60s), /api/watcher-status (90s), cache-warmer (3min).

## PRE-EXISTING BUGS FOUND (not migration-caused; note for later)
1. `/api/rtd_debug` reads `rtd.get('options')` (app.py:4213) — key never exists (_ice_to_rtd_shape emits
   'live_options') → rtd_atm_vol/strad/dte always None from RTD path. Latent.
2. `_load_generic_data` settle-window guard uses CT window for KC/SB/CC (app.py:2816).

## SEQUENCING (non-breaking)
Phase 1 FUTURES (low-risk: ~8 syms = 1 batch, sub-second) → build read_ice_api filling `futures`,
swap behind a flag, parallel-validate vs Excel for N sessions. Phase 2 OPTIONS (heavy: every strike,
multiple 150-batches) later. settle_watcher shares ice_rtd_reader → same read_ice_api fixes it too
([[project_settle_watcher_shares_reader]]). Constraints: QUOTE_BATCH=150, main-thread COM, RecSet not Settle.
