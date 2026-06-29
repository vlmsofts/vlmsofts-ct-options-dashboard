# producer/ — ICE Connect futures+options producer (version-controlled copy)

`dashboard_futures_producer.py` is the standalone 32-bit producer that pulls CT
(and KC/SB/CC) **futures** + the **ATM:15 option band** via icepython and writes
`../api_feed/futures_api_<C>.json`, which the dashboard reads when `ICE_USE_API=1`.

## THIS IS THE SOURCE OF TRUTH (git-tracked). The RUNTIME copy lives at:
`C:\Ice eod records\dashboard_futures_producer.py`

It MUST run from `C:\Ice eod records\` because it does `from ice_com_hardening
import ensure_ice` (the analyzer's hardened 32-bit ICE connector). After editing
the tracked copy here, sync it to the runtime location:

```
copy "producer\dashboard_futures_producer.py" "C:\Ice eod records\dashboard_futures_producer.py"
```

## Run
```
py -3.13-32 dashboard_futures_producer.py --commodity CT          # futures + ATM:15 options
py -3.13-32 dashboard_futures_producer.py --commodity CT --no-options   # futures only
py -3.13-32 dashboard_futures_producer.py --commodity CT --force        # ignore 14:18-16:30 stand-down
```
(ICE XL must be open + logged in.)

## Key design (verified 2026-06-29)
- **ATM:15 band** per live tenor (`***<c> <tenor> ATM:15`), batched quotes <=150.
- **NO ImpVol** — server-computed, ~5x slower (1.76s vs 0.28s/30-sym call). The
  dashboard computes its own B76/SOFR IV; ICE IV is cross-checked vs the 4:30
  settled_surface. Raw fields (Bid/Offer/Last/RecSet/OpenInt) are cheap.
- **Option-expiry filter**: a contract past OPTION expiry (but pre-FND) gets
  futures-only — its option chain is gone and autolisting it HANGS. Schedule in
  `_OPT_EXPIRY` mirrors app.py `ICE_*_EXPIRY`.
- **Self-heal + alert**: unknown contract -> pull `Last Trading Day` from the API
  (no hang) + append to `_MISSING_OPTION_EXPIRY.txt` so the schedule can be
  updated. The API serves option LTD per symbol (`Last Trading Day`/`Expiration`).
- **set_timeout(8)**: any stuck call raises instead of hanging.
- **Stand-down 14:18-16:30 ET**: yields COM to the analyzer's settle_watcher.

## Alert file
`C:\Ice eod records\_MISSING_OPTION_EXPIRY.txt` — if a contract lacks a hardcoded
option-expiry, the producer self-heals from the API AND logs it here (deduped).
Add listed contracts to `_OPT_EXPIRY` (here) and `ICE_*_EXPIRY` (app.py).
