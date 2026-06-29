# ICE Connect Coexistence Probe — GATE for the icepython producer

**Status:** NOT YET RUN. This probe gates the entire icepython-producer replacement
(see memory `project_ice_connect_eval`). Do not build the producer until this passes.

**Owner:** options-flow-analyzer side (they own the proven icepython capture +
`C:\Ice eod records\ice_com_hardening.py`).

**Why this gates everything:** the producer plan assumes icepython can run as a
*separate process* while the dashboard's Excel RTD feed is *still open* against the
same ICE Data Services entitlement during market hours. If the ICE publisher does
NOT allow two concurrent consumers (Excel RTD add-in + a `start_publisher()`
process) on one entitlement, the whole architecture changes:

- **PASS** → build the producer as scoped (runs beside Excel; migrate the dashboard
  read to JSON at leisure; Excel stays as a fallback during cutover).
- **FAIL** → icepython must *replace* Excel outright (no overlap). That means a
  hard cutover, no parallel-run safety net, and the settle_watcher (which today
  reads the same Excel) must move to icepython at the same time. Much bigger blast
  radius — re-scope before committing.

---

## The single question

> During CT/KC market hours, with the ICE RTD FEED CT.xlsx / KC.xlsx workbooks
> OPEN and live in Excel, can a separate 32-bit Python process call
> `icepython.start_publisher()` and `get_quotes(...)` successfully — without
> either feed losing data, throwing, or hanging?

Three sub-checks, all must hold:

1. **icepython connects** while Excel RTD is live (no entitlement-in-use error).
2. **Excel RTD keeps updating** while icepython is polling (open a watched cell,
   confirm it still ticks — the add-in must not freeze or drop the feed).
3. **icepython values match Excel** for the same symbol at the same instant
   (within normal tick latency), confirming both see the real feed, not a stale
   or degraded copy.

---

## Procedure (run on the dashboard machine, market hours, NON-settle window)

Pre-conditions:
- A CT trading day, current ET time is **before 14:25** (outside the settle window
  so `settle_watcher` is not contending — see `app.py:_in_ct_settle_window`).
- Excel open with ICE RTD FEED CT.xlsx live and ticking (confirm a futures cell
  is updating).
- 32-bit Python with icepython available (`ice_com_hardening.py` assumes 32-bit —
  see its flag text). Run this in a SEPARATE interpreter from the dashboard's.

Steps:

1. Note Excel's live `CT Z26` (front future) Last/Bid/Offer at a timestamp.
2. In the separate process, `ensure_ice()` (or `icepython.start_publisher()`),
   then `get_quotes(["CT Z26"], ["Last","Bid","Offer","Settle"], False)`.
   - Record: did it connect? did it return a row? values vs Excel?
3. Leave icepython polling `get_quotes` every ~2s for 60s. During that window,
   watch the Excel cell — does it keep ticking, or freeze/blank?
4. Pull a small option batch via icepython (≤150 symbols) and confirm no hang,
   while Excel options sheets are open.
5. Stop icepython. Confirm Excel RTD still live (didn't get killed by the
   publisher teardown).

Record PASS/FAIL for each of the three sub-checks above, plus any COM error codes
(esp. `-2147221008 CoInitialize`, `-2147467259 Unspecified`, or a silent hang).

---

## Known constraints to design around (already verified by the analyzer team)

- `get_quotes` **hangs uncatchably** above ~200 symbols → batch ≤150
  (`QUOTE_BATCH=150`). A try/except cannot rescue a hang.
- ICE COM must run on the **main thread** (threaded wrapper throws CoInitialize).
- `get_quotes` ImpVol is **decimal** (0.21887); timeseries Implied Volatility is
  **×100** (22.08) — normalize if reconciling.
- Post-close `get_quotes` returns frozen settled values (Last==Settle==RecSet).
- The producer MUST honor the settle windows: CT 14:25–16:00 ET
  (`app.py:_in_ct_settle_window`), KC its own window — stand down so it never
  contends with `settle_watcher`.

---

## If PASS — next artifact

Write `ICE_CONNECT_PRODUCER_CONTRACT.md`: the exact `{mode, futures{}, spreads{},
options{}}` JSON the producer emits (must be byte-identical to
`ice_rtd_reader.read_ice_workbook()`'s return dict, ice_rtd_reader.py:558-563),
the file path + atomic-write convention, the freshness stamp, and the one
JSON-load branch the dashboard adds beside `_read_ice_workbook_safe`
(app.py:27-37). Downstream (`_ice_to_rtd_shape`, straddle math, EOD) stays
untouched.
