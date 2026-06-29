# Resident Streaming Producer — Design Doc (RTD model restored)

**Status: DESIGN / not started. Stops for Lou's approval + analyzer coexistence sign-off before any code.**
**Produced 2026-06-29. Supersedes the snapshot-on-spawn model in `dashboard_futures_producer.py`.**

---

## 0. THE PROBLEM (verified this session)

The current producer (`C:\Ice eod records\dashboard_futures_producer.py`) is **run-once**: argparse → `run_once()` → exit. It calls `ice.get_quotes(syms, fields, False)` — `subscribe=False`, a **one-shot snapshot**. So every invocation cold-pulls from scratch:

- Spawn 32-bit Python + reconnect ICE + autolist all 17 strips + batch-quote ≈ **13s every time**.
- Verified back-to-back: **cold 13.5s, warm 12.7s — NO warm cache.**

The dashboard reads the JSON file this writes. The file is only as fresh as the last producer run. On-demand spawn fires every ~3 min (page auto-refresh) but the fresh-TTL is 60s → the file is "stale" most of the time → dashboard falls back to settle/CSV. **A background loop over the run-once producer just re-pays 13s forever — same staleness, faster.** That is NOT a fix.

## 1. THE INSIGHT — icepython streams, like RTD did

The old Excel RTD was never a timed snapshot: the workbook held **live RTD cells**, ICE pushed updates continuously, the dashboard read in-place. icepython exposes the same:

- `get_quotes(syms, fields, True)` — **`subscribe=True`** opens a persistent subscription. ICE pushes; values stay live in the client.
- **Live-proved 2026-06-29:** subscribe call registers in 0.27s (`get_subscriptions()` shows N active), and a **re-read = 0.027s** (18× faster than one-shot) with live values (CT Z26 last 76.64 / settle 76.38).
- API surface: `get_subscriptions`, `clear_subscriptions`, `start_publisher` / `restart_publisher`, `get_hibernation` / `set_hibernation`.

## 2. TARGET ARCHITECTURE — resident subscribing producer

Replace run-once-on-spawn with a **single long-lived process** that holds the subscription and writes the file on a fast cadence:

```
Live ICE ──push──> [resident producer: subscribe once, read live, write JSON ~1-2s]
                                    │
                            api_feed/futures_api_<C>.json  (never >~2s old)
                                    │
                   dashboard reads the file (instant, 0.03s) — NEVER spawns
```

1. **Startup:** connect via `ensure_ice()`; `get_quotes(all_futures + ATM:15 option strips, fields, True)` to register subscriptions.
2. **Loop:** every ~1–2s, read the now-live values (0.03s) and atomically rewrite the JSON. No re-autolist, no re-pull.
3. **Dashboard:** becomes a **pure file reader** — the inline `_spawn_producer` and its 20s timeout become moot (kept as a dead fallback or removed at cutover).

**Result:** file is always <~2s old → dashboard always serves live → no settle fallback during market hours, no 13s page hang, ever.

## 3. OPEN DESIGN POINTS (resolve before coding)

| # | Question | Notes / leaning |
|---|---|---|
| 1 | **Analyzer coexistence** | settle_watcher holds COM 14:18–16:00; producer stand-down 14:18–16:30. A SECOND resident icepython subscription on the same box, simultaneously with settle_watcher's icepython, is the **UNTESTED 2-client case**. **GATE: analyzer sign-off required.** Options: (a) resident producer also stands down 14:18–16:30 (file ages ~132 min/day → settle takes over, acceptable since settle_watcher owns that window); (b) single shared subscriber both consume. Leaning (a) — minimal change, preserves the proven stand-down. |
| 2 | **ATM-band re-subscription** | The ±15 strikes around ATM shift as the future moves. Need periodic re-autolist (cheap, ~0.5s/strip) to add new strikes + `clear_subscriptions` on ones that fell out, WITHOUT churning the whole book each tick. Cadence: re-autolist every N seconds (e.g. 30–60s), keep the quote read every 1–2s. |
| 3 | **Write cadence** | ~1–2s. Atomic write (temp + rename) — already the pattern. Must not write mid-update partials. |
| 4 | **Startup / restart / self-heal** | ICE outage → subscription drops → detect (stale values / publisher down) → `restart_publisher` or full re-`ensure_ice` + re-subscribe. Process supervisor (bat loop OR Task Scheduler) restarts the whole process if it dies. |
| 5 | **Shape stability** | `read_ice_api` MUST keep returning `{mode, futures, spreads, options}` byte-compatibly — settle_watcher re-source + 5 cross-repo consumers depend on it. The file schema does NOT change; only HOW it's produced. |
| 6 | **Hibernation** | `set_hibernation` — does ICE hibernate idle subscriptions and slow the first read after quiet? Test whether to disable for the resident case. |
| 7 | **One process per commodity, or one for all?** | CT today; KC/SB/CC later. One resident process handling all 4 (more subs, one COM client) vs 4 processes (isolation, more COM clients = more coexistence risk). Leaning one process, all commodities. |

## 4. BLAST RADIUS

- **File schema unchanged** → dashboard render, settle_watcher re-source, and all `read_ice_api` consumers are untouched IF the JSON stays byte-compatible. This is the contract to protect.
- **COM coexistence** is the real risk (point 1) — gated on analyzer.
- **The dashboard inline spawn** (`_spawn_producer`, 20s timeout) is removed/neutered only AT cutover, after the resident producer is proven to keep the file warm across a full session.

## 5. SEQUENCING

1. **Analyzer coexistence sign-off** (point 1) — can a resident subscriber run alongside settle_watcher, or must it stand down 14:18–16:30?
2. Build resident producer for **CT only**, behind its own run mode (NOT replacing the run-once entry — add a `--stream` mode so the old path stays as fallback).
3. Run alongside the dashboard; verify file stays <~2s old for a full session; verify live straddles/smile never drop to settle intra-session.
4. Validate stand-down handoff (14:18–16:30) + settle_watcher coexistence on a live session.
5. Extend to KC/SB/CC.
6. Cutover: dashboard → pure file reader; retire inline spawn.

## 6. STOPGAPS IN PLACE (until this ships)

- CTF6 expired-serial guard — `e081989` (no dead chain → no 20s hang → pull 30s→14s).
- Spawn timeout 8s→20s — `77decf8` (lets the ~13s cold pull finish instead of being killed).
- EOD email + PNG — verified working (settle-based, independent of live feed).
