# Local Data Migration Plan — CT Options Dashboard

**Goal:** Replace all GitHub CSV fetches with local file reads.  
**Result:** Zero network I/O on page load; no external dependency; deterministic, fast data.

---

## Current State

| Source | URL | Used by |
|---|---|---|
| `options_oi.csv` | GitHub raw | `load_data()`, `compute_skew_history()` |
| `oi_data.csv` | GitHub raw | `load_data()`, `compute_skew_history()` |
| Bloomberg RTD | Excel COM (local) | `load_data()` — already local |

**Local files that already exist (partial):**
- `local_options_history.csv` — columns: `date, ticker, pc, strike, px` (simplified — missing `oi, oi_chg, vol`)
- `local_futures_history.csv` — columns: `date, contract, settle` (simplified — missing `last_trade, first_notice`)

**Problem with current local files:** They are missing columns that `load_data()` uses:
- Options: `oi`, `oi_chg`, `vol` — used for C/P ratio, OI bar, skew direction
- Futures: `last_trade`, `first_notice` — used for DTE and `fut_lookup` construction

The migration must fix this or the dashboard loses OI/DTE data.

---

## Architecture Decision

Store local data in **GitHub-compatible raw format** — same column names as the GitHub CSVs.
This means:
- Zero changes to parsing logic inside `load_data()` and `compute_skew_history()`
- `_persist_today()` stores raw rows (not pre-parsed)
- Switch is a one-line change per fetch call: `fetch_csv(URL)` → `read_local_csv(path)`

---

## Phase 0 — Pre-Migration Audit (Do First, Before Any Code Changes)

**Purpose:** Know exactly what you have before touching anything.

### Step 0.1 — Inspect existing local CSVs
```python
import csv
for path in ['local_options_history.csv', 'local_futures_history.csv']:
    with open(path) as f:
        rows = list(csv.DictReader(f))
    dates = sorted(set(r['date'] for r in rows))
    print(f"{path}: {len(rows)} rows, {len(dates)} dates, {dates[0]} → {dates[-1]}")
    print("  columns:", list(rows[0].keys()) if rows else 'EMPTY')
```
Record the earliest and latest date in each file.

### Step 0.2 — Inspect GitHub CSVs
```python
import requests, csv, io
for url, name in [(OPT_CSV_URL, 'options'), (OI_CSV_URL, 'futures')]:
    r = requests.get(url); rows = list(csv.DictReader(io.StringIO(r.text)))
    dates = sorted(set(row.get('date','') for row in rows))
    print(f"GitHub {name}: {len(rows)} rows, {len(dates)} dates, {dates[0]} → {dates[-1]}")
    print("  columns:", list(rows[0].keys()) if rows else 'EMPTY')
```
Record columns exactly — needed to confirm local format matches.

### Step 0.3 — Identify the gap
Compare earliest GitHub date vs earliest local date. The gap is the history that needs bootstrapping.

**CHECKPOINT:** Do not proceed until you have both date ranges documented.

---

## Phase 1 — Bootstrap (One-Time Historical Copy)

**Purpose:** Populate local files with full historical data from GitHub before switching.

### Step 1.1 — Back up existing local files
```
copy local_options_history.csv local_options_history.BACKUP.csv
copy local_futures_history.csv local_futures_history.BACKUP.csv
```
Keep these. If anything goes wrong they are your rollback.

### Step 1.2 — Define target local formats

**`local_options_history.csv` (new full format):**
```
date, commodity, security_des, strike_px, px_settle, open_int, oi_chg, px_volume
```
Matches GitHub `options_oi.csv` exactly — zero parsing changes needed.

**`local_futures_history.csv` (new full format):**
```
date, commodity, contract, settle, last_trade, first_notice
```
Matches GitHub `oi_data.csv` exactly.

### Step 1.3 — Write bootstrap script (`bootstrap_local.py`)

```python
"""
Run once. Downloads full GitHub CSV history and writes to local files.
Safe to re-run — deduplicates by date.
"""
import requests, csv, io, os

OPT_CSV_URL = "https://raw.githubusercontent.com/vlmsofts/oi-dashboard/main/data/options_oi.csv"
OI_CSV_URL  = "https://raw.githubusercontent.com/vlmsofts/oi-dashboard/main/data/oi_data.csv"
LOCAL_OPT   = "local_options_history.csv"
LOCAL_FUT   = "local_futures_history.csv"

def bootstrap(url, local_path, commodity_filter='CT'):
    print(f"Fetching {url}...")
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    rows = list(csv.DictReader(io.StringIO(r.text)))
    ct_rows = [row for row in rows
               if row.get('commodity','').strip().upper() == commodity_filter]
    print(f"  {len(ct_rows)} CT rows from {len(rows)} total")

    # Load existing local dates to avoid duplicates
    existing_dates = set()
    if os.path.exists(local_path):
        with open(local_path, 'r', newline='') as f:
            for row in csv.DictReader(f):
                existing_dates.add(row.get('date','').strip())
        print(f"  {len(existing_dates)} dates already in {local_path}")

    new_rows = [r for r in ct_rows if r.get('date','').strip() not in existing_dates]
    if not new_rows:
        print(f"  Nothing new to add.")
        return

    # Sort chronologically and append
    new_rows.sort(key=lambda r: r.get('date',''))
    new_dates = sorted(set(r.get('date','') for r in new_rows))
    print(f"  Writing {len(new_rows)} rows across {len(new_dates)} new dates...")

    need_header = not os.path.exists(local_path)
    with open(local_path, 'a', newline='') as f:
        w = csv.DictWriter(f, fieldnames=new_rows[0].keys())
        if need_header:
            w.writeheader()
        w.writerows(new_rows)
    print(f"  Done. {local_path} now has {len(existing_dates) + len(new_dates)} dates.")

bootstrap(OPT_CSV_URL, LOCAL_OPT)
bootstrap(OI_CSV_URL,  LOCAL_FUT)
print("\nBootstrap complete.")
```

### Step 1.4 — Run bootstrap and verify
```
python bootstrap_local.py
```

**Verification checks:**
1. Row count in local file ≥ row count of GitHub CT rows (should be equal if no prior local data was CT-format compatible, or local dates may have been in old format)
2. Date range in local file spans GitHub history fully
3. Spot-check: pick 3 random dates, compare option prices between local and GitHub

**CHECKPOINT:** Do not modify `app.py` until local files have verified data back to the earliest GitHub date.

---

## Phase 2 — Update `_persist_today()` (Forward-Looking Fix)

**Purpose:** Ensure new daily data written to local files is in the full GitHub-compatible format.

### Step 2.1 — Change `_persist_today()` signature

```python
# BEFORE:
def _persist_today(last_date, today_opts, ct_fut_rows):

# AFTER:
def _persist_today(last_date, raw_opt_rows, raw_fut_rows):
    """Write today's raw CT rows (GitHub format) to local CSVs. Deduplicates by date."""
```

### Step 2.2 — New implementation

```python
def _persist_today(last_date, raw_opt_rows, raw_fut_rows):
    # Options — write raw rows matching GitHub format
    try:
        existing = set()
        if os.path.exists(LOCAL_OPT_HISTORY):
            with open(LOCAL_OPT_HISTORY, 'r', newline='') as f:
                for row in csv.DictReader(f):
                    existing.add(row.get('date', '').strip())
        if last_date not in existing:
            today_raw = [r for r in raw_opt_rows
                         if r.get('date','').strip() == last_date
                         and r.get('commodity','').strip().upper() == 'CT']
            if today_raw:
                need_header = not os.path.exists(LOCAL_OPT_HISTORY)
                with open(LOCAL_OPT_HISTORY, 'a', newline='') as f:
                    w = csv.DictWriter(f, fieldnames=today_raw[0].keys())
                    if need_header:
                        w.writeheader()
                    w.writerows(today_raw)
    except Exception as e:
        log.warning('opt persist failed: %s', e)

    # Futures — write raw rows matching GitHub format
    try:
        existing = set()
        if os.path.exists(LOCAL_FUT_HISTORY):
            with open(LOCAL_FUT_HISTORY, 'r', newline='') as f:
                for row in csv.DictReader(f):
                    existing.add(row.get('date', '').strip())
        if last_date not in existing:
            today_raw = [r for r in raw_fut_rows
                         if r.get('date','').strip() == last_date
                         and r.get('commodity','').strip().upper() == 'CT']
            if today_raw:
                need_header = not os.path.exists(LOCAL_FUT_HISTORY)
                with open(LOCAL_FUT_HISTORY, 'a', newline='') as f:
                    w = csv.DictWriter(f, fieldnames=today_raw[0].keys())
                    if need_header:
                        w.writeheader()
                    w.writerows(today_raw)
    except Exception as e:
        log.warning('fut persist failed: %s', e)
```

### Step 2.3 — Update call site in `load_data()`
```python
# BEFORE (passing parsed objects):
_persist_today(last_date, today_opts, ct_fut)

# AFTER (passing raw rows from fetch):
_persist_today(last_date, opt_rows, oi_rows)
```

`opt_rows` and `oi_rows` are the raw dicts from `fetch_csv()` — already in scope at the call site.

---

## Phase 3 — Add Local File Reader

**Purpose:** Drop-in replacement for `fetch_csv()` that reads local files.

```python
def read_local_csv(path):
    """Read a local CSV file and return list of dicts (same interface as fetch_csv)."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"Local data file not found: {path}")
    with open(path, 'r', newline='', encoding='utf-8') as f:
        return list(csv.DictReader(f))
```

No caching needed (disk reads are fast; data only changes once per day).

---

## Phase 4 — Parallel Run / Validation (2–5 Trading Days)

**Purpose:** Run both sources simultaneously, compare outputs. Do NOT cut over until validation passes.

### Step 4.1 — Add shadow comparison to `load_data()`

```python
def load_data():
    # Primary: GitHub (still active during validation)
    opt_rows = fetch_csv(OPT_CSV_URL)
    oi_rows  = fetch_csv(OI_CSV_URL)

    # Shadow: local files
    try:
        local_opt = read_local_csv(LOCAL_OPT_HISTORY)
        local_fut = read_local_csv(LOCAL_FUT_HISTORY)
        _validate_sources(opt_rows, local_opt, oi_rows, local_fut)
    except Exception as e:
        log.warning('Shadow validation failed: %s', e)
    # ... rest of load_data() uses opt_rows / oi_rows (GitHub) ...
```

### Step 4.2 — Validation function

```python
def _validate_sources(gh_opt, lc_opt, gh_fut, lc_fut):
    """Compare row counts and latest-date prices between GitHub and local."""
    # Date coverage
    gh_dates = set(r.get('date','') for r in gh_opt if r.get('commodity','').strip().upper()=='CT')
    lc_dates = set(r.get('date','') for r in lc_opt if r.get('commodity','').strip().upper()=='CT')
    missing = gh_dates - lc_dates
    if missing:
        log.warning('LOCAL missing %d dates vs GitHub: %s', len(missing), sorted(missing)[-5:])
    else:
        log.info('LOCAL VALIDATION PASS: date coverage matches GitHub (%d dates)', len(gh_dates))

    # Spot-check: latest date ATM prices for CTN6/CTZ6
    latest = max(gh_dates) if gh_dates else None
    if latest:
        def get_atm(rows, date):
            return {(r.get('security_des','').split()[0], r.get('security_des','').split()[-1]): r.get('px_settle')
                    for r in rows if r.get('date','').strip()==date
                    and r.get('commodity','').strip().upper()=='CT'
                    and r.get('px_settle','0') not in ('','0')}
        gh_atm = get_atm(gh_opt, latest)
        lc_atm = get_atm(lc_opt, latest)
        matches = sum(1 for k in gh_atm if lc_atm.get(k) == gh_atm[k])
        log.info('VALIDATION: %d/%d prices match on %s', matches, len(gh_atm), latest)
```

### Step 4.3 — Monitoring cadence
Run the dashboard normally for 2–5 trading days. Check logs each day:
- `LOCAL VALIDATION PASS` → local is in sync
- `LOCAL missing N dates` → persistence is not writing; investigate before continuing
- Price mismatches → column mapping issue; fix before continuing

**CHECKPOINT:** Do not proceed to Phase 5 until 2+ consecutive trading days show `VALIDATION PASS` with zero missing dates.

---

## Phase 5 — Full Cutover

**Purpose:** Switch `load_data()` and `compute_skew_history()` to local files. Remove GitHub dependency.

### Step 5.1 — `load_data()` switch (2 lines)
```python
# BEFORE:
opt_rows = fetch_csv(OPT_CSV_URL)
oi_rows  = fetch_csv(OI_CSV_URL)

# AFTER:
opt_rows = read_local_csv(LOCAL_OPT_HISTORY)
oi_rows  = read_local_csv(LOCAL_FUT_HISTORY)
```

Remove shadow validation block. Remove `_validate_sources()`. Remove `fetch_csv()` calls at these sites.

### Step 5.2 — `compute_skew_history()` switch (2 lines, same pattern)
```python
# BEFORE:
opt_rows = fetch_csv(OPT_CSV_URL)
oi_rows  = fetch_csv(OI_CSV_URL)

# AFTER:
opt_rows = read_local_csv(LOCAL_OPT_HISTORY)
oi_rows  = read_local_csv(LOCAL_FUT_HISTORY)
```

### Step 5.3 — Keep `fetch_csv()` function
Do not delete it yet. It is still used by the `/api/debug-rtd` endpoint and may be needed for emergency fallback.

### Step 5.4 — Update constants (optional cleanup)
```python
# Can be kept as documentation even after cutover
OPT_CSV_URL = "https://..."  # kept as fallback reference only
OI_CSV_URL  = "https://..."  # kept as fallback reference only
```

---

## Phase 6 — Post-Cutover Verification

Run the dashboard for one full week:

| Check | How |
|---|---|
| Page load time | Should be noticeably faster (no HTTP roundtrip) |
| IV percentile | Compare to a date you know from before — should be identical |
| Skew history chart | Should have same or more data points than before cutover |
| Futures DTE | Check CTN6/CTZ6 DTE matches Bloomberg expiry dates |
| Daily persistence | After each trading day, verify new date appears in local CSVs |

---

## Failsafes and Rollback

### Backup policy
Before Phase 1: `BACKUP_YYYYMMDD_options.csv` and `BACKUP_YYYYMMDD_futures.csv`  
Before Phase 5: another backup set  
Keep backups for 30 days minimum.

### Emergency rollback (any phase)
Revert the 2-line change in `load_data()` and `compute_skew_history()`:
```python
# Rollback: restore these 4 lines
opt_rows = fetch_csv(OPT_CSV_URL)
oi_rows  = fetch_csv(OI_CSV_URL)
# (in both functions)
```
The GitHub CSVs remain live. As long as `fetch_csv()` is not deleted, rollback is instant.

### If GitHub CSV becomes unavailable before cutover
Bootstrap script can be re-run at any time against a cached copy of the GitHub file.

### If local CSV gets corrupted
Restore from backup and re-run bootstrap to fill any missing dates.
Since `_persist_today()` deduplicates by date, re-running bootstrap is always safe.

### OneDrive sync warning
The local CSVs are in an OneDrive folder. OneDrive may lock files during sync.
**Mitigation:** `_persist_today()` uses `'a'` (append) mode with a short window open time — conflict risk is low. If sync conflicts appear, move local CSVs out of the OneDrive subtree to a pure local path (e.g. `C:\ct-data\`).

---

## Implementation Sequence Summary

```
Phase 0  Audit               — read only, no changes           ~15 min
Phase 1  Bootstrap           — run bootstrap_local.py          ~20 min + verify
Phase 2  Update persist()    — app.py change + call site        ~30 min
Phase 3  Add reader func     — 5-line function in app.py        ~10 min
Phase 4  Parallel run        — monitor 2–5 trading days         2–5 days
Phase 5  Cutover             — 4-line change in app.py          ~15 min
Phase 6  Post-verify         — monitor 1 week                   1 week
```

Total active coding: ~90 minutes across two sessions.  
Total elapsed: 1–2 weeks (dominated by validation monitoring, not work).

---

## Files Affected

| File | Change |
|---|---|
| `bootstrap_local.py` | New file — one-time use |
| `app.py` | `_persist_today()` rewrite; `read_local_csv()` add; 4 fetch lines swapped |
| `local_options_history.csv` | Format upgrade (add missing columns via bootstrap) |
| `local_futures_history.csv` | Format upgrade (add missing columns via bootstrap) |
| `templates/index.html` | No changes |
| `rtd_reader.py` | No changes |

---

## DO NOT

- **Do not delete `fetch_csv()`** until 30+ days post-cutover with zero incidents
- **Do not run bootstrap on a date when GitHub data is not yet updated** (before ~9:30am day after trade date)
- **Do not skip the parallel run phase** — the format change to `_persist_today()` means old local files are incompatible; you need the bootstrap to cover the gap before the format change is live
- **Do not store the backup files inside the same OneDrive folder** without confirming sync is not mid-flight

---

*Last updated: 2026-05-19*
