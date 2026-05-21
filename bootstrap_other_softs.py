"""
bootstrap_other_softs.py — seed local KC/SB/CC history files
Copies KC, SB, CC rows from the OI dashboard CSVs into the CT dashboard
local files.  Run once (safe to re-run — deduplicates by date+key).

  python bootstrap_other_softs.py

If Bloomberg later provides pre-2026-05-01 options history, re-run after
appending those rows to the OI dashboard options_oi.csv.
"""

import csv, os, pathlib, shutil

# ── Source files (OI dashboard) ───────────────────────────────────────────────
OI_DASH = pathlib.Path(r'C:\Users\Louis\OneDrive - VLM Commodities LTD\Desktop\Open interest dashboard\data')
SRC_OPT = OI_DASH / 'options_oi.csv'
SRC_FUT = OI_DASH / 'oi_data.csv'

# ── Destination files (CT dashboard) ─────────────────────────────────────────
HERE    = pathlib.Path(__file__).parent
TARGETS = {
    'KC': {
        'opt': HERE / 'local_kc_options_history.csv',
        'fut': HERE / 'local_kc_futures_history.csv',
    },
    'SB': {
        'opt': HERE / 'local_sb_options_history.csv',
        'fut': HERE / 'local_sb_futures_history.csv',
    },
    'CC': {
        'opt': HERE / 'local_cc_options_history.csv',
        'fut': HERE / 'local_cc_futures_history.csv',
    },
}


def write_csv(path, rows, fieldnames):
    """Write rows to path, backup if file already exists."""
    if path.exists():
        shutil.copy(path, path.with_suffix('.bak'))
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)
    print(f'  Wrote {len(rows):,} rows -> {path.name}')


def bootstrap():
    for src, label in [(SRC_OPT, 'options'), (SRC_FUT, 'futures')]:
        if not src.exists():
            print(f'ERROR: source not found: {src}')
            return

    # ── Options ───────────────────────────────────────────────────────────────
    print('Reading options_oi.csv ...')
    with open(SRC_OPT, newline='', encoding='utf-8') as f:
        opt_rows = list(csv.DictReader(f))
    opt_fields = list(opt_rows[0].keys()) if opt_rows else []

    for comm, paths in TARGETS.items():
        rows = [r for r in opt_rows
                if r.get('commodity', '').strip().upper() == comm]
        rows.sort(key=lambda r: (r.get('date', ''), r.get('security_des', '')))
        print(f'  {comm} options: {len(rows):,} rows '
              f'({rows[0]["date"] if rows else "?"} to {rows[-1]["date"] if rows else "?"})')
        write_csv(paths['opt'], rows, opt_fields)

    # ── Futures ───────────────────────────────────────────────────────────────
    print('Reading oi_data.csv ...')
    with open(SRC_FUT, newline='', encoding='utf-8') as f:
        fut_rows = list(csv.DictReader(f))
    fut_fields = list(fut_rows[0].keys()) if fut_rows else []

    for comm, paths in TARGETS.items():
        rows = [r for r in fut_rows
                if r.get('commodity', '').strip().upper() == comm]
        rows.sort(key=lambda r: (r.get('date', ''), r.get('contract', '')))
        print(f'  {comm} futures: {len(rows):,} rows '
              f'({rows[0]["date"] if rows else "?"} to {rows[-1]["date"] if rows else "?"})')
        write_csv(paths['fut'], rows, fut_fields)

    print('Done.')


if __name__ == '__main__':
    bootstrap()
