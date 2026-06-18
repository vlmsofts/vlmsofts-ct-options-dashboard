"""
migrate_add_vlm_key.py — one-time, idempotent M4 migration
===========================================================
Adds the trailing `vlm_key` column to local_options_history.csv and backfills it
on every existing row from that row's OWN fields (security_des / strike_px /
put_call), so the on-disk file is never left ragged (old 10-col header + new
11-col rows).

Key derivation mirrors how the readers already parse this file (see app.py
parse_security_des and the flow-analyzer joins):
  - contract = security_des[:4]            e.g. 'CTN6'
  - put_call = security_des[4] if C/P, else the put_call column
  - strike   = strike_px column when present; otherwise the embedded strike in
               security_des ('CTN6P    75' -> 75). ~30k rows use the embedded form.
A row that resolves to neither a valid strike nor C/P (malformed / non-option)
gets an EMPTY vlm_key rather than crashing — readers use .get() and tolerate it.

Idempotent: safe to re-run. It always rewrites the header to 11-col and
recomputes vlm_key for every row from source fields (so a partial prior run is
healed). A timestamped .bak is written once per run before any change.

Run from the ct-options-dashboard dir:
    python migrate_add_vlm_key.py            # migrate in place (writes .bak)
    python migrate_add_vlm_key.py --dry-run  # report only, no write
"""

import argparse
import csv
import os
import shutil
import sys
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
TARGET = os.path.join(HERE, 'local_options_history.csv')

# Trailing column appended to whatever the file's existing header is.
NEW_COL = 'vlm_key'

_PUT_CALL_NORM = {'C': 'C', 'P': 'P', 'CALL': 'C', 'PUT': 'P'}


def _make_option_key(contract, strike, put_call):
    """VERBATIM logic of contract_calendar.make_option_key (and app.py's copy)."""
    code = str(contract).strip().upper()
    pc = _PUT_CALL_NORM.get(str(put_call).strip().upper())
    if pc is None:
        raise ValueError(f'unrecognized put_call {put_call!r}')
    strike_int = int(round(float(strike) * 100))
    return f'{code}-{strike_int:05d}-{pc}'


def _row_vlm_key(row):
    """Derive vlm_key for an existing CSV row from its own fields. Returns '' if
    the row is not a resolvable C/P option (kept blank, never raises)."""
    sec = (row.get('security_des') or '').strip()
    if len(sec) < 5:
        return ''
    contract = sec[:4]
    # put/call: prefer the embedded char at position 4, fall back to the column
    pc_raw = sec[4] if sec[4] in ('C', 'P') else (row.get('put_call') or '').strip()
    if pc_raw not in ('C', 'P'):
        return ''
    # strike: prefer the strike_px column, else the embedded strike in security_des
    strike = None
    px = (row.get('strike_px') or '').strip()
    if px:
        try:
            strike = float(px)
        except ValueError:
            strike = None
    if strike is None:
        parts = sec.split()
        if len(parts) >= 2:
            try:
                strike = float(parts[1])
            except ValueError:
                strike = None
    if strike is None or strike <= 0:
        return ''
    try:
        return _make_option_key(contract, strike, pc_raw)
    except ValueError:
        return ''


def migrate(dry_run=False):
    if not os.path.exists(TARGET):
        print(f'[migrate] target not found: {TARGET}')
        return 1

    with open(TARGET, 'r', newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        fieldnames = list(reader.fieldnames or [])
        rows = list(reader)

    if not fieldnames:
        print('[migrate] empty/headerless file — nothing to do')
        return 1

    out_fields = fieldnames + ([NEW_COL] if NEW_COL not in fieldnames else [])

    filled = blank = already = 0
    for row in rows:
        prior = (row.get(NEW_COL) or '').strip()
        key = _row_vlm_key(row)
        row[NEW_COL] = key
        if key:
            filled += 1
            if prior == key:
                already += 1
        else:
            blank += 1

    print(f'[migrate] rows={len(rows)} keyed={filled} blank(unresolvable)={blank} '
          f'already-correct={already}')
    print(f'[migrate] header: {len(fieldnames)}-col -> {len(out_fields)}-col '
          f'({"+"+NEW_COL if NEW_COL not in fieldnames else NEW_COL+" already present"})')

    if dry_run:
        print('[migrate] --dry-run: no file written')
        return 0

    bak = f'{TARGET}.{datetime.now():%Y%m%d_%H%M%S}.bak'
    shutil.copy2(TARGET, bak)
    print(f'[migrate] backup written: {os.path.basename(bak)}')

    tmp = TARGET + '.tmp'
    with open(tmp, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=out_fields, extrasaction='ignore')
        w.writeheader()
        w.writerows(rows)
    os.replace(tmp, TARGET)
    print(f'[migrate] rewrote {TARGET} ({len(rows)} rows, {len(out_fields)} cols)')
    return 0


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--dry-run', action='store_true',
                    help='report only; do not write the file or a backup')
    args = ap.parse_args()
    sys.exit(migrate(dry_run=args.dry_run))
