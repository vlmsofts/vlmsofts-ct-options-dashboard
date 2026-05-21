"""
normalize_bbg_softs.py
======================
Converts Bloomberg softs pull CSVs to the column format expected by app.py.
Run once (safe to re-run — rewrites output files).

  python normalize_bbg_softs.py

Inputs (C:\\Users\\Louis\\Downloads\\):
  coffee_options_history.csv   coffee_futures_history.csv
  sugar_options_history.csv    sugar_futures_history.csv
  cocoa_options_history.csv    cocoa_futures_history.csv

Outputs (ct-options-dashboard folder — replaces OI dashboard bootstrap):
  local_kc_options_history.csv    local_kc_futures_history.csv
  local_sb_options_history.csv    local_sb_futures_history.csv
  local_cc_options_history.csv    local_cc_futures_history.csv
"""

import csv
import os
import pathlib
import shutil
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
# PATHS
# ─────────────────────────────────────────────────────────────────────────────

DOWNLOADS = pathlib.Path(r'C:\Users\Louis\Downloads')
HERE      = pathlib.Path(__file__).parent

COMMODITIES = {
    'KC': {
        'opt_in':  DOWNLOADS / 'coffee_options_history.csv',
        'fut_in':  DOWNLOADS / 'coffee_futures_history.csv',
        'opt_out': HERE / 'local_kc_options_history.csv',
        'fut_out': HERE / 'local_kc_futures_history.csv',
    },
    'SB': {
        'opt_in':  DOWNLOADS / 'sugar_options_history.csv',
        'fut_in':  DOWNLOADS / 'sugar_futures_history.csv',
        'opt_out': HERE / 'local_sb_options_history.csv',
        'fut_out': HERE / 'local_sb_futures_history.csv',
    },
    'CC': {
        'opt_in':  DOWNLOADS / 'cocoa_options_history.csv',
        'fut_in':  DOWNLOADS / 'cocoa_futures_history.csv',
        'opt_out': HERE / 'local_cc_options_history.csv',
        'fut_out': HERE / 'local_cc_futures_history.csv',
    },
}

# ─────────────────────────────────────────────────────────────────────────────
# LOOKUP TABLES
# ─────────────────────────────────────────────────────────────────────────────

# Month code → month number (ICE/Bloomberg standard)
_CODE_TO_NUM = {
    'F': 1, 'G': 2, 'H': 3, 'J': 4, 'K': 5, 'M': 6,
    'N': 7, 'Q': 8, 'U': 9, 'V': 10, 'X': 11, 'Z': 12,
}

# Month number → 3-char name used in the ordinal contract format (e.g. KCJUL3)
# Must match FUTURES_MONTH_SUFFIX in app.py exactly.
_NUM_TO_3CHAR = {
    1: 'JAN', 2: 'FEB', 3: 'MAR', 4: 'APR', 5: 'MAY', 6: 'JUN',
    7: 'JUL', 8: 'AUG', 9: 'SEP', 10: 'OCT', 11: 'NOV', 12: 'DEC',
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _parse_option_ticker(ticker):
    """
    'KCN6C 257.5 Comdty' -> {'commodity':'KC', 'security_des':'KCN6C', 'strike_px':257.5}
    Returns None on failure.
    """
    parts = ticker.strip().split()
    if len(parts) < 2:
        return None
    raw   = parts[0]          # 'KCN6C'
    try:
        strike = float(parts[1])
    except (ValueError, IndexError):
        return None
    if len(raw) < 5:
        return None
    commodity = raw[:2].upper()
    # Last char must be C or P
    if raw[-1] not in ('C', 'P'):
        return None
    return {'commodity': commodity, 'security_des': raw, 'strike_px': strike}


def _parse_futures_ticker(ticker):
    """
    'KCH6 Comdty' -> {'commodity':'KC', 'month_num':3, 'delivery_year':2026}
    Returns None on failure.
    """
    raw = ticker.strip().replace(' Comdty', '').replace(' CURNCY', '')
    if len(raw) < 4:
        return None
    commodity    = raw[:2].upper()
    month_code   = raw[2].upper()
    month_num    = _CODE_TO_NUM.get(month_code)
    if month_num is None:
        return None
    try:
        year_digit   = int(raw[3])
    except (ValueError, IndexError):
        return None
    delivery_year = 2020 + year_digit
    return {'commodity': commodity, 'month_num': month_num, 'delivery_year': delivery_year}


def _contract_ordinal(row_date_str, month_num, delivery_year):
    """
    Compute the ordinal used by app.py's get_hist_fwd / parse_futures_my_generic.
    ordinal = delivery_year - first_year + 1
    first_year = row_year if row_month <= month_num else row_year + 1
    Returns None if ordinal <= 0 (expired relative to row date).
    """
    try:
        d = datetime.strptime(row_date_str, '%Y-%m-%d')
    except ValueError:
        return None
    first_year = d.year if d.month <= month_num else d.year + 1
    ordinal    = delivery_year - first_year + 1
    return ordinal if ordinal >= 1 else None


# ─────────────────────────────────────────────────────────────────────────────
# NORMALISE OPTIONS
# ─────────────────────────────────────────────────────────────────────────────

OPT_FIELDS = ['date', 'commodity', 'security_des', 'strike_px',
              'px_settle', 'open_int', 'oi_chg', 'px_volume']


def normalize_options(commodity, spec):
    src  = spec['opt_in']
    dest = spec['opt_out']
    if not src.exists():
        print(f'  [SKIP] {src.name} not found')
        return

    print(f'  Reading {src.name} ...')
    with open(src, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        raw_rows = list(reader)

    out_rows   = []
    skip_count = 0
    for r in raw_rows:
        ticker = r.get('ticker', '').strip()
        parsed = _parse_option_ticker(ticker)
        if not parsed:
            skip_count += 1
            continue
        if parsed['commodity'] != commodity:
            skip_count += 1
            continue
        date_s = r.get('date', '').strip()
        if not date_s:
            skip_count += 1
            continue

        px_last   = r.get('PX_LAST', '') or ''
        open_int  = r.get('OPEN_INT', '') or ''

        out_rows.append({
            'date':         date_s,
            'commodity':    parsed['commodity'],
            'security_des': parsed['security_des'],
            'strike_px':    parsed['strike_px'],
            'px_settle':    px_last.strip() if px_last.strip() else '',
            'open_int':     open_int.strip() if open_int.strip() else '',
            'oi_chg':       '',
            'px_volume':    '',
        })

    out_rows.sort(key=lambda r: (r['date'], r['security_des'], r['strike_px']))

    _write_csv(dest, out_rows, OPT_FIELDS)
    print(f'    {commodity} options: {len(out_rows):,} rows written '
          f'({skip_count} skipped)')


# ─────────────────────────────────────────────────────────────────────────────
# NORMALISE FUTURES
# ─────────────────────────────────────────────────────────────────────────────

FUT_FIELDS = ['date', 'commodity', 'contract', 'bbg_ticker',
              'settle', 'open_int', 'oi_chg', 'first_notice', 'last_trade']


def normalize_futures(commodity, spec):
    src  = spec['fut_in']
    dest = spec['fut_out']
    if not src.exists():
        print(f'  [SKIP] {src.name} not found')
        return

    print(f'  Reading {src.name} ...')
    with open(src, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        raw_rows = list(reader)

    out_rows   = []
    skip_count = 0
    for r in raw_rows:
        ticker = r.get('ticker', '').strip()
        parsed = _parse_futures_ticker(ticker)
        if not parsed:
            skip_count += 1
            continue
        if parsed['commodity'] != commodity:
            skip_count += 1
            continue
        date_s = r.get('date', '').strip()
        if not date_s:
            skip_count += 1
            continue

        settle_s   = (r.get('px_settle', '') or '').strip()
        if not settle_s:
            skip_count += 1
            continue

        month_num     = parsed['month_num']
        delivery_year = parsed['delivery_year']
        ordinal       = _contract_ordinal(date_s, month_num, delivery_year)
        if ordinal is None:
            # Contract already expired relative to this row date — skip
            skip_count += 1
            continue

        month_3 = _NUM_TO_3CHAR[month_num]
        contract   = f'{commodity}{month_3}{ordinal}'
        bbg_ticker = f'{contract} Comdty'

        out_rows.append({
            'date':         date_s,
            'commodity':    commodity,
            'contract':     contract,
            'bbg_ticker':   bbg_ticker,
            'settle':       settle_s,
            'open_int':     (r.get('open_int', '') or '').strip(),
            'oi_chg':       '',
            'first_notice': (r.get('first_notice', '') or '').strip(),
            'last_trade':   (r.get('last_trade', '') or '').strip(),
        })

    out_rows.sort(key=lambda r: (r['date'], r['contract']))

    _write_csv(dest, out_rows, FUT_FIELDS)
    print(f'    {commodity} futures: {len(out_rows):,} rows written '
          f'({skip_count} skipped)')


# ─────────────────────────────────────────────────────────────────────────────
# IO
# ─────────────────────────────────────────────────────────────────────────────

def _write_csv(path, rows, fieldnames):
    if path.exists():
        shutil.copy(path, path.with_suffix('.bak'))
        print(f'    (backed up existing {path.name})')
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print('=' * 60)
    print('  normalize_bbg_softs.py')
    print('=' * 60)
    for commodity, spec in COMMODITIES.items():
        print(f'\n--- {commodity} ---')
        normalize_options(commodity, spec)
        normalize_futures(commodity, spec)
    print('\nDone.')
    print()
    print('NOTE: px_settle in options files is mapped from Bloomberg PX_LAST.')
    print('      This is last-traded price, not official exchange settlement.')
    print('      For liquid ATM/near-ATM strikes it is a close proxy.')


if __name__ == '__main__':
    main()
