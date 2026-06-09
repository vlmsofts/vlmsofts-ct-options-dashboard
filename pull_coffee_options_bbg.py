"""
pull_coffee_options_bbg.py
==========================
Pulls the ICE Coffee C (KC) option chain history from Bloomberg and writes
a CSV in the exact format that normalize_bbg_softs.py expects as input.

Output: C:\\Users\\Louis\\Downloads\\coffee_options_history.csv
  columns: ticker, date, PX_LAST, OPEN_INT
  ticker form: 'KCN6C 257.5 Comdty'  (matches normalize_bbg_softs._parse_option_ticker)

Then run:  python normalize_bbg_softs.py
  -> rewrites local_kc_options_history.csv (backs up old to .bak)

Requires a running Bloomberg Terminal on this machine (pdblp + blpapi).

  python pull_coffee_options_bbg.py
  python pull_coffee_options_bbg.py --start 2025-01-03 --end 2026-06-09
  python pull_coffee_options_bbg.py --dry-run     # list option tickers, no BDH

NOTE: px_settle in the dashboard maps from PX_LAST (last trade), not official
exchange settlement. For ATM/near-ATM strikes this is a close proxy.
"""

import argparse
import csv
import datetime as _dt
import pathlib
import sys

import numpy as _np
# pdblp 0.1.8 still references np.NaN, removed in NumPy 2.0. Restore it so
# bulkref()/ref() don't crash on this NumPy 2.x install.
if not hasattr(_np, 'NaN'):
    _np.NaN = _np.nan

import pdblp

# ── Config ────────────────────────────────────────────────────────────────────
OUT_PATH = pathlib.Path(r'C:\Users\Louis\Downloads\coffee_options_history.csv')

# Coffee C futures whose option chains we want history for.
# Only contracts that still resolve as live OPT_CHAIN securities are useful —
# expired fronts (KCH5/K5/N5/U5/Z5) no longer return a chain. Deferred contracts
# still listed (KCN6 onward) were ALSO trading through 2025 as back months, so
# their BDH history covers the 2025 backfill window down to each strike's listing.
# October/V coffee options do not resolve as 'KCV6 Comdty' here (different root),
# so they are pulled via their parent serial chain if at all — omitted from this list.
# Dry-run 2026-06-09 confirmed these return chains (286–406 strikes each):
KC_FUTURES = [
    'KCN6', 'KCU6', 'KCZ6',
    'KCH7', 'KCK7', 'KCN7', 'KCU7', 'KCZ7',
    'KCH8',
]

DEFAULT_START = '2025-01-03'
DEFAULT_END   = _dt.date.today().strftime('%Y-%m-%d')

FIELDS = ['PX_LAST', 'OPEN_INT']


def _bbg_to_yyyymmdd(s):
    return s.replace('-', '')


def get_option_tickers(con, fut_ticker):
    """
    Return the list of option security strings for a coffee future, e.g.
    'KCN6 Comdty' -> ['KCN6C 240 Comdty', 'KCN6P 240 Comdty', ...].

    Uses OPT_CHAIN (bulk reference field). Each row's security is like
    'KCN6C 240 Comdty'. We keep them verbatim — they already match the
    normalizer's expected ticker form.
    """
    sec = f'{fut_ticker} Comdty'
    try:
        df = con.bulkref(sec, 'OPT_CHAIN')
    except Exception as e:
        print(f'   [WARN] OPT_CHAIN failed for {sec}: {e}')
        return []
    # bulkref returns columns: ticker, field, name, value
    vals = df[df['name'].str.contains('Option', case=False, na=False)]['value'].tolist()
    if not vals:
        # Some terminals label the member column differently; fall back to all values
        vals = df['value'].tolist()
    # Keep only well-formed option strings ending in ' Comdty' with C/P + strike
    out = []
    for v in vals:
        v = str(v).strip()
        if not v.endswith('Comdty'):
            continue
        head = v.replace(' Comdty', '').strip()
        parts = head.split()
        if len(parts) < 2:
            continue
        root = parts[0]
        if not root or root[-1] not in ('C', 'P'):
            continue
        out.append(v)
    return sorted(set(out))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--start', default=DEFAULT_START)
    ap.add_argument('--end',   default=DEFAULT_END)
    ap.add_argument('--dry-run', action='store_true',
                    help='Enumerate option tickers only; skip the BDH history pull.')
    ap.add_argument('--host', default='localhost')
    ap.add_argument('--port', type=int, default=8194)
    args = ap.parse_args()

    print('=' * 64)
    print('  Coffee C option-chain history pull (Bloomberg)')
    print(f'  window: {args.start} .. {args.end}')
    print('=' * 64)

    con = pdblp.BCon(host=args.host, port=args.port, timeout=30000)
    con.start()

    # 1) Enumerate option tickers across all listed coffee futures
    all_opts = []
    for fut in KC_FUTURES:
        opts = get_option_tickers(con, fut)
        print(f'  {fut}: {len(opts)} option tickers')
        all_opts.extend(opts)
    all_opts = sorted(set(all_opts))
    print(f'\n  TOTAL distinct option tickers: {len(all_opts)}')

    if not all_opts:
        print('  No option tickers found — nothing to pull. Is the Terminal logged in?')
        con.stop()
        sys.exit(1)

    if args.dry_run:
        for t in all_opts[:50]:
            print('   ', t)
        if len(all_opts) > 50:
            print(f'    ... (+{len(all_opts)-50} more)')
        con.stop()
        return

    # 2) Historical pull. BDH over many securities at once; pdblp returns a
    #    wide frame indexed by date with (ticker, field) columns.
    start = _bbg_to_yyyymmdd(args.start)
    end   = _bbg_to_yyyymmdd(args.end)

    print(f'\n  Pulling BDH {FIELDS} for {len(all_opts)} tickers ...')
    # Chunk to keep request sizes sane.
    CHUNK = 100
    out_rows = []
    for i in range(0, len(all_opts), CHUNK):
        chunk = all_opts[i:i + CHUNK]
        print(f'    chunk {i//CHUNK + 1}/{(len(all_opts)+CHUNK-1)//CHUNK} '
              f'({len(chunk)} tickers)')
        try:
            df = con.bdh(chunk, FIELDS, start, end,
                         longdata=True)  # long: ticker, date, field, value
        except Exception as e:
            print(f'      [WARN] BDH failed for chunk: {e}')
            continue
        # longdata columns: date, ticker, field, value
        # Pivot field -> column per (ticker, date)
        bucket = {}
        for _, r in df.iterrows():
            tkr = str(r['ticker']).strip()
            dt  = r['date']
            dt_s = dt.strftime('%Y-%m-%d') if hasattr(dt, 'strftime') else str(dt)[:10]
            fld = str(r['field']).strip()
            val = r['value']
            bucket.setdefault((tkr, dt_s), {})[fld] = val
        for (tkr, dt_s), fv in bucket.items():
            px = fv.get('PX_LAST', '')
            oi = fv.get('OPEN_INT', '')
            if px in (None, '') and oi in (None, ''):
                continue
            out_rows.append({
                'ticker':   tkr,
                'date':     dt_s,
                'PX_LAST':  '' if px is None else px,
                'OPEN_INT': '' if oi is None else oi,
            })

    out_rows.sort(key=lambda r: (r['date'], r['ticker']))
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT_PATH, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=['ticker', 'date', 'PX_LAST', 'OPEN_INT'])
        w.writeheader()
        w.writerows(out_rows)

    con.stop()
    n_dates = len({r['date'] for r in out_rows})
    print(f'\n  Wrote {len(out_rows):,} rows ({n_dates} dates) -> {OUT_PATH}')
    print('  Next: python normalize_bbg_softs.py')


if __name__ == '__main__':
    main()
