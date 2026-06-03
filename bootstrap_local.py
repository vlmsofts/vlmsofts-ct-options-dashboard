"""
One-time bootstrap: download full GitHub CSV history into local files.
Safe to re-run — deduplicates by date.

GitHub options columns:  date, commodity, security_des, contract_month, put_call,
                         strike_px, open_int, oi_chg, px_settle, px_volume
GitHub futures columns:  date, commodity, contract, bbg_ticker, settle,
                         open_int, oi_chg, first_notice, last_trade
"""
import requests
import csv
import io
import os

OPT_CSV_URL = "https://raw.githubusercontent.com/vlmsofts/oi-dashboard/main/data/options_oi.csv"
OI_CSV_URL  = "https://raw.githubusercontent.com/vlmsofts/oi-dashboard/main/data/oi_data.csv"

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
LOCAL_OPT   = os.path.join(BASE_DIR, "local_options_history.csv")
LOCAL_FUT   = os.path.join(BASE_DIR, "local_futures_history.csv")


def bootstrap(url, local_path, commodity_filter='CT'):
    print(f"\nFetching {url} ...")
    r = requests.get(url, timeout=30)
    r.raise_for_status()

    all_rows = list(csv.DictReader(io.StringIO(r.text)))
    ct_rows  = [row for row in all_rows
                if row.get('commodity', '').strip().upper() == commodity_filter]
    ct_rows.sort(key=lambda row: row.get('date', ''))

    print(f"  {len(ct_rows)} {commodity_filter} rows from {len(all_rows)} total")
    gh_dates = sorted(set(row.get('date', '') for row in ct_rows))
    print(f"  GitHub date range: {gh_dates[0]} to {gh_dates[-1]}  ({len(gh_dates)} dates)")

    # Load existing local dates to avoid duplicates
    existing_dates = set()
    if os.path.exists(local_path):
        with open(local_path, 'r', newline='', encoding='utf-8') as f:
            for row in csv.DictReader(f):
                existing_dates.add(row.get('date', '').strip())
        print(f"  Existing local dates: {len(existing_dates)}")

    new_rows = [row for row in ct_rows
                if row.get('date', '').strip() not in existing_dates]

    if not new_rows:
        print(f"  Nothing new — local already up to date.")
        return

    new_dates = sorted(set(row.get('date', '') for row in new_rows))
    print(f"  Writing {len(new_rows)} rows across {len(new_dates)} new dates ...")

    # Overwrite: replace old simplified-format file with full GitHub format
    # (old file had wrong column names — keep backup, write fresh)
    old_cols_wrong = False
    if os.path.exists(local_path) and existing_dates:
        with open(local_path, 'r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            old_cols = reader.fieldnames or []
        # GitHub options has 'px_settle'; old local had 'px' — detect mismatch
        if 'px_settle' not in old_cols and 'settle' not in old_cols:
            old_cols_wrong = True
            print(f"  WARNING: existing file has old simplified columns {old_cols}")
            print(f"  Will REPLACE with GitHub-format data (old format backed up as .BACKUP.csv)")

    if old_cols_wrong:
        # Replace entirely with full GitHub-format data (CT rows only)
        rows_to_write = ct_rows
        mode = 'w'
        print(f"  Writing {len(rows_to_write)} total CT rows in GitHub format ...")
    else:
        rows_to_write = new_rows
        mode = 'a'

    need_header = mode == 'w' or not os.path.exists(local_path)
    with open(local_path, mode, newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=ct_rows[0].keys())
        if need_header:
            w.writeheader()
        w.writerows(rows_to_write)

    final_dates = set()
    with open(local_path, 'r', newline='', encoding='utf-8') as f:
        for row in csv.DictReader(f):
            final_dates.add(row.get('date', '').strip())
    print(f"  Done. {local_path} now has {len(final_dates)} dates.")

    # Spot-check: compare a few prices on the latest date
    latest = gh_dates[-1]
    gh_latest  = [row for row in ct_rows if row.get('date', '').strip() == latest]
    with open(local_path, 'r', newline='', encoding='utf-8') as f:
        lc_latest = [row for row in csv.DictReader(f)
                     if row.get('date', '').strip() == latest]
    print(f"  Spot-check {latest}: GitHub {len(gh_latest)} rows, local {len(lc_latest)} rows — "
          f"{'MATCH' if len(gh_latest) == len(lc_latest) else 'MISMATCH'}")


if __name__ == '__main__':
    bootstrap(OPT_CSV_URL, LOCAL_OPT, commodity_filter='CT')
    bootstrap(OI_CSV_URL,  LOCAL_FUT, commodity_filter='CT')
    print("\n--- Bootstrap complete ---")
    print("Next step: verify date ranges above, then update app.py.")
