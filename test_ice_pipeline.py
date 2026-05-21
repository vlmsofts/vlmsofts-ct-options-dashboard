"""
test_ice_pipeline.py — comprehensive cross-check of ICE RTD reader vs Bloomberg pipeline
==========================================================================================
Run while 'ICE RTD FEED CT.xlsx' is open in Excel with the ICE feed active.

Checks:
  1. Workbook opens and all sheets readable
  2. Futures: all contracts, live forward vs settle forward
  3. Mode detection: live / today_settle / prior_settle
  4. Every option sheet: settle IV and live mid IV (correct forward per mode)
  5. Put-call parity at every ATM strike (should equal F - K within bid-offer)
  6. Bloomberg CSV comparison: exact strike match using strike_px column
  7. IV surface sanity: monotone term structure, no extreme values

Run:
  cd "C:\\Users\\Louis\\OneDrive - VLM Commodities LTD\\Desktop\\ct-options-dashboard"
  python test_ice_pipeline.py
"""

import math
import sys
import os
import pathlib

HERE = pathlib.Path(__file__).parent
sys.path.insert(0, str(HERE))

import ice_rtd_reader as ice

# ── IV solver ─────────────────────────────────────────────────────────────────
try:
    from scipy.stats import norm
    from scipy.optimize import brentq
except ImportError:
    print("ERROR: scipy not installed.  Run: pip install scipy")
    sys.exit(1)

import pandas as pd

def black76(F, K, T, sigma, is_call):
    if T <= 0 or sigma <= 0 or F <= 0 or K <= 0:
        return max(0.0, (F-K) if is_call else (K-F))
    d1 = (math.log(F/K) + 0.5*sigma**2*T) / (sigma*math.sqrt(T))
    d2 = d1 - sigma*math.sqrt(T)
    if is_call:
        return F*norm.cdf(d1) - K*norm.cdf(d2)
    return K*norm.cdf(-d2) - F*norm.cdf(-d1)

def iv(price, F, K, T, is_call):
    if price is None or price <= 0 or T <= 0 or F is None or F <= 0:
        return None
    intrinsic = max(0.0, (F-K) if is_call else (K-F))
    if price <= intrinsic + 1e-6:
        return None
    try:
        v = brentq(lambda s: black76(F, K, T, s, is_call) - price,
                   1e-6, 10.0, xtol=1e-7, maxiter=300)
        return v if 0.001 < v < 9.99 else None
    except Exception:
        return None

def dte_years(expiry_str):
    from datetime import date
    return max((date.fromisoformat(expiry_str) - date.today()).days, 0) / 365.0

def pct(v, width=8):
    return f"{v*100:.2f}%" if v is not None else "-".center(width)


# ── Exact CT expiry dates from app.py ────────────────────────────────────────
CT_EXPIRY = {
    'CTN6': '2026-06-12',
    'CTQ6': '2026-07-17',
    'CTU6': '2026-08-21',
    'CTV6': '2026-09-11',
    'CTX6': '2026-10-16',
    'CTZ6': '2026-11-13',
    'CTF7': '2026-12-18',
    'CTG7': '2027-01-22',
    'CTH7': '2027-02-05',
    'CTJ7': '2027-03-19',
    'CTK7': '2027-04-16',
    'CTM7': '2027-05-21',
    'CTN7': '2027-06-11',
}


# ── Bloomberg CSV loader ───────────────────────────────────────────────────────
def load_bloomberg_csv():
    path = HERE / 'local_options_history.csv'
    if not path.exists():
        return None, None
    df = pd.read_csv(path)
    latest = df['date'].max()
    df_today = df[df['date'] == latest].copy()
    # Normalise strike_px to float
    df_today['strike_px'] = pd.to_numeric(df_today['strike_px'], errors='coerce')
    return df_today, latest


def bbg_lookup(df, contract, put_call, strike):
    """Exact match on contract code, put_call, and strike value."""
    if df is None:
        return None
    mask = (
        df['security_des'].str.contains(contract, na=False, regex=False) &
        (df['put_call'].str.strip().str.upper() == put_call.upper()) &
        (df['strike_px'].round(4) == round(float(strike), 4))
    )
    rows = df[mask]
    if rows.empty:
        return None
    px = pd.to_numeric(rows['px_settle'].iloc[0], errors='coerce')
    return float(px) if not math.isnan(px) else None


# ─────────────────────────────────────────────────────────────────────────────
# TEST SECTIONS
# ─────────────────────────────────────────────────────────────────────────────

def section(title):
    print(f"\n{'─'*70}")
    print(f"  {title}")
    print(f"{'─'*70}")


def test_workbook_open():
    section("1. WORKBOOK OPEN")
    wb = ice.open_workbook('CT')
    if wb is None:
        print("  FAIL: 'ICE RTD FEED CT.xlsx' not found in running Excel instances.")
        print("  Open the file in Excel with the ICE feed active and re-run.")
        sys.exit(1)
    print(f"  PASS: found '{wb.name}'")
    return wb


def test_sheets(wb):
    section("2. SHEET DISCOVERY")
    sheets = ice.get_option_sheets(wb, 'CT')
    print(f"  Option sheets (sorted by expiry): {sheets}")
    if not sheets:
        print("  WARN: no option sheets found.")
    else:
        print(f"  PASS: {len(sheets)} sheet(s)")
    return sheets


def test_futures(wb, mode):
    section("3. FUTURES (all contracts)")
    futures = ice.read_futures(wb, 'CT')
    if not futures:
        print("  FAIL: no futures rows read")
        return futures

    fm = ice.front_month(futures, 'CT')
    print(f"  Front month: {fm}")
    print()
    print(f"  {'Contract':8}  {'Settle':>8}  {'Last':>8}  {'Bid':>8}  "
          f"{'Offer':>8}  {'LiveFwd':>8}  {'OI':>8}  State")
    print(f"  {'─'*78}")

    for contract in sorted(futures, key=lambda c: ice._contract_sort_key(c,'CT')):
        f = futures[contract]
        live_fwd = ice.get_forward(f, 'live')
        oi_str = f"{int(f['oi']):>8,}" if f['oi'] else "       -"
        print(f"  {contract:8}  "
              f"{(f['settle'] or 0):>8.4f}  "
              f"{(f['last']   or 0):>8.4f}  "
              f"{(f['bid']    or 0):>8.4f}  "
              f"{(f['offer']  or 0):>8.4f}  "
              f"{(live_fwd    or 0):>8.4f}  "
              f"{oi_str}  "
              f"{f['market_state'] or '-'}")

    return futures


def test_mode(wb, futures):
    section("4. MODE DETECTION")
    mode = ice.detect_mode(wb, 'CT', stored_atm_settle=None)
    print(f"  Mode: {mode.upper()}")
    fm = ice.front_month(futures, 'CT')
    if fm:
        f = futures[fm]
        settle_fwd = f.get('settle')
        live_fwd   = ice.get_forward(f, 'live')
        print(f"  Front month {fm}: settle_fwd={settle_fwd}  live_fwd={live_fwd}")
        if mode == 'live' and live_fwd and settle_fwd:
            diff = live_fwd - settle_fwd
            print(f"  Intraday move from settle: {diff:+.4f}")
    return mode


def test_option_chains(wb, sheets, futures, mode):
    section("5. OPTION CHAINS — settle IV + live mid IV + put-call parity")

    atm_ivs = {}   # contract → (settle_iv, live_iv) for ATM, used for term structure check

    for sheet in sheets:
        expiry = CT_EXPIRY.get(sheet)
        if expiry is None:
            print(f"\n  {sheet}: no expiry date in CT_EXPIRY — skipping")
            continue

        T = dte_years(expiry)
        if T <= 0:
            print(f"\n  {sheet}: expired — skipping")
            continue

        rows = ice.read_options(wb, sheet)
        if not rows:
            print(f"\n  {sheet}: no rows read")
            continue

        # Get forwards for this contract
        f_data  = futures.get(sheet)
        settle_fwd = f_data.get('settle') if f_data else None
        live_fwd   = ice.get_forward(f_data, mode) if f_data else None

        # For IV calcs:
        #   settle IV  → always uses settle forward (yesterday's official price)
        #   live mid IV → uses live forward in live mode, settle forward otherwise
        fwd_for_live = live_fwd if mode == 'live' else settle_fwd

        atm = ice.atm_strike(fwd_for_live or settle_fwd)

        print(f"\n  {sheet}  expiry={expiry}  DTE={T*365:.0f}d"
              f"  settle_fwd={settle_fwd}  live_fwd={live_fwd}  ATM={atm}")
        print(f"  {'Strike':>7}  {'C_set':>6}  {'P_set':>6}  "
              f"{'C_set_IV':>9}  {'P_set_IV':>9}  "
              f"{'C_mid_IV':>9}  {'P_mid_IV':>9}  "
              f"{'PC_par':>8}  {'C_OI':>7}  {'P_OI':>7}")
        print(f"  {'─'*95}")

        atm_set_iv = None
        atm_live_iv = None

        # Show ATM ±6 strikes
        atm_rows = [r for r in rows if abs(r['strike'] - atm) <= 6]

        for r in atm_rows:
            K = r['strike']

            c_set_iv  = iv(r['call_settle'], settle_fwd, K, T, True)
            p_set_iv  = iv(r['put_settle'],  settle_fwd, K, T, False)

            c_mid = ((r['call_bid'] + r['call_offer']) / 2
                     if r['call_bid'] is not None and r['call_offer'] is not None else None)
            p_mid = ((r['put_bid']  + r['put_offer'])  / 2
                     if r['put_bid']  is not None and r['put_offer']  is not None else None)

            c_live_iv = iv(c_mid, fwd_for_live, K, T, True)  if c_mid else None
            p_live_iv = iv(p_mid, fwd_for_live, K, T, False) if p_mid else None

            # Put-call parity: C - P = F - K  (r=0, undiscounted)
            if r['call_settle'] is not None and r['put_settle'] is not None and settle_fwd:
                pcp_check  = r['call_settle'] - r['put_settle']
                pcp_theory = settle_fwd - K
                pcp_diff   = pcp_check - pcp_theory
                pcp_str    = f"{pcp_diff:>+7.3f}"
            else:
                pcp_str    = "      -"

            marker = " <ATM" if K == atm else ""

            c_oi_str = f"{int(r['call_oi']):>7,}" if r['call_oi'] else "      -"
            p_oi_str = f"{int(r['put_oi']):>7,}"  if r['put_oi']  else "      -"

            print(f"  {K:>7.2f}  "
                  f"{(r['call_settle'] or 0):>6.3f}  "
                  f"{(r['put_settle']  or 0):>6.3f}  "
                  f"{pct(c_set_iv):>9}  "
                  f"{pct(p_set_iv):>9}  "
                  f"{pct(c_live_iv):>9}  "
                  f"{pct(p_live_iv):>9}  "
                  f"{pcp_str}  "
                  f"{c_oi_str}  "
                  f"{p_oi_str}"
                  f"{marker}")

            if K == atm:
                atm_set_iv  = c_set_iv
                atm_live_iv = c_live_iv

        # Put-call parity note
        print(f"\n  PC_par column = (call_settle - put_settle) - (settle_fwd - K)")
        print(f"  Should be near 0.  Values outside ±0.05 warrant investigation.")

        if atm_set_iv:
            atm_ivs[sheet] = (atm_set_iv, atm_live_iv)

    return atm_ivs


def test_term_structure(atm_ivs):
    section("6. ATM IV TERM STRUCTURE")
    if not atm_ivs:
        print("  No data")
        return
    print(f"  {'Contract':8}  {'Settle_IV':>10}  {'Live_IV':>10}")
    print(f"  {'─'*35}")
    prev_set = None
    for contract in sorted(atm_ivs, key=lambda c: ice._contract_sort_key(c,'CT')):
        s_iv, l_iv = atm_ivs[contract]
        flag = ""
        if prev_set is not None and s_iv is not None and s_iv < prev_set - 0.02:
            flag = " WARN: IV inversion"
        print(f"  {contract:8}  {pct(s_iv):>10}  {pct(l_iv):>10}{flag}")
        if s_iv is not None:
            prev_set = s_iv


def test_bloomberg_comparison(wb, sheets, futures, mode):
    section("7. BLOOMBERG CSV COMPARISON (exact strike match)")

    bbg_df, bbg_date = load_bloomberg_csv()
    if bbg_df is None:
        print("  Bloomberg CSV not found — skipping")
        return

    print(f"  Bloomberg CSV latest date: {bbg_date}")

    # Compare first two sheets that have Bloomberg data
    checked = 0
    for sheet in sheets:
        if checked >= 2:
            break
        expiry = CT_EXPIRY.get(sheet)
        if not expiry:
            continue
        T = dte_years(expiry)
        if T <= 0:
            continue

        f_data     = futures.get(sheet)
        settle_fwd = f_data.get('settle') if f_data else None
        live_fwd   = ice.get_forward(f_data, mode) if f_data else None
        fwd_live   = live_fwd if mode == 'live' else settle_fwd
        atm        = ice.atm_strike(fwd_live or settle_fwd)

        ice_rows = ice.read_options(wb, sheet)
        ice_atm  = next((r for r in ice_rows if r['strike'] == atm), None)

        if not ice_atm or not settle_fwd:
            continue

        # Bloomberg exact lookup
        bbg_c = bbg_lookup(bbg_df, sheet, 'C', atm)
        bbg_p = bbg_lookup(bbg_df, sheet, 'P', atm)

        ice_c = ice_atm['call_settle']
        ice_p = ice_atm['put_settle']

        bbg_c_iv = iv(bbg_c, settle_fwd, atm, T, True)  if bbg_c else None
        bbg_p_iv = iv(bbg_p, settle_fwd, atm, T, False) if bbg_p else None
        ice_c_iv = iv(ice_c, settle_fwd, atm, T, True)  if ice_c else None
        ice_p_iv = iv(ice_p, settle_fwd, atm, T, False) if ice_p else None

        print(f"\n  {sheet}  ATM={atm}  settle_fwd={settle_fwd}  DTE={T*365:.0f}d")
        print(f"  {'':14}  {'Bloomberg CSV':>14}  {'ICE settle':>14}  {'Diff':>8}")
        print(f"  {'─'*56}")

        def row(label, b, i):
            b_s = f"{b:.4f}" if b else "-"
            i_s = f"{i:.4f}" if i else "-"
            d_s = f"{i-b:+.4f}" if (b and i) else "-"
            print(f"  {label:14}  {b_s:>14}  {i_s:>14}  {d_s:>8}")

        row("Call price",   bbg_c,    ice_c)
        row("Put price",    bbg_p,    ice_p)
        row("Call IV",
            bbg_c_iv, ice_c_iv)  # shows as decimals
        row("Put IV",
            bbg_p_iv, ice_p_iv)

        # IV comparison in bps
        if bbg_c_iv and ice_c_iv:
            diff_bps = (ice_c_iv - bbg_c_iv) * 10000
            status = "PASS" if abs(diff_bps) < 100 else "WARN"
            print(f"\n  Call IV diff: {diff_bps:+.1f} bps  [{status}]")
        if bbg_p_iv and ice_p_iv:
            diff_bps = (ice_p_iv - bbg_p_iv) * 10000
            status = "PASS" if abs(diff_bps) < 100 else "WARN"
            print(f"  Put  IV diff: {diff_bps:+.1f} bps  [{status}]")

        # Note if Bloomberg date < today (prior settle comparison)
        from datetime import date
        if bbg_date < str(date.today()):
            print(f"\n  NOTE: Bloomberg CSV is {bbg_date}, ICE settle is today.")
            print(f"  Price difference reflects today's market move, not data error.")

        checked += 1


def test_oi_sanity(wb, sheets, futures):
    section("8. OI SANITY CHECK")
    print(f"  {'Sheet':8}  {'Strikes':>8}  {'Calls w/ OI':>12}  "
          f"{'Puts w/ OI':>11}  {'Total call OI':>14}  {'Total put OI':>13}")
    print(f"  {'─'*75}")

    for sheet in sheets:
        rows = ice.read_options(wb, sheet)
        if not rows:
            print(f"  {sheet:8}  no data")
            continue
        n_c_oi = sum(1 for r in rows if r['call_oi'] and r['call_oi'] > 0)
        n_p_oi = sum(1 for r in rows if r['put_oi']  and r['put_oi']  > 0)
        tot_c  = sum(r['call_oi'] for r in rows if r['call_oi'])
        tot_p  = sum(r['put_oi']  for r in rows if r['put_oi'])
        print(f"  {sheet:8}  {len(rows):>8}  {n_c_oi:>12}  "
              f"{n_p_oi:>11}  {tot_c:>14,.0f}  {tot_p:>13,.0f}")

    # Futures OI
    print()
    fut = ice.read_futures(wb, 'CT')
    tot_fut_oi = sum(f['oi'] for f in fut.values() if f['oi'])
    print(f"  Total futures OI across all contracts: {tot_fut_oi:,.0f}")


# ─────────────────────────────────────────────────────────────────────────────
# RUN ALL
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("=" * 70)
    print("  ICE RTD PIPELINE — COMPREHENSIVE CROSS-CHECK (CT)")
    print("=" * 70)

    wb      = test_workbook_open()
    sheets  = test_sheets(wb)
    futures = test_futures(wb, mode=None)   # pass None — mode not needed for display
    mode    = test_mode(wb, futures)
    atm_ivs = test_option_chains(wb, sheets, futures, mode)
    test_term_structure(atm_ivs)
    test_bloomberg_comparison(wb, sheets, futures, mode)
    test_oi_sanity(wb, sheets, futures)

    print("\n" + "=" * 70)
    print("  CROSS-CHECK COMPLETE")
    print("=" * 70)


if __name__ == '__main__':
    main()
