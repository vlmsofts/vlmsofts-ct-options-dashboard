r"""
dashboard_futures_producer.py — Phase 1 of the ct-options-dashboard ICE-Connect migration.

Pulls CT (and optionally KC/SB/CC) FUTURES via the icepython ICE Connect API and writes
a JSON file in the EXACT `outrights` dict shape the dashboard's `_ice_to_rtd_shape`
produces, so the dashboard can read it in place of the Excel RTD workbook.

WHY A SEPARATE 32-BIT PROCESS:
  icepython's ICEPython.dll is a 32-bit in-process COM server and must run on the MAIN
  thread. The dashboard is 64-bit Python 3.14 with Flask threaded=True — it cannot host
  icepython. So this producer runs standalone under `py -3.13-32`, writes JSON, and the
  dashboard reads the JSON.

RUN:  py -3.13-32 dashboard_futures_producer.py            # CT, one shot
      py -3.13-32 dashboard_futures_producer.py --commodity CT --once
  (ICE XL must be open + logged in.)

OUTPUT:  <OUT_DIR>\futures_api_<COMMODITY>.json   (atomic write)
  {
    "ts": "2026-06-29T10:15:00",          # local ISO, stamped by the dashboard reader if absent
    "commodity": "CT",
    "mode": "live" | "settle",
    "outrights": { "CTZ6": { last, settle, yest_settle, change, pct_chg, oi, oi_chg,
                             volume, high, low, block_vol, efs_vol, efp_vol,
                             hv10, hv30, hv60, hv90 }, ... }
  }

FIELD MAP (live-verified 2026-06-29 — see DOCS\ICE_CONNECT_MIGRATION_MAP.md in the dashboard repo):
  settle      = RecSet      (NOT Settle — Settle is None intraday)
  yest_settle = PrevPrice   (native; no prior-day CSV diff needed)
  change      = Change      (native net change)
  volume      = Cumulative Volume
  oi          = OpenInt
  block/efs/efp_vol = Block/EFS/EFP Volume (native)
  last, high, low, bid, offer = native
  oi_chg, hv10/30/60/90 = None here (dashboard fills oi_chg from CSV; HV is CSV-computed)

CONSTRAINTS: QUOTE_BATCH<=150 (get_quotes hangs above ~200). Main-thread only. October (V)
excluded for CT per house rules.
"""

import os
import sys
import json
import argparse

# ── Output location: write into the dashboard repo so it reads a local file ──
OUT_DIR = r"C:\Users\Louis\OneDrive - VLM Commodities LTD\Desktop\ct-options-dashboard\api_feed"

QUOTE_BATCH = 150  # hang ceiling is ~200; stay well under

# Connect field request codes, in order. Index positions used when unpacking rows.
FUT_FIELDS = [
    "Last", "Bid", "Offer", "RecSet", "PrevPrice", "Change",
    "High", "Low", "Cumulative Volume", "OpenInt",
    "Block Volume", "EFS Volume", "EFP Volume", "Last Trading Day",
]

# Month code -> number, and the reverse, matching ice_rtd_reader._MON_CODE.
_MON_CODE = {1:'F',2:'G',3:'H',4:'J',5:'K',6:'M',7:'N',8:'Q',9:'U',10:'V',11:'X',12:'Z'}
_CODE_MON = {v: k for k, v in _MON_CODE.items()}

# Per-commodity standard months (mirror COMMODITY_CONFIG in app.py). October (V=10)
# excluded for CT. SB has no Dec.
_STD_MONTHS = {
    'CT': (3, 5, 7, 12),                 # Oct excluded by house rule
    'KC': (3, 5, 7, 9, 12),
    'SB': (3, 5, 7, 10),                 # no Dec
    'CC': (3, 5, 7, 9, 12),
}

# Option last-trading-day (expiry) schedule — MIRRORS app.py ICE_*_EXPIRY (source
# of truth). A future stays listed until its FND, but its OPTIONS expire ~2 months
# earlier; once the option-expiry date has passed there is NO live option chain and
# autolisting '***<contract> ATM:N' HANGS / returns garbage. So: after option
# expiry (but while the future is still live, pre-FND), pull FUTURES ONLY for that
# contract — never attempt its option chain. Keep this in sync with app.py.
_OPT_EXPIRY = {
    'CT': {
        'CTN6':'2026-06-12','CTQ6':'2026-07-17','CTU6':'2026-08-21','CTV6':'2026-09-11',
        'CTX6':'2026-10-16','CTZ6':'2026-11-13',
        'CTF7':'2026-12-18','CTG7':'2027-01-22','CTH7':'2027-02-05','CTJ7':'2027-03-19',
        'CTK7':'2027-04-16','CTM7':'2027-05-21','CTN7':'2027-06-11','CTQ7':'2027-07-16',
        'CTU7':'2027-08-20','CTV7':'2027-09-10','CTX7':'2027-10-15','CTZ7':'2027-11-12',
        'CTF8':'2027-12-17','CTH8':'2028-02-11','CTK8':'2028-04-13','CTN8':'2028-06-09',
        'CTU8':'2028-08-18','CTZ8':'2028-11-10',  # CTZ8 added 2026-06-29 from API LTD
    },
    'KC': {
        'KCN6':'2026-06-12','KCQ6':'2026-07-10','KCU6':'2026-08-14','KCZ6':'2026-11-12',
        'KCH7':'2027-02-10','KCK7':'2027-04-09','KCN7':'2027-06-11','KCU7':'2027-08-13',
        'KCZ7':'2027-11-12','KCH8':'2028-02-11',
    },
    'SB': {
        'SBN6':'2026-06-15','SBQ6':'2026-07-15','SBU6':'2026-08-17','SBV6':'2026-09-15',
        'SBF7':'2026-12-15','SBH7':'2027-02-16','SBK7':'2027-04-15','SBN7':'2027-06-15',
        'SBV7':'2027-09-15','SBF8':'2027-12-15','SBH8':'2028-02-15','SBK8':'2028-04-17',
        'SBN8':'2028-06-15','SBV8':'2028-09-15',
    },
    'CC': {
        'CCN6':'2026-06-12','CCQ6':'2026-07-10','CCU6':'2026-08-14','CCZ6':'2026-11-13',
    },
}


# Alert flag-file: written when a contract has NO hardcoded option-expiry date.
# Lou checks/watches this; the producer also tries the API as a self-heal.
_EXPIRY_ALERT_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), '_MISSING_OPTION_EXPIRY.txt')


def _alert_missing_expiry(dash_code, resolved):
    """Append a line to the alert file so a missing hardcoded expiry is visible.
    Deduped — a contract already alerted is not re-logged. Best-effort; never raises."""
    try:
        if os.path.exists(_EXPIRY_ALERT_FILE):
            with open(_EXPIRY_ALERT_FILE, encoding='utf-8') as f:
                if any(line.startswith(dash_code + '\t') for line in f):
                    return  # already alerted
        with open(_EXPIRY_ALERT_FILE, 'a', encoding='utf-8') as f:
            f.write(f"{dash_code}\tno hardcoded option-expiry; "
                    f"API LastTradingDay={resolved or 'UNKNOWN'} "
                    f"-> add to _OPT_EXPIRY (and app.py ICE_*_EXPIRY)\n")
    except Exception:
        pass


def _api_option_ltd(ice, prefix, dash_code):
    """Fetch this contract's option Last Trading Day from ICE (self-heal when the
    hardcoded dict lacks it). Quotes ONE ATM-ish option symbol's 'Last Trading Day'.
    Returns 'YYYY-MM-DD' or None. Bounded by the producer's set_timeout."""
    # Build a probe symbol: '***<prefix> <tenor> ATM:1' -> one live option symbol.
    yd = dash_code[-1]
    mc = dash_code[2]            # month code char
    yy = (2020 + int(yd)) % 100
    tenor = f"{mc}{yy:02d}"      # 'Z26'
    try:
        band = list(ice.get_autolist(f'***{prefix.upper()} {tenor} ATM:1'))
        if not band:
            return None
        q = list(ice.get_quotes([band[0]], ["Last Trading Day"], False))
        return q[1][1] if len(q) > 1 else None
    except Exception:
        return None


def _options_expired(dash_code, today_iso, ice=None, prefix=None):
    """True if this contract's OPTIONS have expired (skip its chain, pull futures
    only). Resolution order: (1) hardcoded _OPT_EXPIRY (fast); (2) if missing and
    `ice` given, the API LastTradingDay (self-heal) + write an alert so Lou can add
    it permanently. Unknown + no API -> treat as NOT expired (attempt it) but alert."""
    exp = _OPT_EXPIRY.get(dash_code[:2].upper(), {}).get(dash_code)
    if exp:
        return today_iso > exp
    # Unknown contract — self-heal from the API and alert.
    resolved = _api_option_ltd(ice, prefix, dash_code) if ice else None
    _alert_missing_expiry(dash_code, resolved)
    if resolved:
        return today_iso > resolved
    return False   # can't determine — let build_options try it (set_timeout guards a hang)

# ── CT calendar — LOCKED, canonical (source: flow-analyzer ice_eod_capture.py:47-62,
#    handed over 2026-06-29 as the single source of truth so the dashboard and the
#    analyzer pipeline agree on cotton month rules). Cotton trades ONLY F H K N U X Z;
#    G J M Q V (incl. Oct=V) are excluded; serials F/U/X roll to bona-fide H/Z. ──
_CT_BONAFIDE = {'H', 'K', 'N', 'Z'}
_CT_SERIAL   = {'F': 'H', 'U': 'Z', 'X': 'Z'}
_CT_EXCLUDED = {'G', 'J', 'M', 'Q', 'V'}

# COM stand-down window (ET) — the producer must NOT fire icepython while the
# analyzer's settle_watcher (14:18-16:00) or the 2:25 cotton blotter / 4:00 surface
# pulls are running. Two icepython clients at once is the one untested case.
# (Analyzer's definitive call, 2026-06-29: stand down 14:18-16:30 ET.)
_STANDDOWN_START = (14, 18)
_STANDDOWN_END   = (16, 30)


def _in_standdown_window():
    """True during 14:18-16:30 ET — producer must not call icepython then."""
    try:
        try:
            from zoneinfo import ZoneInfo as _ZI
        except ImportError:
            from backports.zoneinfo import ZoneInfo as _ZI
        import datetime as _dt
        now = _dt.datetime.now(_ZI('America/New_York'))
        hm = (now.hour, now.minute)
        return _STANDDOWN_START <= hm <= _STANDDOWN_END
    except Exception:
        # Fail SAFE: if we can't tell the time, assume stand-down (do not contend COM).
        return True


def _safe_float(v):
    try:
        if v is None:
            return None
        f = float(v)
        # ICE sometimes returns sentinel non-finite values
        if f != f or f in (float('inf'), float('-inf')):
            return None
        return f
    except (TypeError, ValueError):
        return None


def _enumerate_symbols(prefix, years=(6, 7, 8)):
    """Return [(dashboard_code, ice_symbol), ...] for the standard months of `prefix`.
    dashboard_code e.g. 'CTZ6'; ice_symbol e.g. 'CT Z26'. years are single digits
    (6 -> 2026). We over-enumerate; symbols with no live data are dropped after the pull."""
    out = []
    months = _STD_MONTHS.get(prefix.upper(), (3, 5, 7, 12))
    for yd in years:
        full_year = 2020 + yd  # 6 -> 2026
        yy = full_year % 100   # 26
        for m in months:
            mc = _MON_CODE[m]
            dash_code = f"{prefix.upper()}{mc}{yd}"      # CTZ6
            ice_sym   = f"{prefix.upper()} {mc}{yy:02d}" # 'CT Z26'
            out.append((dash_code, ice_sym))
    return out


def _batched(seq, n):
    for i in range(0, len(seq), n):
        yield seq[i:i + n]


def build_outrights(ice, prefix):
    """Pull futures via get_quotes and return (mode, outrights_dict)."""
    pairs = _enumerate_symbols(prefix)
    sym_to_code = {ice_sym: code for code, ice_sym in pairs}
    symbols = [ice_sym for _, ice_sym in pairs]

    outrights = {}
    any_live = False

    for batch in _batched(symbols, QUOTE_BATCH):
        rows = list(ice.get_quotes(batch, FUT_FIELDS, False))
        # rows[0] is the header; rows[1:] are data, each [symbol, *values]
        for rec in rows[1:]:
            ice_sym = rec[0]
            code = sym_to_code.get(ice_sym)
            if not code:
                continue
            vals = dict(zip(FUT_FIELDS, rec[1:]))

            settle      = _safe_float(vals.get("RecSet"))
            yest_settle = _safe_float(vals.get("PrevPrice"))
            change      = _safe_float(vals.get("Change"))
            last        = _safe_float(vals.get("Last"))
            bid         = _safe_float(vals.get("Bid"))
            offer       = _safe_float(vals.get("Offer"))
            volume      = _safe_float(vals.get("Cumulative Volume"))
            oi          = _safe_float(vals.get("OpenInt"))
            high        = _safe_float(vals.get("High"))
            low         = _safe_float(vals.get("Low"))
            block_vol   = _safe_float(vals.get("Block Volume"))
            efs_vol     = _safe_float(vals.get("EFS Volume"))
            efp_vol     = _safe_float(vals.get("EFP Volume"))

            # Drop a contract that has no data at all (not listed / not trading).
            if settle is None and last is None and bid is None and offer is None:
                continue

            # pct_chg from native Change / PrevPrice (mirrors the dashboard compute).
            pct_chg = (round(change / yest_settle * 100, 4)
                       if (change is not None and yest_settle not in (None, 0)) else None)

            if bid is not None or offer is not None or last is not None:
                any_live = True

            outrights[code] = {
                'last':        last,
                'bid':         bid,
                'offer':       offer,
                'settle':      settle,
                'yest_settle': yest_settle,
                'change':      change,
                'pct_chg':     pct_chg,
                'oi':          oi,
                'oi_chg':      None,   # dashboard computes vs yesterday's CSV OI
                'volume':      volume,
                'high':        high,
                'low':         low,
                'block_vol':   block_vol,
                'efs_vol':     efs_vol,
                'efp_vol':     efp_vol,
                'hv10': None, 'hv30': None, 'hv60': None, 'hv90': None,
            }

    mode = 'live' if any_live else 'settle'
    return mode, outrights


# ── Phase 2: option chains — ATM BAND (±N strikes), batched ──────────────────
# Option symbol form: 'CT Z26C76' / 'CT Z26P76.5' = <prefix> <tenor><C|P><strike>.
#
# WHY ATM BAND, NOT FULL CHAIN: pulling every strike of every tenor (500+ symbols)
# intermittently HANGS get_quotes (observed 319s / indefinite). The ±15-strike ATM
# band (~30 symbols/tenor) is small, fast (~1s/strip), and hang-free in testing —
# and it contains the ATM straddle AND the liquid 25-delta skew region the dashboard
# actually uses. Wings beyond the band fall back to settled values (settled surface).
#
# EFFICIENCY: gather ALL strips' band symbols first (cheap autolist per strip), then
# QUOTE THEM IN BATCHES OF <=150 across strips — 2 quote calls for ~240 symbols
# instead of one per strip. (autolist is the slow part; quotes are ~1.5s for all.)
# Autolist is via '***<prefix> <tenor> ATM:N' (NEVER ** — that hangs). ICE returns
# the real listed strikes, so commodity grids (CT 1c, KC 2.5c, SB mixed 0.25/0.50,
# CC scaled) are all handled correctly with no grid assumptions.
#
# Output shape _ice_to_rtd_shape expects:
#   options[<DASHCODE>] = [{strike, call_bid, call_offer, call_last, call_settle,
#                           call_oi, call_iv, put_bid, put_offer, put_last,
#                           put_settle, put_oi, put_iv}, ...]
# NOTE: ImpVol is DELIBERATELY NOT pulled. It is a server-COMPUTED field that
# costs ~1.5s per 30-symbol quote (5x the cost of raw book fields) and would turn
# the 8-strip pull from ~6s into ~18s. The dashboard computes its own B76/SOFR IV
# anyway, and IV is cross-checked against the 4:30 settled_surface (ImpVol_decimal,
# computed once at EOD). Raw fields (Bid/Offer/Last/RecSet/OpenInt) are all cheap.
import re as _re

ATM_BAND = int(os.getenv('ICE_ATM_BAND', '15'))      # strikes each side of ATM
OPT_FIELDS = ["Bid", "Offer", "Last", "RecSet", "OpenInt"]   # NO ImpVol — see note above

# Parse 'Z26C76.5' (prefix+space already stripped) -> (right='C', strike=76.5).
_OPT_RE = _re.compile(r'^[A-Z]\d{2}([CP])(\d+(?:\.\d+)?)$')


def build_options(ice, prefix, fut_pairs):
    """Pull the ATM band for each listed futures tenor. Returns {dashcode: [rows]}.
    fut_pairs is the same (dashcode, ice_sym) list build_outrights enumerated.

    Two-phase: (1) autolist the ATM:N band per strip, mapping each symbol back to
    its dashcode; (2) quote ALL gathered symbols in <=150 batches across strips."""
    # Phase 1: gather band symbols per strip -> {ice_option_symbol: dash_code}
    sym_to_dash = {}
    for dash_code, fut_ice in fut_pairs:
        tenor = fut_ice.split(' ', 1)[1] if ' ' in fut_ice else fut_ice  # 'Z26'
        try:
            band = list(ice.get_autolist(f'***{prefix.upper()} {tenor} ATM:{ATM_BAND}'))
        except Exception:
            continue
        for sym in band:
            sym_to_dash[sym] = dash_code
    if not sym_to_dash:
        return {}

    # Phase 2: quote every band symbol in batches of <=150, across all strips.
    # nested {dash_code: {strike: {'call':{...},'put':{...}}}}
    by_dash = {}
    all_syms = list(sym_to_dash.keys())
    for batch in _batched(all_syms, QUOTE_BATCH):
        try:
            rows = list(ice.get_quotes(batch, OPT_FIELDS, False))
        except Exception:
            continue
        for rec in rows[1:]:
            sym = rec[0]
            dash_code = sym_to_dash.get(sym)
            if not dash_code:
                continue
            tail = sym.split(' ', 1)[1] if ' ' in sym else sym  # 'Z26C76'
            m = _OPT_RE.match(tail)
            if not m:
                continue
            right, strike_s = m.groups()
            strike = _safe_float(strike_s)
            if strike is None:
                continue
            vals = dict(zip(OPT_FIELDS, rec[1:]))
            leg = {
                'bid':    _safe_float(vals.get("Bid")),
                'offer':  _safe_float(vals.get("Offer")),
                'last':   _safe_float(vals.get("Last")),
                'settle': _safe_float(vals.get("RecSet")),
                'oi':     _safe_float(vals.get("OpenInt")),
            }
            slot = by_dash.setdefault(dash_code, {}).setdefault(strike, {})
            slot['call' if right == 'C' else 'put'] = leg

    # Build the per-strike rows in the shape _ice_to_rtd_shape consumes.
    options = {}
    for dash_code, by_strike in by_dash.items():
        rows_out = []
        for strike in sorted(by_strike):
            c = by_strike[strike].get('call', {})
            p = by_strike[strike].get('put', {})
            rows_out.append({
                'strike':      strike,
                'call_bid':    c.get('bid'),    'call_offer':  c.get('offer'),
                'call_last':   c.get('last'),   'call_settle': c.get('settle'),
                'call_oi':     c.get('oi'),
                'put_bid':     p.get('bid'),    'put_offer':   p.get('offer'),
                'put_last':    p.get('last'),   'put_settle':  p.get('settle'),
                'put_oi':      p.get('oi'),
            })
        if rows_out:
            options[dash_code] = rows_out
    return options


def _atomic_write(path, payload):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(payload, f)
    os.replace(tmp, path)  # atomic on the same volume


def run_once(commodity, force=False, with_options=True):
    # Stand down during the analyzer's COM-ownership window (14:18-16:30 ET): two
    # icepython clients (this + settle_watcher) at once is the one untested case.
    # We do NOT pull and do NOT overwrite the JSON — the last-good file simply ages
    # out past the dashboard's 120s staleness guard, which then falls back to
    # Excel/CSV (correct: settle_watcher owns that window). --force overrides.
    if not force and _in_standdown_window():
        print('[producer] in COM stand-down window (14:18-16:30 ET) — skipping pull; '
              'last-good JSON left to age out')
        return None, None
    import icepython as _ice_mod
    from ice_com_hardening import ensure_ice
    ice = ensure_ice()  # 32-bit, ICE XL open; retries + RegAsm + loud flag on failure
    # Per-call timeout: a stuck get_autolist/get_quotes RAISES after this many
    # seconds instead of hanging forever, so the except-continue in build_options
    # skips one bad strip rather than wedging the whole pull. Safety net only —
    # the option-expiry filter below is the real fix for the known hang (expired
    # chains). Env-tunable.
    try:
        _ice_mod.set_timeout(int(os.getenv('ICE_CALL_TIMEOUT_SEC', '8')))
    except Exception:
        pass
    mode, outrights = build_outrights(ice, commodity)

    # Phase 2: live option chains (default on). The live smile / ATM vol / live
    # straddles need these; without them the dashboard labels vol SETTLE. Pulling
    # the chains is heavier (autolist + 150-batch quotes), so it can be skipped
    # with --no-options for a fast futures-only pull.
    options = {}
    if with_options:
        fut_pairs = _enumerate_symbols(commodity)
        # ET date for the option-expiry check.
        try:
            try:
                from zoneinfo import ZoneInfo as _ZI
            except ImportError:
                from backports.zoneinfo import ZoneInfo as _ZI
            import datetime as _dt
            _today = _dt.datetime.now(_ZI('America/New_York')).strftime('%Y-%m-%d')
        except Exception:
            _today = ''
        # Pull chains ONLY for contracts that (a) came back live in the futures pull
        # AND (b) whose OPTIONS have not expired. A contract past option-expiry but
        # pre-FND still trades as a future (already captured above) — its option
        # chain is gone and autolisting it HANGS. Futures-only for those.
        live_pairs = [
            (c, s) for (c, s) in fut_pairs
            if c in outrights
            and not (_today and _options_expired(c, _today, ice=ice, prefix=commodity))
        ]
        options = build_options(ice, commodity, live_pairs)

    payload = {
        # ts intentionally omitted — the dashboard stamps read-time freshness itself.
        'commodity': commodity.upper(),
        'mode': mode,
        'outrights': outrights,
        'options': options,   # {} when --no-options
    }
    out_path = os.path.join(OUT_DIR, f"futures_api_{commodity.upper()}.json")
    _atomic_write(out_path, payload)
    _n_opt_strikes = sum(len(v) for v in options.values())
    print(f"[producer] {commodity.upper()}: wrote {len(outrights)} futures + "
          f"{len(options)} chains ({_n_opt_strikes} strikes) (mode={mode}) -> {out_path}")
    return out_path, payload


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--commodity', default='CT')
    ap.add_argument('--once', action='store_true', help='single pull then exit (default)')
    ap.add_argument('--force', action='store_true',
                    help='override the 14:18-16:30 ET COM stand-down (manual testing only)')
    ap.add_argument('--no-options', action='store_true',
                    help='skip the heavier option-chain pull (futures only)')
    args = ap.parse_args()
    try:
        run_once(args.commodity, force=args.force, with_options=not args.no_options)
    except Exception as e:
        print(f"[producer] FAILED: {e!r}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
