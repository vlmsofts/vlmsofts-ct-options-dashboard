from flask import Flask, render_template, jsonify, request
import os, requests, csv, io, json, time, math, logging
from datetime import datetime, timedelta

log = logging.getLogger(__name__)
try:
    import rtd_reader as _rtd_reader
except ImportError:
    _rtd_reader = None

server = Flask(__name__)

OPT_CSV_URL = "https://raw.githubusercontent.com/vlmsofts/oi-dashboard/main/data/options_oi.csv"
OI_CSV_URL  = "https://raw.githubusercontent.com/vlmsofts/oi-dashboard/main/data/oi_data.csv"

MONTH_CODE = {'F':1,'G':2,'H':3,'J':4,'K':5,'M':6,'N':7,'Q':8,'U':9,'V':10,'X':11,'Z':12}
MONTH_NAME = {1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',
              7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}

# Futures generic ticker suffix (e.g. CTJUL1 -> month 7)
FUTURES_MONTH_SUFFIX = {
    1:'JAN',2:'FEB',3:'MAR',4:'APR',5:'MAY',6:'JUN',
    7:'JUL',8:'AUG',9:'SEP',10:'OCT',11:'NOV',12:'DEC',
}
FUTURES_MONTH_FROM_SUFFIX = {v: k for k, v in FUTURES_MONTH_SUFFIX.items()}
# Cotton only has futures for these months; serial month options map to the next one
CT_STANDARD_MONTHS = (3, 5, 7, 12)
CT_EXCLUDED_MONTHS = {10}   # October options are illiquid; omit from all displays

_cache = {}
CACHE_TTL = 60
RISK_FREE  = 0.045

LOCAL_OPT_HISTORY = os.path.join(os.path.dirname(__file__), 'local_options_history.csv')
LOCAL_FUT_HISTORY = os.path.join(os.path.dirname(__file__), 'local_futures_history.csv')
LOCAL_KC_OPT_HISTORY = os.path.join(os.path.dirname(__file__), 'local_kc_options_history.csv')
LOCAL_KC_FUT_HISTORY = os.path.join(os.path.dirname(__file__), 'local_kc_futures_history.csv')
LOCAL_SB_OPT_HISTORY = os.path.join(os.path.dirname(__file__), 'local_sb_options_history.csv')
LOCAL_SB_FUT_HISTORY = os.path.join(os.path.dirname(__file__), 'local_sb_futures_history.csv')
LOCAL_CC_OPT_HISTORY = os.path.join(os.path.dirname(__file__), 'local_cc_options_history.csv')
LOCAL_CC_FUT_HISTORY = os.path.join(os.path.dirname(__file__), 'local_cc_futures_history.csv')

_skew_hist_cache = {}   # keyed by commodity

# ICE Cotton No. 2 option last-trading-day dates sourced from Bloomberg (OMON <Go>).
# RTD 'all options' sheet will override these when Excel is live; this dict is the
# authoritative fallback so DTE is never wrong when RTD is unavailable.
ICE_CT_EXPIRY = {
    # 2026
    'CTN6': '2026-06-12',
    'CTQ6': '2026-07-17',
    'CTU6': '2026-08-21',
    'CTV6': '2026-09-11',
    'CTX6': '2026-10-16',
    'CTZ6': '2026-11-13',
    # 2027
    'CTF7': '2026-12-18',
    'CTG7': '2027-01-22',
    'CTH7': '2027-02-05',
    'CTJ7': '2027-03-19',
    'CTK7': '2027-04-16',
    'CTM7': '2027-05-21',
    'CTN7': '2027-06-11',
    'CTQ7': '2027-07-16',
    'CTU7': '2027-08-20',
    'CTV7': '2027-09-10',
    'CTX7': '2027-10-15',
    'CTZ7': '2027-11-12',
    # 2028
    'CTF8': '2027-12-17',
    'CTH8': '2028-02-11',
    'CTK8': '2028-04-13',
    'CTN8': '2028-06-09',
    'CTU8': '2028-08-18',
}

COMMODITY_CONFIG = {
    'CT': {
        'prefix': 'CT', 'name': 'ICE Cotton No. 2',
        'opt_csv': LOCAL_OPT_HISTORY, 'fut_csv': LOCAL_FUT_HISTORY,
        'std_months': frozenset({3, 5, 7, 12}), 'excl_months': frozenset({10}),
        'serial_map': {1:3, 2:3, 4:5, 6:7, 8:12, 9:12, 11:12},
        'expiry_override': ICE_CT_EXPIRY,
    },
    'KC': {
        'prefix': 'KC', 'name': 'ICE Coffee C',
        'opt_csv': LOCAL_KC_OPT_HISTORY, 'fut_csv': LOCAL_KC_FUT_HISTORY,
        'std_months': frozenset({3, 5, 7, 9, 12}), 'excl_months': frozenset(),
        'serial_map': {8: 9},
        'expiry_override': {},
    },
    'SB': {
        'prefix': 'SB', 'name': 'ICE Sugar No. 11',
        'opt_csv': LOCAL_SB_OPT_HISTORY, 'fut_csv': LOCAL_SB_FUT_HISTORY,
        'std_months': frozenset({3, 5, 7, 10}), 'excl_months': frozenset(),
        'serial_map': {1:3, 8:10, 9:10},
        'expiry_override': {},
    },
    'CC': {
        'prefix': 'CC', 'name': 'ICE Cocoa',
        'opt_csv': LOCAL_CC_OPT_HISTORY, 'fut_csv': LOCAL_CC_FUT_HISTORY,
        'std_months': frozenset({3, 5, 7, 9, 12}), 'excl_months': frozenset(),
        'serial_map': {8: 9},
        'expiry_override': {},
    },
}

# ── HTTP cache ────────────────────────────────────────────────────────────────

def fetch_csv(url):
    now = time.time()
    if url in _cache and now - _cache[url]['ts'] < CACHE_TTL:
        return _cache[url]['data']
    r = requests.get(url, timeout=15)
    r.raise_for_status()
    reader = csv.DictReader(io.StringIO(r.text))
    data = list(reader)
    _cache[url] = {'data': data, 'ts': now}
    return data

def read_local_csv(path):
    """Read a local CSV and return list of dicts — same interface as fetch_csv()."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"Local data file not found: {path}")
    with open(path, 'r', newline='', encoding='utf-8') as f:
        return list(csv.DictReader(f))

# ── Black-76 (Python) ─────────────────────────────────────────────────────────

def _ncdf(x):
    return 0.5 * (1.0 + math.erf(x / math.sqrt(2.0)))

def _npdf(x):
    return math.exp(-0.5 * x * x) / math.sqrt(2.0 * math.pi)

def b76_price(F, K, T, r, sigma, is_call):
    if T <= 0:
        return max(0.0, (F - K) if is_call else (K - F))
    try:
        d1 = (math.log(F / K) + 0.5 * sigma * sigma * T) / (sigma * math.sqrt(T))
    except (ValueError, ZeroDivisionError):
        return 0.0
    d2 = d1 - sigma * math.sqrt(T)
    df = math.exp(-r * T)
    if is_call:
        return df * (F * _ncdf(d1) - K * _ncdf(d2))
    return df * (K * _ncdf(-d2) - F * _ncdf(-d1))

def b76_vega(F, K, T, r, sigma):
    if T <= 0 or sigma <= 0:
        return 0.0
    try:
        d1 = (math.log(F / K) + 0.5 * sigma * sigma * T) / (sigma * math.sqrt(T))
    except (ValueError, ZeroDivisionError):
        return 0.0
    return math.exp(-r * T) * F * _npdf(d1) * math.sqrt(T)

def b76_theta(F, K, T, r, sigma, is_call):
    """Daily theta (cents/day). Black-76: decay + interest on discounted premium."""
    if T <= 0 or sigma <= 0:
        return 0.0
    try:
        d1 = (math.log(F / K) + 0.5 * sigma * sigma * T) / (sigma * math.sqrt(T))
    except (ValueError, ZeroDivisionError):
        return 0.0
    d2  = d1 - sigma * math.sqrt(T)
    df  = math.exp(-r * T)
    t1  = -(F * df * _npdf(d1) * sigma) / (2.0 * math.sqrt(T))
    if is_call:
        t2 = -r * df * (F * _ncdf(d1) - K * _ncdf(d2))
    else:
        t2 = -r * df * (K * _ncdf(-d2) - F * _ncdf(-d1))
    return (t1 + t2) / 365.0

def b76_delta(F, K, T, r, sigma, is_call):
    if T <= 0:
        return (1.0 if F > K else 0.0) if is_call else (-1.0 if F < K else 0.0)
    try:
        d1 = (math.log(F / K) + 0.5 * sigma * sigma * T) / (sigma * math.sqrt(T))
    except (ValueError, ZeroDivisionError):
        return 0.0
    df = math.exp(-r * T)
    return df * _ncdf(d1) if is_call else -df * _ncdf(-d1)

def implied_vol(mkt_price, F, K, T, r, is_call):
    if T <= 0 or not mkt_price or mkt_price <= 0:
        return None
    if F <= 0 or K <= 0:
        return None
    intrinsic = max(0.0, (F - K) if is_call else (K - F))
    if mkt_price <= intrinsic * 0.999:
        return None
    sigma = 0.20
    for _ in range(100):
        price = b76_price(F, K, T, r, sigma, is_call)
        vega  = b76_vega(F, K, T, r, sigma)
        if vega < 1e-10:
            break
        diff = price - mkt_price
        if abs(diff) < 1e-8:
            break
        sigma -= diff / vega
        sigma = max(0.001, min(sigma, 5.0))
    if abs(b76_price(F, K, T, r, sigma, is_call) - mkt_price) > 0.05:
        return None
    if sigma > 1.50 or sigma < 0.01:
        return None
    return sigma

def _delta_interp(points, target_delta):
    """Linear interpolation of IV at target_delta from a list of (delta, iv) pairs."""
    if len(points) < 2:
        return None
    pts = sorted(points, key=lambda x: x[0])
    for i in range(len(pts) - 1):
        d0, iv0 = pts[i]
        d1, iv1 = pts[i+1]
        if d0 <= target_delta <= d1:
            if abs(d1 - d0) < 1e-6:
                return iv0
            return iv0 + (target_delta - d0) / (d1 - d0) * (iv1 - iv0)
    return None

# ── Ticker / security_des parsing ─────────────────────────────────────────────
# Two CSV formats exist in the same file:
#   Old: security_des="CTN6P    62", strike_px=""  (strike embedded in sec_des)
#   New: security_des="CTN6P",       strike_px="83" (strike in its own column)
# put_call column is 'P', 'C', or blank — unreliable; use security_des suffix instead.

def parse_security_des(sec_des, fallback_strike=None):
    """
    Handles both CSV formats.
    'CTN6P    62' or ('CTN6P', fallback_strike='83')
      -> {'ticker': 'CTN6', 'pc': 'Put',  'strike': 62.0}
    'CTN6C    90' or ('CTN6C', fallback_strike='90')
      -> {'ticker': 'CTN6', 'pc': 'Call', 'strike': 90.0}
    Returns None on parse failure.
    """
    try:
        parts = sec_des.strip().split()
        if not parts:
            return None
        raw = parts[0]  # e.g. 'CTN6P' or 'CTN6C'

        if raw.endswith('P'):
            pc, ticker = 'Put', raw[:-1]
        elif raw.endswith('C'):
            pc, ticker = 'Call', raw[:-1]
        else:
            return None

        # Strike: from security_des (old format) or from strike_px column (new format)
        if len(parts) >= 2:
            strike = float(parts[1])
        elif fallback_strike not in (None, '', '0', 0):
            strike = float(fallback_strike)
        else:
            return None

        return {'ticker': ticker, 'pc': pc, 'strike': strike}
    except Exception:
        pass
    return None

def parse_ct_ticker(ticker):
    """'CTN6' -> (month_code='N', year=2026, month_num=7) or None"""
    if not ticker or len(ticker) < 4 or not ticker.startswith('CT'):
        return None
    code = ticker[2]
    if code not in MONTH_CODE:
        return None
    try:
        year_digit = int(ticker[3])
    except ValueError:
        return None
    return code, 2020 + year_digit, MONTH_CODE[code]

def parse_futures_my(contract, date_s):
    """
    'CTJUL1', '2025-01-02' -> (month=7, year=2025)
    'CTJUL2', '2025-01-02' -> (month=7, year=2026)
    Uses ordinal numbering: ordinal 1 = nearest upcoming occurrence of that month.
    Returns None on failure.
    """
    if not contract or not contract.startswith('CT') or len(contract) < 6:
        return None
    suffix  = contract[2:5].upper()
    ord_str = contract[5:]
    month_num = FUTURES_MONTH_FROM_SUFFIX.get(suffix)
    if not month_num:
        return None
    try:
        ordinal = int(ord_str)
    except (ValueError, TypeError):
        return None
    try:
        d = datetime.strptime(date_s, '%Y-%m-%d')
    except ValueError:
        return None
    first_year = d.year if d.month <= month_num else d.year + 1
    return month_num, first_year + (ordinal - 1)


# Explicit anchor for CT serial month options -> underlying standard futures month
CT_SERIAL_FUTURES = {
    1: 3,   # Jan -> Mar
    2: 3,   # Feb -> Mar
    4: 5,   # Apr -> May
    6: 7,   # Jun -> Jul
    8: 12,  # Aug -> Dec (Oct excluded as illiquid)
    9: 12,  # Sep -> Dec
    11: 12, # Nov -> Dec
}

def next_standard_month(month_num, year):
    """Map a serial CT month to its anchor standard futures month."""
    std_m = CT_SERIAL_FUTURES.get(month_num)
    if std_m:
        return std_m, year
    # Already a standard month — shouldn't be called, but safe fallback
    return month_num, year

def option_expiry_date(month_num, year):
    """
    ICE Cotton options expire on the last Friday of the month preceding the
    delivery month — applies to both standard and serial option months.
    e.g. CTN6 (Jul delivery) -> last Friday of Jun 2026
         CTU6 (Sep serial)   -> last Friday of Aug 2026  (NOT Dec first_notice)
    """
    prev_m, prev_y = (month_num - 1, year) if month_num > 1 else (12, year - 1)
    # Last day of prev_m
    if prev_m == 12:
        last_day = datetime(prev_y + 1, 1, 1) - timedelta(days=1)
    else:
        last_day = datetime(prev_y, prev_m + 1, 1) - timedelta(days=1)
    # Roll back to Friday (Mon=0 … Fri=4)
    days_back = (last_day.weekday() - 4) % 7
    return (last_day - timedelta(days=days_back)).strftime('%Y-%m-%d')


def ticker_label(ticker):
    p = parse_ct_ticker(ticker)
    if not p:
        return ticker
    _, year, month_num = p
    return f"{MONTH_NAME[month_num]} {str(year)[-2:]}"

def parse_generic_ticker(ticker, prefix):
    """'KCN6' -> ('N', 2026, 7) using given prefix, or None."""
    plen = len(prefix)
    if not ticker or len(ticker) < plen + 2 or not ticker.startswith(prefix):
        return None
    code = ticker[plen]
    if code not in MONTH_CODE:
        return None
    try:
        year_digit = int(ticker[plen + 1])
    except (ValueError, IndexError):
        return None
    return code, 2020 + year_digit, MONTH_CODE[code]


def parse_futures_my_generic(contract, date_s, prefix):
    """'KCJUL1', '2026-05-01', 'KC' -> (7, 2026) — ordinal-based like parse_futures_my."""
    plen = len(prefix)
    if not contract or not contract.startswith(prefix) or len(contract) < plen + 4:
        return None
    suffix  = contract[plen:plen + 3].upper()
    ord_str = contract[plen + 3:]
    month_num = FUTURES_MONTH_FROM_SUFFIX.get(suffix)
    if not month_num:
        return None
    try:
        ordinal = int(ord_str)
    except (ValueError, TypeError):
        return None
    try:
        d = datetime.strptime(date_s, '%Y-%m-%d')
    except ValueError:
        return None
    first_year = d.year if d.month <= month_num else d.year + 1
    return month_num, first_year + (ordinal - 1)


def _kc_opt_expiry(month_num, year):
    """KC options expire 8 biz days before first biz day of delivery month."""
    d = datetime(year, month_num, 1)
    while d.weekday() >= 5:
        d += timedelta(days=1)
    count = 0
    while count < 8:
        d -= timedelta(days=1)
        if d.weekday() < 5:
            count += 1
    return d.strftime('%Y-%m-%d')


# ── Core data load ────────────────────────────────────────────────────────────

def load_data(commodity='CT'):
    if commodity != 'CT':
        return _load_generic_data(commodity)

    try:
        opt_rows = read_local_csv(LOCAL_OPT_HISTORY)
        oi_rows  = read_local_csv(LOCAL_FUT_HISTORY)
    except Exception as e:
        return {'error': str(e)}

    # Live Bloomberg RTD overlay — non-blocking; None if Excel not open / file missing
    rtd = None
    if _rtd_reader:
        try:
            rtd = _rtd_reader.read_live()
        except Exception as e:
            log.debug('RTD fetch skipped: %s', e)

    # Parse every CT options row upfront — extract ticker, pc, strike from security_des
    ct_opts = []
    for r in opt_rows:
        if r.get('commodity', '').strip().upper() != 'CT':
            continue
        parsed = parse_security_des(r.get('security_des', ''), r.get('strike_px'))
        if not parsed:
            continue
        try:
            px  = float(r.get('px_settle', 0) or 0)
            oi  = int(r.get('open_int', 0) or 0)
            oic = int(r.get('oi_chg', 0) or 0)
            vol = float(r.get('px_volume', 0) or 0)
        except (ValueError, TypeError):
            continue
        ct_opts.append({
            'date':   r.get('date', '').strip(),
            'ticker': parsed['ticker'],
            'pc':     parsed['pc'],       # 'Call' or 'Put'
            'strike': parsed['strike'],
            'px':     px,
            'oi':     oi,
            'oi_chg': oic,
            'vol':    vol,
        })

    ct_fut = [r for r in oi_rows
              if r.get('commodity', '').strip().upper() == 'CT']

    if not ct_opts:
        return {'error': 'No CT options data found'}

    # Dates
    all_dates = sorted(set(r['date'] for r in ct_opts if r['date']))
    last_date = all_dates[-1]
    prev_date = all_dates[-2] if len(all_dates) >= 2 else last_date
    last_dt   = datetime.strptime(last_date, '%Y-%m-%d')
    week_target = last_dt - timedelta(days=7)
    week_date = max(
        (d for d in all_dates if datetime.strptime(d, '%Y-%m-%d') <= week_target),
        default=prev_date
    )

    # generic_settle['CTJUL1']['2025-01-03'] = 69.89  — every row, no filtering
    generic_settle = {}
    for row in ct_fut:
        contract = (row.get('contract') or '').strip()
        settle_s = (row.get('settle')   or '').strip()
        date_s   = (row.get('date')     or '').strip()
        if not contract or not settle_s or not date_s:
            continue
        try:
            val = float(settle_s)
        except (ValueError, TypeError):
            continue
        if contract not in generic_settle:
            generic_settle[contract] = {}
        generic_settle[contract][date_s] = val

    # Futures lookup: (month_num, year) -> {settle, last_trade, first_notice}
    # Uses last_trade when available (accurate); falls back to ordinal parsing.
    fut_lookup = {}
    for row in ct_fut:
        contract = (row.get('contract')     or '').strip()
        lt_str   = (row.get('last_trade')   or '').strip()
        fn_str   = (row.get('first_notice') or '').strip()
        settle_s = (row.get('settle')       or '').strip()
        date_s   = (row.get('date')         or '').strip()
        if not contract or not settle_s or not date_s:
            continue
        key = None
        if lt_str:
            try:
                lt_dt = datetime.strptime(lt_str, '%Y-%m-%d')
                key = (lt_dt.month, lt_dt.year)
            except ValueError:
                pass
        if key is None:
            result = parse_futures_my(contract, date_s)
            if result:
                key = result
        if key is None:
            continue
        try:
            settle_f = float(settle_s)
        except (ValueError, TypeError):
            continue
        if key not in fut_lookup or date_s > fut_lookup[key]['date']:
            fut_lookup[key] = {
                'settle':       settle_f,
                'last_trade':   lt_str or None,
                'first_notice': fn_str or None,
                'date':         date_s,
            }

    # Active expiries from today's data (base tickers: CTN6, CTZ6 etc.)
    today_opts = [r for r in ct_opts if r['date'] == last_date]
    seen = {}
    for row in today_opts:
        t = row['ticker']
        if t in seen:
            continue
        p = parse_ct_ticker(t)
        if p:
            seen[t] = p

    expiry_list = sorted(
        (t for t, p in seen.items() if p[2] not in CT_EXCLUDED_MONTHS),
        key=lambda t: (seen[t][1], seen[t][2])
    )

    expiry_labels = {}
    futures       = {}
    last_trade    = {}

    for ticker in expiry_list:
        _, year, month_num = seen[ticker]
        expiry_labels[ticker] = ticker_label(ticker)
        key = (month_num, year)
        entry = fut_lookup.get(key)
        if entry is None and month_num not in CT_STANDARD_MONTHS:
            # Serial month: use next standard CT futures month as the underlying
            std_m, std_y = next_standard_month(month_num, year)
            entry = fut_lookup.get((std_m, std_y))
        if entry:
            futures[ticker] = entry['settle']
        # Use hardcoded Bloomberg expiry dates as the primary source; the computed
        # "last Friday of preceding month" formula is ~14 days wrong for Cotton.
        if ticker in ICE_CT_EXPIRY:
            last_trade[ticker] = ICE_CT_EXPIRY[ticker]
        else:
            last_trade[ticker] = option_expiry_date(month_num, year)

    # RTD 'all options' sheet can further refine expiry dates (e.g. new contracts).
    if rtd:
        live_opts_map = rtd.get('live_options') or {}
        for ticker in expiry_list:
            lo = live_opts_map.get(ticker)
            if lo and lo.get('expiry'):
                last_trade[ticker] = lo['expiry']

    # Override futures price with put-call parity implied forward from today's options.
    # This corrects for any timing lag between options and futures CSV updates.
    for ticker in expiry_list:
        lt = last_trade.get(ticker)
        if not lt:
            continue
        try:
            dte_t = max(0, (datetime.strptime(lt, '%Y-%m-%d') -
                            datetime.strptime(last_date, '%Y-%m-%d')).days)
        except ValueError:
            continue
        if dte_t <= 0:
            continue
        T = dte_t / 365.0
        by_strike = {}
        for row in today_opts:
            if row['ticker'] != ticker or row['px'] <= 0:
                continue
            k = row['strike']
            if k not in by_strike:
                by_strike[k] = {}
            by_strike[k][row['pc']] = row['px']
        implied_Fs = []
        for k, pcs in by_strike.items():
            if 'Call' in pcs and 'Put' in pcs:
                implied_Fs.append(k + (pcs['Call'] - pcs['Put']) * math.exp(RISK_FREE * T))
        if len(implied_Fs) >= 3:
            implied_Fs.sort()
            futures[ticker] = implied_Fs[len(implied_Fs) // 2]  # median

    # Historical put-call parity forwards for prev/week dates.
    # Using today's forward with historical option prices breaks IV when the market moves.
    def parity_fwd(ticker, date_str):
        lt = last_trade.get(ticker)
        if not lt:
            return None
        try:
            dte_t = max(0, (datetime.strptime(lt, '%Y-%m-%d') -
                            datetime.strptime(date_str, '%Y-%m-%d')).days)
        except ValueError:
            return None
        if dte_t <= 0:
            return None
        T = dte_t / 365.0
        by_k = {}
        for row in ct_opts:
            if row['date'] != date_str or row['ticker'] != ticker or row['px'] <= 0:
                continue
            by_k.setdefault(row['strike'], {})[row['pc']] = row['px']
        implied_Fs = [
            k + (pcs['Call'] - pcs['Put']) * math.exp(RISK_FREE * T)
            for k, pcs in by_k.items()
            if 'Call' in pcs and 'Put' in pcs
        ]
        if len(implied_Fs) >= 3:
            implied_Fs.sort()
            return implied_Fs[len(implied_Fs) // 2]
        return None

    prev_futures = {t: (parity_fwd(t, prev_date) or futures.get(t)) for t in expiry_list}
    week_futures = {t: (parity_fwd(t, week_date) or futures.get(t)) for t in expiry_list}

    # ATM strike per expiry
    atm_strike = {}
    for ticker in expiry_list:
        fwd = futures.get(ticker)
        if fwd is None:
            continue
        strikes = set(r['strike'] for r in today_opts if r['ticker'] == ticker)
        if strikes:
            atm_strike[ticker] = min(strikes, key=lambda k: abs(k - fwd))

    # RTD ATM strike override (live Bloomberg computed ATM)
    if rtd:
        for tkr, d in (rtd.get('options') or {}).items():
            if tkr in expiry_list and d.get('atm_strike'):
                atm_strike[tkr] = d['atm_strike']

    # ── IV helpers ────────────────────────────────────────────────────────────
    def get_dte(ticker, ref_date):
        lt = last_trade.get(ticker)
        if not lt:
            return 0
        try:
            return max(0, (datetime.strptime(lt, '%Y-%m-%d') -
                           datetime.strptime(ref_date, '%Y-%m-%d')).days)
        except ValueError:
            return 0

    def solve_iv(row, fwd, dte):
        if dte <= 0 or row['px'] <= 0:
            return None
        T = dte / 365.0
        is_call = (row['pc'] == 'Call')
        return implied_vol(row['px'], fwd, row['strike'], T, RISK_FREE, is_call)

    def atm_iv_for_date(ticker, date_str, fwd_override=None):
        fwd = fwd_override if fwd_override is not None else futures.get(ticker)
        if fwd is None:
            return None
        dte = get_dte(ticker, date_str)
        if dte <= 0:
            return None
        T  = dte / 365.0
        df = math.exp(-RISK_FREE * T)
        # For today use the pre-computed ATM strike (includes Bloomberg RTD override).
        # For historical dates use a per-date dynamic ATM from available strikes.
        if fwd_override is None:
            atm = atm_strike.get(ticker)
            if atm is None:
                return None
            rows = [r for r in ct_opts
                    if r['date'] == date_str and r['ticker'] == ticker
                    and abs(r['strike'] - atm) < 0.01]
        else:
            date_rows = [r for r in ct_opts if r['date'] == date_str and r['ticker'] == ticker]
            if not date_rows:
                return None
            avail = set(r['strike'] for r in date_rows)
            atm   = min(avail, key=lambda k: abs(k - fwd))
            rows  = [r for r in date_rows if abs(r['strike'] - atm) < 0.01]
        # Use straddle (call + put) for ATM IV — same method as live Bloomberg feed.
        # Single-option IV is noisier when F ≠ K; straddle cancels the bias.
        call_px = next((r['px'] for r in rows if r['pc'] == 'Call' and r['px'] > 0), None)
        put_px  = next((r['px'] for r in rows if r['pc'] == 'Put'  and r['px'] > 0), None)
        if call_px and put_px:
            strad   = call_px + put_px
            call_eq = (strad + (fwd - atm) * df) / 2.0
            if call_eq > 0:
                iv = implied_vol(call_eq, fwd, atm, T, RISK_FREE, True)
                if iv is not None:
                    return iv
        # Fallback to single option if one side is missing
        for pc in ('Call', 'Put'):
            for row in rows:
                if row['pc'] == pc and row['px'] > 0:
                    iv = solve_iv(row, fwd, dte)
                    if iv is not None:
                        return iv
        return None

    # ── ATM IV today / prev / week ────────────────────────────────────────────
    atm_iv        = {}
    atm_iv_1d_chg = {}
    atm_iv_1w_chg = {}

    for ticker in expiry_list:
        iv_t = atm_iv_for_date(ticker, last_date)
        if iv_t is None:
            continue
        atm_iv[ticker] = round(iv_t * 100, 2)
        iv_p = atm_iv_for_date(ticker, prev_date, prev_futures.get(ticker))
        if iv_p is not None:
            atm_iv_1d_chg[ticker] = round((iv_t - iv_p) * 100, 2)
        iv_w = atm_iv_for_date(ticker, week_date, week_futures.get(ticker))
        if iv_w is not None:
            atm_iv_1w_chg[ticker] = round((iv_t - iv_w) * 100, 2)

    # RTD straddles sheet override — use Bloomberg's live ATM vol directly.
    # Bloomberg computes and displays atm_vol live on the Cotton Straddles RTD sheet.
    # This is more reliable than the settle-based atm_iv_for_date() and provides
    # live data even when the bid/ask straddle computation below cannot run.
    # The live smile bid/ask computation (line ~888) may further refine this value.
    if rtd:
        for tkr, d in (rtd.get('options') or {}).items():
            if tkr not in expiry_list:
                continue
            vol_rtd = d.get('atm_vol_rtd')
            if vol_rtd and vol_rtd > 0:
                # Bloomberg reports ATM vol as a percentage (e.g. 24.4 for 24.4%);
                # guard against decimal form (< 1) just in case.
                atm_iv[tkr] = round(vol_rtd if vol_rtd > 1 else vol_rtd * 100, 2)

    # ATM IV comes entirely from CSV settle + Bloomberg expiry DTE (computed above).
    # The "all options" live sheet feeds the smile chart only — it does not override ATM IV.

    # ── IV Percentile ─────────────────────────────────────────────────────────
    # Uses per-date historical forward (from generic_settle) + per-date ATM strike
    # so history is accurate regardless of today's price level.

    def get_hist_fwd(month_num, year, date_str):
        """
        Look up the historical futures settle for (month_num, year) on date_str
        using generic contract tickers (CTJUL1, CTMAR1 etc.).
        Serial CT months are mapped to the next standard futures month.
        """
        std_m, std_y = (month_num, year) if month_num in CT_STANDARD_MONTHS \
                       else next_standard_month(month_num, year)
        try:
            d = datetime.strptime(date_str, '%Y-%m-%d')
        except ValueError:
            return None
        first_year = d.year if d.month <= std_m else d.year + 1
        ordinal = std_y - first_year + 1
        if ordinal < 1 or ordinal > 3:
            return None
        suffix   = FUTURES_MONTH_SUFFIX.get(std_m)
        contract = f'CT{suffix}{ordinal}'
        return generic_settle.get(contract, {}).get(date_str)

    iv_percentile  = {}
    history_months = {}

    # Pre-group options by (ticker, date) to avoid repeated full-list scans
    opts_by_ticker_date = {}
    for row in ct_opts:
        key = (row['ticker'], row['date'])
        if key not in opts_by_ticker_date:
            opts_by_ticker_date[key] = []
        opts_by_ticker_date[key].append(row)

    for ticker in expiry_list:
        iv_pct = atm_iv.get(ticker)
        if iv_pct is None:
            continue
        p = parse_ct_ticker(ticker)
        if not p:
            continue
        _, t_year, t_month = p

        date_ivs = {}
        for d in all_dates:
            rows_d = opts_by_ticker_date.get((ticker, d), [])
            if not rows_d:
                continue
            fwd_d = get_hist_fwd(t_month, t_year, d)
            if not fwd_d:
                continue
            strikes_d = set(r['strike'] for r in rows_d)
            atm_d = min(strikes_d, key=lambda k: abs(k - fwd_d))
            dte_d = get_dte(ticker, d)
            for pc in ('Call', 'Put'):
                for row in rows_d:
                    if row['pc'] == pc and abs(row['strike'] - atm_d) < 0.01:
                        iv = solve_iv(row, fwd_d, dte_d)
                        if iv is not None:
                            date_ivs[d] = iv * 100
                            break
                if d in date_ivs:
                    break

        if len(date_ivs) < 2:
            continue

        iv_vals = sorted(date_ivs.values())
        rank = sum(1 for v in iv_vals if v <= iv_pct)
        iv_percentile[ticker] = round(rank / len(iv_vals) * 100)

        sorted_dates = sorted(date_ivs.keys())
        d0 = datetime.strptime(sorted_dates[0],  '%Y-%m-%d')
        d1 = datetime.strptime(sorted_dates[-1], '%Y-%m-%d')
        history_months[ticker] = max(1, (d1.year - d0.year) * 12 + (d1.month - d0.month))

    # ── Skew + C/P ratio ──────────────────────────────────────────────────────
    skew_direction = {}
    skew_value     = {}
    cp_ratio       = {}
    call_oi_total  = {}
    put_oi_total   = {}

    for ticker in expiry_list:
        fwd = futures.get(ticker)
        if fwd is None:
            continue
        dte = get_dte(ticker, last_date)
        if dte <= 0:
            continue
        T = dte / 365.0

        c_oi, p_oi = 0, 0
        strikes_data = []

        for row in today_opts:
            if row['ticker'] != ticker:
                continue
            if row['pc'] == 'Call':
                c_oi += row['oi']
            else:
                p_oi += row['oi']
            if row['px'] <= 0:
                continue
            is_call = (row['pc'] == 'Call')
            iv = implied_vol(row['px'], fwd, row['strike'], T, RISK_FREE, is_call)
            if iv is None:
                continue
            delta_val = b76_delta(fwd, row['strike'], T, RISK_FREE, iv, is_call)
            theta_val = b76_theta(fwd, row['strike'], T, RISK_FREE, iv, is_call)
            strikes_data.append({'k': row['strike'], 'pc': row['pc'],
                                  'iv': iv, 'delta': delta_val, 'theta': theta_val})

        call_oi_total[ticker] = c_oi
        put_oi_total[ticker]  = p_oi
        cp_ratio[ticker] = round(c_oi / p_oi, 2) if p_oi > 0 else None

        calls_25 = [s for s in strikes_data if s['pc'] == 'Call' and 0.10 <= s['delta'] <= 0.50]
        puts_25  = [s for s in strikes_data if s['pc'] == 'Put'  and -0.50 <= s['delta'] <= -0.10]

        call_iv = min(calls_25, key=lambda s: abs(s['delta'] - 0.25))['iv'] * 100 if calls_25 else None
        put_iv  = min(puts_25,  key=lambda s: abs(s['delta'] + 0.25))['iv'] * 100 if puts_25  else None

        if call_iv is not None and put_iv is not None:
            diff = put_iv - call_iv
            skew_value[ticker] = round(diff, 2)
            skew_direction[ticker] = 'PUTS BID' if diff > 0.5 else 'CALLS BID' if diff < -0.5 else 'NEUTRAL'
        else:
            skew_direction[ticker] = 'NEUTRAL'
            skew_value[ticker]     = 0.0

    # ── RTD-sourced supplementary data ───────────────────────────────────────
    hv_data      = {}
    live_futures = {}
    rtd_spreads  = {}
    data_source  = 'csv_only'

    if rtd:
        data_source = rtd.get('source', 'csv_only')
        for tkr, d in (rtd.get('outrights') or {}).items():
            hv_data[tkr]      = {k: d.get(k) for k in ('hv10', 'hv30', 'hv60', 'hv90')}
            live_futures[tkr] = {k: d.get(k) for k in
                                 ('last', 'settle', 'change', 'pct_chg',
                                  'oi', 'oi_chg', 'volume', 'high', 'low')}
        rtd_spreads = {tkr: {k: d.get(k) for k in ('settle', 'change', 'legs')}
                       for tkr, d in (rtd.get('spreads') or {}).items()}

    # ── Live smile from Bloomberg 'all options' sheet ─────────────────────────
    # Parity forward derived from live mid prices — consistent with every strike.
    # OTM convention: puts for K < ATM, calls for K > ATM, average at ATM.
    live_smile     = {}   # {ticker: {strike: iv_pct}}
    live_smile_fwd = {}   # {ticker: live_parity_forward}
    if rtd:
        live_opts_map = rtd.get('live_options') or {}
        for ticker in expiry_list:
            lo = live_opts_map.get(ticker)
            if not lo or not lo.get('strikes'):
                continue
            dte = get_dte(ticker, last_date)
            if dte <= 0:
                continue
            T = dte / 365.0

            # Group mid prices by strike for parity forward computation.
            # Prefer (bid+ask)/2 — same source as the live smile section — so the
            # ATM straddle IV override uses live market prices, not stale last-trade.
            by_k = {}
            for s in lo['strikes']:
                bid, ask = s.get('bid'), s.get('ask')
                if bid and ask and bid > 0 and ask > 0:
                    px = (bid + ask) / 2.0
                elif s.get('mid') and s['mid'] > 0:
                    px = s['mid']
                else:
                    px = s.get('last')
                if px and px > 0:
                    by_k.setdefault(s['strike'], {})[s['pc']] = px

            # Live parity forward from mid prices
            impl_Fs = sorted([
                k + (pcs['Call'] - pcs['Put']) * math.exp(RISK_FREE * T)
                for k, pcs in by_k.items() if 'Call' in pcs and 'Put' in pcs
            ])
            if len(impl_Fs) >= 3:
                live_F = impl_Fs[len(impl_Fs) // 2]
            else:
                live_F = futures.get(ticker)
            if not live_F:
                continue

            live_smile_fwd[ticker] = round(live_F, 4)

            # ATM strike from live forward
            avail = sorted(by_k.keys())
            live_atm = min(avail, key=lambda k: abs(k - live_F)) if avail else atm_strike.get(ticker)
            if not live_atm:
                continue

            # OTM smile: collect IVs per strike, average call+put at ATM
            # Filters: delta 0.03–0.55, min price 0.05¢, IV within [50%, 250%] of ATM IV
            # The ATM IV bound catches stale/no-market quotes that survive the delta filter
            # (especially far-dated options where even deep OTM strikes have delta > 0.03)
            settle_atm_iv = (atm_iv.get(ticker) or 20) / 100.0
            iv_lo = settle_atm_iv * 0.50
            iv_hi = settle_atm_iv * 2.50
            smile_ivs = {}
            for s in lo['strikes']:
                K   = s['strike']
                bid = s.get('bid')
                ask = s.get('ask')
                # Require a live two-sided market — stale last-trade prices produce
                # artificial dips/spikes that survive the IV bound filter.
                if not bid or not ask or bid < 0.02 or ask < 0.02 or ask > bid * 8:
                    continue
                px = (bid + ask) / 2.0
                if px < 0.05:
                    continue
                if K < live_atm and s['pc'] != 'Put':
                    continue
                if K > live_atm and s['pc'] != 'Call':
                    continue
                iv = implied_vol(px, live_F, K, T, RISK_FREE, s['pc'] == 'Call')
                if not iv:
                    continue
                d = abs(b76_delta(live_F, K, T, RISK_FREE, iv, s['pc'] == 'Call'))
                if d < 0.03:
                    continue
                # Two-tier IV coherence: near-ATM (d>0.15) must stay within 25% of
                # settle ATM IV — catches stale market-maker quotes at liquid strikes
                # where a Bloomberg model would clearly reject them. Wing strikes
                # (d<=0.15) keep the wider [50%, 250%] band.
                if d > 0.15:
                    if not (settle_atm_iv * 0.75 < iv < settle_atm_iv * 1.30):
                        continue
                else:
                    if not (iv_lo < iv < iv_hi):
                        continue
                smile_ivs.setdefault(K, []).append(iv)

            # Require at least 8 clean strike points before publishing the live smile.
            # Far-dated / illiquid contracts (e.g. CTZ7) have too few reliable BQL
            # quotes to draw a meaningful shape — fewer points means the data cannot
            # be trusted and produces a misleading curve vs the settle smile.
            if len(smile_ivs) < 8:
                pass  # not enough quotes — skip live smile for this ticker
            else:
                # Cast strike keys to int so JSON serialises as "83" not "83.0"
                # — JS object lookup liveSmileMap[83] requires key "83", not "83.0"
                smile = {int(k): round(sum(ivs) / len(ivs) * 100, 2)
                         for k, ivs in smile_ivs.items()}
                if smile:
                    live_smile[ticker] = smile

            # ATM IV from live straddle: call_mid + put_mid at ATM strike, live forward
            atm_pxs = by_k.get(live_atm, {})
            call_mid_atm = atm_pxs.get('Call')
            put_mid_atm  = atm_pxs.get('Put')
            if call_mid_atm and put_mid_atm and call_mid_atm > 0 and put_mid_atm > 0:
                strad_mid = call_mid_atm + put_mid_atm
                df_atm    = math.exp(-RISK_FREE * T)
                call_eq   = (strad_mid + (live_F - live_atm) * df_atm) / 2
                if call_eq > 0:
                    live_iv = implied_vol(call_eq, live_F, live_atm, T, RISK_FREE, True)
                    if live_iv and 0.05 < live_iv < 1.50:
                        atm_iv[ticker] = round(live_iv * 100, 2)
                        iv_p = atm_iv_for_date(ticker, prev_date, prev_futures.get(ticker))
                        if iv_p is not None:
                            atm_iv_1d_chg[ticker] = round(live_iv * 100 - iv_p * 100, 2)
                        iv_w = atm_iv_for_date(ticker, week_date, week_futures.get(ticker))
                        if iv_w is not None:
                            atm_iv_1w_chg[ticker] = round(live_iv * 100 - iv_w * 100, 2)

    # RTD straddles sheet — Bloomberg streaming ATM vol overrides stale BQL value.
    # The straddles sheet uses "Dec26"/"Jul26" keys; map to "CTZ6"/"CTN6" etc.
    _STRAD_MONTH = {
        'Jan': 'F', 'Feb': 'G', 'Mar': 'H', 'Apr': 'J', 'May': 'K', 'Jun': 'M',
        'Jul': 'N', 'Aug': 'Q', 'Sep': 'U', 'Oct': 'V', 'Nov': 'X', 'Dec': 'Z',
    }
    if rtd:
        for code, sd in (rtd.get('straddles') or {}).items():
            mo, yr = code[:3], code[3:]
            mc = _STRAD_MONTH.get(mo)
            if not mc or not yr.isdigit():
                continue
            tkr = 'CT' + mc + yr[-1]
            if tkr not in expiry_list:
                continue
            vol = sd.get('atm_vol')
            if vol and vol > 0:
                atm_iv[tkr] = round(vol * 100 if vol < 1 else vol, 2)

    # ── Serialize options for client ──────────────────────────────────────────
    def filter_date(date_str):
        return [r for r in ct_opts if r['date'] == date_str]

    _persist_today(last_date)

    return {
        'last_date':      last_date,
        'prev_date':      prev_date,
        'week_date':      week_date,
        'expiries':       expiry_list,
        'expiry_labels':  expiry_labels,
        'futures':        futures,
        'prev_futures':   prev_futures,
        'week_futures':   week_futures,
        'last_trade':     last_trade,
        'atm_strike':     atm_strike,
        'atm_iv':         atm_iv,
        'atm_iv_1d_chg':  atm_iv_1d_chg,
        'atm_iv_1w_chg':  atm_iv_1w_chg,
        'iv_percentile':  iv_percentile,
        'history_months': history_months,
        'skew_direction': skew_direction,
        'skew_value':     skew_value,
        'cp_ratio':       cp_ratio,
        'call_oi':        call_oi_total,
        'put_oi':         put_oi_total,
        'options_today':  filter_date(last_date),
        'options_prev':   filter_date(prev_date),
        'options_week':   filter_date(week_date),
        'hv_data':        hv_data,
        'live_futures':   live_futures,
        'rtd_spreads':    rtd_spreads,
        'data_source':    data_source,
        'live_smile':     live_smile,
        'live_smile_fwd': live_smile_fwd,
        'commodity':      'CT',
        'commodity_name': 'ICE Cotton No. 2',
    }

# ── Daily persistence ─────────────────────────────────────────────────────────

def _persist_today(_unused=None):
    """Append any dates from GitHub that are newer than local files.
    Called on every load_data() but only fetches when GitHub has dates local doesn't."""
    today_str = datetime.today().strftime('%Y-%m-%d')

    try:
        opt_existing = set()
        if os.path.exists(LOCAL_OPT_HISTORY):
            with open(LOCAL_OPT_HISTORY, 'r', newline='', encoding='utf-8') as f:
                for row in csv.DictReader(f):
                    opt_existing.add(row.get('date', '').strip())
        opt_latest = max(opt_existing) if opt_existing else ''
    except Exception as e:
        log.warning('opt persist check failed: %s', e)
        opt_latest = today_str  # skip on error

    try:
        fut_existing = set()
        if os.path.exists(LOCAL_FUT_HISTORY):
            with open(LOCAL_FUT_HISTORY, 'r', newline='', encoding='utf-8') as f:
                for row in csv.DictReader(f):
                    fut_existing.add(row.get('date', '').strip())
        fut_latest = max(fut_existing) if fut_existing else ''
    except Exception as e:
        log.warning('fut persist check failed: %s', e)
        fut_latest = today_str

    if opt_latest >= today_str and fut_latest >= today_str:
        return  # local is current — no network call needed

    # Local is behind today — fetch GitHub and append any missing dates
    try:
        if opt_latest < today_str:
            all_rows = fetch_csv(OPT_CSV_URL)
            new_rows = [r for r in all_rows
                        if r.get('date', '').strip() > opt_latest
                        and r.get('commodity', '').strip().upper() == 'CT']
            if new_rows:
                new_rows.sort(key=lambda r: r.get('date', ''))
                need_header = not os.path.exists(LOCAL_OPT_HISTORY)
                with open(LOCAL_OPT_HISTORY, 'a', newline='', encoding='utf-8') as f:
                    w = csv.DictWriter(f, fieldnames=new_rows[0].keys())
                    if need_header:
                        w.writeheader()
                    w.writerows(new_rows)
                new_dates = sorted(set(r.get('date','') for r in new_rows))
                log.info('Persisted opt rows for %d new dates: %s', len(new_dates), new_dates)
    except Exception as e:
        log.warning('opt persist failed: %s', e)

    try:
        if fut_latest < today_str:
            all_rows = fetch_csv(OI_CSV_URL)
            today_raw = [r for r in all_rows
                         if r.get('date', '').strip() > fut_latest
                         and r.get('commodity', '').strip().upper() == 'CT']
            if today_raw:
                today_raw.sort(key=lambda r: r.get('date', ''))
                need_header = not os.path.exists(LOCAL_FUT_HISTORY)
                with open(LOCAL_FUT_HISTORY, 'a', newline='', encoding='utf-8') as f:
                    w = csv.DictWriter(f, fieldnames=today_raw[0].keys())
                    if need_header:
                        w.writeheader()
                    w.writerows(today_raw)
                new_dates = sorted(set(r.get('date','') for r in today_raw))
                log.info('Persisted fut rows for %d new dates: %s', len(new_dates), new_dates)
    except Exception as e:
        log.warning('fut persist failed: %s', e)


def _persist_today_generic(commodity):
    """Append any dates from GitHub that are newer than local files for KC/SB/CC."""
    cfg = COMMODITY_CONFIG[commodity]
    today_str = datetime.today().strftime('%Y-%m-%d')

    try:
        opt_existing = set()
        if os.path.exists(cfg['opt_csv']):
            with open(cfg['opt_csv'], 'r', newline='', encoding='utf-8') as f:
                for row in csv.DictReader(f):
                    opt_existing.add(row.get('date', '').strip())
        opt_latest = max(opt_existing) if opt_existing else ''
    except Exception as e:
        log.warning('%s opt persist check failed: %s', commodity, e)
        opt_latest = today_str

    try:
        fut_existing = set()
        if os.path.exists(cfg['fut_csv']):
            with open(cfg['fut_csv'], 'r', newline='', encoding='utf-8') as f:
                for row in csv.DictReader(f):
                    fut_existing.add(row.get('date', '').strip())
        fut_latest = max(fut_existing) if fut_existing else ''
    except Exception as e:
        log.warning('%s fut persist check failed: %s', commodity, e)
        fut_latest = today_str

    if opt_latest >= today_str and fut_latest >= today_str:
        return

    try:
        if opt_latest < today_str:
            all_rows = fetch_csv(OPT_CSV_URL)
            new_rows = [r for r in all_rows
                        if r.get('date', '').strip() > opt_latest
                        and r.get('commodity', '').strip().upper() == commodity]
            if new_rows:
                new_rows.sort(key=lambda r: r.get('date', ''))
                need_header = not os.path.exists(cfg['opt_csv'])
                with open(cfg['opt_csv'], 'a', newline='', encoding='utf-8') as f:
                    w = csv.DictWriter(f, fieldnames=new_rows[0].keys())
                    if need_header:
                        w.writeheader()
                    w.writerows(new_rows)
                log.info('%s: persisted %d opt rows', commodity, len(new_rows))
    except Exception as e:
        log.warning('%s opt persist failed: %s', commodity, e)

    try:
        if fut_latest < today_str:
            all_rows = fetch_csv(OI_CSV_URL)
            new_rows = [r for r in all_rows
                        if r.get('date', '').strip() > fut_latest
                        and r.get('commodity', '').strip().upper() == commodity]
            if new_rows:
                new_rows.sort(key=lambda r: r.get('date', ''))
                need_header = not os.path.exists(cfg['fut_csv'])
                with open(cfg['fut_csv'], 'a', newline='', encoding='utf-8') as f:
                    w = csv.DictWriter(f, fieldnames=new_rows[0].keys())
                    if need_header:
                        w.writeheader()
                    w.writerows(new_rows)
                log.info('%s: persisted %d fut rows', commodity, len(new_rows))
    except Exception as e:
        log.warning('%s fut persist failed: %s', commodity, e)


def _load_generic_data(commodity):
    """Load and compute options analytics for KC, SB, or CC (settle-only, no RTD)."""
    cfg        = COMMODITY_CONFIG[commodity]
    prefix     = cfg['prefix']
    std_months = cfg['std_months']
    excl_months = cfg['excl_months']
    serial_map = cfg['serial_map']

    try:
        opt_rows = read_local_csv(cfg['opt_csv'])
        oi_rows  = read_local_csv(cfg['fut_csv'])
    except Exception as e:
        return {'error': str(e)}

    comm_opts = []
    for r in opt_rows:
        if r.get('commodity', '').strip().upper() != commodity:
            continue
        parsed = parse_security_des(r.get('security_des', ''), r.get('strike_px'))
        if not parsed:
            continue
        try:
            px  = float(r.get('px_settle', 0) or 0)
            oi  = int(r.get('open_int', 0) or 0)
            oic = int(r.get('oi_chg', 0) or 0)
            vol = float(r.get('px_volume', 0) or 0)
        except (ValueError, TypeError):
            continue
        comm_opts.append({
            'date':   r.get('date', '').strip(),
            'ticker': parsed['ticker'],
            'pc':     parsed['pc'],
            'strike': parsed['strike'],
            'px':     px,
            'oi':     oi,
            'oi_chg': oic,
            'vol':    vol,
        })

    comm_fut = [r for r in oi_rows if r.get('commodity', '').strip().upper() == commodity]

    if not comm_opts:
        return {'error': f'No {commodity} options data found'}

    all_dates = sorted(set(r['date'] for r in comm_opts if r['date']))
    last_date = all_dates[-1]
    prev_date = all_dates[-2] if len(all_dates) >= 2 else last_date
    last_dt   = datetime.strptime(last_date, '%Y-%m-%d')
    week_target = last_dt - timedelta(days=7)
    week_date = max(
        (d for d in all_dates if datetime.strptime(d, '%Y-%m-%d') <= week_target),
        default=prev_date
    )

    generic_settle = {}
    for row in comm_fut:
        contract = (row.get('contract') or '').strip()
        settle_s = (row.get('settle')   or '').strip()
        date_s   = (row.get('date')     or '').strip()
        if not contract or not settle_s or not date_s:
            continue
        try:
            generic_settle.setdefault(contract, {})[date_s] = float(settle_s)
        except (ValueError, TypeError):
            pass

    fut_lookup = {}
    for row in comm_fut:
        contract = (row.get('contract')     or '').strip()
        lt_str   = (row.get('last_trade')   or '').strip()
        fn_str   = (row.get('first_notice') or '').strip()
        settle_s = (row.get('settle')       or '').strip()
        date_s   = (row.get('date')         or '').strip()
        if not contract or not settle_s or not date_s:
            continue
        plen = len(prefix)
        if len(contract) < plen + 3:
            continue
        suffix_3   = contract[plen:plen + 3].upper()
        month_num  = FUTURES_MONTH_FROM_SUFFIX.get(suffix_3)
        if not month_num:
            continue
        # Derive delivery year from first_notice (reliable: fn_dt.month <= delivery month always).
        # Falls back to last_trade year, then to ordinal parse.
        # This corrects for Bloomberg ordinal rollover when the front contract expires.
        year = None
        if fn_str:
            try:
                fn_dt = datetime.strptime(fn_str, '%Y-%m-%d')
                year = fn_dt.year if fn_dt.month <= month_num else fn_dt.year + 1
            except ValueError:
                pass
        if year is None and lt_str:
            try:
                lt_dt = datetime.strptime(lt_str, '%Y-%m-%d')
                year = lt_dt.year if lt_dt.month <= month_num else lt_dt.year + 1
            except ValueError:
                pass
        if year is None:
            res = parse_futures_my_generic(contract, date_s, prefix)
            year = res[1] if res else None
        if year is None:
            continue
        key = (month_num, year)
        try:
            settle_f = float(settle_s)
        except (ValueError, TypeError):
            continue
        if key not in fut_lookup or date_s > fut_lookup[key]['date']:
            fut_lookup[key] = {
                'settle':       settle_f,
                'last_trade':   lt_str or None,
                'first_notice': fn_str or None,
                'date':         date_s,
            }

    today_opts = [r for r in comm_opts if r['date'] == last_date]
    seen = {}
    for row in today_opts:
        t = row['ticker']
        if t in seen:
            continue
        p = parse_generic_ticker(t, prefix)
        if p:
            seen[t] = p

    expiry_list = sorted(
        (t for t, p in seen.items() if p[2] not in excl_months),
        key=lambda t: (seen[t][1], seen[t][2])
    )

    expiry_labels = {}
    futures       = {}
    last_trade    = {}

    for ticker in expiry_list:
        _, year, month_num = seen[ticker]
        expiry_labels[ticker] = f"{MONTH_NAME[month_num]} {str(year)[-2:]}"

        key = (month_num, year)
        entry = fut_lookup.get(key)
        if entry is None and month_num not in std_months:
            std_m = serial_map.get(month_num, month_num)
            entry = fut_lookup.get((std_m, year))
        if entry:
            futures[ticker] = entry['settle']

        exp = cfg['expiry_override'].get(ticker)
        if not exp:
            if commodity == 'KC':
                exp = _kc_opt_expiry(month_num, year)
            else:
                std_m = serial_map.get(month_num, month_num) if month_num not in std_months else month_num
                for k in [(month_num, year), (std_m, year)]:
                    e = fut_lookup.get(k)
                    if e and e.get('last_trade'):
                        exp = e['last_trade']
                        break
        if exp:
            last_trade[ticker] = exp

    for ticker in expiry_list:
        lt = last_trade.get(ticker)
        if not lt:
            continue
        try:
            dte_t = max(0, (datetime.strptime(lt, '%Y-%m-%d') -
                            datetime.strptime(last_date, '%Y-%m-%d')).days)
        except ValueError:
            continue
        if dte_t <= 0:
            continue
        T = dte_t / 365.0
        by_strike = {}
        for row in today_opts:
            if row['ticker'] != ticker or row['px'] <= 0:
                continue
            by_strike.setdefault(row['strike'], {})[row['pc']] = row['px']
        implied_Fs = [
            k + (pcs['Call'] - pcs['Put']) * math.exp(RISK_FREE * T)
            for k, pcs in by_strike.items() if 'Call' in pcs and 'Put' in pcs
        ]
        if len(implied_Fs) >= 3:
            implied_Fs.sort()
            futures[ticker] = implied_Fs[len(implied_Fs) // 2]

    def parity_fwd(ticker, date_str):
        lt = last_trade.get(ticker)
        if not lt:
            return None
        try:
            dte_t = max(0, (datetime.strptime(lt, '%Y-%m-%d') -
                            datetime.strptime(date_str, '%Y-%m-%d')).days)
        except ValueError:
            return None
        if dte_t <= 0:
            return None
        T = dte_t / 365.0
        by_k = {}
        for row in comm_opts:
            if row['date'] != date_str or row['ticker'] != ticker or row['px'] <= 0:
                continue
            by_k.setdefault(row['strike'], {})[row['pc']] = row['px']
        implied_Fs = [
            k + (pcs['Call'] - pcs['Put']) * math.exp(RISK_FREE * T)
            for k, pcs in by_k.items() if 'Call' in pcs and 'Put' in pcs
        ]
        if len(implied_Fs) >= 3:
            implied_Fs.sort()
            return implied_Fs[len(implied_Fs) // 2]
        return None

    prev_futures = {t: (parity_fwd(t, prev_date) or futures.get(t)) for t in expiry_list}
    week_futures = {t: (parity_fwd(t, week_date) or futures.get(t)) for t in expiry_list}

    atm_strike = {}
    for ticker in expiry_list:
        fwd = futures.get(ticker)
        if fwd is None:
            continue
        strikes = set(r['strike'] for r in today_opts if r['ticker'] == ticker)
        if strikes:
            atm_strike[ticker] = min(strikes, key=lambda k: abs(k - fwd))

    def get_dte(ticker, ref_date):
        lt = last_trade.get(ticker)
        if not lt:
            return 0
        try:
            return max(0, (datetime.strptime(lt, '%Y-%m-%d') -
                           datetime.strptime(ref_date, '%Y-%m-%d')).days)
        except ValueError:
            return 0

    def solve_iv(row, fwd, dte):
        if dte <= 0 or row['px'] <= 0:
            return None
        T = dte / 365.0
        is_call = (row['pc'] == 'Call')
        return implied_vol(row['px'], fwd, row['strike'], T, RISK_FREE, is_call)

    def atm_iv_for_date(ticker, date_str, fwd_override=None):
        fwd = fwd_override if fwd_override is not None else futures.get(ticker)
        if fwd is None:
            return None
        dte = get_dte(ticker, date_str)
        if dte <= 0:
            return None
        T  = dte / 365.0
        df = math.exp(-RISK_FREE * T)
        if fwd_override is None:
            atm = atm_strike.get(ticker)
            if atm is None:
                return None
            rows = [r for r in comm_opts
                    if r['date'] == date_str and r['ticker'] == ticker
                    and abs(r['strike'] - atm) < 0.01]
        else:
            date_rows = [r for r in comm_opts if r['date'] == date_str and r['ticker'] == ticker]
            if not date_rows:
                return None
            avail = set(r['strike'] for r in date_rows)
            atm   = min(avail, key=lambda k: abs(k - fwd))
            rows  = [r for r in date_rows if abs(r['strike'] - atm) < 0.01]
        call_px = next((r['px'] for r in rows if r['pc'] == 'Call' and r['px'] > 0), None)
        put_px  = next((r['px'] for r in rows if r['pc'] == 'Put'  and r['px'] > 0), None)
        if call_px and put_px:
            strad   = call_px + put_px
            call_eq = (strad + (fwd - atm) * df) / 2.0
            if call_eq > 0:
                iv = implied_vol(call_eq, fwd, atm, T, RISK_FREE, True)
                if iv is not None:
                    return iv
        for pc in ('Call', 'Put'):
            for row in rows:
                if row['pc'] == pc and row['px'] > 0:
                    iv = solve_iv(row, fwd, dte)
                    if iv is not None:
                        return iv
        return None

    atm_iv        = {}
    atm_iv_1d_chg = {}
    atm_iv_1w_chg = {}

    for ticker in expiry_list:
        iv_t = atm_iv_for_date(ticker, last_date)
        if iv_t is None:
            continue
        atm_iv[ticker] = round(iv_t * 100, 2)
        iv_p = atm_iv_for_date(ticker, prev_date, prev_futures.get(ticker))
        if iv_p is not None:
            atm_iv_1d_chg[ticker] = round((iv_t - iv_p) * 100, 2)
        iv_w = atm_iv_for_date(ticker, week_date, week_futures.get(ticker))
        if iv_w is not None:
            atm_iv_1w_chg[ticker] = round((iv_t - iv_w) * 100, 2)

    def get_hist_fwd_generic(month_num, year, date_str):
        std_m = month_num if month_num in std_months else serial_map.get(month_num, month_num)
        try:
            d = datetime.strptime(date_str, '%Y-%m-%d')
        except ValueError:
            return None
        first_year = d.year if d.month <= std_m else d.year + 1
        ordinal = year - first_year + 1
        if ordinal < 1 or ordinal > 3:
            return None
        suffix = FUTURES_MONTH_SUFFIX.get(std_m)
        contract = f'{prefix}{suffix}{ordinal}'
        return generic_settle.get(contract, {}).get(date_str)

    iv_percentile  = {}
    history_months = {}

    opts_by_ticker_date = {}
    for row in comm_opts:
        opts_by_ticker_date.setdefault((row['ticker'], row['date']), []).append(row)

    for ticker in expiry_list:
        iv_pct = atm_iv.get(ticker)
        if iv_pct is None:
            continue
        p = parse_generic_ticker(ticker, prefix)
        if not p:
            continue
        _, t_year, t_month = p

        date_ivs = {}
        for d in all_dates:
            rows_d = opts_by_ticker_date.get((ticker, d), [])
            if not rows_d:
                continue
            fwd_d = get_hist_fwd_generic(t_month, t_year, d)
            if not fwd_d:
                continue
            strikes_d = set(r['strike'] for r in rows_d)
            atm_d = min(strikes_d, key=lambda k: abs(k - fwd_d))
            dte_d = get_dte(ticker, d)
            for pc in ('Call', 'Put'):
                for row in rows_d:
                    if row['pc'] == pc and abs(row['strike'] - atm_d) < 0.01:
                        iv = solve_iv(row, fwd_d, dte_d)
                        if iv is not None:
                            date_ivs[d] = iv * 100
                            break
                if d in date_ivs:
                    break

        if len(date_ivs) < 2:
            continue
        iv_vals = sorted(date_ivs.values())
        rank = sum(1 for v in iv_vals if v <= iv_pct)
        iv_percentile[ticker] = round(rank / len(iv_vals) * 100)

        sorted_dates = sorted(date_ivs.keys())
        d0 = datetime.strptime(sorted_dates[0],  '%Y-%m-%d')
        d1 = datetime.strptime(sorted_dates[-1], '%Y-%m-%d')
        history_months[ticker] = max(1, (d1.year - d0.year) * 12 + (d1.month - d0.month))

    skew_direction = {}
    skew_value     = {}
    cp_ratio       = {}
    call_oi_total  = {}
    put_oi_total   = {}

    for ticker in expiry_list:
        fwd = futures.get(ticker)
        if fwd is None:
            continue
        dte = get_dte(ticker, last_date)
        if dte <= 0:
            continue
        T = dte / 365.0

        c_oi, p_oi = 0, 0
        strikes_data = []
        for row in today_opts:
            if row['ticker'] != ticker:
                continue
            if row['pc'] == 'Call':
                c_oi += row['oi']
            else:
                p_oi += row['oi']
            if row['px'] <= 0:
                continue
            is_call = (row['pc'] == 'Call')
            iv = implied_vol(row['px'], fwd, row['strike'], T, RISK_FREE, is_call)
            if iv is None:
                continue
            delta_val = b76_delta(fwd, row['strike'], T, RISK_FREE, iv, is_call)
            strikes_data.append({'k': row['strike'], 'pc': row['pc'],
                                  'iv': iv, 'delta': delta_val})

        call_oi_total[ticker] = c_oi
        put_oi_total[ticker]  = p_oi
        cp_ratio[ticker] = round(c_oi / p_oi, 2) if p_oi > 0 else None

        calls_25 = [s for s in strikes_data if s['pc'] == 'Call' and 0.10 <= s['delta'] <= 0.50]
        puts_25  = [s for s in strikes_data if s['pc'] == 'Put'  and -0.50 <= s['delta'] <= -0.10]

        call_iv = min(calls_25, key=lambda s: abs(s['delta'] - 0.25))['iv'] * 100 if calls_25 else None
        put_iv  = min(puts_25,  key=lambda s: abs(s['delta'] + 0.25))['iv'] * 100 if puts_25  else None

        if call_iv is not None and put_iv is not None:
            diff = put_iv - call_iv
            skew_value[ticker] = round(diff, 2)
            skew_direction[ticker] = 'PUTS BID' if diff > 0.5 else 'CALLS BID' if diff < -0.5 else 'NEUTRAL'
        else:
            skew_direction[ticker] = 'NEUTRAL'
            skew_value[ticker]     = 0.0

    def filter_date(date_str):
        return [r for r in comm_opts if r['date'] == date_str]

    _persist_today_generic(commodity)

    return {
        'last_date':      last_date,
        'prev_date':      prev_date,
        'week_date':      week_date,
        'expiries':       expiry_list,
        'expiry_labels':  expiry_labels,
        'futures':        futures,
        'prev_futures':   prev_futures,
        'week_futures':   week_futures,
        'last_trade':     last_trade,
        'atm_strike':     atm_strike,
        'atm_iv':         atm_iv,
        'atm_iv_1d_chg':  atm_iv_1d_chg,
        'atm_iv_1w_chg':  atm_iv_1w_chg,
        'iv_percentile':  iv_percentile,
        'history_months': history_months,
        'skew_direction': skew_direction,
        'skew_value':     skew_value,
        'cp_ratio':       cp_ratio,
        'call_oi':        call_oi_total,
        'put_oi':         put_oi_total,
        'options_today':  filter_date(last_date),
        'options_prev':   filter_date(prev_date),
        'options_week':   filter_date(week_date),
        'hv_data':        {},
        'live_futures':   {},
        'rtd_spreads':    {},
        'data_source':    'csv_only',
        'live_smile':     {},
        'live_smile_fwd': {},
        'commodity':      commodity,
        'commodity_name': cfg['name'],
    }


# ── Historical skew ───────────────────────────────────────────────────────────

def compute_skew_history(commodity='CT'):
    """Compute rolling front-month and per-ticker delta-IV time series."""
    cfg        = COMMODITY_CONFIG.get(commodity, COMMODITY_CONFIG['CT'])
    prefix     = cfg['prefix']
    std_months = cfg['std_months']
    excl_months = cfg['excl_months']
    serial_map = cfg['serial_map']
    expiry_override = cfg['expiry_override']

    now = time.time()
    _cache_entry = _skew_hist_cache.setdefault(commodity, {'data': None, 'ts': 0})
    if _cache_entry['data'] is not None and now - _cache_entry['ts'] < 3600:
        return _cache_entry['data']

    try:
        opt_rows = read_local_csv(cfg['opt_csv'])
        oi_rows  = read_local_csv(cfg['fut_csv'])
    except Exception as e:
        return {'error': str(e)}

    ct_opts = []
    seen_keys = set()
    for r in opt_rows:
        if r.get('commodity', '').strip().upper() != commodity:
            continue
        parsed = parse_security_des(r.get('security_des', ''), r.get('strike_px'))
        if not parsed:
            continue
        try:
            px = float(r.get('px_settle', 0) or 0)
        except (ValueError, TypeError):
            continue
        key = (r.get('date','').strip(), parsed['ticker'], parsed['pc'], parsed['strike'])
        if key in seen_keys:
            continue
        seen_keys.add(key)
        ct_opts.append({'date': key[0], 'ticker': parsed['ticker'],
                        'pc': parsed['pc'], 'strike': parsed['strike'], 'px': px})

    generic_settle = {}
    for row in oi_rows:
        if row.get('commodity', '').strip().upper() != commodity:
            continue
        contract = (row.get('contract') or '').strip()
        settle_s = (row.get('settle') or '').strip()
        date_s   = (row.get('date') or '').strip()
        if not contract or not settle_s or not date_s:
            continue
        try:
            generic_settle.setdefault(contract, {})[date_s] = float(settle_s)
        except (ValueError, TypeError):
            pass

    # fut_lt: (delivery_month, year) -> last_trade_str; used for SB/CC option expiry lookup
    # Key by delivery month (from first_notice if available, else last_trade) because
    # SB futures expire in the month before delivery, so last_trade.month ≠ delivery month.
    fut_lt = {}
    for row in oi_rows:
        if row.get('commodity', '').strip().upper() != commodity:
            continue
        lt_str = (row.get('last_trade')   or '').strip()
        fn_str = (row.get('first_notice') or '').strip()
        contract = (row.get('contract') or '').strip()
        if not lt_str:
            continue
        plen = len(prefix)
        if len(contract) < plen + 3:
            continue
        suffix_3  = contract[plen:plen + 3].upper()
        month_num = FUTURES_MONTH_FROM_SUFFIX.get(suffix_3)
        if not month_num:
            continue
        year = None
        if fn_str:
            try:
                fn_dt = datetime.strptime(fn_str, '%Y-%m-%d')
                year = fn_dt.year if fn_dt.month <= month_num else fn_dt.year + 1
            except ValueError:
                pass
        if year is None:
            try:
                lt_dt = datetime.strptime(lt_str, '%Y-%m-%d')
                year = lt_dt.year if lt_dt.month <= month_num else lt_dt.year + 1
            except ValueError:
                continue
        key = (month_num, year)
        if key not in fut_lt:
            fut_lt[key] = lt_str

    opts_by = {}
    for row in ct_opts:
        opts_by.setdefault((row['ticker'], row['date']), []).append(row)

    all_dates = sorted(set(r['date'] for r in ct_opts if r['date']))

    # Pre-compute per-ticker metadata (expiry, futures contract info)
    ticker_meta = {}
    for (ticker, _) in opts_by:
        if ticker in ticker_meta:
            continue
        p = parse_generic_ticker(ticker, prefix)
        if not p:
            continue
        _, yr, mo = p
        if mo in excl_months:
            continue
        # Expiry date
        exp_str = expiry_override.get(ticker)
        if not exp_str:
            if commodity == 'CT':
                exp_str = option_expiry_date(mo, yr)
            elif commodity == 'KC':
                exp_str = _kc_opt_expiry(mo, yr)
            else:
                std_m2 = serial_map.get(mo, mo) if mo not in std_months else mo
                exp_str = fut_lt.get((mo, yr)) or fut_lt.get((std_m2, yr))
        if not exp_str:
            continue
        std_m = mo if mo in std_months else serial_map.get(mo, mo)
        std_y = yr
        suffix = FUTURES_MONTH_SUFFIX.get(std_m)
        if not suffix:
            continue
        ticker_meta[ticker] = {
            'expiry': exp_str, 'std_m': std_m, 'std_y': std_y,
            'suffix': suffix, 'mo': mo, 'yr': yr,
        }

    SERIES_KEYS = ['atm', 'call_10', 'call_25', 'call_35', 'put_10', 'put_25', 'put_35']

    # Compute IV data for every (ticker, date) in one pass
    # ticker_date_ivs[ticker][date] = {atm, call_10, ..., put_35, _dte}
    ticker_date_ivs = {}

    for (ticker, d), rows_d in opts_by.items():
        meta = ticker_meta.get(ticker)
        if not meta:
            continue
        try:
            dte = (datetime.strptime(meta['expiry'], '%Y-%m-%d') -
                   datetime.strptime(d, '%Y-%m-%d')).days
        except ValueError:
            continue
        if dte <= 0:
            continue
        T = dte / 365.0

        try:
            d_dt = datetime.strptime(d, '%Y-%m-%d')
        except ValueError:
            continue
        first_year = d_dt.year if d_dt.month <= meta['std_m'] else d_dt.year + 1
        ordinal = meta['std_y'] - first_year + 1
        if not (1 <= ordinal <= 3):
            continue
        fwd = generic_settle.get(f"{prefix}{meta['suffix']}{ordinal}", {}).get(d)
        if not fwd or fwd <= 0:
            continue

        # Override with put-call parity implied forward, same as load_data()
        by_strike_parity = {}
        for r in rows_d:
            if r['px'] <= 0:
                continue
            k = r['strike']
            if k not in by_strike_parity:
                by_strike_parity[k] = {}
            by_strike_parity[k][r['pc']] = r['px']
        implied_Fs = []
        for k, pcs in by_strike_parity.items():
            if 'Call' in pcs and 'Put' in pcs:
                implied_Fs.append(k + (pcs['Call'] - pcs['Put']) * math.exp(RISK_FREE * T))
        if len(implied_Fs) >= 3:
            implied_Fs.sort()
            fwd = implied_Fs[len(implied_Fs) // 2]

        strikes = set(r['strike'] for r in rows_d if r['px'] > 0)
        if not strikes:
            continue
        atm_k = min(strikes, key=lambda k: abs(k - fwd))
        df = math.exp(-RISK_FREE * T)

        c_px = next((r['px'] for r in rows_d if r['pc']=='Call' and abs(r['strike']-atm_k)<0.01 and r['px']>0), None)
        p_px = next((r['px'] for r in rows_d if r['pc']=='Put'  and abs(r['strike']-atm_k)<0.01 and r['px']>0), None)
        atm_iv_val = None
        if c_px and p_px:
            call_eq = ((c_px + p_px) + (fwd - atm_k) * df) / 2.0
            if call_eq > 0:
                atm_iv_val = implied_vol(call_eq, fwd, atm_k, T, RISK_FREE, True)
        if atm_iv_val is None:
            for pc in ('Call', 'Put'):
                for r in rows_d:
                    if r['pc'] == pc and abs(r['strike'] - atm_k) < 0.01 and r['px'] > 0:
                        atm_iv_val = implied_vol(r['px'], fwd, r['strike'], T, RISK_FREE, pc == 'Call')
                        if atm_iv_val:
                            break
                if atm_iv_val:
                    break
        if atm_iv_val is None:
            continue

        call_pts, put_pts = [], []
        for r in rows_d:
            if r['px'] <= 0:
                continue
            is_call = (r['pc'] == 'Call')
            iv = implied_vol(r['px'], fwd, r['strike'], T, RISK_FREE, is_call)
            if iv is None:
                continue
            dv = b76_delta(fwd, r['strike'], T, RISK_FREE, iv, is_call)
            if is_call and 0.03 < dv < 0.70:
                call_pts.append((dv, iv))
            elif not is_call and -0.70 < dv < -0.03:
                put_pts.append((dv, iv))

        c25 = _delta_interp(call_pts, 0.25)
        p25 = _delta_interp(put_pts, -0.25)
        if c25 is None or p25 is None:
            continue

        c10 = _delta_interp(call_pts, 0.10)
        c35 = _delta_interp(call_pts, 0.35)
        p10 = _delta_interp(put_pts, -0.10)
        p35 = _delta_interp(put_pts, -0.35)

        if ticker not in ticker_date_ivs:
            ticker_date_ivs[ticker] = {}
        ticker_date_ivs[ticker][d] = {
            'atm':     round(atm_iv_val * 100, 2),
            'call_10': round(c10 * 100, 2) if c10 else None,
            'call_25': round(c25 * 100, 2),
            'call_35': round(c35 * 100, 2) if c35 else None,
            'put_10':  round(p10 * 100, 2) if p10 else None,
            'put_25':  round(p25 * 100, 2),
            'put_35':  round(p35 * 100, 2) if p35 else None,
            '_dte':    dte,
        }

    def _empty_series():
        return {k: [] for k in ['dates'] + SERIES_KEYS}

    # Rolling: nearest standard month (H/K/N/Z) with DTE >= 30 per date
    rolling = _empty_series()
    for d in all_dates:
        best = None
        best_dte = None
        for ticker, date_map in ticker_date_ivs.items():
            if ticker_meta.get(ticker, {}).get('mo') not in std_months:
                continue
            ivs = date_map.get(d)
            if not ivs or ivs['_dte'] < 30:
                continue
            if best is None or ivs['_dte'] < best_dte:
                best = ivs
                best_dte = ivs['_dte']
        if best is None:
            continue
        rolling['dates'].append(d)
        for k in SERIES_KEYS:
            rolling[k].append(best.get(k))

    # Per-ticker series (only include if >= 10 dates)
    ticker_series = {}
    for ticker, date_map in ticker_date_ivs.items():
        sorted_dates = sorted(date_map.keys())
        if len(sorted_dates) < 10:
            continue
        s = _empty_series()
        for d in sorted_dates:
            ivs = date_map[d]
            s['dates'].append(d)
            for k in SERIES_KEYS:
                s[k].append(ivs.get(k))
        ticker_series[ticker] = s

    result = {'rolling': rolling, 'tickers': ticker_series}
    _cache_entry['data'] = result
    _cache_entry['ts'] = now
    return result


# ── Routes ────────────────────────────────────────────────────────────────────

def _no_cache(response):
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

@server.route('/')
def index():
    data = load_data()
    return _no_cache(server.make_response(render_template('index.html', data=data)))

@server.route('/api/data')
def api_data():
    commodity = request.args.get('commodity', 'CT').upper()
    if commodity not in COMMODITY_CONFIG:
        return _no_cache(jsonify({'error': f'Unknown commodity: {commodity}'}))
    return _no_cache(jsonify(load_data(commodity)))

@server.route('/api/debug')
def api_debug():
    """Show exact inputs and outputs of the IV computation for every active ticker."""
    try:
        opt_rows = read_local_csv(LOCAL_OPT_HISTORY)
        oi_rows  = read_local_csv(LOCAL_FUT_HISTORY)
    except Exception as e:
        return jsonify({'error': str(e)})

    rtd = None
    if _rtd_reader:
        try:
            rtd = _rtd_reader.read_live()
        except Exception:
            pass

    ct_opts = []
    for r in opt_rows:
        if r.get('commodity', '').strip().upper() != 'CT':
            continue
        parsed = parse_security_des(r.get('security_des', ''), r.get('strike_px'))
        if not parsed:
            continue
        try:
            px  = float(r.get('px_settle', 0) or 0)
            oi  = int(r.get('open_int', 0) or 0)
        except (ValueError, TypeError):
            continue
        ct_opts.append({'date': r.get('date','').strip(), 'ticker': parsed['ticker'],
                        'pc': parsed['pc'], 'strike': parsed['strike'], 'px': px, 'oi': oi})

    ct_fut = [r for r in oi_rows if r.get('commodity','').strip().upper() == 'CT']
    all_dates = sorted(set(r['date'] for r in ct_opts if r['date']))
    if not all_dates:
        return jsonify({'error': 'no dates'})
    last_date = all_dates[-1]
    today_opts = [r for r in ct_opts if r['date'] == last_date]

    # Futures lookup
    fut_lookup = {}
    for row in ct_fut:
        contract = (row.get('contract') or '').strip()
        lt_str   = (row.get('last_trade') or '').strip()
        fn_str   = (row.get('first_notice') or '').strip()
        settle_s = (row.get('settle') or '').strip()
        date_s   = (row.get('date') or '').strip()
        if not contract or not settle_s or not date_s:
            continue
        key = None
        if lt_str:
            try:
                lt_dt = datetime.strptime(lt_str, '%Y-%m-%d')
                key = (lt_dt.month, lt_dt.year)
            except ValueError:
                pass
        if key is None:
            result = parse_futures_my(contract, date_s)
            if result:
                key = result
        if key is None:
            continue
        try:
            settle_f = float(settle_s)
        except (ValueError, TypeError):
            continue
        if key not in fut_lookup or date_s > fut_lookup[key]['date']:
            fut_lookup[key] = {'settle': settle_f, 'last_trade': lt_str or None,
                               'first_notice': fn_str or None, 'date': date_s}

    seen = {}
    for row in today_opts:
        t = row['ticker']
        if t in seen:
            continue
        p = parse_ct_ticker(t)
        if p:
            seen[t] = p

    expiry_list = sorted(
        (t for t, p in seen.items() if p[2] not in CT_EXCLUDED_MONTHS),
        key=lambda t: (seen[t][1], seen[t][2])
    )

    live_opts_dbg = (rtd.get('live_options') or {}) if rtd else {}
    sheet_names = _rtd_reader.workbook_sheet_names() if _rtd_reader else []
    out = {'csv_last_date': last_date, 'rtd_available': rtd is not None,
           'rtd_source': rtd.get('source') if rtd else None,
           'workbook_sheets': sheet_names,
           'live_options_tickers': {t: v.get('expiry') for t,v in live_opts_dbg.items()},
           'tickers': {}}

    for ticker in expiry_list:
        _, year, month_num = seen[ticker]
        key = (month_num, year)
        entry = fut_lookup.get(key)
        if entry is None and month_num not in CT_STANDARD_MONTHS:
            std_m, std_y = next_standard_month(month_num, year)
            entry = fut_lookup.get((std_m, std_y))
        if not entry:
            out['tickers'][ticker] = {'error': 'no futures entry'}
            continue

        futures_csv_settle = entry['settle']
        # Use Bloomberg's expiry if available, else our computed date
        lo_dbg = live_opts_dbg.get(ticker)
        lt = (lo_dbg['expiry'] if lo_dbg and lo_dbg.get('expiry') else None) \
             or option_expiry_date(month_num, year)

        try:
            dte = max(0, (datetime.strptime(lt, '%Y-%m-%d') -
                          datetime.strptime(last_date, '%Y-%m-%d')).days)
        except (ValueError, TypeError):
            dte = 0

        T = dte / 365.0

        # Put-call parity forward
        by_k = {}
        for row in today_opts:
            if row['ticker'] == ticker and row['px'] > 0:
                by_k.setdefault(row['strike'], {})[row['pc']] = row['px']
        implied_Fs = sorted([k + (pcs['Call'] - pcs['Put']) * math.exp(RISK_FREE * T)
                             for k, pcs in by_k.items() if 'Call' in pcs and 'Put' in pcs])
        fwd = implied_Fs[len(implied_Fs) // 2] if len(implied_Fs) >= 3 else futures_csv_settle

        # ATM strike
        strikes = set(r['strike'] for r in today_opts if r['ticker'] == ticker)
        atm = min(strikes, key=lambda k: abs(k - fwd)) if strikes else None

        # ATM option prices and IVs
        atm_call = atm_put = call_iv = put_iv = None
        for row in today_opts:
            if row['ticker'] == ticker and atm and abs(row['strike'] - atm) < 0.01:
                if row['pc'] == 'Call':
                    atm_call = row['px']
                    if T > 0 and atm_call > 0:
                        call_iv = implied_vol(atm_call, fwd, atm, T, RISK_FREE, True)
                elif row['pc'] == 'Put':
                    atm_put = row['px']
                    if T > 0 and atm_put > 0:
                        put_iv = implied_vol(atm_put, fwd, atm, T, RISK_FREE, False)

        # RTD override fields
        rtd_atm_vol = rtd_strad_value = rtd_strad_settle = rtd_atm_strike = \
            rtd_dte = rtd_live_last = rtd_live_settle = None
        if rtd:
            opt_d = (rtd.get('options') or {}).get(ticker, {})
            rtd_atm_vol      = opt_d.get('atm_vol_rtd')
            rtd_strad_value  = opt_d.get('strad_value')
            rtd_strad_settle = opt_d.get('strad_settle') or opt_d.get('strad_settle_rtd')
            rtd_atm_strike   = opt_d.get('atm_strike') or opt_d.get('current_strike') or opt_d.get('atm_strike_rtd')
            rtd_dte          = opt_d.get('dte_rtd')
            lo = (rtd.get('outrights') or {}).get(ticker, {})
            rtd_live_last    = lo.get('last')
            rtd_live_settle  = lo.get('settle')

        out['tickers'][ticker] = {
            'expiry_date_used': lt,
            'dte': dte,
            'T': round(T, 4),
            'futures_csv_settle': round(futures_csv_settle, 4),
            'parity_forward': round(fwd, 4),
            'atm_strike': atm,
            'atm_call_px': atm_call,
            'atm_put_px': atm_put,
            'call_iv_pct': round(call_iv * 100, 2) if call_iv else None,
            'put_iv_pct': round(put_iv * 100, 2) if put_iv else None,
            'num_parity_pairs': len(implied_Fs),
            'rtd_atm_vol_raw': rtd_atm_vol,
            'rtd_atm_vol_pct': round(rtd_atm_vol * 100 if rtd_atm_vol and rtd_atm_vol < 2.0 else rtd_atm_vol, 2) if rtd_atm_vol else None,
            'rtd_strad_value': rtd_strad_value,
            'rtd_strad_settle': rtd_strad_settle,
            'rtd_atm_strike': rtd_atm_strike,
            'rtd_dte': rtd_dte,
            'rtd_live_last': rtd_live_last,
            'rtd_live_settle': rtd_live_settle,
        }

    return _no_cache(jsonify(out))

@server.route('/api/debug-rtd')
def api_debug_rtd():
    """Show raw RTD data: straddles sheet + options sheet + live_smile sample."""
    rtd = None
    if _rtd_reader:
        try:
            rtd = _rtd_reader.read_live()
        except Exception as e:
            return jsonify({'error': str(e)})
    if not rtd:
        return jsonify({'error': 'RTD not available'})
    straddles = rtd.get('straddles') or {}
    options   = rtd.get('options')   or {}
    out = {
        'source': rtd.get('source'),
        'straddles': {k: v for k, v in straddles.items()},
        'options_atm': {k: {f: v[f] for f in ('atm_strike','atm_vol_rtd','strad_val_rtd','call_value','put_value') if f in v}
                        for k, v in options.items()},
    }
    return _no_cache(jsonify(out))

@server.route('/api/skew-history')
def api_skew_history():
    commodity = request.args.get('commodity', 'CT').upper()
    if commodity not in COMMODITY_CONFIG:
        return _no_cache(jsonify({'error': f'Unknown commodity: {commodity}'}))
    return _no_cache(jsonify(compute_skew_history(commodity)))

@server.route('/api/debug/sheet')
def api_debug_sheet():
    """Show raw first rows of the 'all options' sheet for diagnostics."""
    if not _rtd_reader:
        return jsonify({'error': 'rtd_reader not available'})
    result = _rtd_reader.read_sheet_sample('all options')
    return _no_cache(jsonify(result))

if __name__ == '__main__':
    server.run(debug=True, port=5050, use_reloader=True, reloader_type='stat')
