from flask import Flask, render_template, jsonify, request
import os, re, requests, csv, io, json, time, math, logging, threading
from datetime import datetime, timedelta

log = logging.getLogger(__name__)

# ICE RTD — exclusive live data source
try:
    import ice_rtd_reader as _ice_rtd_reader
except ImportError:
    _ice_rtd_reader = None

# Timeout wrapper for COM reads — a stuck Excel (cell-edit / modal) hangs the
# synchronous xlwings call indefinitely. Run in a thread pool with a hard
# deadline so a stuck workbook never freezes the Flask request handler.
import concurrent.futures as _cf
_RTD_TIMEOUT = int(os.getenv('COM_TIMEOUT_SECONDS', '8'))  # seconds; COM reads normally return in < 1s

def _read_ice_workbook_safe(wb_key):
    """Return read_ice_workbook() result or None if it times out / raises."""
    if not _ice_rtd_reader:
        return None
    try:
        with _cf.ThreadPoolExecutor(max_workers=1) as _ex:
            fut = _ex.submit(_ice_rtd_reader.read_ice_workbook, wb_key)
            return fut.result(timeout=_RTD_TIMEOUT)
    except (_cf.TimeoutError, Exception) as _e:
        log.warning('RTD COM read timed out or failed (%s) — falling back to CSV', _e)
        return None

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

# ICE Cotton No. 2 closed dates — settlement fetch must not write rows on these days
# (holiday rows carry forward prior session's settle prices with today's date, breaking
#  the T used to back-calculate settlement IV in the straddle vol-change calc)
# 2026 source: ICE Futures US Trading Holiday Calendar (uploaded by user 2026-05-23)
# 2027 calendar expected June 2026 — add entries below when received
_ICE_CT_CLOSED = {
    '2026-01-01', '2026-01-19', '2026-02-16', '2026-04-03',
    '2026-05-25', '2026-06-19', '2026-07-03', '2026-09-07',
    '2026-11-26', '2026-12-25', '2027-01-01',
    # 2027 entries to be added when ICE releases the 2027 calendar
}

def _is_ct_trading_day(date_str: str) -> bool:
    """False on weekends and ICE Cotton holidays."""
    try:
        d = datetime.strptime(date_str, '%Y-%m-%d')
    except ValueError:
        return False
    return d.weekday() < 5 and date_str not in _ICE_CT_CLOSED

_cache = {}
CACHE_TTL = 60
_ld_cache    = {}   # load_data() result cache — keyed by commodity
LD_CACHE_TTL = 30   # seconds; also invalidated when CSV files change on disk
RISK_FREE  = 0.045

LOCAL_OPT_HISTORY = os.path.join(os.path.dirname(__file__), 'local_options_history.csv')
LOCAL_FUT_HISTORY = os.path.join(os.path.dirname(__file__), 'local_futures_history.csv')
LOCAL_SPR_HISTORY = os.path.join(os.path.dirname(__file__), 'local_futures_spreads_history.csv')

# Bloomberg shadow backup — GitHub/Bloomberg data written here as failsafe only.
# Nothing in the pipeline reads these files. Manual reference if settle_watcher fails.
_BBG_BACKUP_DIR = os.path.join(os.path.dirname(__file__), 'bloomberg_backup')
_BBG_OPT_BACKUP = os.path.join(_BBG_BACKUP_DIR, 'ct_options_bloomberg.csv')
_BBG_FUT_BACKUP = os.path.join(_BBG_BACKUP_DIR, 'ct_futures_bloomberg.csv')
LOCAL_KC_OPT_HISTORY = os.path.join(os.path.dirname(__file__), 'local_kc_options_history.csv')
LOCAL_KC_FUT_HISTORY = os.path.join(os.path.dirname(__file__), 'local_kc_futures_history.csv')
LOCAL_KC_SPR_HISTORY = os.path.join(os.path.dirname(__file__), 'local_kc_futures_spreads_history.csv')
LOCAL_SB_OPT_HISTORY = os.path.join(os.path.dirname(__file__), 'local_sb_options_history.csv')
LOCAL_SB_FUT_HISTORY = os.path.join(os.path.dirname(__file__), 'local_sb_futures_history.csv')
LOCAL_CC_OPT_HISTORY = os.path.join(os.path.dirname(__file__), 'local_cc_options_history.csv')
LOCAL_CC_FUT_HISTORY = os.path.join(os.path.dirname(__file__), 'local_cc_futures_history.csv')

_skew_hist_cache = {}   # keyed by commodity

# ICE Cotton No. 2 option last-trading-day dates — sourced from ICE rulebook and exchange calendar.
# These are the authoritative expiry dates; the ICE RTD workbook may add new contracts at runtime.
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

# ICE Coffee C option last-trading-day dates — source: ice.com/products/14/Coffee-C-Options/expiry
ICE_KC_EXPIRY = {
    'KCN6': '2026-06-12', 'KCQ6': '2026-07-10', 'KCU6': '2026-08-14',
    'KCZ6': '2026-11-12',
    'KCH7': '2027-02-10', 'KCK7': '2027-04-09', 'KCN7': '2027-06-11',
    'KCU7': '2027-08-13', 'KCZ7': '2027-11-12',
    'KCH8': '2028-02-11',
}

# ICE Sugar No. 11 option last-trading-day dates — source: ice.com/products/22/Sugar-No-11-Options/expiry
ICE_SB_EXPIRY = {
    'SBN6': '2026-06-15', 'SBQ6': '2026-07-15', 'SBU6': '2026-08-17',
    'SBV6': '2026-09-15',
    'SBF7': '2026-12-15', 'SBH7': '2027-02-16', 'SBK7': '2027-04-15',
    'SBN7': '2027-06-15', 'SBV7': '2027-09-15',
    'SBF8': '2027-12-15', 'SBH8': '2028-02-15', 'SBK8': '2028-04-17',
    'SBN8': '2028-06-15', 'SBV8': '2028-09-15',
}

# ICE Cocoa option last-trading-day dates — source: ice.com/products/8/Cocoa-Options/expiry
ICE_CC_EXPIRY = {
    'CCN6': '2026-06-12', 'CCQ6': '2026-07-10', 'CCU6': '2026-08-14',
    'CCZ6': '2026-11-13',
    'CCH7': '2027-02-12', 'CCK7': '2027-04-09', 'CCN7': '2027-06-11',
    'CCU7': '2027-08-13', 'CCZ7': '2027-11-12',
}


COMMODITY_CONFIG = {
    'CT': {
        'prefix': 'CT', 'name': 'ICE Cotton No. 2',
        'opt_csv': LOCAL_OPT_HISTORY, 'fut_csv': LOCAL_FUT_HISTORY,
        'std_months': frozenset({3, 5, 7, 12}), 'excl_months': frozenset({10}),
        'serial_map': {1:3, 2:3, 9:12, 11:12},
        'expiry_override': ICE_CT_EXPIRY,
    },
    'KC': {
        'prefix': 'KC', 'name': 'ICE Coffee C',
        'opt_csv': LOCAL_KC_OPT_HISTORY, 'fut_csv': LOCAL_KC_FUT_HISTORY,
        'spr_csv': LOCAL_KC_SPR_HISTORY,
        'std_months': frozenset({3, 5, 7, 9, 12}), 'excl_months': frozenset(),
        'serial_map': {8: 9},
        'expiry_override': ICE_KC_EXPIRY,
        # Straddle DISPLAY filter (display + EOD only — backend data unchanged).
        # When the RTD workbook is open, the straddle table shows exactly the
        # option tabs present in it (live_options keys). When RTD is offline,
        # falls back to this pinned list so the table never blanks. Add a tab in
        # Excel → it shows live with no restart; to change the OFFLINE fallback,
        # edit this set and restart. Serial months (Q/V etc.) and sparse far
        # months are intentionally excluded by not opening tabs for them.
        'straddle_tickers': frozenset({'KCN6', 'KCU6', 'KCZ6', 'KCH7', 'KCK7'}),
        # KC options trade on a 2.5-cent strike grid (verified: ICE OMON, ICE
        # platform, RTD sheet). Force any computed straddle ATM onto this grid so
        # an integer fallback can never display an off-grid strike (e.g. 243).
        'strike_increment': 2.5,
    },
    'SB': {
        'prefix': 'SB', 'name': 'ICE Sugar No. 11',
        'opt_csv': LOCAL_SB_OPT_HISTORY, 'fut_csv': LOCAL_SB_FUT_HISTORY,
        'std_months': frozenset({3, 5, 7, 10}), 'excl_months': frozenset(),
        'serial_map': {1:3, 8:10, 9:10},
        'expiry_override': ICE_SB_EXPIRY,
    },
    'CC': {
        'prefix': 'CC', 'name': 'ICE Cocoa',
        'opt_csv': LOCAL_CC_OPT_HISTORY, 'fut_csv': LOCAL_CC_FUT_HISTORY,
        'std_months': frozenset({3, 5, 7, 9, 12}), 'excl_months': frozenset(),
        'serial_map': {8: 9},
        'expiry_override': ICE_CC_EXPIRY,
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

def _decade_year(year_digit):
    """Convert single-digit year to full year, safe through 2039."""
    now_yr = datetime.now().year
    decade = (now_yr // 10) * 10
    year = decade + year_digit
    if year < now_yr - 5:
        year += 10
    return year

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
    return code, _decade_year(year_digit), MONTH_CODE[code]

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


def _generic_to_ice_code(contract, date_str):
    """'CTJUL1', '2026-06-02' -> 'CTN6'  (old ordinal format -> ICE RTD code)."""
    parsed = parse_futures_my(contract, date_str)
    if not parsed:
        return None
    month_num, year = parsed
    _inv = {v: k for k, v in MONTH_CODE.items()}
    mc = _inv.get(month_num)
    if not mc:
        return None
    return f"{contract[:2]}{mc}{str(year)[-1]}"


def _ordinal_to_ice_code(contract, date_str, prefix):
    """Prefix-aware ordinal->ICE: 'KCMAY1','2026-06-09','KC' -> 'KCK6'.
    _generic_to_ice_code is CT-hardcoded (parse_futures_my requires 'CT'); this
    works for any prefix via parse_futures_my_generic."""
    parsed = parse_futures_my_generic(contract, date_str, prefix)
    if not parsed:
        return None
    month_num, year = parsed
    _inv = {v: k for k, v in MONTH_CODE.items()}
    mc = _inv.get(month_num)
    if not mc:
        return None
    return f"{prefix}{mc}{str(year)[-1]}"


# CT serial option months and their underlying standard futures month.
# Cotton only has 4 serial option months — no Apr, Jun, or Aug serials exist.
CT_SERIAL_FUTURES = {
    1:  3,  # Jan -> Mar
    2:  3,  # Feb -> Mar
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
    return code, _decade_year(year_digit), MONTH_CODE[code]


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


def _nth_friday_of_month(month, year, n=2):
    """Return the nth Friday of the given month/year."""
    d = datetime(year, month, 1)
    days_to_friday = (4 - d.weekday()) % 7
    return d + timedelta(days=days_to_friday + (n - 1) * 7)

def _kc_opt_expiry(month_num, year):
    """KC/CC options expire on 2nd Friday of month preceding delivery month."""
    prev_month = month_num - 1 or 12
    prev_year  = year if month_num > 1 else year - 1
    return _nth_friday_of_month(prev_month, prev_year, 2).strftime('%Y-%m-%d')

def _cc_opt_expiry(month_num, year):
    """CC options expire on 2nd Friday of month preceding delivery month."""
    return _kc_opt_expiry(month_num, year)

def _sb_opt_expiry(month_num, year):
    """SB options expire on 15th of month preceding delivery month, next BD if weekend."""
    prev_month = month_num - 1 or 12
    prev_year  = year if month_num > 1 else year - 1
    d = datetime(prev_year, prev_month, 15)
    while d.weekday() >= 5:
        d += timedelta(days=1)
    return d.strftime('%Y-%m-%d')


# ── ICE RTD adapter ───────────────────────────────────────────────────────────

def _ice_to_rtd_shape(ice_data):
    """
    Convert read_ice_workbook() output to the shape rtd_reader.read_live() returns
    so the rest of load_data() is unchanged.
    Returns None when ICE workbook is unavailable.
    """
    if not ice_data or ice_data.get('mode') == 'unavailable':
        return None

    mode        = ice_data['mode']
    futures_raw = ice_data.get('futures', {})
    options_raw = ice_data.get('options', {})

    outrights = {}
    for tkr, f in futures_raw.items():
        bid, offer = f.get('bid'), f.get('offer')
        if mode == 'live' and bid and offer:
            last = (bid + offer) / 2.0
        else:
            last = f.get('last') or f.get('settle')
        outrights[tkr] = {
            'last':        last,
            'settle':      f.get('settle'),
            'yest_settle': f.get('yest_settle'),
            'change':      f.get('change'),
            'pct_chg':     f.get('pct_chg'),
            'oi':          f.get('oi'),
            'oi_chg':      None,
            'volume':      f.get('volume'),
            'high':        f.get('high'),
            'low':         f.get('low'),
            'block_vol':   f.get('block_vol'),
            'efs_vol':     f.get('efs_vol'),
            'efp_vol':     f.get('efp_vol'),
            'hv10':        None, 'hv30': None, 'hv60': None, 'hv90': None,
        }

    live_opts = {}
    for sheet_name, rows in options_raw.items():
        ticker  = sheet_name.upper()
        strikes = []
        for r in rows:
            k = r['strike']
            for pc, bk, ok, lk, sk in [
                ('Call', 'call_bid', 'call_offer', 'call_last', 'call_settle'),
                ('Put',  'put_bid',  'put_offer',  'put_last',  'put_settle'),
            ]:
                bid    = r.get(bk)
                ask    = r.get(ok)
                last   = r.get(lk)
                settle = r.get(sk)
                mid  = (bid + ask) / 2.0 if (bid and ask and bid > 0 and ask > 0) else last
                strikes.append({'strike': k, 'pc': pc,
                                'bid': bid, 'ask': ask, 'mid': mid,
                                'last': last, 'settle': settle, 'vol': None})
        if strikes:
            live_opts[ticker] = {'expiry': None, 'strikes': strikes}

    return {
        'outrights':    outrights,
        'spreads':      ice_data.get('spreads', {}),
        'live_options': live_opts,
        'source':       f'ice_rtd_{mode}',
    }


# ── Core data load ────────────────────────────────────────────────────────────

def _in_ct_settle_window():
    """True during 14:25–16:00 ET on a CT trading day.
    settle_watcher.py owns the ICE COM interface exclusively in this window.
    The dashboard must not call read_ice_workbook() during this period."""
    try:
        try:
            from zoneinfo import ZoneInfo as _ZI
        except ImportError:
            from backports.zoneinfo import ZoneInfo as _ZI
        _now = datetime.now(_ZI('America/New_York'))
        return (_is_ct_trading_day(_now.strftime('%Y-%m-%d'))
                and (14, 25) <= (_now.hour, _now.minute) <= (16, 0))
    except Exception:
        return False


def load_data(commodity='CT'):
    if commodity != 'CT':
        return _load_generic_data(commodity)

    # ── Result cache — invalidated on CSV file changes or after LD_CACHE_TTL ──
    try:
        _om = os.path.getmtime(LOCAL_OPT_HISTORY)
        _fm = os.path.getmtime(LOCAL_FUT_HISTORY)
    except OSError:
        _om = _fm = 0
    _now = time.time()
    _cached = _ld_cache.get('CT')
    if (_cached
            and _now - _cached['ts'] < LD_CACHE_TTL
            and _cached.get('om') == _om
            and _cached.get('fm') == _fm):
        return _cached['data']

    try:
        opt_rows = read_local_csv(LOCAL_OPT_HISTORY)
        oi_rows  = read_local_csv(LOCAL_FUT_HISTORY)
    except Exception as e:
        return {'error': str(e)}

    # Live data — ICE RTD exclusive source.
    # Blocked during 14:25–16:00 ET: settle_watcher owns the COM interface then.
    rtd = None
    if _ice_rtd_reader and not _in_ct_settle_window():
        try:
            rtd = _ice_to_rtd_shape(_read_ice_workbook_safe('CT'))
        except Exception as e:
            log.debug('ICE RTD fetch skipped: %s', e)

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
            oi  = int(float(r.get('open_int', 0) or 0))
            oic = int(float(r.get('oi_chg', 0) or 0))
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

    # flow_rtd.json — written at futures settlement time (before options settle),
    # holds yesterday's ICE call_settle/put_settle for every contract/strike.
    # Used as prev_c/prev_p when post-settlement (RTD has flipped to today's values).
    # Try today's date first: today's file has yesterday's settle prices (written at
    # futures settlement before options settled). Fall back to last_date if not found.
    _flow_rtd_opts = {}
    try:
        import json as _json
        _today_for_rtd = datetime.now().strftime('%Y-%m-%d')
        _flow_rtd_path = None
        for _rtd_date in (_today_for_rtd, last_date):
            _candidate = os.path.normpath(os.path.join(
                os.path.dirname(__file__), '..', 'Options_flow_analyzer',
                'data', _rtd_date, 'flow_rtd.json'
            ))
            if os.path.exists(_candidate):
                _flow_rtd_path = _candidate
                break
        if _flow_rtd_path:
            _frd = _json.load(open(_flow_rtd_path, encoding='utf-8'))
            for _tkr, _cdata in (_frd.get('contracts') or {}).items():
                for _row in (_cdata.get('options') or []):
                    _k = int(float(_row.get('strike', 0)))
                    _cs = _row.get('call_settle')
                    _ps = _row.get('put_settle')
                    if _cs and _ps and _cs > 0 and _ps > 0:
                        _flow_rtd_opts[(_tkr, _k)] = (_cs, _ps)
            log.debug('flow_rtd.json loaded: %d option rows', len(_flow_rtd_opts))
    except Exception as _e:
        log.debug('flow_rtd.json load skipped: %s', _e)

    # CSV option settle lookup — today's settled prices from local_options_history.csv.
    # Used when RTD is offline post-settlement so straddle values don't go blank.
    # Format: {(ticker, strike_int): {'C': call_settle, 'P': put_settle}}
    # Two security_des formats exist:
    #   settle_watcher: 'CTN6P    75' (embedded strike) + separate strike_px column
    #   _persist_ct_options_ice: 'CTN6P' (short) + separate strike_px column
    # Always use strike_px as primary; fall back to embedded parsing when absent.
    _csv_opt_settle = {}
    try:
        for _r in read_local_csv(LOCAL_OPT_HISTORY):
            if _r.get('date', '').strip() != last_date:
                continue
            _sd = _r.get('security_des', '').strip()
            _px = _r.get('px_settle', '').strip()
            if not _sd or not _px or len(_sd) < 5 or _sd[4] not in ('C', 'P'):
                continue
            try:
                _px_f = float(_px)
                if _px_f <= 0:
                    continue
                _t = _sd[:4]
                _pc = _sd[4]
                _strike_raw = _r.get('strike_px', '').strip()
                if _strike_raw:
                    _k = int(float(_strike_raw))
                elif len(_sd) >= 6:
                    _k = int(float(_sd[5:].strip()))
                else:
                    continue
                _csv_opt_settle.setdefault((_t, _k), {})[_pc] = _px_f
            except (ValueError, TypeError):
                continue
    except Exception as _e:
        log.debug('csv_opt_settle load skipped: %s', _e)

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

    # Previous-trading-day settlement from CSV — needed to fix yest_settle in live_futures.
    # RTD 'Change' column is last-trade based (last_trade − S_{t-1}), not settle-to-settle,
    # so yest_settle = rtd_settle − rtd_change gives the wrong value.
    csv_prev_settle = {}
    for row in ct_fut:
        if (row.get('date') or '').strip() != prev_date:
            continue
        contract = (row.get('contract') or '').strip()
        settle_s = (row.get('settle')   or '').strip()
        date_s   = (row.get('date')     or '').strip()
        if not contract or not settle_s:
            continue
        lt_str = (row.get('last_trade') or '').strip()
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
        csv_prev_settle[key] = settle_f

    # Futures-CSV-based prior settle — uses fut_prev (futures CSV second-most-recent date).
    # Needed for post_settle detection which must use futures CSV date, not options CSV date.
    # Built after fut_dates/fut_prev are computed further below; populated in a second pass.
    csv_fut_prev_settle = {}   # populated after _build_csv_oi section below

    # Yesterday's OI from CSV — used to compute OI chg against RTD live OI.
    # The CSV oi_chg field is always blank; compute it as rtd_oi - csv_prev_oi instead.
    # Use same date logic as yest_settle: prev_date when post-settle, last_date otherwise.
    def _build_csv_oi(target_date):
        result = {}
        for row in ct_fut:
            if (row.get('date') or '').strip() != target_date:
                continue
            contract = (row.get('contract') or '').strip()
            oi_s     = (row.get('open_int') or '').strip()
            date_s   = (row.get('date')     or '').strip()
            if not contract or not oi_s:
                continue
            lt_str = (row.get('last_trade') or '').strip()
            key = None
            if lt_str:
                try:
                    lt_dt = datetime.strptime(lt_str, '%Y-%m-%d')
                    key = (lt_dt.month, lt_dt.year)
                except ValueError:
                    pass
            if key is None:
                r2 = parse_futures_my(contract, date_s)
                if r2:
                    key = r2
            if key is None:
                continue
            try:
                result[key] = float(oi_s)
            except (ValueError, TypeError):
                continue
        return result
    # Futures CSV has its own date sequence — may lag behind options CSV.
    # Derive dates from ct_fut directly so we always find OI rows.
    fut_dates   = sorted(set(r.get('date','').strip() for r in ct_fut if r.get('date','').strip()))
    fut_last    = fut_dates[-1]  if fut_dates            else last_date
    fut_prev    = fut_dates[-2]  if len(fut_dates) >= 2  else fut_last
    csv_last_oi = _build_csv_oi(fut_last)
    csv_prev_oi = _build_csv_oi(fut_prev)

    # Populate csv_fut_prev_settle now that fut_prev is known.
    # Uses futures CSV second-most-recent date so post_settle yest_settle
    # always resolves to the actual prior trading session, independent of
    # whether the options CSV has caught up yet.
    for row in ct_fut:
        if (row.get('date') or '').strip() != fut_prev:
            continue
        contract = (row.get('contract') or '').strip()
        settle_s = (row.get('settle')   or '').strip()
        date_s   = (row.get('date')     or '').strip()
        if not contract or not settle_s:
            continue
        lt_str = (row.get('last_trade') or '').strip()
        key = None
        if lt_str:
            try:
                lt_dt = datetime.strptime(lt_str, '%Y-%m-%d')
                key = (lt_dt.month, lt_dt.year)
            except ValueError:
                pass
        if key is None:
            r2 = parse_futures_my(contract, date_s)
            if r2:
                key = r2
        if key is None:
            continue
        try:
            csv_fut_prev_settle[key] = float(settle_s)
        except (ValueError, TypeError):
            continue

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

    # ── ATM IV today / week ───────────────────────────────────────────────────
    # 1D change is computed in the straddle loop (live IV − settlement IV via B76).
    atm_iv        = {}
    atm_iv_1d_chg = {}
    atm_iv_1w_chg = {}

    for ticker in expiry_list:
        iv_t = atm_iv_for_date(ticker, last_date)
        if iv_t is None:
            continue
        atm_iv[ticker] = round(iv_t * 100, 2)
        iv_w = atm_iv_for_date(ticker, week_date, week_futures.get(ticker))
        if iv_w is not None:
            atm_iv_1w_chg[ticker] = round((iv_t - iv_w) * 100, 2)

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
    today_str    = datetime.now().strftime('%Y-%m-%d')
    post_settle  = (fut_last == today_str)

    if rtd:
        data_source = rtd.get('source', 'csv_only')
        for tkr, d in (rtd.get('outrights') or {}).items():
            hv_data[tkr]      = {k: d.get(k) for k in ('hv10', 'hv30', 'hv60', 'hv90')}
            live_futures[tkr] = {k: d.get(k) for k in (
                'last', 'settle', 'yest_settle', 'change', 'pct_chg',
                'oi', 'oi_chg', 'volume', 'high', 'low',
                'block_vol', 'efs_vol', 'efp_vol',
            )}
        rtd_spreads = {key: d for key, d in (rtd.get('spreads') or {}).items()}

        # Fix yest_settle for outrights and spreads.
        #
        # Pre-settlement: RTD 'settle' col is static = yesterday's settlement.
        #   Use it directly — always correct before ICE publishes today's.
        #
        # Post-settlement: futures CSV (written by settle_watcher) has today's
        #   row with yest_settle stored correctly from the 14:18 snapshot.
        #   Read it directly — no derived computation needed.

        # Build yest_settle lookup from today's futures CSV rows when available
        csv_today_yest = {}
        if post_settle:
            for row in ct_fut:
                if (row.get('date') or '').strip() != fut_last:
                    continue
                c = (row.get('contract') or '').strip()
                ys = (row.get('yest_settle') or '').strip()
                if not c or not ys:
                    continue
                p = parse_ct_ticker(c)
                if p:
                    try:
                        csv_today_yest[(p[2], p[1])] = float(ys)
                    except (ValueError, TypeError):
                        pass

        for tkr, d in live_futures.items():
            parsed = parse_ct_ticker(tkr)
            if not parsed:
                continue
            _, yr, mn = parsed
            if post_settle:
                yest = csv_today_yest.get((mn, yr))
            else:
                # RTD settle = yesterday's published settlement (static all day)
                yest = d.get('settle')
            if yest is not None:
                d['yest_settle'] = round(yest, 4)
                chg_base = d.get('last') if d.get('last') is not None else d.get('settle')
                if chg_base is not None:
                    chg = round(chg_base - yest, 4)
                    d['change'] = chg
                    d['pct_chg'] = round(chg / yest * 100, 4) if yest else None
            # OI chg = RTD live OI − yesterday's CSV OI
            yest_oi_src = csv_prev_oi if post_settle else csv_last_oi
            yest_oi = yest_oi_src.get((mn, yr))
            live_oi = d.get('oi')
            if live_oi is not None and yest_oi is not None:
                d['oi_chg'] = round(live_oi - yest_oi)

        for skey, sd in rtd_spreads.items():
            parts = skey.split('/')
            if len(parts) != 2:
                continue
            pn = parse_ct_ticker(parts[0])
            pf = parse_ct_ticker(parts[1])
            if not pn or not pf:
                continue
            if post_settle:
                yn = csv_prev_settle.get((pn[2], pn[1]))
                yf = csv_prev_settle.get((pf[2], pf[1]))
            else:
                yn = (live_futures.get(parts[0]) or {}).get('settle')
                yf = (live_futures.get(parts[1]) or {}).get('settle')
            if yn is not None and yf is not None and sd.get('settle') is not None:
                yest_spr = round(yn - yf, 4)
                sd['yest_settle'] = yest_spr
                chg_base_spr = sd.get('last') if sd.get('last') is not None else sd.get('settle')
                if chg_base_spr is not None:
                    chg_spr = round(chg_base_spr - yest_spr, 4)
                    sd['change'] = chg_spr
                    sd['pct_chg'] = round(chg_spr / yest_spr * 100, 4) if yest_spr else None

    # ── CSV fallback for futures fields when RTD is offline post-settlement ──────
    # settle_watcher writes settle, high, low, volume, efp, efs, block, OI, OI_chg
    # to local_futures_history.csv at ~14:45 ET. When RTD is unavailable after
    # close, fill any None/missing live_futures fields from that CSV row so the
    # EOD email always has complete data regardless of workbook state.
    _today_str_fb = datetime.now().strftime('%Y-%m-%d')
    if fut_last == _today_str_fb:
        def _fv(v):
            try: return float(v) if v not in (None, '') else None
            except (ValueError, TypeError): return None
        for _row in ct_fut:
            if (_row.get('date') or '').strip() != fut_last:
                continue
            _tkr_raw = (_row.get('contract') or '').strip()
            if not _tkr_raw:
                continue
            # Translate old-format key (CTJUL1) to ICE code (CTN6) so live_futures
            # is always keyed consistently regardless of whether RTD was available.
            _tkr = _generic_to_ice_code(_tkr_raw, fut_last) or _tkr_raw
            if _tkr not in live_futures:
                live_futures[_tkr] = {}
            _lf = live_futures[_tkr]
            for _csv_key, _lf_key in [
                ('settle',    'settle'),
                ('yest_settle','yest_settle'),
                ('high',      'high'),
                ('low',       'low'),
                ('volume',    'volume'),
                ('efp_vol',   'efp_vol'),
                ('efs_vol',   'efs_vol'),
                ('block_vol', 'block_vol'),
                ('open_int',  'oi'),
                ('oi_chg',    'oi_chg'),
            ]:
                if _lf.get(_lf_key) is None:
                    _lf[_lf_key] = _fv(_row.get(_csv_key))
            # compute change/pct_chg if RTD didn't supply them
            _s, _y = _lf.get('settle'), _lf.get('yest_settle')
            if _lf.get('change') is None and _s is not None and _y is not None and _y:
                _lf['change']  = round(_s - _y, 4)
                _lf['pct_chg'] = round((_s - _y) / _y * 100, 4)

    # ── CSV fallback for spreads H/L/V when RTD is offline or read_spreads misses rows ──
    # settle_watcher writes spreads to local_futures_spreads_history.csv at ~14:31 ET.
    # RTD spreads are empty whenever the workbook is unavailable or product-name check
    # fails. Load CSV and fill rtd_spreads so the frontend has H/L/V regardless.
    if fut_last == datetime.now().strftime('%Y-%m-%d'):
        _spr_by_key_fb = {}
        try:
            with open(LOCAL_SPR_HISTORY, encoding='utf-8') as _ssf:
                for _ssr in csv.DictReader(_ssf):
                    _sk = (_ssr.get('contract') or '').strip()
                    _sd = (_ssr.get('date') or '').strip()
                    if _sk and _sd and (_sk not in _spr_by_key_fb or _sd > _spr_by_key_fb[_sk]['date']):
                        _spr_by_key_fb[_sk] = _ssr
        except Exception:
            pass
        def _sfv(v):
            try: return float(v) if v not in (None, '') else None
            except (ValueError, TypeError): return None
        for _sk, _scr in _spr_by_key_fb.items():
            _sh = _sfv(_scr.get('high')); _sl = _sfv(_scr.get('low')); _svol = _sfv(_scr.get('volume'))
            if _sk not in rtd_spreads:
                _sparts = _sk.split('/')
                if len(_sparts) == 2:
                    _spn = parse_ct_ticker(_sparts[0]); _spf = parse_ct_ticker(_sparts[1])
                    _sdisp = (f"{MONTH_NAME[_spn[2]]}{str(_spn[1])[-2:]}/{MONTH_NAME[_spf[2]]}{str(_spf[1])[-2:]}"
                              if _spn and _spf else _sk)
                    _ss_stt = _sfv(_scr.get('settle')); _ss_ys = _sfv(_scr.get('yest_settle'))
                    _ss_chg = _sfv(_scr.get('change'))
                    _ss_pct = round(_ss_chg / _ss_ys * 100, 2) if (_ss_chg is not None and _ss_ys and _ss_ys != 0) else None
                    rtd_spreads[_sk] = {
                        'display': _sdisp, 'settle': _ss_stt, 'yest_settle': _ss_ys,
                        'change': _ss_chg, 'pct_chg': _ss_pct,
                        'high': _sh, 'low': _sl, 'volume': _svol,
                        'block_vol': _sfv(_scr.get('block_vol')),
                        'efs_vol': _sfv(_scr.get('efs_vol')),
                        'efp_vol': _sfv(_scr.get('efp_vol')),
                    }
            else:
                for _sfk, _scv in [('high', _sh), ('low', _sl), ('volume', _svol)]:
                    if rtd_spreads[_sk].get(_sfk) is None:
                        rtd_spreads[_sk][_sfk] = _scv

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
            dte = get_dte(ticker, datetime.now().strftime('%Y-%m-%d'))
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
                        iv_w = atm_iv_for_date(ticker, week_date, week_futures.get(ticker))
                        if iv_w is not None:
                            atm_iv_1w_chg[ticker] = round(live_iv * 100 - iv_w * 100, 2)

    # ── HV from local futures history (replaces Bloomberg RTD HV) ────────────
    csv_hv = _compute_hv(LOCAL_FUT_HISTORY, 'CT')
    for tkr in expiry_list:
        if tkr in csv_hv:
            hv_data.setdefault(tkr, {}).update(csv_hv[tkr])

    # ── Straddle run (Jul–Dec 2026 only) ─────────────────────────────────────
    # Read settle prices directly from ICE RTD workbook when available.
    # Blocked during 14:25–16:00 ET: settle_watcher owns the COM interface then.
    # Straddle loop falls back to _flow_rtd_opts / B76 when _ice_raw is None.
    _ice_raw = None
    if _ice_rtd_reader and not _in_ct_settle_window():
        try:
            _ice_raw = _read_ice_workbook_safe('CT')
            if _ice_raw and _ice_raw.get('mode') == 'unavailable':
                _ice_raw = None
        except Exception:
            _ice_raw = None

    # Snapshot freeze — use the 14:16 RTD capture for straddle/EOD computations.
    # Avoids all mode-transition issues between futures and options settlement.
    # Freeze lifts once options_settled=true in settle_status.json.
    # Does not affect any other dashboard function (live futures, vol smile, skew, etc.).
    _snap_path = os.path.normpath(os.path.join(
        os.path.dirname(__file__), '..', 'Options_flow_analyzer',
        'data', today_str, 'rtd_snap.json'
    ))
    _opts_settled_now = False
    try:
        _ss_path = os.path.join(os.path.dirname(__file__), 'settle_status.json')
        with open(_ss_path, encoding='utf-8') as _ssf:
            _ss_data = _json.load(_ssf)
        _opts_settled_now = bool(_ss_data.get('date') == today_str and _ss_data.get('options_settled'))
    except Exception:
        pass
    # Dashboard-side backup: if live RTD read succeeded and it's >= 14:16, save snapshot
    # so the freeze works even if settle_watcher missed the 14:16 window.
    if _ice_raw and not _opts_settled_now and not os.path.exists(_snap_path):
        try:
            try:
                from zoneinfo import ZoneInfo as _ZI_snap
            except ImportError:
                from backports.zoneinfo import ZoneInfo as _ZI_snap
            _snap_now = datetime.now(_ZI_snap('America/New_York'))
            if (_snap_now.hour, _snap_now.minute) >= (14, 16):
                os.makedirs(os.path.dirname(_snap_path), exist_ok=True)
                with open(_snap_path, 'w', encoding='utf-8') as _sf:
                    _json.dump(_ice_raw, _sf)
        except Exception:
            pass
    if not _opts_settled_now and os.path.exists(_snap_path):
        try:
            with open(_snap_path, encoding='utf-8') as _f14:
                _ice_raw = _json.load(_f14)
        except Exception:
            pass  # fall through to live RTD result or None

    # Supplement straddle list with any contracts in the live workbook not yet in the CSV
    straddle_tickers = list(expiry_list)
    if _ice_raw:
        for sheet_name in (_ice_raw.get('options') or {}):
            t = sheet_name.upper()
            if t in straddle_tickers:
                continue
            p = parse_ct_ticker(t)
            if not p or p[2] in CT_EXCLUDED_MONTHS:
                continue
            straddle_tickers.append(t)
            expiry_labels[t] = ticker_label(t)
            if t in ICE_CT_EXPIRY:
                last_trade[t] = ICE_CT_EXPIRY[t]
            else:
                _, yr, mo = p
                last_trade[t] = option_expiry_date(mo, yr)
            ice_f = (_ice_raw.get('futures') or {}).get(t, {})
            if ice_f.get('settle'):
                futures[t] = ice_f['settle']

    # Sep27 and Nov27 excluded until liquidity warrants inclusion
    straddle_tickers = [t for t in straddle_tickers if t not in {'CTU7', 'CTX7'}]

    # Re-sort chronologically by (year, month). expiry_list is already sorted, but
    # contracts supplemented from the live workbook above (not yet in the CSV) get
    # appended at the end — without this re-sort a newly added contract (e.g. CTX6)
    # displays out of sequence at the bottom of the straddle tab and EOD run.
    def _strad_sort_key(t):
        p = parse_ct_ticker(t)
        return (p[1], p[2]) if p else (9999, 99)
    straddle_tickers.sort(key=_strad_sort_key)

    # Price tape fallback: most-recent live mid per contract from ct_price_tape.csv.
    # Used when _ice_raw is None (COM contention) so straddle ATM reflects live prices.
    _tape_live = {}
    try:
        _tape_path = os.path.normpath(os.path.join(
            os.path.dirname(__file__), '..', 'Options_flow_analyzer',
            'data', today_str, 'ct_price_tape.csv'
        ))
        if os.path.exists(_tape_path):
            import csv as _csv_tape
            with open(_tape_path, newline='', encoding='utf-8') as _tf:
                for _tr in reversed(list(_csv_tape.DictReader(_tf))):
                    _tc = _tr.get('contract', '')
                    if _tc and _tc not in _tape_live:
                        _tmid = _tr.get('mid') or _tr.get('last')
                        if _tmid:
                            try: _tape_live[_tc] = float(_tmid)
                            except ValueError: pass
    except Exception:
        pass

    # The GitHub options CSV is published one business day late — last_date is the
    # publication date, not the trading date.  Settlement was actually priced on the
    # last CT trading session before today.  Use that date for T_settle so the
    # back-computed settle IV uses the correct DTE.
    _prev_day = datetime.now() - timedelta(days=1)
    while not _is_ct_trading_day(_prev_day.strftime('%Y-%m-%d')):
        _prev_day -= timedelta(days=1)
    _actual_settle_str = _prev_day.strftime('%Y-%m-%d')

    straddles = []
    for ticker in straddle_tickers:
        my = _contract_code_to_month_year(ticker, 'CT')
        if not my:
            continue
        lt    = last_trade.get(ticker)
        label = expiry_labels.get(ticker)
        if not lt:
            continue

        # ── Forward from ICE RTD workbook (exclusive source) ─────────────
        ice_chain   = (_ice_raw or {}).get('options', {}).get(ticker, [])
        ice_fut_row = (_ice_raw or {}).get('futures', {}).get(ticker, {})

        # Live forward: bid/offer mid → last → price tape → settle
        _fb, _fo = ice_fut_row.get('bid'), ice_fut_row.get('offer')
        if _fb and _fo and _fb > 0 and _fo > 0:
            fwd = (_fb + _fo) / 2.0
        elif ice_fut_row.get('last') and ice_fut_row['last'] > 0:
            fwd = ice_fut_row['last']
        elif _tape_live.get(ticker):
            fwd = _tape_live[ticker]
        else:
            fwd = ice_fut_row.get('settle') or futures.get(ticker)

        atm_row = None
        if fwd and ice_chain:
            atm = min((r['strike'] for r in ice_chain), key=lambda k: abs(k - fwd))
            atm_row = next((r for r in ice_chain if r['strike'] == atm), None)
            if atm_row:
                _rtd_mode = (_ice_raw or {}).get('mode', 'live')
                # Always try live bid/ask first — options trade until ~14:50 ET regardless
                # of futures settlement. prior_settle mode only means futures settled, not
                # that options bid/ask is stale. Only fall back to call_settle when absent.
                _val_from_live_bid_ask = False
                cb, co = atm_row.get('call_bid'), atm_row.get('call_offer')
                pb, po = atm_row.get('put_bid'),  atm_row.get('put_offer')
                if cb and co and cb > 0 and co > 0 and pb and po and pb > 0 and po > 0:
                    today_c = (cb + co) / 2.0
                    today_p = (pb + po) / 2.0
                    _val_from_live_bid_ask = True
                else:
                    today_c = atm_row.get('call_last')
                    today_p = atm_row.get('put_last')
                    if _rtd_mode != 'live':
                        # call_last is stale in prior_settle mode — fall back to call_settle
                        if today_c is None: today_c = atm_row.get('call_settle')
                        if today_p is None: today_p = atm_row.get('put_settle')
            else:
                today_c = today_p = None
        elif fwd:
            # ice_chain empty (COM miss) — compute ATM from live fwd directly
            _atm_frac = fwd % 1.0
            atm = float(math.ceil(fwd) if _atm_frac >= 0.50 else math.floor(fwd))
            today_c = today_p = None
        else:
            fwd = atm = today_c = today_p = None

        if not all([fwd, atm]):
            continue

        try:
            _lt_dt   = datetime.strptime(lt, '%Y-%m-%d')
            dte      = max(0, (_lt_dt - datetime.strptime(datetime.now().strftime('%Y-%m-%d'), '%Y-%m-%d')).days)
            T        = dte / 365.0
            T_settle = max(0, (_lt_dt - datetime.strptime(_actual_settle_str, '%Y-%m-%d')).days) / 365.0
        except ValueError:
            continue
        if dte <= 0:
            continue
        if T_settle <= 0:
            T_settle = T

        month_num, yr = my
        if month_num not in CT_STANDARD_MONTHS:
            # Serial month (Aug/Sep/Nov/Jan/Feb/Apr/Jun): always derive the live straddle
            # from the underlying standard month's live IV+forward so the value reflects
            # today's move rather than being frozen at yesterday's settle.
            std_m, std_y = next_standard_month(month_num, yr)
            _inv_mc = {v: k for k, v in MONTH_CODE.items()}
            std_tkr = f"CT{_inv_mc.get(std_m, '')}{str(std_y)[-1:]}"
            std_iv  = atm_iv.get(std_tkr)
            std_row = (_ice_raw or {}).get('futures', {}).get(std_tkr, {})
            _sfb, _sfo = std_row.get('bid'), std_row.get('offer')
            if _sfb and _sfo and _sfb > 0 and _sfo > 0:
                std_fwd = (_sfb + _sfo) / 2.0
            elif std_row.get('last') and std_row['last'] > 0:
                std_fwd = std_row['last']
            elif _tape_live.get(std_tkr):
                std_fwd = _tape_live[std_tkr]
            else:
                std_fwd = std_row.get('settle') or futures.get(std_tkr)
            if std_fwd:
                # ATM rule: fractional ≥ 0.50 → upper strike (ceil), < 0.50 → lower (floor)
                _frac = std_fwd % 1.0
                atm = float(math.ceil(std_fwd) if _frac >= 0.50 else math.floor(std_fwd))
                atm_row = next((r for r in ice_chain if abs(r['strike'] - atm) < 0.01), None)
                fwd = std_fwd
                # Use actual live bid/ask at the correct ATM — same source as the header card.
                # B76 derivation from the standard month vol is only a fallback when no live
                # market exists for the serial month.
                _tc = _tp = None
                if atm_row:
                    _cb, _co = atm_row.get('call_bid'), atm_row.get('call_offer')
                    _pb, _po = atm_row.get('put_bid'),  atm_row.get('put_offer')
                    if _cb and _co and _cb > 0 and _co > 0:
                        _tc = (_cb + _co) / 2.0
                    if _pb and _po and _pb > 0 and _po > 0:
                        _tp = (_pb + _po) / 2.0
                if _tc and _tp:
                    val = round(_tc + _tp, 2)
                elif atm_iv.get(ticker) or std_iv:
                    _serial_iv = (atm_iv.get(ticker) or std_iv) / 100.0
                    val = round(b76_price(std_fwd, atm, T, RISK_FREE, _serial_iv, True) +
                                b76_price(std_fwd, atm, T, RISK_FREE, _serial_iv, False), 2)
                else:
                    continue
            elif today_c is not None and today_p is not None:
                val = round(today_c + today_p, 2)  # last resort: frozen settle
            else:
                continue
        elif today_c is not None and today_p is not None:
            val = round(today_c + today_p, 2)
            # prior_settle mode + no live bid/ask: val came from call_settle (yesterday's).
            # Override with CSV settled prices if available, else B76 at today's settled futures.
            # Skip when live bid/ask was used — val already reflects today's market.
            if _rtd_mode != 'live' and post_settle and not _val_from_live_bid_ask:
                if _csv_opt_settle:
                    _opt_csv = _csv_opt_settle.get((ticker, int(atm)), {})
                    _tc_csv, _tp_csv = _opt_csv.get('C'), _opt_csv.get('P')
                    if _tc_csv and _tp_csv:
                        val = round(_tc_csv + _tp_csv, 2)
                else:
                    _fut_today = futures.get(ticker)
                    _prior_iv_ps = atm_iv.get(ticker)
                    if _fut_today and _prior_iv_ps and T > 0:
                        val = round(b76_price(_fut_today, atm, T, RISK_FREE, _prior_iv_ps / 100.0, True) +
                                    b76_price(_fut_today, atm, T, RISK_FREE, _prior_iv_ps / 100.0, False), 2)
        else:
            # No live bid/offer/last from options chain.
            # B76 theoretical: use whenever fwd + prior IV are available — works
            # during the settle window (_ice_raw=None) because fwd comes from the
            # futures CSV (today's settled price) and _prior_iv from yesterday's CSV.
            # Fall back to CSV settled prices only when B76 inputs are unavailable.
            _prior_iv = atm_iv.get(ticker)
            if _prior_iv and fwd and T > 0:
                val = round(b76_price(fwd, atm, T, RISK_FREE, _prior_iv / 100.0, True) +
                            b76_price(fwd, atm, T, RISK_FREE, _prior_iv / 100.0, False), 2)
            elif atm is not None and _csv_opt_settle:
                _opt = _csv_opt_settle.get((ticker, int(atm)), {})
                _tc, _tp = _opt.get('C'), _opt.get('P')
                if _tc and _tp:
                    val = round(_tc + _tp, 2)
                else:
                    continue
            else:
                continue  # no prices available

        # Implied vol from the settle straddle (exact, using workbook forward)
        df_t   = math.exp(-RISK_FREE * T)
        call_eq = (val + (fwd - atm) * df_t) / 2.0
        iv_pct = None
        if call_eq > 0:
            iv = implied_vol(call_eq, fwd, atm, T, RISK_FREE, True)
            if iv:
                iv_pct = round(iv * 100, 2)
        if iv_pct is None:
            iv_pct = atm_iv.get(ticker)

        # Settlement straddle = yesterday's published settlement straddle.
        # fwd_settle = yesterday's futures settle — used only for settle_iv_pct back-solve.
        # Pre-settlement: RTD 'settle' is static (yesterday's); post-settlement it flips
        # to today's so use csv_prev_settle to stay anchored to yesterday.
        if post_settle:
            if month_num not in CT_STANDARD_MONTHS:
                _sm_fs, _sy_fs = next_standard_month(month_num, yr)
                fwd_settle = csv_prev_settle.get((_sm_fs, _sy_fs)) or (ice_fut_row.get('settle') if ice_fut_row else None) or futures.get(ticker)
            else:
                fwd_settle = csv_prev_settle.get((month_num, yr)) or (ice_fut_row.get('settle') if ice_fut_row else None) or futures.get(ticker)
        else:
            fwd_settle = (ice_fut_row.get('settle') if ice_fut_row else None) or futures.get(ticker)

        # Settlement lookup uses TODAY's ATM strike (prev_atm = atm).
        # Settlement = yesterday's straddle price at today's ATM → Change and % CHG Vol
        # are same-strike comparisons. fwd_settle is only used in the settle_iv_pct
        # back-solve to derive yesterday's implied vol at today's ATM strike.
        prev_atm = atm
        prev_atm_row = atm_row

        # Pre-settlement: RTD call_settle/put_settle = yesterday's values → use directly.
        # Post-settlement: RTD has flipped to today's values.
        #   Priority 1: flow_rtd.json (written at futures settlement, before options
        #               settled) — holds yesterday's ICE call_settle/put_settle exactly.
        #   Priority 2: CSV prev_date px_settle (Bloomberg approx, fallback only).
        prev_c = prev_p = None
        if not post_settle and prev_atm_row:
            cs = prev_atm_row.get('call_settle')
            ps = prev_atm_row.get('put_settle')
            if cs and cs > 0: prev_c = cs
            if ps and ps > 0: prev_p = ps

        if (prev_c is None or prev_p is None) and post_settle:
            # flow_rtd.json — correct ICE yesterday settle
            _fk = (ticker, int(prev_atm))
            if _fk in _flow_rtd_opts:
                prev_c, prev_p = _flow_rtd_opts[_fk]

        # Final fallback: CSV px_settle
        if prev_c is None or prev_p is None:
            settle_ref = prev_date if (post_settle and last_date == today_str) else last_date
            for r in ct_opts:
                if r['ticker'] != ticker or abs(r['strike'] - prev_atm) >= 0.01:
                    continue
                if r['date'] == settle_ref:
                    if r['pc'] == 'Call' and r['px'] > 0 and prev_c is None:
                        prev_c = r['px']
                    elif r['pc'] == 'Put' and r['px'] > 0 and prev_p is None:
                        prev_p = r['px']

        # Serial-month B76 fallback: if CSV has no ATM strike for this serial,
        # derive settlement straddle from the standard month's prior settlement IV.
        # Use _prev_atm_s2 (from yesterday's standard month settle) not live atm.
        if (prev_c is None or prev_p is None) and month_num not in CT_STANDARD_MONTHS:
            _settle_ref2 = prev_date if (post_settle and last_date == today_str) else last_date
            std_m2, std_y2 = next_standard_month(month_num, yr)
            _inv_mc2 = {v: k for k, v in MONTH_CODE.items()}
            std_tkr2 = f"CT{_inv_mc2.get(std_m2, '')}{str(std_y2)[-1:]}"
            _std_settle_iv = atm_iv_for_date(std_tkr2, _settle_ref2)
            _fwd_s2 = csv_prev_settle.get((std_m2, std_y2)) or futures.get(std_tkr2) or futures.get(ticker)
            if _fwd_s2 and _fwd_s2 > 0:
                _prev_atm_s2 = float(math.ceil(_fwd_s2) if (_fwd_s2 % 1.0) >= 0.50 else math.floor(_fwd_s2))
            else:
                _prev_atm_s2 = prev_atm
            if _std_settle_iv and _fwd_s2 and _fwd_s2 > 0 and T_settle > 0:
                _half = b76_price(_fwd_s2, _prev_atm_s2, T_settle, RISK_FREE, _std_settle_iv / 100.0, True)
                if _half and _half > 0:
                    prev_c = _half
                    prev_p = _half

        prev_val = round(prev_c + prev_p, 2) if (prev_c and prev_p) else None
        chg      = round(val - prev_val, 2) if prev_val is not None else None

        # % CHG on day = live IV − settlement IV
        # Settlement IV uses T_settle (DTE as of last_date) not today's T — the settlement
        # straddle price was set when the option had T_settle days left, not T days.
        # prev_atm used throughout so settlement IV is computed at yesterday's strike.
        settle_iv_pct = None
        if prev_val and fwd_settle and fwd_settle > 0 and T_settle > 0:
            df_s = math.exp(-RISK_FREE * T_settle)
            call_eq_s = (prev_val + (fwd_settle - prev_atm) * df_s) / 2.0
            if call_eq_s > 0:
                iv_s = implied_vol(call_eq_s, fwd_settle, prev_atm, T_settle, RISK_FREE, True)
                if iv_s:
                    settle_iv_pct = round(iv_s * 100, 2)
        chg_vol = round(iv_pct - settle_iv_pct, 2) if (iv_pct is not None and settle_iv_pct is not None) else None
        if chg_vol is not None:
            atm_iv_1d_chg[ticker] = chg_vol

        try:
            breakeven = round(atm * (iv_pct / 100.0) / math.sqrt(252), 2) if iv_pct else None
        except (ValueError, ZeroDivisionError):
            breakeven = None

        straddles.append({
            'ticker':    ticker,
            'label':     label,
            'forward':   round(fwd, 2),
            'strike':    atm,
            'value':     val,
            'atm_vol':   iv_pct,
            'chg_vol':   chg_vol,
            'prev':      prev_val,
            'change':    chg,
            'dte':       dte,
            'expiry':    lt,
            'breakeven': breakeven,
        })

    # ── Serialize options for client ──────────────────────────────────────────
    def filter_date(date_str):
        return [r for r in ct_opts if r['date'] == date_str]

    _persist_today(last_date)

    # Auto-persist (cold-start bootstrap only) — settle_watcher owns the main CSVs
    # during trading hours. Only run after 16:30 ET AND settle_watcher has no record
    # for today (i.e. it clearly did not run). Never write during the settlement window.
    if _ice_raw and _ice_raw.get('options'):
        ice_tickers_set  = set(s.upper() for s in _ice_raw['options'])
        today_tickers_set = set(r['ticker'] for r in today_opts)
        if ice_tickers_set - today_tickers_set:
            _allow_persist = False
            try:
                try:
                    from zoneinfo import ZoneInfo as _ZI2
                except ImportError:
                    from backports.zoneinfo import ZoneInfo as _ZI2
                _now_et2 = datetime.now(_ZI2('America/New_York'))
                _today_cal2 = _now_et2.strftime('%Y-%m-%d')
                if (_now_et2.hour, _now_et2.minute) >= (16, 30):
                    _status_path2 = os.path.join(os.path.dirname(__file__), 'settle_status.json')
                    try:
                        with open(_status_path2, encoding='utf-8') as _sf2:
                            _ss2 = json.load(_sf2)
                        _sw_ran_today = (_ss2.get('date') == _today_cal2)
                    except Exception:
                        _sw_ran_today = False
                    if not _sw_ran_today:
                        _allow_persist = True
            except Exception:
                pass
            if _allow_persist:
                _snap = _ice_raw
                _td   = last_date
                def _auto_persist():
                    try:
                        _persist_ct_options_ice(_snap, _td)
                        _persist_futures_ice(LOCAL_FUT_HISTORY, 'CT', _snap, _td)
                        log.info('Cold-start auto-persist: %s', ice_tickers_set - today_tickers_set)
                    except Exception as e:
                        log.warning('Auto-persist failed: %s', e)
                threading.Thread(target=_auto_persist, daemon=True).start()
            else:
                log.debug('Auto-persist skipped: before 16:30 ET or settle_watcher ran today')

    _result = {
        'last_date':      last_date,
        'today_str':      datetime.now().strftime('%Y-%m-%d'),
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
        'straddles':      straddles,
        'commodity':      'CT',
        'commodity_name': 'ICE Cotton No. 2',
    }
    # Stale straddle guard: if RTD is offline and straddles are empty,
    # serve the last cached straddles from today so the dashboard doesn't
    # go blank between market close and settlement publication.
    if not _result.get('straddles') and _ice_raw is None:
        _prev = _ld_cache.get('CT')
        if _prev and _prev['data'].get('last_date') == last_date:
            _prev_strads = _prev['data'].get('straddles')
            if _prev_strads:
                _result['straddles'] = _prev_strads
                log.debug('Straddles: serving cached values (RTD offline, same day)')

    _ld_cache['CT'] = {'data': _result, 'ts': _now, 'om': _om, 'fm': _fm}
    return _result

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

    # Local is behind today — fetch GitHub and write to Bloomberg shadow backup ONLY.
    # The main pipeline CSVs (LOCAL_OPT_HISTORY, LOCAL_FUT_HISTORY) are written
    # exclusively by settle_watcher.py from ICE RTD. Bloomberg data never enters
    # the live pipeline. The backup files are a manual failsafe only.
    os.makedirs(_BBG_BACKUP_DIR, exist_ok=True)

    try:
        if opt_latest < today_str:
            all_rows = fetch_csv(OPT_CSV_URL)
            new_rows = [r for r in all_rows
                        if r.get('date', '').strip() > opt_latest
                        and r.get('commodity', '').strip().upper() == 'CT']
            if new_rows:
                new_rows.sort(key=lambda r: r.get('date', ''))
                need_header = not os.path.exists(_BBG_OPT_BACKUP)
                with open(_BBG_OPT_BACKUP, 'a', newline='', encoding='utf-8') as f:
                    w = csv.DictWriter(f, fieldnames=new_rows[0].keys())
                    if need_header:
                        w.writeheader()
                    w.writerows(new_rows)
                new_dates = sorted(set(r.get('date','') for r in new_rows))
                log.info('Bloomberg backup: opt rows for %d dates: %s', len(new_dates), new_dates)
    except Exception as e:
        log.warning('opt bloomberg backup failed: %s', e)

    try:
        if fut_latest < today_str:
            all_rows = fetch_csv(OI_CSV_URL)
            today_raw = [r for r in all_rows
                         if r.get('date', '').strip() > fut_latest
                         and r.get('commodity', '').strip().upper() == 'CT']
            if today_raw:
                today_raw.sort(key=lambda r: r.get('date', ''))
                need_header = not os.path.exists(_BBG_FUT_BACKUP)
                with open(_BBG_FUT_BACKUP, 'a', newline='', encoding='utf-8') as f:
                    w = csv.DictWriter(f, fieldnames=today_raw[0].keys(), extrasaction='ignore')
                    if need_header:
                        w.writeheader()
                    w.writerows(today_raw)
                new_dates = sorted(set(r.get('date','') for r in today_raw))
                log.info('Bloomberg backup: fut rows for %d dates: %s', len(new_dates), new_dates)
    except Exception as e:
        log.warning('fut bloomberg backup failed: %s', e)


def _append_projected(path, new_rows):
    """Append GitHub feed rows to a local CSV, projected onto that file's own
    column schema. The GitHub feed carries more columns (and a different order)
    than the compact local KC/SB/CC files; writing raw rows would misalign data
    when the file is later read by DictReader against its short header. Read the
    existing header and write only those fields, by name, dropping extras."""
    if not os.path.exists(path):
        # No local file yet — write with the feed's own columns + header.
        with open(path, 'a', newline='', encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=list(new_rows[0].keys()),
                               extrasaction='ignore')
            w.writeheader()
            w.writerows(new_rows)
        return
    with open(path, 'r', newline='', encoding='utf-8') as f:
        header = next(csv.reader(f))
    with open(path, 'a', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=header, extrasaction='ignore')
        # Project each feed row onto the local header; missing cols -> ''.
        w.writerows({k: r.get(k, '') for k in header} for r in new_rows)


def _kc_watcher_owns_today(today_str):
    """True if settle_watcher_kc has confirmed today's KC futures settlement.
    When True, the GitHub feed must not append today's KC futures row — the
    watcher's true-ICE-settle row is authoritative. Reads settle_status_kc.json
    written by settle_watcher_kc.py. Fails open (returns False) so the feed
    still runs if the status file is missing/unreadable (watcher didn't run)."""
    try:
        status_path = os.path.join(os.path.dirname(__file__), 'settle_status_kc.json')
        if not os.path.exists(status_path):
            return False
        with open(status_path, 'r', encoding='utf-8') as f:
            st = json.load(f)
        return st.get('date') == today_str and bool(st.get('futures_settled'))
    except Exception:
        return False


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
                _append_projected(cfg['opt_csv'], new_rows)
                log.info('%s: persisted %d opt rows', commodity, len(new_rows))
    except Exception as e:
        log.warning('%s opt persist failed: %s', commodity, e)

    # KC failsafe: settle_watcher_kc owns today's KC futures row (true ICE
    # settle + high/low/spreads from RTD). If it has confirmed today's futures,
    # do NOT let the GitHub feed (last-trade) append a duplicate/overwriting row.
    # SB/CC have no watcher yet, so this gate only applies to KC.
    if commodity == 'KC' and _kc_watcher_owns_today(today_str):
        return

    try:
        if fut_latest < today_str:
            all_rows = fetch_csv(OI_CSV_URL)
            new_rows = [r for r in all_rows
                        if r.get('date', '').strip() > fut_latest
                        and r.get('commodity', '').strip().upper() == commodity]
            if new_rows:
                new_rows.sort(key=lambda r: r.get('date', ''))
                # KC's local futures file uses cotton's wide schema where the
                # spread-volume columns are efp_vol/efs_vol/block_vol, but the OI
                # feed names them efp_volume/efs_volume/block_volume. Alias them
                # so the feed's high/low/volume/efp/efs/block flow in by name
                # when the watcher did NOT run. (high/low/volume already match.)
                if commodity == 'KC':
                    _alias = {'efp_volume': 'efp_vol',
                              'efs_volume': 'efs_vol',
                              'block_volume': 'block_vol'}
                    for r in new_rows:
                        for src, dst in _alias.items():
                            if src in r and not r.get(dst):
                                r[dst] = r.get(src, '')
                _append_projected(cfg['fut_csv'], new_rows)
                log.info('%s: persisted %d fut rows', commodity, len(new_rows))
    except Exception as e:
        log.warning('%s fut persist failed: %s', commodity, e)


# ── ICE RTD daily settle persistence ──────────────────────────────────────────




def _delete_date_from_csv(path, date_str, commodity=None):
    """Remove all rows matching date_str (and optional commodity) from a CSV in-place."""
    if not os.path.exists(path):
        return
    try:
        with open(path, 'r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            fieldnames = reader.fieldnames
            rows = list(reader)
        kept = [r for r in rows
                if r.get('date', '').strip() != date_str
                or (commodity and r.get('commodity', '').strip().upper() != commodity.upper())]
        with open(path, 'w', newline='', encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            w.writeheader()
            w.writerows(kept)
    except Exception as e:
        log.warning('_delete_date_from_csv(%s, %s) failed: %s', path, date_str, e)


def _contract_code_to_month_year(contract_code, prefix):
    """'CTN6' → (month_num=7, year=2026)."""
    plen = len(prefix)
    code = contract_code[plen:]
    if len(code) < 2:
        return None
    mc = code[0].upper()
    month_num = MONTH_CODE.get(mc)
    if month_num is None:
        return None
    try:
        year = _decade_year(int(code[1]))
    except (ValueError, IndexError):
        return None
    return month_num, year


def _futures_ordinal_contract(month_num, year, today_str, prefix):
    """Return ordinal contract code matching get_hist_fwd formula, e.g. 'CTJUL1'."""
    try:
        d = datetime.strptime(today_str, '%Y-%m-%d')
    except ValueError:
        return None
    first_year = d.year if d.month <= month_num else d.year + 1
    ordinal = year - first_year + 1
    if ordinal < 1 or ordinal > 6:
        return None
    suffix = FUTURES_MONTH_SUFFIX.get(month_num)
    if not suffix:
        return None
    return f'{prefix}{suffix}{ordinal}'


def _build_lt_fn_lookup(fut_path, prefix):
    """Read existing futures CSV → {(month_num, year): {'last_trade': ..., 'first_notice': ...}}"""
    lookup = {}
    if not os.path.exists(fut_path):
        return lookup
    try:
        with open(fut_path, 'r', newline='', encoding='utf-8') as f:
            for row in csv.DictReader(f):
                lt = (row.get('last_trade')   or '').strip()
                fn = (row.get('first_notice') or '').strip()
                if not lt and not fn:
                    continue
                contract = (row.get('contract') or '').strip()
                date_s   = (row.get('date')     or '').strip()
                if not contract or not date_s:
                    continue
                res = parse_futures_my_generic(contract, date_s, prefix)
                if not res:
                    continue
                month_num, year = res
                key = (month_num, year)
                if key not in lookup:
                    lookup[key] = {'last_trade': lt or None, 'first_notice': fn or None}
    except Exception:
        pass
    return lookup


_CT_OPT_FIELDS      = ['date','commodity','security_des','contract_month','put_call',
                        'strike_px','open_int','oi_chg','px_settle','px_volume']
_GENERIC_OPT_FIELDS = ['date','commodity','security_des','strike_px',
                        'px_settle','open_int','oi_chg','px_volume']
_FUT_FIELDS         = ['date','commodity','contract',
                        'settle','open_int','oi_chg','first_notice','last_trade']


def _persist_ct_options_ice(ice_data, today_str):
    """Upsert today's CT option settles from ICE RTD into local_options_history.csv."""
    # Never overwrite data that settle_watcher.py has already confirmed as settled.
    # settle_watcher owns the settled rows — live RTD prices must not replace them.
    _status_path = os.path.join(os.path.dirname(__file__), 'settle_status.json')
    try:
        with open(_status_path, encoding='utf-8') as _sf:
            _st = json.load(_sf)
        if _st.get('date') == today_str and _st.get('options_settled'):
            log.info('CT options already settled for %s — skipping auto-persist', today_str)
            return 0
    except Exception:
        pass

    options = ice_data.get('options', {})
    if not options:
        return 0
    _delete_date_from_csv(LOCAL_OPT_HISTORY, today_str, commodity='CT')

    rows = []
    for sheet_name, chain in options.items():
        res = _contract_code_to_month_year(sheet_name, 'CT')
        if not res:
            continue
        month_num, year = res
        contract_month = f'{MONTH_NAME[month_num]} {year}'
        for r in chain:
            strike = r.get('strike')
            if strike is None:
                continue
            for pc_char, settle_key, oi_key in [
                ('C', 'call_settle', 'call_oi'),
                ('P', 'put_settle',  'put_oi'),
            ]:
                px = r.get(settle_key)
                if px is None:
                    continue
                rows.append({
                    'date':           today_str,
                    'commodity':      'CT',
                    'security_des':   f'{sheet_name.upper()}{pc_char}',
                    'contract_month': contract_month,
                    'put_call':       pc_char,
                    'strike_px':      strike,
                    'open_int':       r.get(oi_key) or '',
                    'oi_chg':         '',
                    'px_settle':      round(float(px), 4),
                    'px_volume':      '',
                })

    if not rows:
        return 0
    need_header = not os.path.exists(LOCAL_OPT_HISTORY)
    with open(LOCAL_OPT_HISTORY, 'a', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=_CT_OPT_FIELDS)
        if need_header:
            w.writeheader()
        w.writerows(rows)
    log.info('CT: persisted %d opt rows for %s', len(rows), today_str)
    return len(rows)


def _persist_generic_options_ice(opt_path, commodity, ice_data, today_str):
    """Upsert today's option settles from ICE RTD into local_{kc/sb/cc}_options_history.csv."""
    options = ice_data.get('options', {})
    if not options:
        return 0
    _delete_date_from_csv(opt_path, today_str, commodity=commodity)

    rows = []
    prefix = commodity.upper()
    for sheet_name, chain in options.items():
        for r in chain:
            strike = r.get('strike')
            if strike is None:
                continue
            for pc_char, settle_key, oi_key in [
                ('C', 'call_settle', 'call_oi'),
                ('P', 'put_settle',  'put_oi'),
            ]:
                px = r.get(settle_key)
                if px is None:
                    continue
                rows.append({
                    'date':         today_str,
                    'commodity':    prefix,
                    'security_des': f'{sheet_name.upper()}{pc_char}',
                    'strike_px':    strike,
                    'px_settle':    round(float(px), 4),
                    'open_int':     r.get(oi_key) or '',
                    'oi_chg':       '',
                    'px_volume':    '',
                })

    if not rows:
        return 0
    need_header = not os.path.exists(opt_path)
    with open(opt_path, 'a', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=_GENERIC_OPT_FIELDS)
        if need_header:
            w.writeheader()
        w.writerows(rows)
    log.info('%s: persisted %d opt rows for %s', commodity, len(rows), today_str)
    return len(rows)


def _persist_futures_ice(fut_path, commodity, ice_data, today_str):
    """Upsert today's futures settles from ICE RTD into the local futures CSV."""
    # Never overwrite data that settle_watcher.py has already confirmed as settled.
    if commodity.upper() == 'CT':
        _status_path = os.path.join(os.path.dirname(__file__), 'settle_status.json')
        try:
            with open(_status_path, encoding='utf-8') as _sf:
                _st = json.load(_sf)
            if _st.get('date') == today_str and _st.get('futures_settled'):
                log.info('CT futures already settled for %s — skipping auto-persist', today_str)
                return 0
        except Exception:
            pass

    futures = ice_data.get('futures', {})
    if not futures:
        return 0
    _delete_date_from_csv(fut_path, today_str, commodity=commodity)

    prefix  = commodity.upper()
    lt_fn   = _build_lt_fn_lookup(fut_path, prefix)

    # Build previous OI lookup from the most recent CSV entry per contract
    # so oi_chg is computed and stored at write time rather than patched at read time.
    prev_oi = {}
    try:
        for row in read_local_csv(fut_path):
            if row.get('commodity', '').strip().upper() != prefix:
                continue
            if (row.get('date', '').strip()) == today_str:
                continue
            contract = row.get('contract', '').strip()
            oi_s     = row.get('open_int', '').strip()
            date_s   = row.get('date', '').strip()
            if not contract or not oi_s or not date_s:
                continue
            res2 = _contract_code_to_month_year(contract, prefix)
            if not res2:
                continue
            try:
                oi_f = float(oi_s)
            except (ValueError, TypeError):
                continue
            if res2 not in prev_oi or date_s > prev_oi[res2][0]:
                prev_oi[res2] = (date_s, oi_f)
    except Exception:
        pass
    prev_oi = {k: v[1] for k, v in prev_oi.items()}

    rows    = []
    for contract_code, f in futures.items():
        settle = f.get('settle')
        if settle is None:
            continue
        res = _contract_code_to_month_year(contract_code, prefix)
        if not res:
            continue
        month_num, year = res
        ordinal_contract = _futures_ordinal_contract(month_num, year, today_str, prefix)
        if not ordinal_contract:
            continue
        lt_fn_info = lt_fn.get((month_num, year), {})
        new_oi  = f.get('oi')
        old_oi  = prev_oi.get((month_num, year))
        try:
            oi_chg = round(float(new_oi) - old_oi) if (new_oi is not None and old_oi is not None) else ''
        except (TypeError, ValueError):
            oi_chg = ''
        rows.append({
            'date':         today_str,
            'commodity':    prefix,
            'contract':     ordinal_contract,
            'settle':       round(float(settle), 4),
            'open_int':     new_oi or '',
            'oi_chg':       oi_chg,
            'first_notice': lt_fn_info.get('first_notice') or '',
            'last_trade':   lt_fn_info.get('last_trade') or '',
        })

    if not rows:
        return 0
    need_header = not os.path.exists(fut_path)
    with open(fut_path, 'a', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=_FUT_FIELDS)
        if need_header:
            w.writeheader()
        w.writerows(rows)
    log.info('%s: persisted %d fut rows for %s', commodity, len(rows), today_str)
    return len(rows)


# Settlement persistence is handled by Options_flow_analyzer/settle_watcher.py.
# That process runs daily from 14:25 ET, detects settlement via RTD workbook,
# and writes to local_futures_history.csv, local_futures_spreads_history.csv,
# and local_options_history.csv. The _ld_cache invalidates on file mtime change
# so the dashboard picks up new data automatically on the next browser refresh.

# ── Background schedulers ─────────────────────────────────────────────────────

# Pre-close cache flush: 19:15 BST = 14:15 ET — forces a fresh live data load
# 5 minutes before the 14:20 CT close so cached values are current at EOD.
PRECLOSE_FLUSH_HOUR   = 19
PRECLOSE_FLUSH_MINUTE = 15

_preclose_timer      = None
_preclose_timer_lock = threading.Lock()


def _run_preclose_flush():
    """Cache flush at PRECLOSE_FLUSH_HOUR:PRECLOSE_FLUSH_MINUTE. No CSV write — live data only."""
    log.info('Pre-close cache flush firing')
    _cache.clear()
    _ld_cache.clear()
    _schedule_preclose_flush()


def _schedule_preclose_flush():
    global _preclose_timer
    with _preclose_timer_lock:
        if _preclose_timer is not None:
            _preclose_timer.cancel()
        now    = datetime.now()
        target = now.replace(hour=PRECLOSE_FLUSH_HOUR, minute=PRECLOSE_FLUSH_MINUTE,
                             second=0, microsecond=0)
        if target <= now:
            target += timedelta(days=1)
        delay = (target - now).total_seconds()
        _preclose_timer = threading.Timer(delay, _run_preclose_flush)
        _preclose_timer.daemon = True
        _preclose_timer.start()
        log.info('Next pre-close flush in %.0f s (at %s)', delay, target.strftime('%H:%M'))




def _compute_hv(fut_path, commodity, windows=(10, 30, 60, 90)):
    """
    Historical volatility from local futures CSV, computed per ordinal contract series.
    CTJUL1 = current nearest-Jul contract (e.g. CTN6), CTDEC1 = CTZ6, CTJUL2 = CTN7, etc.
    Each series is computed independently; roll jumps (>15% daily) are excluded.
    Returns {ticker: {'hv10': ..., 'hv30': ..., 'hv60': ..., 'hv90': ...}} as decimals.
    """
    if not os.path.exists(fut_path):
        return {}
    prefix = commodity.upper()
    plen   = len(prefix)
    _month_letter = {v: k for k, v in MONTH_CODE.items()}  # int → letter

    # Collect settle prices and last_trade per ordinal contract name.
    # Also collect individual contract rows (e.g. CTN6) separately so we can
    # append them to the ordinal series tail (settle_watcher writes individual
    # rows for recent dates once the ordinal series stops rolling).
    contract_prices = {}   # {contract_name: [(date_str, settle), ...]}
    contract_lt     = {}   # {contract_name: last_trade_str}
    indiv_prices    = {}   # {ticker: [(date_str, settle), ...]} e.g. {'CTN6': [...]}

    # Individual contract pattern: exactly plen+2 chars, e.g. CTN6 (CT + N + 6)
    _indiv_re = re.compile(r'^' + re.escape(prefix) + r'([FGHJKMNQUVXZ])(\d)$')

    try:
        with open(fut_path, 'r', newline='', encoding='utf-8') as f:
            for row in csv.DictReader(f):
                if (row.get('commodity') or '').strip().upper() != commodity:
                    continue
                contract = (row.get('contract') or '').strip()
                settle_s = (row.get('settle') or '').strip()
                date_s   = (row.get('date') or '').strip()
                lt_s     = (row.get('last_trade') or '').strip()
                if not contract or not settle_s or not date_s:
                    continue
                try:
                    settle = float(settle_s)
                    if settle <= 0:
                        continue
                except (ValueError, TypeError):
                    continue
                if _indiv_re.match(contract):
                    indiv_prices.setdefault(contract, []).append((date_s, settle))
                elif len(contract) >= plen + 4:
                    contract_prices.setdefault(contract, []).append((date_s, settle))
                    if lt_s:
                        contract_lt[contract] = lt_s
    except Exception:
        return {}

    result = {}
    max_w = max(windows)

    for contract, price_list in contract_prices.items():
        if len(price_list) < 5:
            continue
        # Identify ticker from last_trade date + month suffix
        lt_str   = contract_lt.get(contract)
        if not lt_str:
            continue
        suffix_3 = contract[plen:plen + 3].upper()
        month_num = FUTURES_MONTH_FROM_SUFFIX.get(suffix_3)
        if not month_num:
            continue
        try:
            lt_dt = datetime.strptime(lt_str, '%Y-%m-%d')
            year  = lt_dt.year if lt_dt.month <= month_num else lt_dt.year + 1
        except ValueError:
            continue
        month_letter = _month_letter.get(month_num)
        if not month_letter:
            continue
        ticker = f"{prefix}{month_letter}{str(year)[-1:]}"

        # Extend ordinal series with any individual contract rows (e.g. CTN6)
        # that come after the last ordinal row — settle_watcher writes these
        # for the most recent trading days once the ordinal series stops.
        indiv = sorted(indiv_prices.get(ticker, []), key=lambda x: x[0])
        if indiv:
            last_ord_date = sorted(price_list, key=lambda x: x[0])[-1][0]
            price_list = price_list + [(d, s) for d, s in indiv if d > last_ord_date]

        # Compute log returns from the tail of the series (enough for max window)
        sorted_prices = sorted(price_list, key=lambda x: x[0])
        tail = sorted_prices[-(max_w + 10):]   # grab a bit more than needed
        log_returns = []
        for i in range(1, len(tail)):
            prev_s, curr_s = tail[i-1][1], tail[i][1]
            if prev_s > 0 and curr_s > 0:
                try:
                    lr = math.log(curr_s / prev_s)
                    if abs(lr) < 0.15:   # exclude roll-day jumps
                        log_returns.append(lr)
                except (ValueError, ZeroDivisionError):
                    pass

        hv = {}
        for w in windows:
            sample = log_returns[-w:]
            if len(sample) >= max(w // 2, 5):
                n    = len(sample)
                mean = sum(sample) / n
                var  = sum((x - mean) ** 2 for x in sample) / max(n - 1, 1)
                hv[f'hv{w}'] = math.sqrt(var) * math.sqrt(252)
            else:
                hv[f'hv{w}'] = None
        result[ticker] = hv

    return result


def _load_generic_data(commodity):
    """Load and compute options analytics for KC, SB, or CC with full ICE RTD live data."""
    cfg        = COMMODITY_CONFIG[commodity]
    prefix     = cfg['prefix']
    std_months = cfg['std_months']
    excl_months = cfg['excl_months']
    serial_map = cfg['serial_map']

    try:
        # Mirror cotton: read deep history from the local commodity files.
        # The OI-dashboard GitHub feed is used only by _persist_today_generic
        # to append the daily tail into these same local files.
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
            oi  = int(float(r.get('open_int', 0) or 0))
            oic = int(float(r.get('oi_chg', 0) or 0))
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

    rtd = None
    if _ice_rtd_reader and not _in_ct_settle_window():
        try:
            rtd = _ice_to_rtd_shape(_read_ice_workbook_safe(commodity))
        except Exception as e:
            log.debug('ICE RTD fetch skipped for %s: %s', commodity, e)

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
            elif commodity == 'CC':
                exp = _cc_opt_expiry(month_num, year)
            elif commodity == 'SB':
                exp = _sb_opt_expiry(month_num, year)
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

    # ── RTD live data overlay ─────────────────────────────────────────────────
    hv_data      = {}
    live_futures = {}
    rtd_spreads  = {}
    data_source  = 'csv_only'

    if rtd:
        data_source = rtd.get('source', 'csv_only')
        for tkr, d in (rtd.get('outrights') or {}).items():
            hv_data[tkr]      = {k: d.get(k) for k in ('hv10', 'hv30', 'hv60', 'hv90')}
            live_futures[tkr] = {k: d.get(k) for k in (
                'last', 'settle', 'yest_settle', 'change', 'pct_chg',
                'oi', 'oi_chg', 'volume', 'high', 'low',
                'block_vol', 'efs_vol', 'efp_vol',
            )}
        rtd_spreads = {key: d for key, d in (rtd.get('spreads') or {}).items()}

    # ── KC post-settle true-settle re-lock (KC only; mirrors cotton 1311-1351) ──
    # settle_watcher_kc writes the TRUE ICE settle (+ high/low/vol/OI) to the
    # local KC futures CSV after close. When RTD is offline/missing post-settle,
    # fill any None live_futures field from that CSV row so settled values are
    # correct regardless of workbook state. Fill-if-None: live RTD wins when
    # present; the watcher's true settle fills the gaps. CT/SB/CC skip this.
    if commodity == 'KC':
        _fut_dates = [(_r.get('date') or '').strip() for _r in comm_fut
                      if (_r.get('date') or '').strip()]
        _fut_last  = max(_fut_dates) if _fut_dates else ''
        _today_fb  = datetime.now().strftime('%Y-%m-%d')
        if _fut_last == _today_fb:
            def _fv(v):
                try: return float(v) if v not in (None, '') else None
                except (ValueError, TypeError): return None
            for _row in comm_fut:
                if (_row.get('date') or '').strip() != _fut_last:
                    continue
                _tkr_raw = (_row.get('contract') or '').strip()
                if not _tkr_raw:
                    continue
                # Watcher writes ICE codes (KCN6, 4 chars); feed-history rows are
                # old ordinal format (KCMAR1, 6 chars) — translate the latter to
                # ICE code so the key matches live_futures (ICE-keyed from RTD).
                # Must use the prefix-aware translator (_generic_to_ice_code is
                # CT-hardcoded and returns None for KC -> orphan keys / no-op).
                _tkr = _tkr_raw if len(_tkr_raw) == 4 \
                       else (_ordinal_to_ice_code(_tkr_raw, _fut_last, prefix) or _tkr_raw)
                if _tkr not in live_futures:
                    live_futures[_tkr] = {}
                _lf = live_futures[_tkr]
                for _csv_key, _lf_key in [
                    ('settle',     'settle'),
                    ('yest_settle','yest_settle'),
                    ('high',       'high'),
                    ('low',        'low'),
                    ('volume',     'volume'),
                    ('efp_vol',    'efp_vol'),
                    ('efs_vol',    'efs_vol'),
                    ('block_vol',  'block_vol'),
                    ('open_int',   'oi'),
                    ('oi_chg',     'oi_chg'),
                ]:
                    if _lf.get(_lf_key) is None:
                        _lf[_lf_key] = _fv(_row.get(_csv_key))
                _s, _y = _lf.get('settle'), _lf.get('yest_settle')
                if _lf.get('change') is None and _s is not None and _y is not None and _y:
                    _lf['change']  = round(_s - _y, 4)
                    _lf['pct_chg'] = round((_s - _y) / _y * 100, 4)

        # ── KC spreads CSV fallback (mirrors cotton 1353-1393, generic) ──────
        # settle_watcher_kc writes calendar spreads to the KC spreads CSV after
        # close. The KC GitHub feed carries NO spread rows, so when RTD is
        # offline post-settle this is the only source of spread H/L/V/settle.
        if cfg.get('spr_csv') and _fut_last == _today_fb:
            _spr_by_key_fb = {}
            try:
                with open(cfg['spr_csv'], encoding='utf-8') as _ssf:
                    for _ssr in csv.DictReader(_ssf):
                        _sk = (_ssr.get('contract') or '').strip()
                        _sd = (_ssr.get('date') or '').strip()
                        if _sk and _sd and (_sk not in _spr_by_key_fb or _sd > _spr_by_key_fb[_sk]['date']):
                            _spr_by_key_fb[_sk] = _ssr
            except Exception:
                pass
            def _sfv(v):
                try: return float(v) if v not in (None, '') else None
                except (ValueError, TypeError): return None
            for _sk, _scr in _spr_by_key_fb.items():
                if (_scr.get('date') or '').strip() != _fut_last:
                    continue
                _sh = _sfv(_scr.get('high')); _sl = _sfv(_scr.get('low')); _svol = _sfv(_scr.get('volume'))
                if _sk not in rtd_spreads:
                    _sparts = _sk.split('/')
                    if len(_sparts) == 2:
                        _spn = parse_generic_ticker(_sparts[0], prefix)
                        _spf = parse_generic_ticker(_sparts[1], prefix)
                        _sdisp = (f"{MONTH_NAME[_spn[2]]}{str(_spn[1])[-2:]}/{MONTH_NAME[_spf[2]]}{str(_spf[1])[-2:]}"
                                  if _spn and _spf else _sk)
                        _ss_stt = _sfv(_scr.get('settle')); _ss_ys = _sfv(_scr.get('yest_settle'))
                        _ss_chg = _sfv(_scr.get('change'))
                        _ss_pct = round(_ss_chg / _ss_ys * 100, 2) if (_ss_chg is not None and _ss_ys and _ss_ys != 0) else None
                        rtd_spreads[_sk] = {
                            'display': _sdisp, 'settle': _ss_stt, 'yest_settle': _ss_ys,
                            'change': _ss_chg, 'pct_chg': _ss_pct,
                            'high': _sh, 'low': _sl, 'volume': _svol,
                            'block_vol': _sfv(_scr.get('block_vol')),
                            'efs_vol': _sfv(_scr.get('efs_vol')),
                            'efp_vol': _sfv(_scr.get('efp_vol')),
                        }
                else:
                    for _sfk, _scv in [('high', _sh), ('low', _sl), ('volume', _svol)]:
                        if rtd_spreads[_sk].get(_sfk) is None:
                            rtd_spreads[_sk][_sfk] = _scv

    # ── Live smile from ICE RTD option bid/ask ────────────────────────────────
    live_smile     = {}
    live_smile_fwd = {}
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

            impl_Fs = sorted([
                k + (pcs['Call'] - pcs['Put']) * math.exp(RISK_FREE * T)
                for k, pcs in by_k.items() if 'Call' in pcs and 'Put' in pcs
            ])
            live_F = impl_Fs[len(impl_Fs) // 2] if len(impl_Fs) >= 3 else futures.get(ticker)
            if not live_F:
                continue

            live_smile_fwd[ticker] = round(live_F, 4)
            avail    = sorted(by_k.keys())
            live_atm = min(avail, key=lambda k: abs(k - live_F)) if avail else atm_strike.get(ticker)
            if not live_atm:
                continue

            settle_atm_iv = (atm_iv.get(ticker) or 20) / 100.0
            iv_lo = settle_atm_iv * 0.50
            iv_hi = settle_atm_iv * 2.50
            smile_ivs = {}
            for s in lo['strikes']:
                K   = s['strike']
                bid = s.get('bid')
                ask = s.get('ask')
                if not bid or not ask or bid < 0.005 or ask < 0.005 or ask > bid * 8:
                    continue
                px = (bid + ask) / 2.0
                if px < 0.005:
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
                if d > 0.15:
                    if not (settle_atm_iv * 0.75 < iv < settle_atm_iv * 1.30):
                        continue
                else:
                    if not (iv_lo < iv < iv_hi):
                        continue
                smile_ivs.setdefault(K, []).append(iv)

            if len(smile_ivs) >= 5:
                smile = {round(k, 2): round(sum(ivs) / len(ivs) * 100, 2)
                         for k, ivs in smile_ivs.items()}
                if smile:
                    live_smile[ticker] = smile

            # ATM IV from live straddle mid
            atm_pxs      = by_k.get(live_atm, {})
            call_mid_atm = atm_pxs.get('Call')
            put_mid_atm  = atm_pxs.get('Put')
            if call_mid_atm and put_mid_atm and call_mid_atm > 0 and put_mid_atm > 0:
                strad_mid = call_mid_atm + put_mid_atm
                df_atm    = math.exp(-RISK_FREE * T)
                call_eq   = (strad_mid + (live_F - live_atm) * df_atm) / 2
                if call_eq > 0:
                    live_iv = implied_vol(call_eq, live_F, live_atm, T, RISK_FREE, True)
                    if live_iv and 0.01 < live_iv < 2.00:
                        atm_iv[ticker] = round(live_iv * 100, 2)
                        # 1D chg: live IV − settlement IV (from last_date CSV, not prev_date)
                        settle_iv_base = atm_iv_for_date(ticker, last_date)
                        if settle_iv_base is not None:
                            atm_iv_1d_chg[ticker] = round(live_iv * 100 - settle_iv_base * 100, 2)
                        iv_w = atm_iv_for_date(ticker, week_date, week_futures.get(ticker))
                        if iv_w is not None:
                            atm_iv_1w_chg[ticker] = round(live_iv * 100 - iv_w * 100, 2)

    # ── HV from local futures history ─────────────────────────────────────────
    csv_hv = _compute_hv(cfg['fut_csv'], commodity)
    for tkr in expiry_list:
        if tkr in csv_hv:
            hv_data.setdefault(tkr, {}).update(csv_hv[tkr])

    # ── Straddle run — live bid/offer mid from ICE RTD ────────────────────────
    # KC pre-close straddle freeze (13:20 ET): lock the straddle SOURCE 10 min
    # before the 13:30 KC close so strikes + values reflect the final liquid
    # market and stop drifting/blanking after the close. KC settle publishes
    # ~12:25 (BEFORE this freeze), so — unlike CT's freeze — it does NOT lift on
    # options_settled; it holds the 13:20 snapshot through EOD (cleared naturally
    # by the per-day directory). Scoped to the straddle source only: the futures
    # bar, vol smile and skew keep reading live `rtd`. KC-only via the
    # 'straddle_tickers' config that only KC defines (CT/SB/CC: _strad_src = rtd).
    _strad_src = rtd
    if cfg.get('straddle_tickers'):
        try:
            from zoneinfo import ZoneInfo as _ZI_kf
        except ImportError:
            from backports.zoneinfo import ZoneInfo as _ZI_kf
        _kf_now   = datetime.now(_ZI_kf('America/New_York'))
        _kf_hm    = (_kf_now.hour, _kf_now.minute)
        _kf_path  = os.path.normpath(os.path.join(
            os.path.dirname(__file__), '..', 'Options_flow_analyzer',
            'data', _kf_now.strftime('%Y-%m-%d'), 'rtd_snap_kc.json'))
        # Write once, only inside the 13:20–13:30 pre-close window, from a live
        # read that actually has option quotes — never freeze post-close/stale data.
        if ((13, 20) <= _kf_hm <= (13, 30) and rtd and rtd.get('live_options')
                and not os.path.exists(_kf_path)):
            try:
                os.makedirs(os.path.dirname(_kf_path), exist_ok=True)
                with open(_kf_path, 'w', encoding='utf-8') as _kf:
                    json.dump({'live_options': rtd.get('live_options'),
                               'outrights':    rtd.get('outrights')}, _kf)
            except Exception:
                pass
        # Once the freeze exists, use it as the straddle source for the rest of day.
        if os.path.exists(_kf_path):
            try:
                with open(_kf_path, encoding='utf-8') as _kf:
                    _strad_src = json.load(_kf)
            except Exception:
                _strad_src = rtd

    straddles = []
    live_opts_map_s = (_strad_src.get('live_options') or {}) if _strad_src else {}
    outrights_map   = (_strad_src.get('outrights')    or {}) if _strad_src else {}

    # Straddle DISPLAY filter (display + EOD only; expiry_list/CSV/skew untouched).
    # When the RTD workbook is open, show exactly the option tabs it contains
    # (live_options keys) — a tab added in Excel appears with no restart. When
    # RTD is offline, fall back to the pinned cfg['straddle_tickers'] so the
    # table never blanks. Only commodities with a 'straddle_tickers' key are
    # filtered (KC today); CT/SB/CC have no key → allow=None → unchanged.
    allow = None
    if cfg.get('straddle_tickers'):
        live_tabs = {t.upper() for t in live_opts_map_s.keys()}
        allow = live_tabs or cfg['straddle_tickers']

    for ticker in expiry_list:
        if allow is not None and ticker.upper() not in allow:
            continue
        lt    = last_trade.get(ticker)
        label = expiry_labels.get(ticker)
        if not lt:
            continue

        lo          = live_opts_map_s.get(ticker)
        out_row     = outrights_map.get(ticker, {})
        strike_list = (lo.get('strikes') or []) if lo else []

        # Forward: bid/offer mid → last → settle
        _fb = out_row.get('bid') if hasattr(out_row, 'get') else None
        _fo = out_row.get('offer') if hasattr(out_row, 'get') else None
        if _fb and _fo and _fb > 0 and _fo > 0:
            fwd = (_fb + _fo) / 2.0
        elif out_row.get('last') and out_row['last'] > 0:
            fwd = out_row['last']
        else:
            fwd = out_row.get('settle') or futures.get(ticker)

        if not fwd:
            continue

        # ATM from live strikes or CSV fallback
        avail = sorted(set(s['strike'] for s in strike_list)) if strike_list else []
        atm   = min(avail, key=lambda k: abs(k - fwd)) if avail else atm_strike.get(ticker)
        if not atm:
            continue
        # Snap ATM onto the commodity's strike grid (KC = 2.5c). Guards against an
        # integer fallback ever displaying an off-grid strike. KC-only via config.
        _inc = cfg.get('strike_increment')
        if _inc:
            atm = round(atm / _inc) * _inc

        # Live straddle value: bid/offer mid → last → settle at ATM
        today_c = today_p = None
        for s in strike_list:
            if abs(s['strike'] - atm) >= 0.01:
                continue
            bid, ask = s.get('bid'), s.get('ask')
            mid = (bid + ask) / 2.0 if (bid and ask and bid > 0 and ask > 0) else s.get('mid')
            if not (mid and mid > 0):
                mid = s.get('settle')  # deferred contracts: use RTD settle price
            if mid and mid > 0:
                if s['pc'] == 'Call': today_c = mid
                elif s['pc'] == 'Put': today_p = mid

        try:
            dte = max(0, (datetime.strptime(lt, '%Y-%m-%d') -
                          datetime.strptime(last_date, '%Y-%m-%d')).days)
        except ValueError:
            continue
        if dte <= 0:
            continue
        T = dte / 365.0

        if today_c is not None and today_p is not None:
            val = round(today_c + today_p, 2)
        else:
            iv_pct_s = atm_iv.get(ticker)
            if not iv_pct_s:
                continue
            val = round(b76_price(fwd, atm, T, RISK_FREE, iv_pct_s / 100.0, True) +
                        b76_price(fwd, atm, T, RISK_FREE, iv_pct_s / 100.0, False), 2)

        df_t    = math.exp(-RISK_FREE * T)
        call_eq = (val + (fwd - atm) * df_t) / 2.0
        iv_pct  = None
        if call_eq > 0:
            iv = implied_vol(call_eq, fwd, atm, T, RISK_FREE, True)
            if iv:
                iv_pct = round(iv * 100, 2)
        if iv_pct is None:
            iv_pct = atm_iv.get(ticker)

        # Settlement straddle from last_date CSV at live ATM strike
        prev_c = prev_p = None
        for r in comm_opts:
            if r['ticker'] != ticker or r['date'] != last_date or abs(r['strike'] - atm) >= 0.01:
                continue
            if r['pc'] == 'Call' and r['px'] > 0: prev_c = r['px']
            elif r['pc'] == 'Put'  and r['px'] > 0: prev_p = r['px']

        # Fallback: use RTD settle prices when deferred contract not in local CSV
        if prev_c is None:
            for s in strike_list:
                if abs(s['strike'] - atm) >= 0.01: continue
                if s['pc'] == 'Call' and s.get('settle') and s['settle'] > 0:
                    prev_c = s['settle']
                    break
        if prev_p is None:
            for s in strike_list:
                if abs(s['strike'] - atm) >= 0.01: continue
                if s['pc'] == 'Put' and s.get('settle') and s['settle'] > 0:
                    prev_p = s['settle']
                    break

        prev_val = round(prev_c + prev_p, 2) if (prev_c and prev_p) else None
        chg      = round(val - prev_val, 2) if prev_val is not None else None

        # Settlement IV → 1D chg_vol
        fwd_settle   = out_row.get('settle') or futures.get(ticker)
        settle_iv_pct = None
        if prev_val and fwd_settle and fwd_settle > 0 and T > 0:
            df_s = math.exp(-RISK_FREE * T)
            ceq_s = (prev_val + (fwd_settle - atm) * df_s) / 2.0
            if ceq_s > 0:
                iv_s = implied_vol(ceq_s, fwd_settle, atm, T, RISK_FREE, True)
                if iv_s:
                    settle_iv_pct = round(iv_s * 100, 2)
        chg_vol = round(iv_pct - settle_iv_pct, 2) if (iv_pct is not None and settle_iv_pct is not None) else None
        if chg_vol is not None:
            atm_iv_1d_chg[ticker] = chg_vol

        # EOD day-over-day Δ: today's SETTLE straddle vs YESTERDAY's SETTLE straddle.
        # The live `change`/`chg_vol` above compare live-vs-today's-settle, which read
        # 0.00 after the close. The EOD summary wants the overnight settle move instead.
        # Compute a settle straddle (value + IV) for a CSV date at that date's own ATM:
        # value = nearest-strike C+P from the CSV, IV via B76 at that date's forward.
        def _settle_straddle(date_str, m, y):
            d_fwd = get_hist_fwd_generic(m, y, date_str) or fwd_settle
            d_rows = opts_by_ticker_date.get((ticker, date_str), [])
            if not d_rows or not d_fwd or d_fwd <= 0:
                return None, None
            d_strikes = sorted(set(r['strike'] for r in d_rows))
            if not d_strikes:
                return None, None
            d_atm = min(d_strikes, key=lambda k: abs(k - d_fwd))
            _inc = cfg.get('strike_increment')
            if _inc:
                d_atm = round(d_atm / _inc) * _inc
            d_c = d_p = None
            for r in d_rows:
                if abs(r['strike'] - d_atm) >= 0.01 or r['px'] <= 0:
                    continue
                if r['pc'] == 'Call':
                    d_c = r['px']
                elif r['pc'] == 'Put':
                    d_p = r['px']
            if not (d_c and d_p):
                return None, None
            d_val = round(d_c + d_p, 2)
            try:
                d_dte = max(0, (datetime.strptime(lt, '%Y-%m-%d') -
                                datetime.strptime(date_str, '%Y-%m-%d')).days)
            except ValueError:
                d_dte = dte
            d_T = d_dte / 365.0 if d_dte > 0 else T
            d_iv = None
            if d_T > 0:
                d_df = math.exp(-RISK_FREE * d_T)
                d_ceq = (d_val + (d_fwd - d_atm) * d_df) / 2.0
                if d_ceq > 0:
                    _iv = implied_vol(d_ceq, d_fwd, d_atm, d_T, RISK_FREE, True)
                    if _iv:
                        d_iv = round(_iv * 100, 2)
            return d_val, d_iv

        _par = parse_generic_ticker(ticker, prefix)
        eod_today_val, eod_today_iv = prev_val, settle_iv_pct
        eod_yest_val = eod_yest_iv = None
        if _par and prev_date != last_date:
            eod_yest_val, eod_yest_iv = _settle_straddle(prev_date, _par[2], _par[1])
        eod_change = round(eod_today_val - eod_yest_val, 2) \
            if (eod_today_val is not None and eod_yest_val is not None) else None
        eod_chg_vol = round(eod_today_iv - eod_yest_iv, 2) \
            if (eod_today_iv is not None and eod_yest_iv is not None) else None

        try:
            breakeven = round(atm * (iv_pct / 100.0) / math.sqrt(252), 2) if iv_pct else None
        except (ValueError, ZeroDivisionError):
            breakeven = None

        straddles.append({
            'ticker':      ticker,
            'label':       label,
            'forward':     round(fwd, 2),
            'strike':      atm,
            'value':       val,
            'atm_vol':     iv_pct,
            'chg_vol':     chg_vol,
            'prev':        prev_val,
            'change':      chg,
            'dte':         dte,
            'expiry':      lt,
            'breakeven':   breakeven,
            'eod_change':  eod_change,
            'eod_chg_vol': eod_chg_vol,
            'eod_yest':    eod_yest_val,
        })

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
        'hv_data':        hv_data,
        'live_futures':   live_futures,
        'rtd_spreads':    rtd_spreads,
        'data_source':    data_source,
        'live_smile':     live_smile,
        'live_smile_fwd': live_smile_fwd,
        'straddles':      straddles,
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
        # Mirror cotton: read deep history from the local commodity files
        # (same files _load_generic_data uses), not the shallow GitHub feed.
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
            elif commodity == 'CC':
                exp_str = _cc_opt_expiry(mo, yr)
            elif commodity == 'SB':
                exp_str = _sb_opt_expiry(mo, yr)
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
        # Seed forward from the futures CSV (ordinal-keyed, e.g. KCJUL1). This may
        # be missing once the settle watcher starts writing ICE-code rows (KCN6)
        # for recent dates — the ordinal key no longer resolves. Do NOT abort on a
        # missing seed: the put-call parity override below derives the forward from
        # the options themselves and does not need the seed. Only abort if BOTH the
        # seed and parity fail to produce a usable forward. (Before this fix the
        # early-continue silently dropped every date after the format switch — it
        # had already clipped cotton's skew history at 2026-05-28.)
        fwd = generic_settle.get(f"{prefix}{meta['suffix']}{ordinal}", {}).get(d)
        if fwd is not None and fwd <= 0:
            fwd = None

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

        # Neither futures seed nor parity gave a forward — can't place strikes.
        if not fwd or fwd <= 0:
            continue

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

    # Rolling: constant-maturity 30-day interpolation across standard months (H/K/N/Z).
    # ATM uses variance interpolation (σ²×T); skew points use linear on vol.
    T_TARGET = 30 / 365.0
    SKEW_KEYS = ['call_10', 'call_25', 'call_35', 'put_10', 'put_25', 'put_35']
    rolling = _empty_series()
    for d in all_dates:
        candidates = []
        for ticker, date_map in ticker_date_ivs.items():
            if ticker_meta.get(ticker, {}).get('mo') not in std_months:
                continue
            ivs = date_map.get(d)
            if not ivs:
                continue
            candidates.append(ivs)
        if not candidates:
            continue
        candidates.sort(key=lambda x: x['_dte'])

        # Case 4: all expiries below 30 DTE — skip
        if candidates[-1]['_dte'] < 30:
            continue

        # Case 3: all expiries above 30 DTE — flat extrapolation, use nearest
        if candidates[0]['_dte'] > 30:
            pt = candidates[0]
            rolling['dates'].append(d)
            rolling['atm'].append(pt['atm'])
            for k in SKEW_KEYS:
                rolling[k].append(pt.get(k))
            continue

        # Case 1: exact 30 DTE match
        exact = next((c for c in candidates if c['_dte'] == 30), None)
        if exact:
            rolling['dates'].append(d)
            rolling['atm'].append(exact['atm'])
            for k in SKEW_KEYS:
                rolling[k].append(exact.get(k))
            continue

        # Case 2: T_TARGET brackets two expiries — interpolate
        short = next((c for c in reversed(candidates) if c['_dte'] < 30), None)
        long_ = next((c for c in candidates if c['_dte'] > 30), None)
        if short is None or long_ is None:
            continue
        T_s = short['_dte'] / 365.0
        T_l = long_['_dte'] / 365.0
        w = (T_TARGET - T_s) / (T_l - T_s)

        # ATM: linear interpolation on variance (σ²×T)
        atm_s = short['atm'] / 100.0
        atm_l = long_['atm'] / 100.0
        var_s = atm_s ** 2 * T_s
        var_l = atm_l ** 2 * T_l
        var_t = var_s + w * (var_l - var_s)
        atm_30 = round(math.sqrt(var_t / T_TARGET) * 100, 2)

        rolling['dates'].append(d)
        rolling['atm'].append(atm_30)

        # Skew points: linear interpolation on vol
        for k in SKEW_KEYS:
            v_s = short.get(k)
            v_l = long_.get(k)
            if v_s is not None and v_l is not None:
                rolling[k].append(round(v_s + w * (v_l - v_s), 2))
            else:
                rolling[k].append(None)

    # Per-ticker series — only active/future contracts (expiry >= today)
    today_iso = datetime.utcnow().strftime('%Y-%m-%d')
    ticker_series = {}
    for ticker, date_map in ticker_date_ivs.items():
        meta = ticker_meta.get(ticker)
        if not meta:
            continue
        if meta['expiry'] < today_iso:
            continue
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
    if _ice_rtd_reader and not _in_ct_settle_window():
        try:
            rtd = _ice_to_rtd_shape(_read_ice_workbook_safe('CT'))
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
            oi  = int(float(r.get('open_int', 0) or 0))
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
    sheet_names = []
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

@server.route('/api/skew-history')
def api_skew_history():
    commodity = request.args.get('commodity', 'CT').upper()
    if commodity not in COMMODITY_CONFIG:
        return _no_cache(jsonify({'error': f'Unknown commodity: {commodity}'}))
    return _no_cache(jsonify(compute_skew_history(commodity)))

def _assemble_eod_data(commodity='CT'):
    """
    Assembles the four-section EOD data for the given commodity (default CT).
    Straddles  — full straddle tab output (all contracts, all columns)
    Futures    — standard delivery months with all ICE RTD columns
    Spreads    — calendar spreads direct from ICE RTD (CT only; other commodities
                 return [] here and derive any spreads downstream in the appender)
    HV         — 10/30/60/90-day historical volatility
    Returns a plain dict (or {'error': ...}); shared by /api/eod-email
    and /api/save-eod-snapshot so the logic lives in one place.

    commodity-aware via parse_generic_ticker(prefix)/cfg['excl_months']; CT is
    byte-identical to the old behaviour (CT prefix + excl_months == {10}).
    """
    commodity = (commodity or 'CT').strip().upper()
    cfg    = COMMODITY_CONFIG.get(commodity, COMMODITY_CONFIG['CT'])
    prefix = cfg['prefix']
    _excl  = cfg['excl_months']
    def _parse(tkr):
        return parse_generic_ticker(tkr, prefix)
    d = load_data(commodity)
    if 'error' in d:
        return {'error': d['error']}

    # Timestamp in ET
    try:
        from zoneinfo import ZoneInfo as _ZI
    except ImportError:
        from backports.zoneinfo import ZoneInfo as _ZI
    now_et  = datetime.now(_ZI('America/New_York'))
    time_et = now_et.strftime('%H:%M ET')

    # ── Straddles: full list as shown on the straddle tab / PNG ──────────────
    straddles = d.get('straddles', [])

    # ── Futures: all contracts present in live_futures, calendar order ───────
    lf = d.get('live_futures', {})
    def _lf_sort_key(tkr):
        p = _parse(tkr)
        return (p[1], p[2]) if p else (9999, 99)
    std_tickers = sorted(
        [t for t in lf if _parse(t)
         and _parse(t)[2] not in _excl],
        key=_lf_sort_key
    )
    futures_rows = []
    for tkr in std_tickers:
        f    = lf.get(tkr, {})
        lbl  = d['expiry_labels'].get(tkr, tkr)
        stt  = f.get('settle')
        yest = f.get('yest_settle')
        chg  = f.get('change')
        pct  = f.get('pct_chg')
        # Fallback: CSV settle when RTD settle not yet available
        if stt is None:
            stt = d['futures'].get(tkr)
        # Compute pct_chg if missing but derivable
        if pct is None and chg is not None and yest and yest != 0:
            pct = round(chg / yest * 100, 2)
        futures_rows.append({
            'label':      lbl,
            'ticker':     tkr,
            'settle':     stt,
            'yest_settle': yest,
            'change':     chg,
            'pct_chg':    pct,
            'high':       f.get('high'),
            'low':        f.get('low'),
            'volume':     f.get('volume'),
            'efp_vol':    f.get('efp_vol'),
            'efs_vol':    f.get('efs_vol'),
            'block_vol':  f.get('block_vol'),
            'oi':         f.get('oi'),
            'oi_chg':     f.get('oi_chg'),
        })

    # ── Spreads: ICE RTD where available, outright-computed fallback ─────────
    # Order: consecutive standard months, then front Dec/back Dec year spread.
    # Spreads from the loader's rtd_spreads (RTD + CSV merge) plus a per-commodity
    # spreads-CSV fallback for high/low/volume. KC/SB/CC use cfg['spr_csv']; CT has
    # no spr_csv key so it falls back to LOCAL_SPR_HISTORY (unchanged behaviour).
    rtd_spr = d.get('rtd_spreads', {})
    spread_rows = []
    seen_pairs  = set()

    # Load spreads CSV — most recent row per contract for high/low/volume fallback
    _spr_csv = {}
    try:
        with open(cfg.get('spr_csv', LOCAL_SPR_HISTORY), encoding='utf-8') as _sf:
            for _sr in csv.DictReader(_sf):
                _k = (_sr.get('contract') or '').strip()
                _dt = (_sr.get('date') or '').strip()
                if _k and _dt:
                    if _k not in _spr_csv or _dt > _spr_csv[_k]['date']:
                        _spr_csv[_k] = _sr
    except Exception:
        pass

    def _computed_spread(near, far):
        """Build a spread row from outright live_futures when RTD has no spread product."""
        fn = lf.get(near, {}); ff = lf.get(far, {})
        stt_n = fn.get('settle'); stt_f = ff.get('settle')
        yst_n = fn.get('yest_settle'); yst_f = ff.get('yest_settle')
        if stt_n is None or stt_f is None:
            return None
        stt  = round(stt_n - stt_f, 4)
        yest = round(yst_n - yst_f, 4) if (yst_n is not None and yst_f is not None) else None
        chg  = round(stt - yest, 4) if yest is not None else None
        pct  = round(chg / yest * 100, 2) if (chg is not None and yest and yest != 0) else None
        p_n = _parse(near); p_f = _parse(far)
        disp = f"{MONTH_NAME[p_n[2]]}{str(p_n[1])[-2:]}/{MONTH_NAME[p_f[2]]}{str(p_f[1])[-2:]}" if p_n and p_f else f'{near}/{far}'
        return {'display': disp, 'settle': stt, 'yest_settle': yest, 'change': chg,
                'pct_chg': pct, 'high': None, 'low': None, 'volume': None,
                'block_vol': None, 'efs_vol': None, 'efp_vol': None}

    def _fv(v):
        try: return float(v) if v not in (None, '') else None
        except (ValueError, TypeError): return None

    def _row_from_csv(key, near, far):
        """Build a spread row entirely from CSV; returns None if key not in CSV."""
        csv_r = _spr_csv.get(key)
        if not csv_r:
            return None
        p_n = _parse(near); p_f = _parse(far)
        disp = (f"{MONTH_NAME[p_n[2]]}{str(p_n[1])[-2:]}/{MONTH_NAME[p_f[2]]}{str(p_f[1])[-2:]}"
                if p_n and p_f else key)
        stt  = _fv(csv_r.get('settle'))
        yest = _fv(csv_r.get('yest_settle'))
        chg  = _fv(csv_r.get('change'))
        pct  = round(chg / yest * 100, 2) if (chg is not None and yest and yest != 0) else None
        return {
            'display':    disp,
            'settle':     stt,
            'yest_settle': yest,
            'change':     chg,
            'pct_chg':    pct,
            'high':       _fv(csv_r.get('high')),
            'low':        _fv(csv_r.get('low')),
            'volume':     _fv(csv_r.get('volume')),
            'block_vol':  _fv(csv_r.get('block_vol')),
            'efs_vol':    _fv(csv_r.get('efs_vol')),
            'efp_vol':    _fv(csv_r.get('efp_vol')),
        }

    # Spread keys driven by what RTD and CSV actually contain — no synthetic
    # consecutive-pair generation. Order: near-leg calendar, then far-leg.
    # Dec/Dec year spread inserted immediately after the first Dec-near pair.
    dec_contracts = sorted(
        [t for t in lf if _parse(t) and _parse(t)[2] == 12],
        key=_lf_sort_key
    )
    year_spread_key = (f'{dec_contracts[0]}/{dec_contracts[1]}'
                       if len(dec_contracts) >= 2 else None)
    year_spread_inserted = False

    def _add_spread(key, near, far):
        row = _row_from_csv(key, near, far)
        if row is None:
            if key in rtd_spr:
                row = dict(rtd_spr[key])
            else:
                row = _computed_spread(near, far)
        if row:
            spread_rows.append(row)

    def _spread_sort_key(key):
        parts = key.split('/')
        if len(parts) != 2:
            return (9999, 99, 9999, 99)
        pn = _parse(parts[0]); pf = _parse(parts[1])
        return ((pn[1], pn[2]) if pn else (9999, 99)) + ((pf[1], pf[2]) if pf else (9999, 99))

    def _has_excluded_leg(key):
        for leg in key.split('/'):
            p = _parse(leg)
            if p and p[2] in _excl:
                return True
        return False

    all_spr_keys = sorted(
        [k for k in (set(rtd_spr.keys()) | set(_spr_csv.keys()))
         if not _has_excluded_leg(k)],
        key=_spread_sort_key
    )

    for key in all_spr_keys:
        if key in seen_pairs or key == year_spread_key:
            continue
        seen_pairs.add(key)
        parts = key.split('/')
        if len(parts) != 2:
            continue
        near, far = parts[0], parts[1]
        _add_spread(key, near, far)
        # Insert Dec/Dec year spread right after first row whose near leg is Dec
        pn = _parse(near)
        if (not year_spread_inserted and year_spread_key
                and pn and pn[2] == 12
                and year_spread_key not in seen_pairs):
            seen_pairs.add(year_spread_key)
            year_spread_inserted = True
            yd = year_spread_key.split('/')
            _add_spread(year_spread_key, yd[0], yd[1])

    # ── Historical Volatility ─────────────────────────────────────────────────
    hv_src = d.get('hv_data', {})
    lbl_map = d.get('expiry_labels', {})
    hv_rows = []
    for tkr in std_tickers:
        h = hv_src.get(tkr, {})
        if any(h.get(f'hv{w}') for w in (10, 30, 60, 90)):
            hv_rows.append({
                'label': lbl_map.get(tkr, tkr),
                'hv10':  h.get('hv10'),
                'hv30':  h.get('hv30'),
                'hv60':  h.get('hv60'),
                'hv90':  h.get('hv90'),
            })

    return {
        'date':      d.get('last_date'),
        # today_et: the live ET session date the RTD settles belong to. The EOD
        # snapshot stamps THIS (not 'date', which is the options-CSV last_date and
        # lags one business day). Email still uses 'date'. now_et computed above.
        'today_et':  now_et.strftime('%Y-%m-%d'),
        'time_et':   time_et,
        'straddles': straddles,
        'futures':   futures_rows,
        'spreads':   spread_rows,
        'hv':        hv_rows,
    }


@server.route('/api/eod-email')
def api_eod_email():
    """Returns the assembled four-section EOD data as JSON (see _assemble_eod_data)."""
    try:
        data = _assemble_eod_data()
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    if 'error' in data:
        return jsonify({'error': data['error']}), 500
    return jsonify(data)


# Destination consumed by market-intelligence/append_backfill.py (path confirmed
# by the market-intelligence dashboard, 2026-06-10). Single source of truth for
# both the manual "Save Snapshot" button and the Push-to-Site side-effect.
_EOD_SNAPSHOT_PATH = os.path.join(
    r'C:\Users\Louis\OneDrive - VLM Commodities LTD\Desktop',
    'market-intelligence', 'data', 'eod_snapshot.json'
)


def _write_eod_snapshot(data, skip_if_same_date=False, commodity='CT'):
    """Extract the compact snapshot from already-assembled EOD `data` and write it
    to the per-commodity snapshot path. Returns (extracted_dict, wrote_bool).

    This is the exact extract+write logic that used to live inline in
    api_save_eod_snapshot(); the endpoint and push_to_vlm() call it so the logic
    has one home. commodity-aware: snapshot path + standard months + output key
    prefix come from _EOD_PIPELINE / COMMODITY_CONFIG. CT is byte-identical to the
    old behaviour (prefix 'ct', std months Mar/May/Jul/Dec, eod_snapshot.json).

    skip_if_same_date: when True, if the file already exists and its 'date' equals
    data['date'], skip the write (one-write-per-date failsafe for the push path).
    The manual button passes False → behaves exactly as before (always writes).
    """
    commodity = (commodity or 'CT').strip().upper()
    cfg       = COMMODITY_CONFIG.get(commodity, COMMODITY_CONFIG['CT'])
    pfx       = commodity.lower()
    snap_path = _EOD_PIPELINE.get(commodity, _EOD_PIPELINE['CT'])['snapshot']
    # Standard delivery months for this commodity (CT: Mar/May/Jul/Dec), matched on
    # label text — the expired front carries a ticker-only label and is skipped.
    _STD_LABELS = tuple(MONTH_NAME[m] for m in sorted(cfg['std_months']))
    def _is_std(row):
        lbl = (row.get('label') or '')
        return any(m in lbl for m in _STD_LABELS)

    std_futs   = [f for f in data.get('futures', [])   if _is_std(f)]
    std_strads = [s for s in data.get('straddles', []) if _is_std(s)]
    hv         = data.get('hv', [])

    def _settle(i):
        return std_futs[i]['settle'] if i < len(std_futs) else None

    # atm_iv_30d: first standard straddle with dte >= 30; atm_vol already in %.
    atm_iv_30d = None
    for s in std_strads:
        dte = s.get('dte')
        if dte is not None and dte >= 30:
            av = s.get('atm_vol')
            atm_iv_30d = round(av, 2) if av is not None else None
            break

    hv0  = hv[0] if hv else {}
    hv30 = round(hv0['hv30'] * 100, 2) if hv0.get('hv30') is not None else None
    hv60 = round(hv0['hv60'] * 100, 2) if hv0.get('hv60') is not None else None

    extracted = {
        # today's ET session date (the date the live RTD settles belong to), not
        # data['date'] which is the options-CSV last_date and lags one business day.
        'date':           data.get('today_et') or data.get('date'),
        f'{pfx}1_settle': _settle(0),
        f'{pfx}2_settle': _settle(1),
        f'{pfx}3_settle': _settle(2),
        f'{pfx}1_ticker': std_futs[0].get('ticker') if len(std_futs) > 0 else None,
        f'{pfx}2_ticker': std_futs[1].get('ticker') if len(std_futs) > 1 else None,
        f'{pfx}3_ticker': std_futs[2].get('ticker') if len(std_futs) > 2 else None,
        'atm_iv_30d':     atm_iv_30d,
        'hv30':           hv30,
        'hv60':           hv60,
    }

    # One-write-per-date failsafe (push path only): if the file already holds this
    # date, leave it untouched.
    if skip_if_same_date and extracted.get('date') is not None:
        try:
            with open(snap_path, encoding='utf-8') as f:
                if json.load(f).get('date') == extracted['date']:
                    return extracted, False
        except (FileNotFoundError, ValueError):
            pass  # no file yet / unreadable → proceed to write

    os.makedirs(os.path.dirname(snap_path), exist_ok=True)
    with open(snap_path, 'w', encoding='utf-8') as f:
        json.dump(extracted, f, indent=2)
    return extracted, True


# append_backfill.py (market-intelligence) reads eod_snapshot.json and appends one
# EOD row to the backfill CSV. Path confirmed by the market-intelligence dashboard
# (2026-06-10). It has its own duplicate-date guard (sys.exit if the date already
# exists), so a re-run for the same date is a safe no-op.
_APPEND_BACKFILL_PY = os.path.join(
    r'C:\Users\Louis\OneDrive - VLM Commodities LTD\Desktop',
    'market-intelligence', 'append_backfill.py'
)

# Per-commodity EOD signal-backfill pipeline registry. CT is the original cotton
# pipeline (reuses the constants above); KC mirrors it with its own files so the
# cotton snapshot/appender/status are never touched. settle_status is a basename
# resolved against this app's directory.
_EOD_PIPELINE = {
    'CT': {
        'snapshot':      _EOD_SNAPSHOT_PATH,
        'append_py':     _APPEND_BACKFILL_PY,
        'settle_status': 'settle_status.json',
    },
    'KC': {
        'snapshot':      os.path.join(os.path.dirname(_EOD_SNAPSHOT_PATH), 'eod_snapshot_kc.json'),
        'append_py':     os.path.join(os.path.dirname(_APPEND_BACKFILL_PY), 'append_backfill_kc.py'),
        'settle_status': 'settle_status_kc.json',
    },
}


def _spawn_append_backfill(commodity='CT'):
    """Fire-and-forget: run the commodity's appender in a detached child process
    after a NEW-date snapshot write, so EOD backfill needs zero manual steps. Runs
    in a daemon thread (request never blocks/joins). Detached + silent on Windows
    (CREATE_NO_WINDOW, output to a log file). Any failure — missing script, error,
    timeout — is swallowed and never surfaces to the push response."""
    commodity = (commodity or 'CT').strip().upper()
    append_py = _EOD_PIPELINE.get(commodity, _EOD_PIPELINE['CT'])['append_py']
    def _run():
        import subprocess, sys
        try:
            if not os.path.exists(append_py):
                return
            _wd  = os.path.dirname(append_py)
            _log = os.path.join(_wd, os.path.splitext(os.path.basename(append_py))[0] + '_last.log')
            with open(_log, 'w', encoding='utf-8') as lf:
                subprocess.run(
                    [sys.executable, append_py],
                    cwd=_wd, stdout=lf, stderr=lf,
                    creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
                    timeout=120,
                )
        except Exception:
            pass  # best-effort; never affect the push
    threading.Thread(target=_run, daemon=True).start()


@server.route('/api/save-eod-snapshot', methods=['POST'])
def api_save_eod_snapshot():
    """
    Assemble EOD data (shared with /api/eod-email), extract a compact
    snapshot, and write it to the market-intelligence data folder.

    Produces a commodity's SIGNAL-MODEL backfill row (eod_snapshot[_kc].json ->
    append_backfill[_kc].py -> that commodity's backfill CSV). Daily RAW history
    for every commodity is already persisted automatically by
    _persist_today_generic; this endpoint is NOT that. Only commodities with a
    pipeline in _EOD_PIPELINE (CT, KC) are accepted — others are refused so a Save
    on an unsupported tab can never silently overwrite another commodity's snapshot.
    """
    commodity = (request.args.get('commodity') or 'CT').strip().upper()
    if commodity not in _EOD_PIPELINE:
        return jsonify({
            'success': False,
            'error': (f'{commodity} has no EOD signal-backfill pipeline — only '
                      f'{", ".join(sorted(_EOD_PIPELINE))} do. Its raw daily data '
                      f'is already saved automatically; nothing to snapshot here.'),
        }), 400
    try:
        data = _assemble_eod_data(commodity)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500
    if 'error' in data:
        return jsonify({'success': False, 'error': data['error']}), 500

    # Gate on OPTIONS settlement: the backfill needs settled options IV (atm_iv_30d
    # comes from the live straddle, which is non-final until options settle). Refuse
    # unless settle_status.json shows options_settled for THIS session date. The
    # settle-watcher writes options_settled=true *before* it POSTs here, so its
    # canonical EOD fire passes; this blocks only a premature manual press.
    _today_et = data.get('today_et')
    _ss_name  = _EOD_PIPELINE[commodity]['settle_status']
    try:
        with open(os.path.join(os.path.dirname(__file__), _ss_name),
                  encoding='utf-8') as _ssf:
            _ss = json.load(_ssf)
        _opts_settled = bool(_ss.get('date') == _today_et and _ss.get('options_settled'))
    except Exception:
        _opts_settled = False
    if not _opts_settled:
        return jsonify({
            'success': False,
            'error': (f'Options have not settled yet for {_today_et} — the snapshot '
                      f'would carry non-final IV. The settle-watcher fires this '
                      f'automatically once options settle; no manual save needed.'),
        }), 409

    try:
        extracted, _ = _write_eod_snapshot(data, commodity=commodity)
        # Canonical EOD writer: settle-watcher POSTs here once options-settlement is
        # confirmed, so options-date == futures-date == today (date-shift eliminated).
        # After a successful write, auto-run the appender (zero manual steps).
        # Fire-and-forget; it has its own duplicate-date guard so an unconditional
        # spawn is a safe no-op on an already-present date.
        _spawn_append_backfill(commodity)
        return jsonify({'success': True, 'path': _EOD_PIPELINE[commodity]['snapshot'], 'data': extracted})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


# Per-commodity watcher metadata: lock-file basename + settlement window (ET).
# CT values are exactly the originals (window 14:25–16:00) so cotton is unchanged.
# KC window 13:28–15:00 mirrors settle_watcher_kc.py (snapshot 13:28 → hard stop 15:00).
_WATCHER_META = {
    'CT': {'lock': 'settle_watcher.lock',    'win_start': (14, 25), 'win_end': (16, 0)},
    'KC': {'lock': 'settle_watcher_kc.lock', 'win_start': (13, 28), 'win_end': (15, 0)},
}


@server.route('/api/watcher-status', methods=['GET'])
def api_watcher_status():
    """Check whether the settle watcher for the requested commodity is running
    (via its lock file). ?commodity=CT|KC (defaults to CT)."""
    commodity = (request.args.get('commodity') or 'CT').strip().upper()
    meta = _WATCHER_META.get(commodity, _WATCHER_META['CT'])
    _lock = os.path.join(os.path.dirname(__file__), '..', 'Options_flow_analyzer', meta['lock'])
    _lock = os.path.normpath(_lock)
    running = False
    pid = None
    if os.path.exists(_lock):
        try:
            with open(_lock) as f:
                pid = int(f.read().strip())
            # os.kill(pid, 0) always raises OSError on Windows (signal 0 unsupported).
            # Use psutil when available; otherwise trust the lock file exists.
            try:
                import psutil
                running = psutil.pid_exists(pid)
            except ImportError:
                running = True  # lock file present, assume live
        except (OSError, ValueError):
            running = False  # unreadable or stale lock
    # Determine if we're in the settlement window (per-commodity, ET, trading day).
    try:
        from zoneinfo import ZoneInfo as _ZI
    except ImportError:
        from backports.zoneinfo import ZoneInfo as _ZI
    _now = datetime.now(_ZI('America/New_York'))
    _today_str = _now.strftime('%Y-%m-%d')
    _hm = (_now.hour, _now.minute)
    _in_window = (_is_ct_trading_day(_today_str)
                  and meta['win_start'] <= _hm <= meta['win_end'])
    # Both settled today → watcher has completed its job, no warning needed
    _both_settled = False
    try:
        _sp = os.path.join(os.path.dirname(__file__), _settle_status_filename(commodity))
        with open(_sp, encoding='utf-8') as _sf:
            _ss = json.load(_sf)
        _both_settled = (_ss.get('date') == _today_str
                         and bool(_ss.get('futures_settled'))
                         and bool(_ss.get('options_settled')))
    except Exception:
        pass
    return jsonify({'running': running, 'pid': pid, 'in_settle_window': _in_window,
                    'both_settled': _both_settled})


def _settle_status_filename(commodity):
    """Status-file basename per commodity. CT keeps the original name so cotton's
    behavior is byte-for-byte unchanged; KC/SB/CC use settle_status_<comm>.json
    (written by settle_watcher_<comm>.py)."""
    c = (commodity or 'CT').strip().upper()
    return 'settle_status.json' if c == 'CT' else f'settle_status_{c.lower()}.json'


@server.route('/api/settle-status', methods=['GET'])
def api_settle_status():
    """Return settlement status written by the settle watcher for the requested
    commodity (?commodity=CT|KC|SB|CC; defaults to CT). Used by the dashboard
    frontend to auto-refresh and show the settlement banner."""
    import json as _json
    commodity = (request.args.get('commodity') or 'CT').strip().upper()
    status_path = os.path.join(os.path.dirname(__file__), _settle_status_filename(commodity))
    if not os.path.exists(status_path):
        return jsonify({'futures_settled': False, 'options_settled': False,
                        'date': None, 'futures_time': None, 'options_time': None})
    try:
        with open(status_path, encoding='utf-8') as f:
            return jsonify(_json.load(f))
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@server.route('/push-to-vlm', methods=['POST'])
def push_to_vlm():
    """Proxy: forwards PNG + metadata to vlmdata.com server-side (avoids browser CORS)."""
    _VLM_PUSH_URL    = 'https://vlmdata.com/api/analysis/push'
    _VLM_PUSH_SECRET = '5c8b8dfb7aef367764d33aea1c19985a7907ae4198bc12be758a316acecabf7d'
    try:
        files = {}
        data  = {}
        if 'image' in request.files:
            f = request.files['image']
            files['image'] = (f.filename or 'export.png', f.read(), f.content_type or 'image/png')
        for key in ('type', 'rows', 'data'):
            if key in request.form:
                data[key] = request.form[key]
        resp = requests.post(_VLM_PUSH_URL,
                             headers={'x-push-secret': _VLM_PUSH_SECRET},
                             files=files, data=data, timeout=30)

        # Side-effect: on a successful upstream push, also write the CT EOD
        # snapshot that market-intelligence/append_backfill.py consumes. This is
        # the reliable daily writer (the manual "Save Snapshot" step gets missed).
        # CT-only by construction (only the CT tab reaches this endpoint; KC/SB/CC
        # use a separate button). One-write-per-date via skip_if_same_date. Fully
        # isolated: any failure here must NEVER break or delay the push response.
        if 200 <= resp.status_code < 300:
            try:
                _eod_data = _assemble_eod_data()
                if 'error' not in _eod_data:
                    _extracted, _wrote = _write_eod_snapshot(_eod_data, skip_if_same_date=True)
                    # New date actually written → auto-run append_backfill.py (zero
                    # manual steps). Skipped on same-date (_wrote False) or error.
                    # Fire-and-forget; never blocks/breaks the push response.
                    if _wrote:
                        _spawn_append_backfill()
            except Exception:
                pass  # snapshot is best-effort; never affect the push result

        return jsonify({}), resp.status_code
    except Exception as e:
        return jsonify({'error': str(e)}), 502


_schedule_preclose_flush()


# ── Options Flow Pipeline ─────────────────────────────────────────────────────

_FLOW_BASE = os.path.join(
    os.path.dirname(__file__), '..', 'Options_flow_analyzer'
)
_FLOW_PROC = os.path.join(_FLOW_BASE, 'processed')


def _flow_daily_path(date_str):
    return os.path.join(_FLOW_PROC, date_str, 'enriched.csv')

def _flow_legs_path(date_str):
    return os.path.join(_FLOW_PROC, date_str, 'enriched_legs.json')

def _flow_flags_path(date_str):
    return os.path.join(_FLOW_PROC, date_str, 'flags.txt')

def _flow_weekly_dir(week_ending):
    return os.path.join(_FLOW_PROC, f'week-ending-{week_ending}')


def _read_flow_csv(path):
    """Read an enriched.csv or strike_flow_weekly.csv into a list of dicts."""
    if not os.path.exists(path):
        return []
    with open(path, newline='', encoding='utf-8') as fh:
        return list(csv.DictReader(fh))


def _coerce_flow_row(row):
    """Type-coerce string fields from the CSV into native Python types."""
    for int_field in ('qty', 'blk_vol', 'exch_vol'):
        if row.get(int_field) not in (None, ''):
            try:
                row[int_field] = int(row[int_field])
            except (ValueError, TypeError):
                row[int_field] = None
    for float_field in ('price_lo', 'price_hi', 'underlying_lo', 'underlying_hi',
                        'high', 'low', 'prev_settle'):
        if row.get(float_field) not in (None, ''):
            try:
                row[float_field] = float(row[float_field])
            except (ValueError, TypeError):
                row[float_field] = None
    if row.get('enrichment_match') not in (None, ''):
        row['enrichment_match'] = str(row['enrichment_match']).lower() == 'true'
    return row


@server.route('/api/flow/daily')
def api_flow_daily():
    """
    Return enriched flow rows for a single trading day.

    Query params:
        date      YYYY-MM-DD (required)
        commodity CT (default, reserved for future multi-commodity support)

    Response:
        {
          date: str,
          rows: [...],          # enriched.csv rows, types coerced
          flags: [...],         # flags.txt lines
          legs_available: bool  # whether enriched_legs.json exists
        }
    """
    date_str  = request.args.get('date', '')
    if not date_str:
        return jsonify({'error': 'date parameter required (YYYY-MM-DD)'}), 400

    csv_path   = _flow_daily_path(date_str)
    flags_path = _flow_flags_path(date_str)
    legs_path  = _flow_legs_path(date_str)

    rows = [_coerce_flow_row(r) for r in _read_flow_csv(csv_path)]

    flags = []
    if os.path.exists(flags_path):
        with open(flags_path, encoding='utf-8') as fh:
            flags = [l.rstrip('\n') for l in fh if l.strip()]

    return _no_cache(jsonify({
        'date':           date_str,
        'rows':           rows,
        'flags':          flags,
        'legs_available': os.path.exists(legs_path),
    }))


@server.route('/api/flow/weekly')
def api_flow_weekly():
    """
    Return aggregated weekly flow data.

    Query params:
        week_ending  YYYY-MM-DD Friday date (required)

    Response:
        {
          week_ending: str,
          rows:        [...],   # weekly_enriched.csv rows
          strike_rows: [...],   # strike_flow_weekly.csv rows
          flags:       [...]    # flags.txt lines
        }
    """
    week_ending = request.args.get('week_ending', '')
    if not week_ending:
        return jsonify({'error': 'week_ending parameter required (YYYY-MM-DD)'}), 400

    week_dir     = _flow_weekly_dir(week_ending)
    enriched_csv = os.path.join(week_dir, 'weekly_enriched.csv')
    strike_csv   = os.path.join(week_dir, 'strike_flow_weekly.csv')
    flags_file   = os.path.join(week_dir, 'flags.txt')

    rows        = [_coerce_flow_row(r) for r in _read_flow_csv(enriched_csv)]
    strike_rows = _read_flow_csv(strike_csv)   # strike rows stay as strings (all numeric)

    for sr in strike_rows:
        for f in ('strike', 'total_contracts', 'trade_count'):
            if sr.get(f) not in (None, ''):
                try:
                    sr[f] = float(sr[f]) if f == 'strike' else int(sr[f])
                except (ValueError, TypeError):
                    pass

    flags = []
    if os.path.exists(flags_file):
        with open(flags_file, encoding='utf-8') as fh:
            flags = [l.rstrip('\n') for l in fh if l.strip()]

    return _no_cache(jsonify({
        'week_ending': week_ending,
        'rows':        rows,
        'strike_rows': strike_rows,
        'flags':       flags,
    }))


@server.route('/api/draft-eod-email', methods=['POST'])
def api_draft_eod_email():
    """Open an Outlook desktop draft with the EOD PNG embedded inline (Windows only)."""
    try:
        import win32com.client
        import pythoncom
        import base64
        import tempfile
        import threading
        import time
    except ImportError:
        return jsonify({'error': 'win32com not available'}), 500

    body = request.get_json(silent=True) or {}
    subject  = body.get('subject', 'VLM Cotton EOD Summary')
    date_str = body.get('date_str', '')
    png_b64  = body.get('png_b64', '')
    if not png_b64:
        return jsonify({'error': 'No PNG data'}), 400

    if ',' in png_b64:
        png_b64 = png_b64.split(',', 1)[1]
    try:
        png_bytes = base64.b64decode(png_b64)
    except Exception as e:
        return jsonify({'error': f'Bad PNG data: {e}'}), 400

    tmp_path = None
    err_box  = []

    try:
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tf:
            tmp_path = tf.name
            tf.write(png_bytes)
    except Exception as e:
        return jsonify({'error': f'Temp file: {e}'}), 500

    html_body = (
        f'<html><body style="font-family:Arial,sans-serif;margin:0;padding:12px">'
        f'<p style="margin:0 0 10px 0">ICE Cotton No. 2 &mdash; End of Day Summary<br>'
        f'Date: {date_str}</p>'
        f'<img src="cid:eod_png" style="max-width:100%">'
        f'</body></html>'
    )

    def _open_draft(subj, body_html, png_path, errors):
        # COM must be initialised per-thread; do NOT call CoUninitialize until
        # after Outlook has finished creating the inspector window.
        pythoncom.CoInitialize()
        try:
            ol = win32com.client.Dispatch('Outlook.Application')
            mail = ol.CreateItem(0)  # olMailItem
            mail.Subject = subj
            mail.HTMLBody = body_html
            att = mail.Attachments.Add(png_path)
            att.PropertyAccessor.SetProperty(
                'http://schemas.microsoft.com/mapi/proptag/0x3712001F', 'eod_png')
            mail.Display(False)
            time.sleep(3)   # keep COM apartment alive while Outlook loads the window
        except Exception as exc:
            errors.append(str(exc))
        finally:
            pythoncom.CoUninitialize()
            try:
                os.unlink(png_path)
            except Exception:
                pass

    t = threading.Thread(target=_open_draft,
                         args=(subject, html_body, tmp_path, err_box),
                         daemon=True)
    t.start()
    t.join(timeout=12)   # wait up to 12 s for the window to open

    if err_box:
        return jsonify({'error': err_box[0]}), 500
    return jsonify({'ok': True})


if __name__ == '__main__':
    server.run(debug=True, port=5050, use_reloader=True, reloader_type='stat', threaded=True)
