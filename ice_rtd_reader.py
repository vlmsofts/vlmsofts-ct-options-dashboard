"""
ice_rtd_reader.py — reads ICE RTD Excel workbooks via xlwings
=============================================================
Standalone module.  No app.py coupling.  New parallel CT pipeline.

The existing CT pipeline (load_data / local_options_history.csv) is
untouched.  This module is tested in isolation via test_ice_pipeline.py
before any integration work begins.

Workbook location:
  C:\\Users\\Louis\\OneDrive - VLM Commodities LTD\\Site Sync\\ICE RTD FEED CT.xlsx
  (and ICE RTD FEED KC.xlsx / SB.xlsx / CC.xlsx when built)

The workbook must be open in Excel with the ICE RTD feed active.
If Excel is not running or the workbook is not open, all read functions
return None and the caller should fall back to the Bloomberg CSV pipeline.

Return shape of read_ice_workbook():
{
  'mode':    'live' | 'today_settle' | 'prior_settle' | 'unavailable',
  'futures': {
      'CTN6': {'settle': 81.60, 'last': 78.83, 'bid': 78.81,
               'offer': 78.83, 'oi': 136875, 'market_state': 'Open'},
      ...
  },
  'options': {
      'CTN6': [
          {'strike': 78.0,
           'call_bid': 1.73, 'call_offer': 1.83, 'call_last': None,
           'call_settle': 0.92, 'call_oi': 5025,
           'put_bid':  2.25, 'put_offer': 2.35, 'put_last': 2.21,
           'put_settle': 1.24, 'put_oi': 4790},
          ...
      ],
      ...
  }
}
"""

import math
import os

try:
    import xlwings as xw
    _XW_OK = True
except ImportError:
    _XW_OK = False


# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

SITE_SYNC = r'C:\Users\Louis\OneDrive - VLM Commodities LTD\Site Sync'

WB_NAME = {
    'CT': 'ICE RTD FEED CT.xlsx',
    'KC': 'ICE RTD FEED KC.xlsx',
    'SB': 'ICE RTD FEED SB.xlsx',
    'CC': 'ICE RTD FEED CC.xlsx',
}

_MON_NUM = {
    'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,
    'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12
}
_MON_CODE = {
    1:'F',2:'G',3:'H',4:'J',5:'K',6:'M',
    7:'N',8:'Q',9:'U',10:'V',11:'X',12:'Z'
}
# Reverse map: month code → month number (for sort key)
_CODE_MON = {v: k for k, v in _MON_CODE.items()}

# ── Option sheet column positions (0-based, positional — header has duplicates)
# Header: Qty Bid Offer Qty Last Volume OptBlock Settlement OI Strike
#         Qty Bid Offer Qty Last Volume OptBlock Settlement OI
_C_BID    =  1
_C_OFFER  =  2
_C_LAST   =  4
_C_VOL    =  5   # call exchange volume (used by options-flow pipeline)
_C_BLOCK  =  6   # call block volume    (used by options-flow pipeline)
_C_SETTLE =  7
_C_OI     =  8
_STRIKE   =  9
_P_BID    = 11
_P_OFFER  = 12
_P_LAST   = 14
_P_VOL    = 15   # put exchange volume  (used by options-flow pipeline)
_P_BLOCK  = 16   # put block volume     (used by options-flow pipeline)
_P_SETTLE = 17
_P_OI     = 18

# ── Futures sheet column positions (confirmed 2026-05-27 from live workbook inspection)
_FUT_COLS = {
    'Strip':        2,
    'bid':          6,
    'offer':        7,
    'Last Price':   9,
    'Vol':         13,
    'High':        14,
    'Low':         15,
    'Settle':      17,
    'Change':      18,
    'Market State': 19,
    'Block Vol':   34,
    'EFS Vol':     35,
    'EFP Vol':     36,
    'OI':          46,
}


# ─────────────────────────────────────────────────────────────────────────────
# WORKBOOK DISCOVERY
# ─────────────────────────────────────────────────────────────────────────────

def _find_workbook(wb_filename):
    if not _XW_OK:
        return None
    try:
        for app in xw.apps:
            for wb in app.books:
                if wb.name.lower() == wb_filename.lower():
                    return wb
    except Exception:
        pass
    return None


def open_workbook(commodity='CT'):
    name = WB_NAME.get(commodity.upper())
    if not name:
        return None
    return _find_workbook(name)


# ─────────────────────────────────────────────────────────────────────────────
# CONTRACT CODE HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _strip_to_contract(strip_str, prefix):
    """
    "Jul26" → "CTN6",  "Dec27" → "CTZ7",  "Mar27" → "CTH7"
    """
    if not strip_str or not isinstance(strip_str, str):
        return None
    s = strip_str.strip()
    if len(s) < 4:
        return None
    mon_str  = s[:3]
    yr_digit = s[-1]
    month_num = _MON_NUM.get(mon_str)
    if month_num is None:
        return None
    mc = _MON_CODE[month_num]
    return f"{prefix.upper()}{mc}{yr_digit}"


def _contract_sort_key(contract_code, prefix):
    """
    Sort contracts by implied expiry.
    CTN6 → (6, 7),  CTZ6 → (6, 12),  CTH7 → (7, 3)
    Year digit as int, then month number — works for 2024-2029.
    """
    code = contract_code[len(prefix):]   # e.g. "N6"
    if len(code) < 2:
        return (99, 99)
    mc = code[0]
    yr = int(code[1]) if code[1].isdigit() else 99
    mon = _CODE_MON.get(mc, 99)
    return (yr, mon)


def get_option_sheets(wb, prefix):
    """
    Returns option sheet names sorted by contract expiry (nearest first).
    e.g. ['CTN6', 'CTU6', 'CTZ6', 'CTF7', 'CTH7']
    """
    result = []
    plen = len(prefix)
    for sh in wb.sheets:
        name = sh.name.strip()
        if (name.upper().startswith(prefix.upper())
                and len(name) == plen + 2
                and name[plen:].isalnum()):
            result.append(name.upper())
    result.sort(key=lambda c: _contract_sort_key(c, prefix))
    return result


def front_month(futures_dict, prefix):
    """
    Returns the contract code of the nearest-expiry contract that has
    live quotes (bid or offer present).  Falls back to nearest with any settle.
    """
    if not futures_dict:
        return None
    # Sort all contracts by expiry
    sorted_contracts = sorted(futures_dict.keys(),
                               key=lambda c: _contract_sort_key(c, prefix))
    # Prefer front contract with live quotes
    for c in sorted_contracts:
        f = futures_dict[c]
        if f.get('bid') is not None or f.get('offer') is not None:
            return c
    # Fallback: nearest with settle
    for c in sorted_contracts:
        if futures_dict[c].get('settle') is not None:
            return c
    return sorted_contracts[0] if sorted_contracts else None


def get_forward(futures_data, mode):
    """
    Returns the appropriate forward price for IV calculation.
    Live mode  : (bid+offer)/2 → last → settle
    Settle mode: settle only
    """
    if futures_data is None:
        return None
    if mode == 'live':
        bid   = futures_data.get('bid')
        offer = futures_data.get('offer')
        if bid is not None and offer is not None:
            return (bid + offer) / 2.0
        last = futures_data.get('last')
        if last is not None:
            return last
    return futures_data.get('settle')


# ─────────────────────────────────────────────────────────────────────────────
# UTILITIES
# ─────────────────────────────────────────────────────────────────────────────

def atm_strike(futures_price):
    """
    Nearest whole-cent strike.
    80.25 → 80,  80.51 → 81,  80.50 → 81  (standard arithmetic rounding)
    """
    if futures_price is None:
        return None
    return int(math.floor(futures_price + 0.5))


def _safe_float(val):
    if val is None:
        return None
    try:
        f = float(val)
        return f if math.isfinite(f) else None
    except (TypeError, ValueError):
        return None


# ─────────────────────────────────────────────────────────────────────────────
# FUTURES READER
# ─────────────────────────────────────────────────────────────────────────────

def read_futures(wb, prefix='CT'):
    """
    Reads the '[PREFIX] Futures' sheet.
    Returns dict keyed by contract code.
    Only plain futures rows (no spreads, no CSOs).
    """
    sheet_name = f'{prefix.upper()} Futures'
    sh = None
    try:
        sh = wb.sheets[sheet_name]
    except Exception:
        for s in wb.sheets:
            if s.name.strip().lower() == sheet_name.lower():
                sh = s
                break
    if sh is None:
        return {}

    data = sh.used_range.value
    if not data or len(data) < 2:
        return {}

    header = [str(h).strip() if h is not None else '' for h in data[0]]

    def _col(name, fallback):
        try:
            return header.index(name)
        except ValueError:
            return fallback

    col_strip  = _col('Strip',        _FUT_COLS['Strip'])
    col_last   = _col('Last Price',   _FUT_COLS['Last Price'])
    col_vol    = _col('Vol',          _FUT_COLS['Vol'])
    col_high   = _col('High',         _FUT_COLS['High'])
    col_low    = _col('Low',          _FUT_COLS['Low'])
    col_settle = _col('Settle',       _FUT_COLS['Settle'])
    col_change = _col('Change',       _FUT_COLS['Change'])
    col_mstate = _col('Market State', _FUT_COLS['Market State'])
    col_block  = _col('Block Vol',    _FUT_COLS['Block Vol'])
    col_efs    = _col('EFS Vol',      _FUT_COLS['EFS Vol'])
    col_efp    = _col('EFP Vol',      _FUT_COLS['EFP Vol'])
    col_oi     = _col('OI',           _FUT_COLS['OI'])
    col_bid    = _FUT_COLS['bid']
    col_offer  = _FUT_COLS['offer']

    result = {}
    for row in data[1:]:
        if not row or len(row) <= col_settle:
            continue
        product = str(row[0]).strip() if row[0] else ''
        # Only plain futures — skip spreads (strip contains '/') and CSOs
        if 'Futures' not in product:
            continue
        strip_val = row[col_strip] if len(row) > col_strip else None
        if strip_val and '/' in str(strip_val):
            continue

        contract = _strip_to_contract(strip_val, prefix)
        if not contract:
            continue

        # Avoid overwriting with a duplicate (spreads with same strip may appear)
        if contract in result:
            continue

        stt = _safe_float(row[col_settle]) if len(row) > col_settle else None
        chg = _safe_float(row[col_change]) if len(row) > col_change else None
        yest = round(stt - chg, 4) if (stt is not None and chg is not None) else None
        pct  = round(chg / yest * 100, 2) if (chg is not None and yest and yest != 0) else None

        result[contract] = {
            'settle':       stt,
            'last':         _safe_float(row[col_last])   if len(row) > col_last   else None,
            'bid':          _safe_float(row[col_bid])    if len(row) > col_bid    else None,
            'offer':        _safe_float(row[col_offer])  if len(row) > col_offer  else None,
            'volume':       _safe_float(row[col_vol])    if len(row) > col_vol    else None,
            'high':         _safe_float(row[col_high])   if len(row) > col_high   else None,
            'low':          _safe_float(row[col_low])    if len(row) > col_low    else None,
            'change':       chg,
            'pct_chg':      pct,
            'yest_settle':  yest,
            'block_vol':    _safe_float(row[col_block])  if len(row) > col_block  else None,
            'efs_vol':      _safe_float(row[col_efs])    if len(row) > col_efs    else None,
            'efp_vol':      _safe_float(row[col_efp])    if len(row) > col_efp    else None,
            'oi':           _safe_float(row[col_oi])     if len(row) > col_oi     else None,
            'market_state': str(row[col_mstate]).strip() if len(row) > col_mstate and row[col_mstate] else None,
        }

    return result


def read_spreads(wb, prefix='CT'):
    """
    Reads calendar spread rows ('Cotton No. 2 Spr') from the '[PREFIX] Futures' sheet.
    Returns dict keyed by contract pair e.g. 'CTN6/CTZ6'.
    Skips TAS spreads.
    """
    sheet_name = f'{prefix.upper()} Futures'
    sh = None
    try:
        sh = wb.sheets[sheet_name]
    except Exception:
        for s in wb.sheets:
            if s.name.strip().lower() == sheet_name.lower():
                sh = s
                break
    if sh is None:
        return {}

    data = sh.used_range.value
    if not data or len(data) < 2:
        return {}

    result = {}
    for row in data[1:]:
        if not row or len(row) <= _FUT_COLS['Settle']:
            continue
        product = str(row[0]).strip() if row[0] else ''
        if 'Spr' not in product or 'TAS' in product:
            continue

        strip_val = row[_FUT_COLS['Strip']] if len(row) > _FUT_COLS['Strip'] else None
        if not strip_val or '/' not in str(strip_val):
            continue

        parts = str(strip_val).strip().split('/')
        if len(parts) != 2:
            continue

        leg1 = _strip_to_contract(parts[0].strip(), prefix)
        leg2 = _strip_to_contract(parts[1].strip(), prefix)
        if not leg1 or not leg2:
            continue

        key = f'{leg1}/{leg2}'
        if key in result:
            continue

        stt = _safe_float(row[_FUT_COLS['Settle']]) if len(row) > _FUT_COLS['Settle'] else None
        chg = _safe_float(row[_FUT_COLS['Change']]) if len(row) > _FUT_COLS['Change'] else None
        yest = round(stt - chg, 4) if (stt is not None and chg is not None) else None
        pct  = round(chg / yest * 100, 2) if (chg is not None and yest and yest != 0) else None

        result[key] = {
            'display':    f"{parts[0].strip()}/{parts[1].strip()}",
            'settle':     stt,
            'last':       _safe_float(row[_FUT_COLS['Last Price']]) if len(row) > _FUT_COLS['Last Price'] else None,
            'change':     chg,
            'pct_chg':    pct,
            'yest_settle': yest,
            'high':       _safe_float(row[_FUT_COLS['High']])      if len(row) > _FUT_COLS['High']      else None,
            'low':        _safe_float(row[_FUT_COLS['Low']])       if len(row) > _FUT_COLS['Low']       else None,
            'volume':     _safe_float(row[_FUT_COLS['Vol']])        if len(row) > _FUT_COLS['Vol']       else None,
            'block_vol':  _safe_float(row[_FUT_COLS['Block Vol']]) if len(row) > _FUT_COLS['Block Vol'] else None,
            'efs_vol':    _safe_float(row[_FUT_COLS['EFS Vol']])   if len(row) > _FUT_COLS['EFS Vol']   else None,
            'efp_vol':    _safe_float(row[_FUT_COLS['EFP Vol']])   if len(row) > _FUT_COLS['EFP Vol']   else None,
        }

    return result


# ─────────────────────────────────────────────────────────────────────────────
# OPTIONS READER
# ─────────────────────────────────────────────────────────────────────────────

def read_options(wb, sheet_name):
    """
    Reads one option chain sheet.
    Returns list of row dicts sorted ascending by strike.
    Rows with no strike value are skipped.
    """
    try:
        sh = wb.sheets[sheet_name]
    except Exception:
        return []

    data = sh.used_range.value
    if not data or len(data) < 2:
        return []

    rows = []
    for row in data[1:]:
        if not row or len(row) <= _P_OI:
            continue
        strike = _safe_float(row[_STRIKE])
        if strike is None:
            continue
        rows.append({
            'strike':      strike,
            'call_bid':    _safe_float(row[_C_BID]),
            'call_offer':  _safe_float(row[_C_OFFER]),
            'call_last':   _safe_float(row[_C_LAST]),
            'call_vol':    _safe_float(row[_C_VOL])    if len(row) > _C_VOL    else None,
            'call_block':  _safe_float(row[_C_BLOCK])  if len(row) > _C_BLOCK  else None,
            'call_settle': _safe_float(row[_C_SETTLE]),
            'call_oi':     _safe_float(row[_C_OI]),
            'put_bid':     _safe_float(row[_P_BID]),
            'put_offer':   _safe_float(row[_P_OFFER]),
            'put_last':    _safe_float(row[_P_LAST]),
            'put_vol':     _safe_float(row[_P_VOL])    if len(row) > _P_VOL    else None,
            'put_block':   _safe_float(row[_P_BLOCK])  if len(row) > _P_BLOCK  else None,
            'put_settle':  _safe_float(row[_P_SETTLE]),
            'put_oi':      _safe_float(row[_P_OI]),
        })

    rows.sort(key=lambda r: r['strike'])
    return rows


# ─────────────────────────────────────────────────────────────────────────────
# MODE DETECTION
# ─────────────────────────────────────────────────────────────────────────────

def detect_mode(wb, prefix, stored_atm_settle=None):
    """
    'live'         — market open (Market State == 'Open' on front month)
    'today_settle' — market closed, today's settlements published
                     (ATM settle differs from stored Bloomberg value)
    'prior_settle' — market closed, using yesterday's settlements
    """
    futures = read_futures(wb, prefix)
    if not futures:
        return 'prior_settle'

    fm = front_month(futures, prefix)
    if fm is None:
        return 'prior_settle'

    if futures[fm].get('market_state') == 'Open':
        return 'live'

    # Market closed — check whether today's settlements have been published
    if stored_atm_settle is None:
        return 'prior_settle'

    option_sheets = get_option_sheets(wb, prefix)
    if not option_sheets:
        return 'prior_settle'

    fwd = futures.get(option_sheets[0], {}).get('settle')
    if fwd is None:
        return 'prior_settle'

    atm = atm_strike(fwd)
    options = read_options(wb, option_sheets[0])
    atm_row = next((r for r in options if r['strike'] == atm), None)
    if atm_row is None:
        return 'prior_settle'

    ice_settle = atm_row.get('call_settle')
    if ice_settle is not None and abs(ice_settle - stored_atm_settle) > 0.001:
        return 'today_settle'

    return 'prior_settle'


# ─────────────────────────────────────────────────────────────────────────────
# TOP-LEVEL READ
# ─────────────────────────────────────────────────────────────────────────────

def read_ice_workbook(commodity='CT', stored_atm_settle=None):
    """
    Main entry point — thread-safe (initialises COM for the calling thread).

    commodity         : 'CT', 'KC', 'SB', or 'CC'
    stored_atm_settle : ATM call settle from Bloomberg CSV (for mode detection)

    Returns structured dict.  If workbook unavailable: {'mode': 'unavailable'}.
    """
    try:
        import pythoncom
        pythoncom.CoInitialize()
        try:
            return _read_ice_workbook_inner(commodity, stored_atm_settle)
        finally:
            pythoncom.CoUninitialize()
    except ImportError:
        # pythoncom not available — call directly (non-Flask / non-Windows contexts)
        return _read_ice_workbook_inner(commodity, stored_atm_settle)


def _read_ice_workbook_inner(commodity='CT', stored_atm_settle=None):
    wb = open_workbook(commodity)
    if wb is None:
        return {'mode': 'unavailable'}

    prefix  = commodity.upper()
    futures = read_futures(wb, prefix)
    spreads = read_spreads(wb, prefix)

    option_sheets = get_option_sheets(wb, prefix)
    options = {}
    for sheet_name in option_sheets:
        rows = read_options(wb, sheet_name)
        if rows:
            options[sheet_name.upper()] = rows

    mode = detect_mode(wb, prefix, stored_atm_settle)

    return {
        'mode':    mode,
        'futures': futures,
        'spreads': spreads,
        'options': options,
    }


# ─────────────────────────────────────────────────────────────────────────────
# ICE CONNECT API SOURCE (Phase 1 — futures only)  ·  ON-DEMAND
# ─────────────────────────────────────────────────────────────────────────────
# read_ice_api() returns the SAME {mode, futures, spreads, options} shape as
# read_ice_workbook(), so it drops in transparently behind
# _read_ice_workbook_safe / _ice_to_rtd_shape. Phase 1 fills 'futures' only.
#
# NO BACKGROUND LOOP. The pull happens ON DEMAND, exactly when the dashboard
# loads data (manual refresh, option-variable input, settlement reload) — the
# moments the futures forward actually needs to be fresh. read_ice_api spawns the
# standalone 32-bit producer as a subprocess when the cached JSON is missing or
# older than the coalescing TTL, waits briefly, then reads the result. ICE
# Connect is sub-second and stable, so the ~1s on-demand pull is invisible; a
# blind background loop would re-pull all day to serve a page that refreshes a
# handful of times. Fewer moving parts: no loop process to babysit.
#
# This stays Excel-free → immune to the corrupt-strike / COM-wedge failure modes.
# EOD usability (straddle freeze rtd_snap.json @14:16, settle_watcher CSVs, the
# 4:30 settled_surface, skew history, surfaces) does NOT depend on this producer —
# those read settled CSVs and only get MORE robust on the API.

import time as _time
import subprocess as _subprocess

# Default location the producer writes to (dashboard repo \api_feed).
API_FEED_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'api_feed')

# Producer script + the 32-bit interpreter that runs it (icepython is 32-bit
# main-thread-only; this dashboard is 64-bit and cannot host it in-process).
# Both env-overridable so the paths aren't hard-locked to one machine.
PRODUCER_SCRIPT = os.getenv(
    'ICE_PRODUCER_SCRIPT', r'C:\Ice eod records\dashboard_futures_producer.py')
PRODUCER_PY = os.getenv('ICE_PRODUCER_PY', 'py')
PRODUCER_PY_ARG = os.getenv('ICE_PRODUCER_PY_ARG', '-3.13-32')

# Coalescing TTL — reuse the existing JSON if it was written within this many
# seconds, else spawn a fresh pull. This is NOT a background cadence; it only
# stops a single page load (which can call read_ice_api more than once) from
# spawning the producer twice. A human refreshing won't out-pace it.
API_FRESH_TTL_SEC = int(os.getenv('ICE_API_FRESH_TTL_SEC', '60'))

# Hard timeout on the producer subprocess so a stuck pull never wedges a request.
PRODUCER_TIMEOUT_SEC = int(os.getenv('ICE_PRODUCER_TIMEOUT_SEC', '8'))

# Serve-staleness ceiling — even after a refresh attempt, never SERVE a JSON
# older than this. If the producer failed or stood down (14:18-16:30) and the
# last-good file is older than this, return unavailable → the dashboard falls
# back to Excel/CSV (correct during the stand-down: the freeze + settle_watcher
# own that window). Must be > API_FRESH_TTL_SEC.
API_MAX_SERVE_AGE_SEC = int(os.getenv('ICE_API_MAX_SERVE_AGE_SEC', '180'))


def _json_age(path):
    """Seconds since the JSON was last written, or None if it doesn't exist."""
    try:
        return _time.time() - os.stat(path).st_mtime
    except OSError:
        return None


def _spawn_producer(commodity):
    """Run the 32-bit producer once to refresh the JSON. Best-effort: any failure
    (ICE XL down, stand-down no-op, timeout) is swallowed — the caller then reads
    whatever last-good file exists, or returns unavailable and the dashboard falls
    back to Excel/CSV. The producer self-guards the 14:18-16:30 COM stand-down."""
    try:
        _subprocess.run(
            [PRODUCER_PY, PRODUCER_PY_ARG, PRODUCER_SCRIPT, '--commodity', commodity.upper()],
            timeout=PRODUCER_TIMEOUT_SEC,
            capture_output=True,
            check=False,
        )
    except Exception:
        pass  # leave the last-good JSON in place for the caller to read


def read_ice_api(commodity='CT'):
    """Return {mode, futures, spreads, options} from the producer JSON, or
    {'mode': 'unavailable'} if no usable file can be produced.

    ON DEMAND: if the cached JSON is missing or older than API_FRESH_TTL_SEC,
    spawn the 32-bit producer to refresh it, then read. Mirrors
    read_ice_workbook()'s contract exactly so it is a drop-in source.
    Phase 1: futures only (spreads/options empty)."""
    import json as _json
    path = os.path.join(API_FEED_DIR, f'futures_api_{commodity.upper()}.json')

    # On-demand refresh: pull only when the cached file is missing or stale.
    age = _json_age(path)
    if age is None or age > API_FRESH_TTL_SEC:
        _spawn_producer(commodity)
        age = _json_age(path)  # re-stat after the spawn

    # If still no file (producer never ran, or stood down with no prior file),
    # or the best file we have is older than the serve ceiling (producer failed /
    # stood down and last-good is stale), serve nothing → dashboard falls back to
    # Excel/CSV. A dead producer must never look like a live feed.
    if age is None or age > API_MAX_SERVE_AGE_SEC:
        return {'mode': 'unavailable'}

    try:
        with open(path, encoding='utf-8') as f:
            data = _json.load(f)
    except Exception:
        return {'mode': 'unavailable'}

    futures = data.get('outrights') or {}
    if not futures:
        return {'mode': 'unavailable'}
    mode = data.get('mode') or 'live'
    if mode not in ('live', 'today_settle', 'prior_settle'):
        mode = 'live'
    # Phase 2: the producer now also emits live option chains keyed by contract
    # code, in the per-strike row shape _ice_to_rtd_shape expects. Pass them
    # through (empty {} if the producer ran with --no-options). Drives the live
    # smile / ATM vol / live straddles.
    return {
        'mode':    mode,
        'futures': futures,   # producer dict already carries the per-contract keys
        'spreads': {},        # Phase 3 (calendar spreads)
        'options': data.get('options') or {},
    }
