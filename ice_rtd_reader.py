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
_C_SETTLE =  7
_C_OI     =  8
_STRIKE   =  9
_P_BID    = 11
_P_OFFER  = 12
_P_LAST   = 14
_P_SETTLE = 17
_P_OI     = 18

# ── Futures sheet fallback column positions (confirmed from workbook inspection)
_FUT_COLS = {
    'Strip':        2,
    'bid':          6,
    'offer':        7,
    'Last Price':   9,
    'Settle':       17,
    'Market State': 19,
    'OI':           46,
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
    col_settle = _col('Settle',       _FUT_COLS['Settle'])
    col_mstate = _col('Market State', _FUT_COLS['Market State'])
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

        result[contract] = {
            'settle':       _safe_float(row[col_settle]) if len(row) > col_settle else None,
            'last':         _safe_float(row[col_last])   if len(row) > col_last   else None,
            'bid':          _safe_float(row[col_bid])    if len(row) > col_bid    else None,
            'offer':        _safe_float(row[col_offer])  if len(row) > col_offer  else None,
            'oi':           _safe_float(row[col_oi])     if len(row) > col_oi     else None,
            'market_state': str(row[col_mstate]).strip() if len(row) > col_mstate and row[col_mstate] else None,
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
            'call_settle': _safe_float(row[_C_SETTLE]),
            'call_oi':     _safe_float(row[_C_OI]),
            'put_bid':     _safe_float(row[_P_BID]),
            'put_offer':   _safe_float(row[_P_OFFER]),
            'put_last':    _safe_float(row[_P_LAST]),
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
        'options': options,
    }
