from flask import Flask, render_template, jsonify
import requests, csv, io, json, time, math
from datetime import datetime, timedelta

server = Flask(__name__)

OPT_CSV_URL = "https://raw.githubusercontent.com/vlmsofts/oi-dashboard/main/data/options_oi.csv"
OI_CSV_URL  = "https://raw.githubusercontent.com/vlmsofts/oi-dashboard/main/data/oi_data.csv"

MONTH_CODE = {'F':1,'G':2,'H':3,'J':4,'K':5,'M':6,'N':7,'Q':8,'U':9,'V':10,'X':11,'Z':12}
MONTH_NAME = {1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',
              7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}

_cache = {}
CACHE_TTL = 300
RISK_FREE  = 0.045

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

def ticker_label(ticker):
    p = parse_ct_ticker(ticker)
    if not p:
        return ticker
    _, year, month_num = p
    return f"{MONTH_NAME[month_num]} {str(year)[-2:]}"

# ── Core data load ────────────────────────────────────────────────────────────

def load_data():
    try:
        opt_rows = fetch_csv(OPT_CSV_URL)
        oi_rows  = fetch_csv(OI_CSV_URL)
    except Exception as e:
        return {'error': str(e)}

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

    # Futures lookup: (month_num, year) -> {settle, last_trade}
    fut_lookup = {}
    for row in ct_fut:
        lt_str = (row.get('last_trade') or '').strip()
        settle = (row.get('settle')     or '').strip()
        date_s = (row.get('date')       or '').strip()
        if not lt_str or not settle or not date_s:
            continue
        try:
            lt_dt = datetime.strptime(lt_str, '%Y-%m-%d')
            key   = (lt_dt.month, lt_dt.year)
            if key not in fut_lookup or date_s > fut_lookup[key]['date']:
                fut_lookup[key] = {
                    'settle':     float(settle),
                    'last_trade': lt_str,
                    'date':       date_s,
                }
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

    expiry_list = sorted(seen.keys(), key=lambda t: (seen[t][1], seen[t][2]))

    expiry_labels = {}
    futures       = {}
    last_trade    = {}

    for ticker in expiry_list:
        _, year, month_num = seen[ticker]
        expiry_labels[ticker] = ticker_label(ticker)
        key = (month_num, year)
        if key in fut_lookup:
            futures[ticker]    = fut_lookup[key]['settle']
            last_trade[ticker] = fut_lookup[key]['last_trade']

    # ATM strike per expiry
    atm_strike = {}
    for ticker in expiry_list:
        fwd = futures.get(ticker)
        if fwd is None:
            continue
        strikes = set(r['strike'] for r in today_opts if r['ticker'] == ticker)
        if strikes:
            atm_strike[ticker] = min(strikes, key=lambda k: abs(k - fwd))

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

    def atm_iv_for_date(ticker, date_str):
        fwd = futures.get(ticker)
        atm = atm_strike.get(ticker)
        if fwd is None or atm is None:
            return None
        dte = get_dte(ticker, date_str)
        rows = [r for r in ct_opts
                if r['date'] == date_str
                and r['ticker'] == ticker
                and abs(r['strike'] - atm) < 0.01]
        # prefer call, fall back to put
        for pc in ('Call', 'Put'):
            for row in rows:
                if row['pc'] == pc:
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
        iv_p = atm_iv_for_date(ticker, prev_date)
        if iv_p is not None:
            atm_iv_1d_chg[ticker] = round((iv_t - iv_p) * 100, 2)
        iv_w = atm_iv_for_date(ticker, week_date)
        if iv_w is not None:
            atm_iv_1w_chg[ticker] = round((iv_t - iv_w) * 100, 2)

    # ── IV Percentile ─────────────────────────────────────────────────────────
    iv_percentile  = {}
    history_months = {}

    for ticker in expiry_list:
        atm     = atm_strike.get(ticker)
        fwd     = futures.get(ticker)
        iv_pct  = atm_iv.get(ticker)
        if atm is None or fwd is None or iv_pct is None:
            continue

        date_ivs = {}
        for row in ct_opts:
            d = row['date']
            if row['ticker'] != ticker or abs(row['strike'] - atm) > 0.01 or d in date_ivs:
                continue
            dte = get_dte(ticker, d)
            iv  = solve_iv(row, fwd, dte)
            if iv is not None:
                date_ivs[d] = iv * 100

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

    # ── Serialize options for client ──────────────────────────────────────────
    def filter_date(date_str):
        return [r for r in ct_opts if r['date'] == date_str]

    return {
        'last_date':      last_date,
        'prev_date':      prev_date,
        'week_date':      week_date,
        'expiries':       expiry_list,
        'expiry_labels':  expiry_labels,
        'futures':        futures,
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
    }

# ── Routes ────────────────────────────────────────────────────────────────────

@server.route('/')
def index():
    data = load_data()
    return render_template('index.html', data=data)

@server.route('/api/data')
def api_data():
    return jsonify(load_data())

if __name__ == '__main__':
    server.run(debug=True, port=5050)
