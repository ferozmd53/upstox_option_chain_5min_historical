# ============================================================
# UPSTOX NIFTY 5-MIN HISTORICAL OPTION CHAIN
# MULTI-EXPIRY SUPPORT (PAST + PRESENT)
# FIXED: Spot data always uses V3 endpoint
# ============================================================

import os
import math
import time
import traceback
import sqlite3
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import requests
import numpy as np
import pandas as pd

# ============================================================
# CONFIGURATION
# ============================================================
BASE_DIR = r"C:\Users\soiku\Desktop"
TOKEN_FILE = os.path.join(BASE_DIR, "access_token.txt")
OUTPUT_FILE = os.path.join(BASE_DIR, "Upstox_NIFTY_5Min_OptionChain.xlsx")
DB_FILE = os.path.join(BASE_DIR, "Upstox_NIFTY_OptionChain.db")

API_V2 = "https://api.upstox.com/v2"
API_V3 = "https://api.upstox.com/v3"

UNDERLYING_KEY = "NSE_INDEX|Nifty 50"
UNDERLYING_NAME = "NIFTY"

# CHANGE THIS DATE TO TEST PAST OR PRESENT EXPIRY
EXPIRY_DATE = "2026-09-01"  # Past expiry
# EXPIRY_DATE = "2026-09-01"  # Present expiry

START_DATE = "2026-07-25"
END_DATE = "2026-09-01"

INTERVAL_UNIT = "minutes"
INTERVAL = "5"
LOT_SIZE = 65

STRIKES_EACH_SIDE = 10
STRIKE_STEP = 50
EXTRA_OI_DATES = []

RISK_FREE_RATE = 0.10
DAYS_PER_YEAR = 365.0
IST = ZoneInfo("Asia/Kolkata")

pd.set_option("display.max_columns", None)
pd.set_option("display.width", 300)
pd.set_option("display.max_rows", 100)

# ============================================================
# HEADER
# ============================================================
print()
print("=" * 80)
print("UPSTOX NIFTY 5-MIN HISTORICAL OPTION CHAIN")
print("MULTI-EXPIRY SUPPORT (PAST + PRESENT)")
print("=" * 80)
print()
print(f"Underlying : {UNDERLYING_KEY}")
print(f"Expiry     : {EXPIRY_DATE}")
print(f"History    : {START_DATE} -> {END_DATE}")
print(f"Interval   : 5 minutes")
print(f"Lot Size   : {LOT_SIZE}")
print()
print(f"Output     : {OUTPUT_FILE}")
print(f"Database   : {DB_FILE}")
print()

# ============================================================
# LOAD TOKEN
# ============================================================
print("=" * 80)
print("LOADING ACCESS TOKEN")
print("=" * 80)
print()

if not os.path.exists(TOKEN_FILE):
    print(f"ERROR: {TOKEN_FILE} was not found.")
    print("Create:", TOKEN_FILE)
    raise SystemExit

with open(TOKEN_FILE, "r", encoding="utf-8") as f:
    ACCESS_TOKEN = f.read().strip()

if not ACCESS_TOKEN:
    print("ERROR: access_token.txt is empty.")
    raise SystemExit

print("Access token loaded.")
print()

# ============================================================
# HTTP SESSION
# ============================================================
session = requests.Session()
session.headers.update({
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "Accept": "application/json",
    "Content-Type": "application/json"
})

# ============================================================
# API GET
# ============================================================
def api_get(url, params=None, timeout=30):
    try:
        response = session.get(url, params=params, timeout=timeout)
    except Exception as error:
        print("NETWORK ERROR:", error)
        return None

    if response.status_code != 200:
        print("UPSTOX API ERROR")
        print("HTTP:", response.status_code)
        print("URL:", response.url)
        try:
            print(response.json())
        except Exception:
            print(response.text[:1000])
        return None

    try:
        return response.json()
    except Exception as error:
        print("JSON ERROR:", error)
        return None

# ============================================================
# SAFE FLOAT
# ============================================================
def safe_float(value):
    try:
        if value is None:
            return np.nan
        return float(value)
    except Exception:
        return np.nan

# ============================================================
# EXPIRY DETECTION
# ============================================================
today = datetime.now(IST).date()
try:
    expiry_date_obj = datetime.strptime(EXPIRY_DATE, "%Y-%m-%d").date()
    is_past_expiry = expiry_date_obj < today
except Exception:
    print("Invalid EXPIRY_DATE format. Use YYYY-MM-DD")
    raise SystemExit

print(f"Detected Expiry Type: {'PAST' if is_past_expiry else 'ACTIVE/FUTURE'}")

# ============================================================
# FETCH CONTRACTS
# ============================================================
print()
print("=" * 80)
print("GETTING NIFTY OPTION CONTRACTS")
print("=" * 80)
print()

print("Underlying:", UNDERLYING_KEY)
print("Expiry:", EXPIRY_DATE)
print()

# Use the correct endpoint for past or present
if is_past_expiry:
    contracts_response = api_get(
        API_V2 + "/expired-instruments/option/contract",
        params={"instrument_key": UNDERLYING_KEY, "expiry_date": EXPIRY_DATE}
    )
else:
    contracts_response = api_get(
        API_V2 + "/option/contract",
        params={"instrument_key": UNDERLYING_KEY, "expiry_date": EXPIRY_DATE}
    )

if contracts_response is None:
    print("FAILED TO GET OPTION CONTRACTS.")
    raise SystemExit

contracts = contracts_response.get("data", [])
print("Contracts received:", len(contracts))

if not contracts:
    print("NO OPTION CONTRACTS.")
    raise SystemExit

# ============================================================
# PARSE CONTRACTS
# ============================================================
contract_rows = []
for contract in contracts:
    try:
        instrument_key = contract.get("instrument_key")
        strike = contract.get("strike_price")
        option_type = contract.get("instrument_type")
        trading_symbol = contract.get("trading_symbol")

        if not instrument_key or strike is None or option_type is None:
            continue

        option_type = str(option_type).upper()
        if option_type not in ("CE", "PE"):
            continue

        contract_rows.append({
            "instrument_key": instrument_key,
            "strike": float(strike),
            "option_type": option_type,
            "trading_symbol": trading_symbol,
            "expiry": contract.get("expiry"),
            "lot_size": contract.get("lot_size", LOT_SIZE)
        })
    except Exception:
        continue

contracts_df = pd.DataFrame(contract_rows)
if contracts_df.empty:
    print("NO VALID CE/PE CONTRACTS.")
    raise SystemExit

contracts_df = contracts_df.drop_duplicates(subset=["instrument_key"])

print()
print("CE contracts:", len(contracts_df[contracts_df["option_type"] == "CE"]))
print("PE contracts:", len(contracts_df[contracts_df["option_type"] == "PE"]))

# ============================================================
# DATE RANGE
# ============================================================
start_date_obj = datetime.strptime(START_DATE, "%Y-%m-%d").date()
end_date_obj = datetime.strptime(END_DATE, "%Y-%m-%d").date()
effective_end_date = min(end_date_obj, today)

print()
print("=" * 80)
print("DATE RANGE")
print("=" * 80)
print()
print("Today:", today)
print("Start:", start_date_obj)
print("End:", effective_end_date)
print()

if effective_end_date < start_date_obj:
    print("INVALID DATE RANGE.")
    raise SystemExit

# ============================================================
# GET HISTORICAL CANDLES - FIXED FOR SPOT DATA
# ============================================================
def get_historical_candles(instrument_key, from_date, to_date, is_spot=False):
    """
    Get historical candles for an instrument.
    
    Args:
        instrument_key: The instrument key
        from_date: Start date (YYYY-MM-DD)
        to_date: End date (YYYY-MM-DD)
        is_spot: If True, always use V3 endpoint (for spot/underlying)
    """
    # SPOT data (underlying index) always uses V3 endpoint
    if is_spot:
        url = (
            API_V3 + "/historical-candle/" +
            requests.utils.quote(instrument_key, safe="") +
            f"/minutes/5/{to_date}/{from_date}"
        )
    # Option data: use appropriate endpoint based on expiry type
    elif is_past_expiry:
        url = (
            API_V2 + "/expired-instruments/historical-candle/" +
            requests.utils.quote(instrument_key, safe="") +
            f"/5minute/{to_date}/{from_date}"
        )
    else:
        url = (
            API_V3 + "/historical-candle/" +
            requests.utils.quote(instrument_key, safe="") +
            f"/minutes/5/{to_date}/{from_date}"
        )

    data = api_get(url)
    if not data:
        return []

    try:
        return data.get("data", {}).get("candles", [])
    except Exception:
        return []

# ============================================================
# DOWNLOAD SPOT DATA - ALWAYS USE V3
# ============================================================
print("=" * 80)
print("DOWNLOADING NIFTY 5-MINUTE DATA")
print("=" * 80)
print()

spot_records = []
chunk_start = start_date_obj
chunk_number = 0

while chunk_start <= effective_end_date:
    chunk_end = min(chunk_start + timedelta(days=27), effective_end_date)
    chunk_number += 1
    print("Spot chunk:", chunk_number, chunk_start, "->", chunk_end)

    # Pass is_spot=True to always use V3 endpoint
    candles = get_historical_candles(
        UNDERLYING_KEY,
        chunk_start.strftime("%Y-%m-%d"),
        chunk_end.strftime("%Y-%m-%d"),
        is_spot=True
    )
    print("Candles:", len(candles))

    for candle in candles:
        try:
            timestamp = pd.to_datetime(candle[0], utc=True).tz_convert(IST)
            # Make timezone-naive for Excel
            timestamp = timestamp.tz_localize(None)
            spot_records.append({
                "timestamp": timestamp,
                "spot_open": safe_float(candle[1]),
                "spot_high": safe_float(candle[2]),
                "spot_low": safe_float(candle[3]),
                "spot_close": safe_float(candle[4]),
                "spot_volume": safe_float(candle[5]) if len(candle) > 5 else np.nan,
                "spot_oi": safe_float(candle[6]) if len(candle) > 6 else np.nan
            })
        except Exception:
            continue

    chunk_start = chunk_end + timedelta(days=1)
    time.sleep(0.20)

print()
print("Total spot candles:", len(spot_records))

if not spot_records:
    print("NO SPOT DATA.")
    raise SystemExit

spot_df = pd.DataFrame(spot_records)
spot_df = spot_df.sort_values("timestamp").drop_duplicates("timestamp").reset_index(drop=True)

# ============================================================
# LATEST SPOT
# ============================================================
latest_spot = spot_df["spot_close"].dropna().iloc[-1]
print()
print("Latest NIFTY:", latest_spot)

# ============================================================
# ATM
# ============================================================
ATM_STRIKE = round(latest_spot / STRIKE_STEP) * STRIKE_STEP
print("ATM Strike:", ATM_STRIKE)

# ============================================================
# SELECT STRIKES
# ============================================================
lower_strike = ATM_STRIKE - STRIKES_EACH_SIDE * STRIKE_STEP
upper_strike = ATM_STRIKE + STRIKES_EACH_SIDE * STRIKE_STEP

selected_strikes = sorted([
    strike for strike in contracts_df["strike"].unique()
    if lower_strike <= strike <= upper_strike
])

selected_contracts = contracts_df[contracts_df["strike"].isin(selected_strikes)].copy()

print()
print("=" * 80)
print("SELECTED STRIKES")
print("=" * 80)
print()
print("Lower:", lower_strike)
print("ATM:", ATM_STRIKE)
print("Upper:", upper_strike)
print()
print(selected_strikes)
print()
print("Selected contracts:", len(selected_contracts))

if selected_contracts.empty:
    print("NO CONTRACTS SELECTED.")
    raise SystemExit

# ============================================================
# DOWNLOAD OPTION CANDLES
# ============================================================
print()
print("=" * 80)
print("DOWNLOADING OPTION 5-MINUTE DATA")
print("=" * 80)
print()

option_records = []
total_contracts = len(selected_contracts)

for counter, (_, contract) in enumerate(selected_contracts.iterrows(), 1):
    instrument_key = contract["instrument_key"]
    strike = contract["strike"]
    option_type = contract["option_type"]
    trading_symbol = contract["trading_symbol"]

    print()
    print(f"[{counter}/{total_contracts}]", trading_symbol)
    print("Key:", instrument_key)

    chunk_start = start_date_obj
    chunk_number = 0

    while chunk_start <= effective_end_date:
        chunk_end = min(chunk_start + timedelta(days=27), effective_end_date)
        chunk_number += 1
        print("Chunk:", chunk_number, chunk_start, "->", chunk_end)

        # is_spot=False for options
        candles = get_historical_candles(
            instrument_key,
            chunk_start.strftime("%Y-%m-%d"),
            chunk_end.strftime("%Y-%m-%d"),
            is_spot=False
        )
        print("Candles:", len(candles))

        for candle in candles:
            try:
                timestamp = pd.to_datetime(candle[0], utc=True).tz_convert(IST)
                timestamp = timestamp.tz_localize(None)
                option_records.append({
                    "timestamp": timestamp,
                    "strike": strike,
                    "option_type": option_type,
                    "instrument_key": instrument_key,
                    "trading_symbol": trading_symbol,
                    "open": safe_float(candle[1]),
                    "high": safe_float(candle[2]),
                    "low": safe_float(candle[3]),
                    "close": safe_float(candle[4]),
                    "volume_raw": safe_float(candle[5]) if len(candle) > 5 else np.nan,
                    "candle_oi_raw": safe_float(candle[6]) if len(candle) > 6 else np.nan
                })
            except Exception:
                continue

        chunk_start = chunk_end + timedelta(days=1)
        time.sleep(0.20)

print()
print("Total raw option candles:", len(option_records))

if not option_records:
    print("NO OPTION DATA.")
    raise SystemExit

# ============================================================
# OPTION DATAFRAME
# ============================================================
options_df = pd.DataFrame(option_records)
options_df = options_df.sort_values(["timestamp", "strike", "option_type"]).drop_duplicates(
    subset=["timestamp", "instrument_key"]
).reset_index(drop=True)
options_df["trade_date"] = options_df["timestamp"].dt.date

# ============================================================
# HISTORICAL OI
# ============================================================
print()
print("=" * 80)
print("DOWNLOADING HISTORICAL OI")
print("=" * 80)
print()

oi_records = []
current_date = start_date_obj

while current_date <= effective_end_date:
    date_string = current_date.strftime("%Y-%m-%d")
    print("OI date:", date_string)

    oi_data = api_get(
        API_V2 + "/market/oi",
        params={
            "instrument_key": UNDERLYING_KEY,
            "expiry": EXPIRY_DATE,
            "date": date_string
        }
    )

    if oi_data and oi_data.get("data"):
        oi_list = oi_data["data"].get("call_put_oi_data_list", [])
        for item in oi_list:
            strike = safe_float(item.get("strike_price"))
            if np.isfinite(strike):
                oi_records.append({
                    "trade_date": current_date,
                    "expiry": EXPIRY_DATE,
                    "strike": float(strike),
                    "CE_OI_RAW": safe_float(item.get("call_oi")),
                    "PE_OI_RAW": safe_float(item.get("put_oi"))
                })
        print("   OI strikes:", len(oi_list))
    else:
        print("   OI not returned")

    current_date += timedelta(days=1)
    time.sleep(0.20)

print()
print("Historical OI rows:", len(oi_records))

# ============================================================
# OI DATAFRAME
# ============================================================
if oi_records:
    oi_df = pd.DataFrame(oi_records)
    oi_df = oi_df.drop_duplicates(subset=["trade_date", "strike"]).reset_index(drop=True)
else:
    oi_df = pd.DataFrame(columns=["trade_date", "expiry", "strike", "CE_OI_RAW", "PE_OI_RAW"])

# ============================================================
# DROP STALE OI DATES
# ============================================================
if not oi_df.empty:
    _sig_df = oi_df.sort_values(["trade_date", "strike"])
    _signatures = _sig_df.groupby("trade_date").apply(
        lambda g: tuple(g["CE_OI_RAW"].fillna(-1)) + tuple(g["PE_OI_RAW"].fillna(-1))
    ).sort_index()
    _is_stale_repeat = _signatures == _signatures.shift(1)
    _stale_dates = set(_signatures.index[_is_stale_repeat.fillna(False)])
    if _stale_dates:
        print()
        print("Skipping stale/non-trading OI snapshot dates:", sorted(str(d) for d in _stale_dates))
        oi_df = oi_df[~oi_df["trade_date"].isin(_stale_dates)].reset_index(drop=True)

# ============================================================
# DAILY OI + DAILY CHANGE OI
# ============================================================
if not oi_df.empty:
    oi_df = oi_df.sort_values(["strike", "trade_date"]).reset_index(drop=True)
    oi_df["CE_OI"] = oi_df["CE_OI_RAW"] / LOT_SIZE
    oi_df["PE_OI"] = oi_df["PE_OI_RAW"] / LOT_SIZE
    oi_df["CE_Change_OI"] = oi_df.groupby("strike")["CE_OI"].diff()
    oi_df["PE_Change_OI"] = oi_df.groupby("strike")["PE_OI"].diff()

# ============================================================
# MERGE HISTORICAL OI INTO OPTION DATA
# ============================================================
print()
print("Merging historical OI...")

options_df = options_df.merge(
    oi_df,
    on=["trade_date", "strike"],
    how="left"
)

# Set OI based on option type
options_df["CE_OI_RAW"] = np.where(options_df["option_type"] == "CE", options_df["CE_OI_RAW"], np.nan)
options_df["PE_OI_RAW"] = np.where(options_df["option_type"] == "PE", options_df["PE_OI_RAW"], np.nan)
options_df["OI_RAW"] = np.where(
    options_df["option_type"] == "CE",
    options_df["CE_OI_RAW"],
    options_df["PE_OI_RAW"]
)
options_df["OI"] = options_df["OI_RAW"] / LOT_SIZE

# Also ensure CE_OI and PE_OI are properly set
options_df["CE_OI"] = np.where(options_df["option_type"] == "CE", options_df["CE_OI_RAW"] / LOT_SIZE, np.nan)
options_df["PE_OI"] = np.where(options_df["option_type"] == "PE", options_df["PE_OI_RAW"] / LOT_SIZE, np.nan)
options_df["CE_Change_OI"] = np.where(options_df["option_type"] == "CE", options_df["CE_Change_OI"], np.nan)
options_df["PE_Change_OI"] = np.where(options_df["option_type"] == "PE", options_df["PE_Change_OI"], np.nan)

# ============================================================
# MERGE SPOT
# ============================================================
print()
print("=" * 80)
print("MERGING NIFTY + OPTION DATA")
print("=" * 80)
print()

options_df = options_df.sort_values("timestamp")
spot_df = spot_df.sort_values("timestamp")

merged = pd.merge_asof(
    options_df,
    spot_df,
    on="timestamp",
    direction="backward"
)

# ============================================================
# IV / GREEKS FUNCTIONS
# ============================================================
def normal_cdf(x):
    return 0.5 * (1.0 + math.erf(x / math.sqrt(2.0)))

def normal_pdf(x):
    return math.exp(-0.5 * x * x) / math.sqrt(2.0 * math.pi)

def bs_call_price(S, K, T, r, sigma):
    if S <= 0 or K <= 0 or T <= 0 or sigma <= 0:
        return np.nan
    try:
        d1 = (math.log(S / K) + (r + 0.5 * sigma * sigma) * T) / (sigma * math.sqrt(T))
        d2 = d1 - sigma * math.sqrt(T)
        return S * normal_cdf(d1) - K * math.exp(-r * T) * normal_cdf(d2)
    except Exception:
        return np.nan

def bs_put_price(S, K, T, r, sigma):
    if S <= 0 or K <= 0 or T <= 0 or sigma <= 0:
        return np.nan
    try:
        d1 = (math.log(S / K) + (r + 0.5 * sigma * sigma) * T) / (sigma * math.sqrt(T))
        d2 = d1 - sigma * math.sqrt(T)
        return K * math.exp(-r * T) * normal_cdf(-d2) - S * normal_cdf(-d1)
    except Exception:
        return np.nan

def solve_iv(option_price, S, K, T, r, option_type):
    if not np.isfinite(option_price) or not np.isfinite(S):
        return np.nan
    if S <= 0 or K <= 0 or T <= 0:
        return np.nan

    intrinsic = max(0.0, (S - K if option_type == "CE" else K - S))
    if option_price <= intrinsic:
        return np.nan

    low = 0.0001
    high = 5.0

    try:
        if option_type == "CE":
            f_low = bs_call_price(S, K, T, r, low) - option_price
            f_high = bs_call_price(S, K, T, r, high) - option_price
        else:
            f_low = bs_put_price(S, K, T, r, low) - option_price
            f_high = bs_put_price(S, K, T, r, high) - option_price

        if not np.isfinite(f_low) or not np.isfinite(f_high):
            return np.nan
        if f_low * f_high > 0:
            return np.nan

        for _ in range(100):
            mid = (low + high) / 2.0
            if option_type == "CE":
                price = bs_call_price(S, K, T, r, mid)
            else:
                price = bs_put_price(S, K, T, r, mid)
            f_mid = price - option_price
            if abs(f_mid) < 1e-7:
                return mid
            if f_low * f_mid <= 0:
                high = mid
                f_high = f_mid
            else:
                low = mid
                f_low = f_mid
        return (low + high) / 2.0
    except Exception:
        return np.nan

def calculate_greeks(S, K, T, r, iv, option_type):
    result = {"Delta": np.nan, "Gamma": np.nan, "Theta": np.nan, "Vega": np.nan}
    if not all(np.isfinite(x) for x in [S, K, T, r, iv]):
        return result
    if S <= 0 or K <= 0 or T <= 0 or iv <= 0:
        return result

    try:
        sqrt_T = math.sqrt(T)
        d1 = (math.log(S / K) + (r + 0.5 * iv * iv) * T) / (iv * sqrt_T)
        d2 = d1 - iv * sqrt_T
        gamma = normal_pdf(d1) / (S * iv * sqrt_T)
        vega = S * normal_pdf(d1) * sqrt_T / 100.0

        if option_type == "CE":
            delta = normal_cdf(d1)
            theta = (-(S * normal_pdf(d1) * iv) / (2 * sqrt_T) - r * K * math.exp(-r * T) * normal_cdf(d2)) / DAYS_PER_YEAR
        else:
            delta = normal_cdf(d1) - 1.0
            theta = (-(S * normal_pdf(d1) * iv) / (2 * sqrt_T) + r * K * math.exp(-r * T) * normal_cdf(-d2)) / DAYS_PER_YEAR

        result = {"Delta": delta, "Gamma": gamma, "Theta": theta, "Vega": vega}
    except Exception:
        pass
    return result

# ============================================================
# CALCULATE IV / GREEKS
# ============================================================
print()
print("Calculating IV + Greeks...")

expiry_datetime = datetime.strptime(EXPIRY_DATE, "%Y-%m-%d").replace(hour=15, minute=30)

iv_values = []
delta_values = []
gamma_values = []
theta_values = []
vega_values = []
dte_values = []

for _, row in merged.iterrows():
    S = safe_float(row.get("spot_close"))
    K = safe_float(row.get("strike"))
    option_price = safe_float(row.get("close"))
    option_type = row["option_type"]
    timestamp = row["timestamp"]

    # Handle timestamp
    if hasattr(timestamp, 'tzinfo'):
        ts_naive = timestamp.replace(tzinfo=None) if timestamp.tzinfo else timestamp
    else:
        ts_naive = timestamp

    expiry_naive = expiry_datetime.replace(tzinfo=None)

    seconds = (expiry_naive - ts_naive).total_seconds()
    dte = max(seconds / (24 * 60 * 60), 0.0)
    dte_values.append(dte)

    T = max(seconds / (DAYS_PER_YEAR * 24 * 60 * 60), 1.0 / DAYS_PER_YEAR)

    iv = solve_iv(option_price, S, K, T, RISK_FREE_RATE, option_type)
    iv_values.append(iv * 100.0 if np.isfinite(iv) else np.nan)

    greeks = calculate_greeks(S, K, T, RISK_FREE_RATE, iv, option_type)
    delta_values.append(greeks["Delta"])
    gamma_values.append(greeks["Gamma"])
    theta_values.append(greeks["Theta"])
    vega_values.append(greeks["Vega"])

merged["DTE"] = dte_values
merged["IV"] = iv_values
merged["Delta"] = delta_values
merged["Gamma"] = gamma_values
merged["Theta"] = theta_values
merged["Vega"] = vega_values

# ============================================================
# DAILY VOLUME
# ============================================================
print()
print("Calculating daily traded quantity...")

merged["Daily_Volume_RAW"] = merged.groupby(["trade_date", "instrument_key"])["volume_raw"].transform("sum")
merged["Daily_Volume"] = merged["Daily_Volume_RAW"] / LOT_SIZE

# ============================================================
# SPLIT CE AND PE - PRESERVE ALL COLUMNS
# ============================================================
ce = merged[merged["option_type"] == "CE"].copy()
pe = merged[merged["option_type"] == "PE"].copy()

# CE columns
ce["CE_OI"] = ce["CE_OI_RAW"] / LOT_SIZE
ce["CE_Volume"] = ce["volume_raw"] / LOT_SIZE
ce["CE_Daily_TOT_TRADED_QTY"] = ce["Daily_Volume_RAW"]
ce["CE_Daily_Volume"] = ce["Daily_Volume"]

# PE columns
pe["PE_OI"] = pe["PE_OI_RAW"] / LOT_SIZE
pe["PE_Volume"] = pe["volume_raw"] / LOT_SIZE
pe["PE_Daily_TOT_TRADED_QTY"] = pe["Daily_Volume_RAW"]
pe["PE_Daily_Volume"] = pe["Daily_Volume"]

# Rename CE
ce = ce.rename(columns={
    "open": "CE_Open",
    "high": "CE_High",
    "low": "CE_Low",
    "close": "CE_Close",
    "volume_raw": "CE_Volume_RAW",
    "IV": "CE_IV",
    "Delta": "CE_Delta",
    "Gamma": "CE_Gamma",
    "Theta": "CE_Theta",
    "Vega": "CE_Vega",
    "DTE": "CE_DTE",
    "trading_symbol": "CE_Symbol",
    "instrument_key": "CE_Instrument_Key"
})

# Rename PE
pe = pe.rename(columns={
    "open": "PE_Open",
    "high": "PE_High",
    "low": "PE_Low",
    "close": "PE_Close",
    "volume_raw": "PE_Volume_RAW",
    "IV": "PE_IV",
    "Delta": "PE_Delta",
    "Gamma": "PE_Gamma",
    "Theta": "PE_Theta",
    "Vega": "PE_Vega",
    "DTE": "PE_DTE",
    "trading_symbol": "PE_Symbol",
    "instrument_key": "PE_Instrument_Key"
})

# ============================================================
# SELECT COLUMNS FOR CE (PRESERVE SPOT COLUMNS)
# ============================================================
ce_columns = [
    "timestamp", "trade_date", "strike",
    "spot_open", "spot_high", "spot_low", "spot_close", "spot_volume", "spot_oi",
    "CE_Open", "CE_High", "CE_Low", "CE_Close",
    "CE_Volume_RAW", "CE_Volume",
    "CE_Daily_TOT_TRADED_QTY", "CE_Daily_Volume",
    "CE_OI_RAW", "CE_OI", "CE_Change_OI",
    "CE_IV", "CE_Delta", "CE_Gamma", "CE_Theta", "CE_Vega", "CE_DTE",
    "CE_Symbol", "CE_Instrument_Key"
]
ce_columns = [c for c in ce_columns if c in ce.columns]
ce = ce[ce_columns]

# ============================================================
# SELECT COLUMNS FOR PE (PRESERVE SPOT COLUMNS)
# ============================================================
pe_columns = [
    "timestamp", "trade_date", "strike",
    "PE_Open", "PE_High", "PE_Low", "PE_Close",
    "PE_Volume_RAW", "PE_Volume",
    "PE_Daily_TOT_TRADED_QTY", "PE_Daily_Volume",
    "PE_OI_RAW", "PE_OI", "PE_Change_OI",
    "PE_IV", "PE_Delta", "PE_Gamma", "PE_Theta", "PE_Vega", "PE_DTE",
    "PE_Symbol", "PE_Instrument_Key"
]
pe_columns = [c for c in pe_columns if c in pe.columns]
pe = pe[pe_columns]

# ============================================================
# MERGE CE + PE
# ============================================================
chain = pd.merge(ce, pe, on=["timestamp", "strike"], how="outer")
chain = chain.sort_values(["timestamp", "strike"]).reset_index(drop=True)

# Ensure Change OI columns exist
if "CE_Change_OI" not in chain.columns:
    chain["CE_Change_OI"] = np.nan
if "PE_Change_OI" not in chain.columns:
    chain["PE_Change_OI"] = np.nan

# ============================================================
# DERIVED COLUMNS
# ============================================================
chain["PCR_OI"] = np.where(
    chain["CE_OI"].notna() & (chain["CE_OI"].abs() > 0),
    chain["PE_OI"] / chain["CE_OI"],
    np.nan
)

chain["PCR_Change_OI"] = np.where(
    chain["CE_Change_OI"].notna() & (chain["CE_Change_OI"].abs() > 0),
    chain["PE_Change_OI"] / chain["CE_Change_OI"],
    np.nan
)

chain["Total_OI"] = chain["CE_OI"].fillna(0) + chain["PE_OI"].fillna(0)
chain["Total_OI_RAW"] = chain["CE_OI_RAW"].fillna(0) + chain["PE_OI_RAW"].fillna(0)
chain["OI_Difference"] = chain["PE_OI"] - chain["CE_OI"]
chain["IV_Difference"] = chain["PE_IV"] - chain["CE_IV"]
chain["ATM_Distance"] = chain["strike"] - chain["spot_close"]

def get_moneyness(strike, spot):
    if not np.isfinite(strike) or not np.isfinite(spot):
        return ""
    if abs(strike - spot) <= 25:
        return "ATM"
    if strike < spot:
        return "ITM"
    return "OTM"

chain["Moneyness"] = [
    get_moneyness(row["strike"], row["spot_close"])
    for _, row in chain.iterrows()
]

# ============================================================
# FINAL SORT
# ============================================================
chain = chain.sort_values(["timestamp", "strike"]).reset_index(drop=True)

# ============================================================
# FINAL COLUMN ORDER - SAME AS ORIGINAL
# ============================================================
final_columns = [
    "timestamp", "trade_date",
    "spot_open", "spot_high", "spot_low", "spot_close", "spot_volume", "spot_oi",
    "CE_Open", "CE_High", "CE_Low", "CE_Close",
    "CE_Volume_RAW", "CE_Volume",
    "CE_Daily_TOT_TRADED_QTY", "CE_Daily_Volume",
    "CE_OI_RAW", "CE_OI", "CE_Change_OI",
    "CE_IV", "CE_Delta", "CE_Gamma", "CE_Theta", "CE_Vega", "CE_DTE",
    "strike", "Moneyness", "ATM_Distance",
    "PCR_OI", "PCR_Change_OI",
    "Total_OI_RAW", "Total_OI",
    "OI_Difference", "IV_Difference",
    "PE_Open", "PE_High", "PE_Low", "PE_Close",
    "PE_Volume_RAW", "PE_Volume",
    "PE_Daily_TOT_TRADED_QTY", "PE_Daily_Volume",
    "PE_OI_RAW", "PE_OI", "PE_Change_OI",
    "PE_IV", "PE_Delta", "PE_Gamma", "PE_Theta", "PE_Vega", "PE_DTE"
]

final_columns = [c for c in final_columns if c in chain.columns]
chain = chain[final_columns]

# ============================================================
# CLEAN INF
# ============================================================
chain = chain.replace([np.inf, -np.inf], np.nan)

# ============================================================
# EXCEL DATETIME
# ============================================================
def make_excel_datetime_naive(df, column):
    if column not in df.columns:
        return
    dt = pd.to_datetime(df[column], errors="coerce")
    try:
        if dt.dt.tz is not None:
            dt = dt.dt.tz_localize(None)
    except Exception:
        pass
    df[column] = dt

make_excel_datetime_naive(chain, "timestamp")
make_excel_datetime_naive(spot_df, "timestamp")

chain = chain.reset_index(drop=True)
spot_df = spot_df.reset_index(drop=True)

# ============================================================
# SAVE TO SQLITE DATABASE
# ============================================================
print()
print("=" * 80)
print("SAVING DATA TO SQLITE DATABASE")
print("=" * 80)
print()

try:
    db = sqlite3.connect(DB_FILE)
    chain.to_sql("optionchain_5min", db, if_exists="replace", index=False)
    spot_df.to_sql("nifty_spot_5min", db, if_exists="replace", index=False)
    selected_contracts.to_sql("contracts_selected", db, if_exists="replace", index=False)
    oi_df.to_sql("historical_oi", db, if_exists="replace", index=False)
    db.commit()
    db.close()
    print("Database saved:", DB_FILE)
except Exception as error:
    print("DATABASE ERROR:", error)

# ============================================================
# WRITE EXCEL
# ============================================================
print()
print("=" * 80)
print("WRITING EXCEL")
print("=" * 80)
print()

try:
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        chain.to_excel(writer, sheet_name="OptionChain_5Min", index=False)
        spot_df.to_excel(writer, sheet_name="NIFTY_Spot_5Min", index=False)
        selected_contracts.to_excel(writer, sheet_name="Contracts", index=False)
        oi_df.to_excel(writer, sheet_name="Historical_OI", index=False)

        # Summary
        summary = pd.DataFrame({
            "Parameter": [
                "Underlying", "Underlying Key", "Expiry", "Type",
                "Requested Start", "Requested End", "Effective End",
                "Interval", "Lot Size", "Latest Historical Spot",
                "ATM Strike", "Lower Strike", "Upper Strike",
                "Selected Strikes", "CE Contracts", "PE Contracts",
                "Raw Option Candles", "Historical OI Rows", "Final Chain Rows",
                "Risk Free Rate"
            ],
            "Value": [
                UNDERLYING_NAME, UNDERLYING_KEY, EXPIRY_DATE,
                "PAST" if is_past_expiry else "ACTIVE",
                START_DATE, END_DATE, effective_end_date.strftime("%Y-%m-%d"),
                "5 minutes", LOT_SIZE, latest_spot,
                ATM_STRIKE, lower_strike, upper_strike,
                len(selected_strikes),
                len(selected_contracts[selected_contracts["option_type"] == "CE"]),
                len(selected_contracts[selected_contracts["option_type"] == "PE"]),
                len(option_records), len(oi_df), len(chain),
                RISK_FREE_RATE
            ]
        })
        summary.to_excel(writer, sheet_name="Summary", index=False)

    print("Excel file saved:", OUTPUT_FILE)

except Exception as error:
    print("EXCEL ERROR:", error)
    traceback.print_exc()
    raise SystemExit

# ============================================================
# FINAL
# ============================================================
print()
print("=" * 80)
print("SUCCESS")
print("=" * 80)
print()
print("Excel file:", OUTPUT_FILE)
print("Database:", DB_FILE)
print()
print("Final rows:", len(chain))
print("Final columns:", len(chain.columns))
print()
print("Sheets:")
print("1. OptionChain_5Min")
print("2. NIFTY_Spot_5Min")
print("3. Contracts")
print("4. Historical_OI")
print("5. Summary")
print()
print("=" * 80)
print("DONE")
print("=" * 80)
