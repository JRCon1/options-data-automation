#!/usr/bin/env python3

import pandas as pd
import yfinance as yf
import openpyxl
from datetime import datetime, timedelta
import os
import pytz
import numpy as np
from scipy.stats import norm

# ── Configuration ──────────────────────────────────────────────
TICKERS   = ["SPY", "UPRO"]
OPT_TYPES = ["c", "p"]  # c = calls, p = puts
BOUND     = 0.20
MAX_DTE   = 120
LOCAL_XLSX = "data_file.xlsx"

def get_options(ticker_symbol: str,
                opt_type: str = "c",
                bound: float = 0.2,
                max_dte: int = 120) -> pd.DataFrame:
    """
    Fetch options data for a given ticker symbol.
    
    Args:
        ticker_symbol: Stock ticker (e.g., 'SPY')
        opt_type: 'c' for calls, 'p' for puts
        bound: Price range bound (e.g., 0.2 for ±20%)
        max_dte: Maximum days to expiration
        
    Returns:
        DataFrame with options data
    """
    try:
        tk   = yf.Ticker(ticker_symbol)
        spot = tk.history(period="1d")["Close"].iloc[-1]

        lo, hi      = round(spot * (1 - bound), 0), round(spot * (1 + bound), 0)
        cutoff_date = datetime.utcnow() + timedelta(days=max_dte)
        eastern     = pytz.timezone('US/Eastern')
        timestamp   = datetime.now(eastern).replace(microsecond=0)

        rows = []
        for exp_str in tk.options:
            exp_date = pd.to_datetime(exp_str)
            if exp_date > cutoff_date:
                continue
            chain = tk.option_chain(exp_str)
            df = chain.calls if opt_type == "c" else chain.puts
            df = df.loc[
                df["strike"].between(lo, hi),
                ["contractSymbol", "strike", "lastPrice", "bid", "ask", "impliedVolatility"]
            ].copy()
            if df.empty:
                continue
            df["symbol"]           = ticker_symbol
            df["expiry"]           = exp_date
            df["downloaded_at"]    = timestamp
            df["underlying_price"] = round(spot, 2)
            rows.append(df)

        if not rows:
            return pd.DataFrame()
        
        out        = pd.concat(rows, ignore_index=True)
        now        = pd.Timestamp.now(tz=None)
        out["dte"] = (out["expiry"] - now).dt.days + 1
        out = out[out["dte"] > 0]
        
        # Calculate Greeks if we have valid data
        if not out.empty:
            out = calculate_greeks(out, opt_type)
            
        return out
        
    except Exception as e:
        print(f"❌ Error fetching options for {ticker_symbol}: {e}")
        return pd.DataFrame()

def calculate_greeks(data: pd.DataFrame, opt_type: str) -> pd.DataFrame:
    """
    Calculate option Greeks for the given data.
    
    Args:
        data: DataFrame with options data
        opt_type: 'c' for calls, 'p' for puts
        
    Returns:
        DataFrame with Greeks added
    """
    try:
        # Filter out options with implied volatility <= 0.000010
        data = data[data['impliedVolatility'] > 0.000010].copy()
        
        if data.empty:
            return data
        
        # Extract time portion into new 'time' column (HH:MM:SS)
        data['time'] = data['downloaded_at'].dt.strftime('%H:%M:%S')
        
        # Convert dte to years
        data['t'] = data['dte'] / 365.0
        
        # Risk-free rate (using 4.5% as in your example)
        r = 0.045
        
        # Calculate d1
        data['d1'] = (
            np.log(data['underlying_price'] / data['strike']) +
            (r + (data['impliedVolatility'] ** 2) / 2) * data['t']
        ) / (data['impliedVolatility'] * np.sqrt(data['t']))
        
        # Calculate d2
        data['d2'] = data['d1'] - data['impliedVolatility'] * np.sqrt(data['t'])
        
        # Calculate Greeks based on option type
        if opt_type == 'c':  # Calls
            # Delta: N(d1)
            data['delta'] = norm.cdf(data['d1']).round(4)
            
            # Theta (per day): [-S * σ * N'(d1) / (2 * sqrt(T)) - r * K * e^(-rT) * N(d2)] / 365
            data['theta'] = ((
                - (data['underlying_price'] * data['impliedVolatility'] * norm.pdf(data['d1'])) / (2 * np.sqrt(data['t'])) -
                r * data['strike'] * np.exp(-r * data['t']) * norm.cdf(data['d2'])
            ) / 365.0).round(4)
            
        else:  # Puts
            # Delta: N(d1) - 1
            data['delta'] = (norm.cdf(data['d1']) - 1).round(4)
            
            # Theta (per day): [-S * σ * N'(d1) / (2 * sqrt(T)) + r * K * e^(-rT) * N(-d2)] / 365
            data['theta'] = ((
                - (data['underlying_price'] * data['impliedVolatility'] * norm.pdf(data['d1'])) / (2 * np.sqrt(data['t'])) +
                r * data['strike'] * np.exp(-r * data['t']) * norm.cdf(-data['d2'])
            ) / 365.0).round(4)
        
        # Gamma: N'(d1) / (S * σ * sqrt(T)) - same for calls and puts
        data['gamma'] = (norm.pdf(data['d1']) / (
            data['underlying_price'] * data['impliedVolatility'] * np.sqrt(data['t'])
        )).round(4)
        
        # Vega (per 1% volatility): S * sqrt(T) * N'(d1) / 100 - same for calls and puts
        data['vega'] = ((
            data['underlying_price'] * np.sqrt(data['t']) * norm.pdf(data['d1'])
        ) / 100.0).round(4)
        
        # Drop temporary columns
        data = data.drop(columns=['t', 'd1', 'd2'])
        
        return data
        
    except Exception as e:
        print(f"⚠️  Error calculating Greeks: {e}")
        return data

def ensure_workbook_exists():
    """Ensure the XLSX workbook exists, create if not."""
    if not os.path.exists(LOCAL_XLSX):
        wb = openpyxl.Workbook()
        wb.save(LOCAL_XLSX)
        print(f"📊 Created new workbook: {LOCAL_XLSX}")
    else:
        print(f"📊 Using existing workbook: {LOCAL_XLSX}")

def main():
    """Main execution function."""
    eastern = pytz.timezone('US/Eastern')
    now_et = datetime.now(eastern)
    print(f"🚀 Starting options data collection at {now_et.strftime('%Y-%m-%d %H:%M:%S %Z')}")
    
    # Ensure workbook exists
    ensure_workbook_exists()
    
    # Generate timestamp for sheet naming
    timestamp = now_et.strftime("%Y-%m-%d_%H%M")
    
    # Collect and save data for each ticker and option type
    with pd.ExcelWriter(LOCAL_XLSX, engine="openpyxl", mode="a", if_sheet_exists="overlay") as xl:
        total_rows = 0
        for tk in TICKERS:
            for opt_type in OPT_TYPES:
                opt_name = "calls" if opt_type == "c" else "puts"
                print(f"📈 Fetching {opt_name} data for {tk}...")
                df = get_options(tk, opt_type, BOUND, MAX_DTE)
                
                if df.empty:
                    print(f"⚠️  {tk} {opt_name}: no rows matched filters")
                    continue
                    
                # Create sheet name (Excel has 31 char limit)
                sheet = f"{tk}_{opt_type}_{timestamp}"
                if len(sheet) > 31:
                    sheet = sheet[:31]
                    
                df.to_excel(xl, sheet_name=sheet, index=False)
                total_rows += len(df)
                print(f"✅ {sheet:<31s} rows={len(df):4d}")
    
    print(f"🎯 Collection complete! Total rows: {total_rows}")
    print(f"💾 Data saved to: {LOCAL_XLSX}")
    print(f"⏰ Completed at: {now_et.strftime('%Y-%m-%d %H:%M:%S %Z')}")

if __name__ == "__main__":
    main()
