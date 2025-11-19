# nse_live_data_fixed.py
from nsepython import nsefetch
import yfinance as yf
import pandas as pd
import time
import json

symbols = ["RELIANCE", "TCS", "INFY", "HDFCBANK", "ICICIBANK", "KOTAKBANK"]

def safe_get_marketcap(quote, symbol):
    # try common places in NSE response
    try:
        sec = quote.get("securityInfo", {}) or {}
        m1 = sec.get("marketCap")
        if m1:
            return m1
    except Exception:
        pass

    try:
        info = quote.get("info", {}) or {}
        m2 = info.get("marketCap") or info.get("mktCap")
        if m2:
            return m2
    except Exception:
        pass

    # fallback to yfinance
    try:
        yf_t = yf.Ticker(symbol + ".NS")
        info = yf_t.info
        return info.get("marketCap")
    except Exception:
        return None

def safe_get_1y_return(quote, symbol):
    # try NSE priceInfo perChange365d
    try:
        priceInfo = quote.get("priceInfo", {}) or {}
        per365 = priceInfo.get("perChange365d")
        if per365 is not None:
            return per365
    except Exception:
        pass

    # try other possible keys
    try:
        sec = quote.get("securityInfo", {}) or {}
        alt = sec.get("oneYearReturn") or sec.get("change365")
        if alt is not None:
            return alt
    except Exception:
        pass

    # fallback: compute from yfinance historical prices (last vs ~1 year ago)
    try:
        yf_t = yf.Ticker(symbol + ".NS")
        hist = yf_t.history(period="1y", interval="1d")
        if hist.shape[0] >= 2:
            first_close = hist["Close"].iloc[0]
            last_close = hist["Close"].iloc[-1]
            return ((last_close / first_close) - 1) * 100
    except Exception:
        pass

    return None

data = []

for i, symbol in enumerate(symbols):
    try:
        url = f"https://www.nseindia.com/api/quote-equity?symbol={symbol}"
        quote = nsefetch(url)

        # debug print for the first symbol so you can inspect structure (only prints once)
        if i == 0:
            print("\nDEBUG sample for", symbol)
            # pretty-print top-level keys and a small JSON sample (safe)
            try:
                print("Top keys:", list(quote.keys()))
                sample = json.dumps({k: quote[k] for k in list(quote.keys())[:5]}, indent=2, default=str)
                print("Sample snippet (first 5 keys):", sample)
            except Exception as e:
                print("Unable to pretty-print quote:", e)

        priceInfo = quote.get("priceInfo", {}) or {}
        securityInfo = quote.get("securityInfo", {}) or {}
        info = quote.get("info", {}) or {}

        one_year_return = safe_get_1y_return(quote, symbol)
        market_cap = safe_get_marketcap(quote, symbol)

        data.append({
            "Company": info.get("companyName") or symbol,
            "Symbol": symbol,
            "CurrentPrice": priceInfo.get("lastPrice"),
            "Open": priceInfo.get("open"),
            "High": priceInfo.get("intraDayHighLow", {}).get("max"),
            "Low": priceInfo.get("intraDayHighLow", {}).get("min"),
            "PreviousClose": priceInfo.get("previousClose"),
            "OneYearReturn": one_year_return,
            "MarketCap": market_cap,
            "Industry": (quote.get("industryInfo") or {}).get("industry") if quote.get("industryInfo") else None
        })

        time.sleep(1)  # polite pause
    except Exception as e:
        print(f"Error fetching {symbol}: {e}")

df = pd.DataFrame(data)
df.to_excel("C:/Users/Ahmad Zargar/OneDrive/Documents/GitHub/fintech-data-automation/data/nse_live_data_fixed.xlsx", index=False)
print("Saved: ../data/nse_live_data.xlsx")
