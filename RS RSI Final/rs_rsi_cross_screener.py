"""
RS CROSS + RSI Stock Screener (daily)
=====================================
Asks for an Excel file with your stock symbols and returns only the stocks where:

  1. RSI(14) on the daily close  >  60
  2. Relative Strength vs S&P 500 (55 days) JUST crossed ABOVE 1
       previous trading day:  RS <= 1
       latest trading day:    RS >  1
       RS for a day = (stock close that day / stock close 55 trading days earlier)
                      / (S&P 500 same ratio)
       RS > 1 means the stock gained more than the S&P 500 over those 55 days.

Output (Excel, saved next to your list):
  Symbol | Stock Name | Sector | Industry | Market Cap ($B) | Date |
  Last Close | RSI(14) | RS | Prev Day RS

Run:  double-click RS_RSI_Cross_Screener.bat
  or: python rs_rsi_cross_screener.py "C:\\path\\to\\stocks.xlsx"
"""

import os
import sys
from concurrent.futures import ThreadPoolExecutor

# ============================ SETTINGS ============================
RSI_LENGTH = 14
RSI_MIN = 60            # RSI must be ABOVE this
RS_LENGTH = 55          # trading days
RS_LEVEL = 1            # RS must cross ABOVE this
BENCHMARK = "^GSPC"     # S&P 500 index on Yahoo Finance
EXCHANGE_SUFFIX = ""    # "" = US stocks, ".NS" = India NSE, ".BO" = India BSE
SKIP_TODAY_IF_MARKET_OPEN = False   # True = use last completed day only
# ==================================================================

import pandas as pd
import yfinance as yf

SYMBOL_HEADERS = {"symbol", "symbols", "ticker", "tickers", "stock", "stocks",
                  "code", "scrip", "stock symbol"}
OUT_COLS = ["Symbol", "Stock Name", "Sector", "Industry", "Market Cap ($B)",
            "Date", "Last Close", f"RSI({RSI_LENGTH})", "RS", "Prev Day RS"]


# ------------------------------------------------------------------ input
def ask_for_file() -> str:
    if len(sys.argv) > 1:
        return sys.argv[1]
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    return filedialog.askopenfilename(
        title="Select the Excel file with your list of stocks",
        filetypes=[("Excel files", "*.xlsx *.xls")])


def read_symbols(path: str) -> list[str]:
    df = pd.read_excel(path)
    col = next((c for c in df.columns
                if str(c).strip().lower() in SYMBOL_HEADERS), df.columns[0])
    syms = df[col].dropna().astype(str).str.strip().str.upper()
    return list(dict.fromkeys(s for s in syms if s and s != "NAN"))


def to_ticker(sym: str) -> str:
    return sym if ("." in sym or not EXCHANGE_SUFFIX) else sym + EXCHANGE_SUFFIX


# ------------------------------------------------------------------ data
def download_closes(tickers: list[str]) -> dict[str, pd.Series]:
    data = yf.download(tickers, period="1y", interval="1d", auto_adjust=False,
                       group_by="ticker", progress=True, threads=True)
    closes = {}
    for t in tickers:
        try:
            s = data[t]["Close"] if isinstance(data.columns, pd.MultiIndex) else data["Close"]
            s = s.dropna()
            if not s.empty:
                closes[t] = s
        except KeyError:
            pass
    return closes


def lookup_info(tickers: list[str]) -> dict[str, dict]:
    def one(t):
        try:
            info = yf.Ticker(t).info or {}
        except Exception:
            info = {}
        return t, {"name": info.get("longName") or info.get("shortName"),
                   "sector": info.get("sector"),
                   "industry": info.get("industry"),
                   "cap": info.get("marketCap")}
    with ThreadPoolExecutor(max_workers=8) as pool:
        return dict(pool.map(one, tickers))


# ------------------------------------------------------------------ math
def rsi_wilder(close: pd.Series, length: int) -> float:
    """RSI with Wilder's smoothing (same as TradingView)."""
    delta = close.diff()
    avg_gain = delta.clip(lower=0).ewm(alpha=1 / length, min_periods=length, adjust=False).mean()
    avg_loss = (-delta.clip(upper=0)).ewm(alpha=1 / length, min_periods=length, adjust=False).mean()
    g, l = avg_gain.iloc[-1], avg_loss.iloc[-1]
    return 100.0 if l == 0 else float(100 - 100 / (1 + g / l))


def relative_strength(stock: pd.Series, bench: pd.Series, length: int) -> pd.Series:
    """RS value for every day (NaN until enough history)."""
    both = pd.concat([stock, bench], axis=1, join="inner").dropna()
    s, b = both.iloc[:, 0], both.iloc[:, 1]
    return (s / s.shift(length)) / (b / b.shift(length))


# ------------------------------------------------------------------ screen
def screen(symbols, closes, bench):
    today = pd.Timestamp.now().normalize()
    rows = []
    for sym in symbols:
        t = to_ticker(sym)
        s = closes.get(t)
        if s is None or len(s) <= max(RSI_LENGTH, RS_LENGTH + 1):
            rows.append({"Symbol": sym, "Ticker": t, "Result": "No data - check symbol"})
            continue
        if SKIP_TODAY_IF_MARKET_OPEN and s.index[-1].normalize() >= today:
            s = s.iloc[:-1]
        rsi = rsi_wilder(s, RSI_LENGTH)
        rs_all = relative_strength(s, bench[bench.index <= s.index[-1]], RS_LENGTH)
        rs, prev_rs = (rs_all.iloc[-1], rs_all.iloc[-2]) if len(rs_all) >= 2 else (None, None)
        fails = []
        if not rsi > RSI_MIN:
            fails.append(f"RSI <= {RSI_MIN}")
        if rs is None or pd.isna(rs) or pd.isna(prev_rs):
            fails.append("RS unknown")
        elif not rs > RS_LEVEL:
            fails.append(f"RS <= {RS_LEVEL}")
        elif not prev_rs <= RS_LEVEL:
            fails.append(f"RS already above {RS_LEVEL} the day before (no fresh cross)")
        rows.append({"Symbol": sym, "Ticker": t,
                     "Date": s.index[-1].date(),
                     "Last Close": round(float(s.iloc[-1]), 2),
                     f"RSI({RSI_LENGTH})": round(rsi, 2),
                     "RS": round(float(rs), 3) if rs is not None and pd.notna(rs) else None,
                     "Prev Day RS": round(float(prev_rs), 3) if prev_rs is not None and pd.notna(prev_rs) else None,
                     "Result": "; ".join(fails) or "MATCH"})
    return pd.DataFrame(rows)


# ------------------------------------------------------------------ output
def save_excel(matches: pd.DataFrame, all_rows: pd.DataFrame, path: str) -> str:
    base, _ = os.path.splitext(path)
    out = f"{base}_RS_RSI_Cross_results.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        matches.to_excel(xw, sheet_name="Matches", index=False)
        all_rows.to_excel(xw, sheet_name="All Stocks", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value or "")) for c in col) + 2
                ws.column_dimensions[col[0].column_letter].width = max(10, min(45, width))
            ws.freeze_panes = "A2"
    return out


def main():
    path = ask_for_file()
    if not path:
        print("No file selected.")
        return

    symbols = read_symbols(path)
    print(f"Found {len(symbols)} stocks in {os.path.basename(path)}")
    print("Downloading daily prices...")
    closes = download_closes([to_ticker(s) for s in symbols] + [BENCHMARK])
    bench = closes.pop(BENCHMARK, None)
    if bench is None:
        print("ERROR: could not download S&P 500 data. Try again later.")
        return

    result = screen(symbols, closes, bench)
    matches = result[result["Result"] == "MATCH"].copy()

    if not matches.empty:
        print(f"Getting name, sector, industry and market cap for {len(matches)} matching stocks...")
        info = lookup_info(list(matches["Ticker"]))
        matches["Stock Name"] = matches["Ticker"].map(lambda t: info[t]["name"])
        matches["Sector"] = matches["Ticker"].map(lambda t: info[t]["sector"])
        matches["Industry"] = matches["Ticker"].map(lambda t: info[t]["industry"])
        matches["Market Cap ($B)"] = matches["Ticker"].map(
            lambda t: round(info[t]["cap"] / 1e9, 2) if info[t]["cap"] else None)
    else:
        for c in ("Stock Name", "Sector", "Industry", "Market Cap ($B)"):
            matches[c] = None
    matches = matches.sort_values("RS", ascending=False)[OUT_COLS]

    out = save_excel(matches, result, path)

    print(f"\n{len(matches)} of {len(symbols)} stocks match "
          f"RSI({RSI_LENGTH}) > {RSI_MIN} AND RS({RS_LENGTH} days) just crossed above {RS_LEVEL}\n")
    if not matches.empty:
        print(matches.to_string(index=False))
    nodata = result["Result"].str.startswith("No data").sum()
    if nodata:
        print(f"\n{nodata} symbol(s) had no data - see the 'All Stocks' sheet.")
    print(f"\nSaved: {out}")


if __name__ == "__main__":
    main()
