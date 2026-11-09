"""
RS CROSS + SECTOR RS + RSI Stock Screener (daily)
=================================================
Asks for an Excel file with your stock symbols AND a Sector column, and returns
only the stocks that pass ALL of these filters (checked in this order):

  1. RSI(14) on the daily close  >  60
  2. RS vs S&P 500 (55 days) JUST crossed ABOVE 1
       previous trading day:  RS <= 1
       latest trading day:    RS >  1
  3. The stock's sector is "green":
       sector ETF's RS vs S&P 500 (55 days) > 1 on the latest day
  4. RS vs its SECTOR (55 days) crossed ABOVE 1 within the last 5 trading days
     (including today) AND is still > 1 today

  RS for a day = (close that day / close 55 trading days earlier)
                 / (same ratio for the comparison index)

Sector -> index: each sector name in your Excel is mapped to a SPDR sector ETF
(see SECTOR_ETF below). Both Yahoo-style ("Consumer Cyclical") and GICS-style
("Consumer Discretionary") names work; case and "&"/"and" do not matter.

Output (Excel, saved in THIS program's folder, name includes date + time):
  Sheet "Matches"    - stocks passing all 4 filters, grouped by Sector (A-Z),
                       highest RS vs S&P first within each sector
  Sheet "All Stocks" - every stock with its values and the first filter it failed
  Sheet "Sectors"    - each sector ETF's RS vs S&P 500 and whether it is green

Run:  double-click RS_RSI_Sector_Screener.bat
  or: python rs_rsi_sector_screener.py "C:\\path\\to\\stocks.xlsx"
"""

import os
import sys
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

# ============================ SETTINGS ============================
RSI_LENGTH = 14
RSI_MIN = 60            # RSI must be ABOVE this
RS_LENGTH = 55          # trading days
RS_LEVEL = 1            # RS (vs S&P and vs sector) must cross ABOVE this
SECTOR_RS_MIN = 1       # sector RS vs S&P must be ABOVE this to be "green"
CROSS_LOOKBACK = 5      # RS vs sector cross must be within the last N trading days
BENCHMARK = "^GSPC"     # S&P 500 index on Yahoo Finance
EXCHANGE_SUFFIX = ""    # "" = US stocks, ".NS" = India NSE, ".BO" = India BSE
SKIP_TODAY_IF_MARKET_OPEN = False   # True = use last completed day only

# Sector name (as written in your Excel) -> sector ETF on Yahoo Finance.
# Add your own spellings here if a sector shows "Unknown sector" in the output.
SECTOR_ETF = {
    "technology": "XLK", "information technology": "XLK", "tech": "XLK",
    "healthcare": "XLV", "health care": "XLV",
    "financial services": "XLF", "financials": "XLF", "financial": "XLF",
    "consumer cyclical": "XLY", "consumer discretionary": "XLY",
    "consumer defensive": "XLP", "consumer staples": "XLP",
    "energy": "XLE",
    "industrials": "XLI", "industrial": "XLI",
    "basic materials": "XLB", "materials": "XLB",
    "real estate": "XLRE",
    "utilities": "XLU",
    "communication services": "XLC", "communications": "XLC",
    "telecommunication services": "XLC",
}
# ==================================================================

import numpy as np
import pandas as pd
import yfinance as yf

SYMBOL_HEADERS = {"symbol", "symbols", "ticker", "tickers", "stock", "stocks",
                  "code", "scrip", "stock symbol"}
SECTOR_HEADERS = {"sector", "sectors", "gics sector", "sector name"}

RSI_COL = f"RSI({RSI_LENGTH})"
ALL_COLS = ["Symbol", "Ticker", "Sector", "Sector ETF", "Date", "Last Close",
            RSI_COL, "RS vs S&P", "Prev RS vs S&P", "Sector RS vs S&P",
            "RS vs Sector", "Prev RS vs Sector", "Days Since Sector Cross", "Result"]
OUT_COLS = ["Symbol", "Stock Name", "Sector", "Sector ETF", "Industry",
            "Market Cap ($B)", "Date", "Last Close", RSI_COL,
            "RS vs S&P", "Prev RS vs S&P", "Sector RS vs S&P",
            "RS vs Sector", "Prev RS vs Sector", "Days Since Sector Cross"]


def norm_sector(name: str) -> str:
    return " ".join(str(name).lower().replace("&", " and ").split())


SECTOR_ETF = {norm_sector(k): v for k, v in SECTOR_ETF.items()}


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
        title="Select the Excel file with your stocks and sectors",
        filetypes=[("Excel files", "*.xlsx *.xls")])


def read_input(path: str) -> dict[str, str]:
    """Returns {symbol: sector} in file order (duplicates removed)."""
    df = pd.read_excel(path)
    headers = {str(c).strip().lower(): c for c in df.columns}
    sym_col = next((c for h, c in headers.items() if h in SYMBOL_HEADERS), df.columns[0])
    sec_col = next((c for h, c in headers.items() if h in SECTOR_HEADERS), None)
    if sec_col is None:
        raise SystemExit("ERROR: no 'Sector' column found in the Excel file.")
    out = {}
    for sym, sec in zip(df[sym_col], df[sec_col]):
        if pd.isna(sym):
            continue
        sym = str(sym).strip().upper()
        if sym and sym != "NAN" and sym not in out:
            out[sym] = "" if pd.isna(sec) else str(sec).strip()
    return out


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


def last_two(series: pd.Series):
    if len(series) >= 2:
        return series.iloc[-1], series.iloc[-2]
    return np.nan, np.nan


def rnd(x, n=3):
    return round(float(x), n) if x is not None and pd.notna(x) else None


# ------------------------------------------------------------------ screen
def screen(stocks: dict[str, str], closes, bench):
    today = pd.Timestamp.now().normalize()
    rows = []
    for sym, sector in stocks.items():
        t = to_ticker(sym)
        etf = SECTOR_ETF.get(norm_sector(sector)) if sector else None
        row = {"Symbol": sym, "Ticker": t, "Sector": sector or None, "Sector ETF": etf}
        s = closes.get(t)
        if s is None or len(s) <= RS_LENGTH + CROSS_LOOKBACK:
            row["Result"] = "No data / not enough history - check symbol"
            rows.append(row)
            continue
        if SKIP_TODAY_IF_MARKET_OPEN and s.index[-1].normalize() >= today:
            s = s.iloc[:-1]
        last = s.index[-1]
        b = bench[bench.index <= last]

        # Filter 1 + 2 values
        rsi = rsi_wilder(s, RSI_LENGTH)
        rs, prev_rs = last_two(relative_strength(s, b, RS_LENGTH))

        # Filter 3 + 4 values
        sec = closes.get(etf) if etf else None
        sec_rs = rs_sec = prev_rs_sec = np.nan
        days = None
        if sec is not None:
            sec = sec[sec.index <= last]
            sec_rs_all = relative_strength(sec, b, RS_LENGTH)
            sec_rs = sec_rs_all.iloc[-1] if len(sec_rs_all) else np.nan
            rs_vs_sec = relative_strength(s, sec, RS_LENGTH)
            rs_sec, prev_rs_sec = last_two(rs_vs_sec)
            crossed = (rs_vs_sec > RS_LEVEL) & (rs_vs_sec.shift(1) <= RS_LEVEL)
            hits = np.flatnonzero(crossed.iloc[-CROSS_LOOKBACK:].to_numpy())
            if len(hits):   # 0 = crossed today, 1 = yesterday, ...
                days = int(min(CROSS_LOOKBACK, len(crossed)) - 1 - hits[-1])

        row.update({"Date": last.date(),
                    "Last Close": rnd(s.iloc[-1], 2),
                    RSI_COL: rnd(rsi, 2),
                    "RS vs S&P": rnd(rs), "Prev RS vs S&P": rnd(prev_rs),
                    "Sector RS vs S&P": rnd(sec_rs),
                    "RS vs Sector": rnd(rs_sec), "Prev RS vs Sector": rnd(prev_rs_sec),
                    "Days Since Sector Cross": days})

        # Filters in order - report the first one that fails
        if not rsi > RSI_MIN:
            res = f"Filter 1: RSI <= {RSI_MIN}"
        elif pd.isna(rs) or pd.isna(prev_rs):
            res = "Filter 2: RS vs S&P unknown"
        elif not rs > RS_LEVEL:
            res = f"Filter 2: RS vs S&P <= {RS_LEVEL}"
        elif not prev_rs <= RS_LEVEL:
            res = f"Filter 2: RS vs S&P already above {RS_LEVEL} the day before (no fresh cross)"
        elif etf is None:
            res = f"Filter 3: unknown sector '{sector}' - add it to SECTOR_ETF"
        elif sec is None or sec.empty:
            res = f"Filter 3: no data for sector ETF {etf}"
        elif pd.isna(sec_rs):
            res = "Filter 3: sector RS unknown"
        elif not sec_rs > SECTOR_RS_MIN:
            res = f"Filter 3: sector not green ({etf} RS vs S&P <= {SECTOR_RS_MIN})"
        elif pd.isna(rs_sec):
            res = "Filter 4: RS vs sector unknown"
        elif days is None:
            res = (f"Filter 4: RS vs sector above {RS_LEVEL} but crossed more than "
                   f"{CROSS_LOOKBACK} days ago" if rs_sec > RS_LEVEL
                   else f"Filter 4: RS vs sector <= {RS_LEVEL}")
        elif not rs_sec > RS_LEVEL:
            res = f"Filter 4: crossed vs sector but fell back to <= {RS_LEVEL}"
        else:
            res = "MATCH"
        row["Result"] = res
        rows.append(row)
    return pd.DataFrame(rows, columns=ALL_COLS)


def sector_table(etfs, closes, bench) -> pd.DataFrame:
    rows = []
    for etf in etfs:
        sec = closes.get(etf)
        if sec is None:
            rows.append({"Sector ETF": etf, "Date": None, "RS vs S&P": None, "Green": "No data"})
            continue
        rs = relative_strength(sec, bench, RS_LENGTH).iloc[-1]
        rows.append({"Sector ETF": etf, "Date": sec.index[-1].date(), "RS vs S&P": rnd(rs),
                     "Green": "YES" if pd.notna(rs) and rs > SECTOR_RS_MIN else "no"})
    return pd.DataFrame(rows, columns=["Sector ETF", "Date", "RS vs S&P", "Green"])


# ------------------------------------------------------------------ output
def save_excel(sheets: dict[str, pd.DataFrame], input_path: str) -> str:
    folder = os.path.dirname(os.path.abspath(__file__))
    base = os.path.splitext(os.path.basename(input_path))[0]
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    out = os.path.join(folder, f"{base}_RS_RSI_Sector_{stamp}.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name, index=False)
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

    stocks = read_input(path)
    if not stocks:
        print("No symbols found in the file.")
        return
    etfs = sorted({e for e in (SECTOR_ETF.get(norm_sector(s)) for s in stocks.values() if s) if e})
    print(f"Found {len(stocks)} stocks in {os.path.basename(path)} "
          f"({len(etfs)} sector ETFs: {', '.join(etfs) or 'none'})")
    print("Downloading daily prices...")
    tickers = list(dict.fromkeys([to_ticker(s) for s in stocks] + [BENCHMARK] + etfs))
    closes = download_closes(tickers)
    bench = closes.get(BENCHMARK)
    if bench is None:
        print("ERROR: could not download S&P 500 data. Try again later.")
        return

    result = screen(stocks, closes, bench)
    sectors = sector_table(etfs, closes, bench)
    matches = result[result["Result"] == "MATCH"].copy()

    if not matches.empty:
        print(f"Getting name, industry and market cap for {len(matches)} matching stocks...")
        info = lookup_info(list(matches["Ticker"]))
        matches["Stock Name"] = matches["Ticker"].map(lambda t: info[t]["name"])
        matches["Industry"] = matches["Ticker"].map(lambda t: info[t]["industry"])
        matches["Market Cap ($B)"] = matches["Ticker"].map(
            lambda t: round(info[t]["cap"] / 1e9, 2) if info[t]["cap"] else None)
    # group by Sector (A-Z, unknown sector last), highest RS vs S&P first within each
    matches = matches.sort_values(["Sector", "RS vs S&P"], ascending=[True, False],
                                  na_position="last").reindex(columns=OUT_COLS)

    out = save_excel({"Matches": matches, "All Stocks": result, "Sectors": sectors}, path)

    print("\nSector ETFs (green = RS vs S&P > %s):" % SECTOR_RS_MIN)
    print(sectors.to_string(index=False))
    print(f"\n{len(matches)} of {len(stocks)} stocks match: RSI({RSI_LENGTH}) > {RSI_MIN}, "
          f"RS vs S&P just crossed above {RS_LEVEL}, sector green, "
          f"RS vs sector crossed above {RS_LEVEL} in last {CROSS_LOOKBACK} days\n")
    if not matches.empty:
        print(matches.to_string(index=False))
    print("\nWhere stocks dropped out:")
    print(result["Result"].str.split(":").str[0].value_counts().to_string())
    print(f"\nSaved: {out}")


if __name__ == "__main__":
    main()
