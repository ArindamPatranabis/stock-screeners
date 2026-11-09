"""
RS CROSS + Ichimoku Stock Screener (with detail) - WEEKLY
=========================================================
Meant to be run on the WEEKEND to plan the next week's trades.
Asks for an Excel file with your stock symbols and returns only the stocks where
BOTH conditions are true for the LAST COMPLETED WEEK:

  1. Relative Strength vs S&P 500 (55 weeks) crossed ABOVE 1
       previous week:  RS <= 1
       last week:      RS >  1
       RS for a week = (stock weekly close / stock weekly close 55 weeks earlier)
                       / (S&P 500 same ratio)

  2. Weekly close crossed ABOVE the Ichimoku Conversion Line (Tenkan-sen, 9)
     recently: the close has been above the line for AT MOST 2 weeks
       Conversion Line = (highest weekly High of last 9 weeks
                          + lowest weekly Low of last 9 weeks) / 2
       last week:  close > Conversion Line
       AND close <= Conversion Line in at least one of the 2 weeks before
       (i.e. the cross happened last week, or the week before last week).
       Stocks above the line for 3 or more weeks are rejected - the move
       may already be underway.

Extra detail shown (does not change which stocks match):
  Cross Week High / Low = High and Low of last week's candle, the week RS
  crossed above 1. Sector and Industry come from Yahoo Finance.
  Conv. Cross Week = the week the close crossed above the Conversion Line.

"Last completed week" = the week that ended on the most recent Friday BEFORE
today. Run on Sat/Sun -> the week just ended. Run midweek (or on Friday) ->
the previous full week, so an unfinished week is never used.

Weekly bars are built from daily prices (Mon-Fri weeks):
  High = highest daily high, Low = lowest daily low, Close = last daily close.

Output (Excel, saved in the same folder as this program):
  Symbol | Stock Name | Sector | Industry | Market Cap ($B) | Cross Week |
  Cross Week High | Cross Week Low | Close | RS | Prev Week RS |
  Conversion Line | % Above Conv. Line | Conv. Cross Week

Run:  double-click RS_Cross_Detail_Weekly.bat
  or: python rs_cross_detail_weekly.py "C:\\path\\to\\stocks.xlsx"
"""

import os
import sys
from concurrent.futures import ThreadPoolExecutor

# ============================ SETTINGS ============================
RS_LENGTH = 55          # weeks
RS_LEVEL = 1            # RS must cross ABOVE this
CONV_LENGTH = 9         # Ichimoku Conversion Line length (weeks)
CONV_CROSS_WEEKS = 2    # max number of weeks the close may have been above
                        # the Conversion Line (2 = crossed last week or the
                        # week before; 1 = crossed last week only)
BENCHMARK = "^GSPC"     # S&P 500 index on Yahoo Finance
EXCHANGE_SUFFIX = ""    # "" = US stocks, ".NS" = India NSE, ".BO" = India BSE
# ==================================================================

import pandas as pd
import yfinance as yf

SYMBOL_HEADERS = {"symbol", "symbols", "ticker", "tickers", "stock", "stocks",
                  "code", "scrip", "stock symbol"}
OUT_COLS = ["Symbol", "Stock Name", "Sector", "Industry", "Market Cap ($B)",
            "Cross Week", "Cross Week High", "Cross Week Low", "Close", "RS",
            "Prev Week RS", "Conversion Line", "% Above Conv. Line",
            "Conv. Cross Week"]


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
def download_bars(tickers: list[str]) -> dict[str, pd.DataFrame]:
    # 2 years of daily data -> ~104 weekly bars, enough for 55 weeks back
    data = yf.download(tickers, period="2y", interval="1d", auto_adjust=False,
                       group_by="ticker", progress=True, threads=True)
    bars = {}
    for t in tickers:
        try:
            d = data[t] if isinstance(data.columns, pd.MultiIndex) else data
            d = d[["High", "Low", "Close"]].dropna()
            if not d.empty:
                bars[t] = completed_weeks(to_weekly(d))
        except KeyError:
            pass
    return bars


def to_weekly(daily: pd.DataFrame) -> pd.DataFrame:
    """Mon-Fri weekly bars, indexed by the week's last actual trading date."""
    g = daily.groupby(daily.index.to_period("W-FRI"))
    weekly = pd.DataFrame({"High": g["High"].max(),
                           "Low": g["Low"].min(),
                           "Close": g["Close"].last()})
    weekly.index = g.apply(lambda x: x.index[-1]).values
    return weekly


def completed_weeks(weekly: pd.DataFrame) -> pd.DataFrame:
    """Drop the current week (Sat-Fri containing today) - it is not finished."""
    this_week = pd.Timestamp.now().to_period("W-FRI")
    return weekly[weekly.index.to_period("W-FRI") < this_week]


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
def relative_strength(stock: pd.Series, bench: pd.Series, length: int) -> pd.Series:
    """RS value for every week (NaN until enough history), aligned on week."""
    s = stock.copy()
    s.index = s.index.to_period("W-FRI")
    b = bench.copy()
    b.index = b.index.to_period("W-FRI")
    both = pd.concat([s, b], axis=1, join="inner").dropna()
    s, b = both.iloc[:, 0], both.iloc[:, 1]
    return (s / s.shift(length)) / (b / b.shift(length))


def conversion_line(bars: pd.DataFrame, length: int) -> pd.Series:
    """Conversion Line value for every week (NaN until enough history)."""
    return (bars["High"].rolling(length).max() + bars["Low"].rolling(length).min()) / 2


def conv_cross_week(close: pd.Series, cl: pd.Series, weeks: int):
    """Date of the week the close crossed above the Conversion Line, if that
    happened within the last `weeks` completed weeks and the close is still
    above the line last week. Otherwise None."""
    if len(close) < weeks + 1 or not close.iloc[-1] > cl.iloc[-1]:
        return None
    # walk back from last week: the first week found with close <= line is
    # the week BEFORE the cross
    for k in range(2, weeks + 2):
        if pd.isna(cl.iloc[-k]):
            return None
        if close.iloc[-k] <= cl.iloc[-k]:
            return close.index[-k + 1].date()
    return None

# ------------------------------------------------------------------ screen
def screen(symbols, bars, bench):
    rows = []
    for sym in symbols:
        t = to_ticker(sym)
        w = bars.get(t)
        if w is None or len(w) <= RS_LENGTH + 1:
            rows.append({"Symbol": sym, "Ticker": t, "Result": "No data - check symbol"})
            continue
        close = float(w["Close"].iloc[-1])
        cl_all = conversion_line(w, CONV_LENGTH)
        cl = cl_all.iloc[-1]
        conv_cross = conv_cross_week(w["Close"], cl_all, CONV_CROSS_WEEKS)
        rs_all = relative_strength(w["Close"], bench[bench.index <= w.index[-1]], RS_LENGTH)
        rs, prev_rs = (rs_all.iloc[-1], rs_all.iloc[-2]) if len(rs_all) >= 2 else (None, None)
        fails = []
        if rs is None or pd.isna(rs) or pd.isna(prev_rs):
            fails.append("RS unknown")
        elif not rs > RS_LEVEL:
            fails.append(f"RS <= {RS_LEVEL}")
        elif not prev_rs <= RS_LEVEL:
            fails.append(f"RS already above {RS_LEVEL} the week before (no fresh cross)")
        if pd.isna(cl):
            fails.append("Conversion Line unknown")
        elif not close > cl:
            fails.append("Close not above Conversion Line")
        elif conv_cross is None:
            fails.append(f"Above Conversion Line for more than {CONV_CROSS_WEEKS} "
                         f"weeks (cross too old)")
        rows.append({"Symbol": sym, "Ticker": t,
                     "Cross Week": w.index[-1].date(),
                     "Cross Week High": round(float(w["High"].iloc[-1]), 2),
                     "Cross Week Low": round(float(w["Low"].iloc[-1]), 2),
                     "Close": round(close, 2),
                     "RS": round(float(rs), 3) if rs is not None and pd.notna(rs) else None,
                     "Prev Week RS": round(float(prev_rs), 3) if prev_rs is not None and pd.notna(prev_rs) else None,
                     "Conversion Line": round(float(cl), 2) if pd.notna(cl) else None,
                     "% Above Conv. Line": round((close / cl - 1) * 100, 2) if pd.notna(cl) else None,
                     "Conv. Cross Week": conv_cross,
                     "Result": "; ".join(fails) or "MATCH"})
    return pd.DataFrame(rows)


# ------------------------------------------------------------------ output
def save_excel(matches: pd.DataFrame, all_rows: pd.DataFrame, path: str) -> str:
    # save in the program's own folder (not the stock list's folder)
    prog_dir = os.path.dirname(os.path.abspath(__file__))
    base = os.path.splitext(os.path.basename(path))[0]
    out = os.path.join(prog_dir, f"{base}_RS{RS_LENGTH}W_CrossDetail_results.xlsx")
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
    print("Downloading prices (2 years, converted to weekly)...")
    bars = download_bars([to_ticker(s) for s in symbols] + [BENCHMARK])
    bench = bars.pop(BENCHMARK, None)
    if bench is None:
        print("ERROR: could not download S&P 500 data. Try again later.")
        return
    print(f"Using last completed week ending {bench.index[-1].date()}")

    result = screen(symbols, bars, bench["Close"])
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
    # group by Sector (A-Z, unknown sector last), highest RS first within each
    matches = matches.sort_values(["Sector", "RS"], ascending=[True, False],
                                  na_position="last")[OUT_COLS]

    out = save_excel(matches, result, path)

    print(f"\n{len(matches)} of {len(symbols)} stocks had RS vs S&P 500 "
          f"({RS_LENGTH} weeks) cross above {RS_LEVEL} last week AND close "
          f"cross above the Conversion Line ({CONV_LENGTH}) within the last "
          f"{CONV_CROSS_WEEKS} weeks\n")
    if not matches.empty:
        print(matches.to_string(index=False))
    nodata = result["Result"].str.startswith("No data").sum()
    if nodata:
        print(f"\n{nodata} symbol(s) had no data - see the 'All Stocks' sheet.")
    print(f"\nSaved: {out}")


if __name__ == "__main__":
    main()
