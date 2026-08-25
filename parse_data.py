# -*- coding: utf-8 -*-
"""Parse collected CSVs into clean structures for the Excel builder."""
import csv, datetime as dt, os

D = "data"

def era_to_date(s):
    """Convert Japanese era date like 'R8.7.31' / 'H31.4.30' / 'S63.1.7' to ISO."""
    s = s.strip()
    era, rest = s[0], s[1:]
    parts = rest.split(".")
    y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
    base = {"R": 2018, "H": 1988, "S": 1925, "T": 1911}.get(era)
    if base is None:
        return None
    year = base + y
    try:
        return dt.date(year, m, d)
    except ValueError:
        return None

def load_jgb(start="2020-01-01"):
    """Return list of (date, {tenor: yield}) from MOF full history."""
    tenors = ["1Y","2Y","3Y","4Y","5Y","6Y","7Y","8Y","9Y","10Y","15Y","20Y","25Y","30Y","40Y"]
    rows = []
    with open(os.path.join(D, "MOF_jgbcm_all.csv"), encoding="cp932") as f:
        rd = csv.reader(f)
        next(rd)  # title
        header = next(rd)
        for r in rd:
            if not r or not r[0].strip():
                continue
            d = era_to_date(r[0])
            if d is None or d.isoformat() < start:
                continue
            vals = {}
            for i, t in enumerate(tenors):
                v = r[i+1].strip() if i+1 < len(r) else ""
                try:
                    vals[t] = float(v)
                except ValueError:
                    vals[t] = None
            rows.append((d, vals))
    return rows

def load_ust():
    """Return list of (date, {tenor: yield}) from Treasury CSVs 2023-2026.
    Header-driven: 2023-24 files lack the 1.5M column, so map by header name."""
    norm = {"1 MO":"1M","1.5 MO":"1.5M","1.5 MONTH":"1.5M","2 MO":"2M","3 MO":"3M",
            "4 MO":"4M","6 MO":"6M","1 YR":"1Y","2 YR":"2Y","3 YR":"3Y","5 YR":"5Y",
            "7 YR":"7Y","10 YR":"10Y","20 YR":"20Y","30 YR":"30Y"}
    rows = []
    for y in [2023, 2024, 2025, 2026]:
        with open(os.path.join(D, f"UST_{y}.csv"), encoding="utf-8-sig") as f:
            rd = csv.reader(f)
            header = next(rd)
            cols = {}
            for i, h in enumerate(header[1:], start=1):
                t = norm.get(h.strip().upper())
                if t:
                    cols[t] = i
            for r in rd:
                if not r or not r[0].strip():
                    continue
                d = dt.datetime.strptime(r[0], "%m/%d/%Y").date()
                vals = {}
                for t, i in cols.items():
                    v = r[i].strip() if i < len(r) else ""
                    try:
                        vals[t] = float(v)
                    except ValueError:
                        vals[t] = None
                rows.append((d, vals))
    rows.sort(key=lambda x: x[0])
    return rows

def load_funding():
    """Return dict date -> {SOFR, EFFR, TGCR, BGCR}."""
    out = {}
    for name, col in [("SOFR","SOFR"),("EFFR","EFFR"),("TGCR","TGCR"),("BGCR","BGCR")]:
        with open(os.path.join(D, f"NYFED_{name}.csv")) as f:
            rd = csv.reader(f); next(rd)
            for r in rd:
                d = r[0]
                out.setdefault(d, {})[col] = float(r[1]) if r[1] else None
    return out

def load_fx(sym):
    rows = []
    with open(os.path.join(D, f"YH_{sym}.csv")) as f:
        rd = csv.reader(f); next(rd)
        for r in rd:
            try:
                rows.append((r[0], float(r[1])))
            except (ValueError, IndexError):
                pass
    return rows

def load_bund_fut():
    """Bund 10Y futures implied yield (Eurex RX1経由の事前構築CSV, %)."""
    rows = []
    with open(os.path.join(D, "BUND_implied_yield.csv"), encoding="utf-8") as f:
        rd = csv.reader(f)
        next(rd)
        for r in rd:
            try:
                rows.append((r[0], float(r[1])))
            except (ValueError, IndexError):
                pass
    return rows

def load_jgb_fut():
    """JGB 10Y futures proxy: NEXT FUNDS 10年国債先物ETF(2510.T)価格 → インプライド利回り。
    価格→利回りはMOFスポット10Yとの2点キャリブレーション
    (2023-01-04: P=934.1,y=0.491 / 2026-08-21: P=802.4,y=2.801)。"""
    P0, Y0 = 934.1, 0.491
    P1, Y1 = 802.4, 2.801
    b = (Y1 - Y0) / (P1 - P0)
    a = Y0 - b * P0
    rows = []
    with open(os.path.join(D, "YH_2510.csv"), encoding="utf-8") as f:
        rd = csv.reader(f); next(rd)
        for r in rd:
            try:
                rows.append((r[0], a + b * float(r[1])))
            except (ValueError, IndexError):
                pass
    return rows

if __name__ == "__main__":
    jgb = load_jgb()
    ust = load_ust()
    fund = load_funding()
    usdjpy = load_fx("USDJPY")
    print("JGB rows:", len(jgb), "first", jgb[0][0], "last", jgb[-1][0])
    print("  latest 10Y/30Y:", jgb[-1][1]["10Y"], jgb[-1][1]["30Y"])
    print("UST rows:", len(ust), "first", ust[0][0], "last", ust[-1][0])
    print("  latest 2Y/10Y/30Y:", ust[-1][1]["2Y"], ust[-1][1]["10Y"], ust[-1][1]["30Y"])
    print("Funding days:", len(fund))
    k = sorted(fund)[-1]
    print("  latest", k, fund[k])
    print("USDJPY rows:", len(usdjpy), "last", usdjpy[-1])
