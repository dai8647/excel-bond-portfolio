# -*- coding: utf-8 -*-
"""Align all series onto a common business-day index -> master history."""
import datetime as dt
from parse_data import load_jgb, load_ust, load_funding, load_fx, load_bund_fut, load_jgb_fut

def build_master(start="2023-01-01"):
    ust = load_ust()
    jgb = load_jgb(start="2023-01-01")
    fund = load_funding()
    usdjpy = dict(load_fx("USDJPY"))
    eurusd = dict(load_fx("EURUSD"))
    zn = dict(load_fx("UST10Y_fut"))
    # Bund先物プロキシETF(BBGなしのwebデータ): IBGM.L(iShares €債7-10年)。
    # JGBは財務省の実利回りが既にあるためETFプロキシ不要。
    bund_etf = dict(load_fx("IBGM"))
    # 先物インプライド利回り: Bund=Eurex RX1派生 / JGB=2510先物連動ETF価格を換算
    bund_fut = dict(load_bund_fut())
    jgb_fut = dict(load_jgb_fut())

    ust_d = {d.isoformat(): v for d, v in ust}
    jgb_d = {d.isoformat(): v for d, v in jgb}

    # common index = UST trading days (Treasury) intersect funding
    dates = sorted(set(ust_d) & set(fund))
    dates = [d for d in dates if d >= start]

    UST_T = ["2Y","5Y","10Y","20Y","30Y"]
    JGB_T = ["2Y","10Y","20Y","30Y"]

    rows = []
    prev_j = {}
    prev_fx = {}
    for d in dates:
        u = ust_d.get(d, {})
        j = jgb_d.get(d, {})
        f = fund.get(d, {})
        # forward-fill JGB (JP holidays differ from US)
        for t in JGB_T:
            if j.get(t) is not None:
                prev_j[t] = j.get(t)
        # forward-fill FX/futures (their calendars differ from UST)
        for k in ("USDJPY", "EURUSD", "ZN", "BUND_ETF", "BUND_FUT", "JGB_FUT"):
            src = {"USDJPY": usdjpy, "EURUSD": eurusd, "ZN": zn,
                   "BUND_ETF": bund_etf, "BUND_FUT": bund_fut,
                   "JGB_FUT": jgb_fut}[k]
            if src.get(d) is not None:
                prev_fx[k] = src.get(d)
        rec = {"date": d}
        for t in UST_T:
            rec[f"UST_{t}"] = u.get(t)
        for t in JGB_T:
            rec[f"JGB_{t}"] = j.get(t) if j.get(t) is not None else prev_j.get(t)
        rec["SOFR"] = f.get("SOFR")
        rec["EFFR"] = f.get("EFFR")
        rec["TGCR"] = f.get("TGCR")
        rec["BGCR"] = f.get("BGCR")
        rec["USDJPY"] = usdjpy.get(d) if usdjpy.get(d) is not None else prev_fx.get("USDJPY")
        rec["EURUSD"] = eurusd.get(d) if eurusd.get(d) is not None else prev_fx.get("EURUSD")
        rec["ZN"] = zn.get(d) if zn.get(d) is not None else prev_fx.get("ZN")
        rec["BUND_ETF"] = bund_etf.get(d) if bund_etf.get(d) is not None else prev_fx.get("BUND_ETF")
        rec["BUND_FUT"] = bund_fut.get(d) if bund_fut.get(d) is not None else prev_fx.get("BUND_FUT")
        rec["JGB_FUT"] = jgb_fut.get(d) if jgb_fut.get(d) is not None else prev_fx.get("JGB_FUT")
        rows.append(rec)
    return rows, UST_T, JGB_T

if __name__ == "__main__":
    rows, UT, JT = build_master()
    print("aligned rows:", len(rows))
    print("first:", rows[0]["date"], "UST10", rows[0]["UST_10Y"], "JGB10", rows[0]["JGB_10Y"])
    print("last:", rows[-1]["date"], "UST10", rows[-1]["UST_10Y"], "JGB10", rows[-1]["JGB_10Y"],
          "SOFR", rows[-1]["SOFR"], "USDJPY", rows[-1]["USDJPY"])
    # count non-null JGB
    nj = sum(1 for r in rows if r["JGB_10Y"] is not None)
    print("JGB coverage:", nj, "/", len(rows))
    # futures coverage
    print("BUND_FUT coverage:", sum(1 for r in rows if r["BUND_FUT"] is not None), "/", len(rows))
    print("JGB_FUT coverage:", sum(1 for r in rows if r["JGB_FUT"] is not None), "/", len(rows))
    lr = rows[-1]
    print("latest BUND_FUT/JGB_FUT:", lr["BUND_FUT"], lr["JGB_FUT"])
