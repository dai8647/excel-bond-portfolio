# -*- coding: utf-8 -*-
"""Build the bond RV model workbook.
Part 1: helpers + DV01 math + instrument master.
"""
import datetime as dt
from parse_data import load_jgb, load_ust, load_funding, load_fx

# ---------- DV01 math ----------
def par_mod_dur(n_years, y_pct, freq=2):
    """Modified duration of a par bond (coupon=yield), semi-annual coupons.
    n_years: maturity in years, y_pct: yield in percent. Returns years."""
    y = y_pct / 100.0
    if y <= 0 or n_years <= 0:
        return 0.0
    n = n_years * freq
    c = y / freq
    # Macaulay duration (periods) for par bond
    # D = (1+c)/c * (1 - (1+c)^-n)  ... in periods
    try:
        mac_periods = ((1 + c) / c) * (1 - (1 + c) ** (-n))
    except OverflowError:
        return n_years
    mac_years = mac_periods / freq
    mod = mac_years / (1 + c)
    return mod

def dv01_per_1mm(mod_dur):
    """DV01 in $ per $1mm notional per 1bp = 100 * ModDur."""
    return 100.0 * mod_dur

if __name__ == "__main__":
    # sanity checks
    for n, y in [(2, 4.19), (5, 4.37), (10, 4.71), (30, 5.28), (10, 2.80), (30, 3.98)]:
        md = par_mod_dur(n, y)
        print(f"{n}Y @ {y}%: ModDur={md:.2f}  DV01/mm=${dv01_per_1mm(md):.0f}")
