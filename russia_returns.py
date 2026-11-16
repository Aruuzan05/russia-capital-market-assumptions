"""
russia_returns.py — return benchmarks + volatility assumptions for the Russian version of the
KEEPP Monte Carlo simulator ("KEEPP Russia" provider).

What it does
  1. Builds 10-year expected returns per asset from building blocks (rate paths, yield paths,
     Grinold-Kroner equity model, FX drift), in local currency, RUB and USD.
  2. Turns the realised statistics from russia_vol_inputs.xlsx (output of russia_vol.py) into
     forward-looking USD volatilities (RVI blend, duration x yield vol, FX overlay).
  3. Maps the estimated factor correlations onto the asset lines (Ledoit-Wolf matrix from
     russia_vol.py + documented judgement), repairs to positive-definite.
  4. Writes russia_cma.xlsx in which EVERY return/vol is an Excel formula on the Assumptions
     sheet, so the committee can change an input and the whole workbook recalculates.
  5. Python computes the same numbers independently; --check compares them with Excel.

Usage (folder with russia_vol_inputs.xlsx):
    python russia_returns.py                     # offline defaults + vol file
    python russia_returns.py --live              # refresh OFZ / corporate yield & duration from MOEX ISS
    python russia_returns.py --vol-file other.xlsx --out russia_cma.xlsx
"""
from __future__ import annotations
import argparse, math
from dataclasses import dataclass, field
from pathlib import Path
import numpy as np

_HERE = Path(__file__).parent
import pandas as pd

H = 10                                   # horizon, years (fixed: the Paths sheet has 10 columns)
START_YEAR = 2027

# ----------------------------------------------------------------------------- assumptions
# (key, label, value, unit, source note, link). Units: "%" stored as fractions.
ASSUMPTIONS = [
    ("__", "MACRO", None, None, None, None),
    ("pi_ru", "Russia CPI inflation, 10y average", 0.045, "%",
     "CBR target 4% + 0.5pp credibility premium (inflation has overshot target in most years since 2015). Committee input.",
     "https://www.cbr.ru/press/keypr/"),
    ("pi_us", "US CPI inflation, 10y average", 0.025, "%", "Judgement, slightly above the Fed 2% target.", ""),
    ("pi_cn", "China CPI inflation, 10y average", 0.015, "%", "Judgement; used only for the PPP reference of CNY.", ""),
    ("__", "RUB RATES", None, None, None, None),
    ("key_y1", "Key rate, average 2027", 0.115, "%",
     "Midpoint of the CBR forecast range 10.5–12.5% for the 2027 average (July 2026 medium-term forecast).",
     "https://www.cbr.ru/press/keypr/"),
    ("key_y2", "Key rate, average 2028", 0.095, "%", "Glide path between the 2027 average and neutral. Committee input.", ""),
    ("r_star", "Real neutral rate", 0.040, "%",
     "Judgement for Russia's real neutral rate; check the CBR's current estimate in «Основные направления ДКП».",
     "https://www.cbr.ru/about_br/publ/ondkp/"),
    ("mmf_spread", "Money-market fund vs key rate", -0.0025, "%",
     "RUSFAR/LQDT-type funds earn slightly below the key rate after fees.", "https://www.moex.com/ru/index/RUSFAR"),
    ("float_spread", "OFZ-PK floater spread over cash", 0.003, "%", "Judgement: RUONIA-linked OFZ trade near par plus a small spread.", ""),
    ("cash_usd", "USD cash, 10y average (for UIP)", 0.035, "%", "Judgement for the US T-bill average; only feeds the FX UIP leg.", ""),
    ("cny_cash", "CNY cash held in Russia", 0.040, "%",
     "Judgement, well below the 2.4y OFZ-CNY yield (7.1%). VERIFY against RUSFARCNY / bank CNY deposit rates.", "https://www.moex.com/ru/index/RUSFARCNY"),
    ("__", "OFZ (RUB, fixed coupon) — RGBI", None, None, None, None),
    ("ofz_y0", "OFZ starting yield (RGBI)", 0.1595, "%",
     "RGBI YIELD, last obs. 02.11.2026 in russia_vol_inputs.xlsx → ofz_rolling. --live refreshes it.",
     "https://www.moex.com/ru/index/RGBI"),
    ("ofz_D", "OFZ modified duration (RGBI)", 3.90, "yrs", "RGBI DURATION/(1+y), same source as the yield.", "https://www.moex.com/ru/index/RGBI"),
    ("ofz_tp", "OFZ term premium over neutral cash", 0.010, "%", "Judgement: long-run OFZ yield = neutral key rate + term premium.", ""),
    ("conv_N", "Years for yields to reach equilibrium", 5, "yrs", "Linear convergence of all yield curves; committee input.", ""),
    ("__", "RUB CORPORATES — RUCBTRNS", None, None, None, None),
    ("corp_y0", "Corporate IG starting yield", 0.160, "%",
     "Largest RUCBTRNS constituents yield ~15.0–16.3% (Jul 2026). --live replaces it with the index YIELD.",
     "https://www.moex.com/ru/index/RUCBTRNS"),
    ("corp_D", "Corporate IG modified duration", 2.0, "yrs", "RUCBTRNS duration ~1.95y (MOEX index factsheet).", "https://fs.moex.com/files/19822"),
    ("corp_spread", "Equilibrium IG spread over OFZ", 0.010, "%", "Judgement: today's index spread is near zero (corp 2y vs OFZ 4y on an inverted curve); 1.0% assumes partial normalisation.", ""),
    ("corp_loss", "IG expected credit loss p.a.", 0.003, "%", "≈ 0.5% default rate × 60% loss given default.", ""),
    ("__", "CNY BONDS — analyst snapshot 20.10.2026", None, None, None, None),
    ("ofzc_y0", "OFZ-CNY basket yield (29/33/36)", 0.0787, "%",
     "Equal-weight YTM of ОФЗ 29/33/36 CNY, from CNYbonds201026.xlsx (cny_basket_summary.csv).", "https://www.moex.com/ru/index/RUCNYTR"),
    ("ofzc_D", "OFZ-CNY basket modified duration", 4.52, "yrs", "Computed from coupon, frequency and maturity in the snapshot.", ""),
    ("ofzc_ystar", "OFZ-CNY equilibrium yield", 0.0787, "%", "No view: set equal to the starting yield (flat curve).", ""),
    ("cnyc_y0", "CNY corporate AAA basket yield", 0.0833, "%", "Equal-weight YTM of 15 AAA issues (3 stale quotes excluded).", ""),
    ("cnyc_D", "CNY corporate AAA modified duration", 1.80, "yrs", "To the put date where there is one.", ""),
    ("cnyc_ystar", "CNY corporate equilibrium yield", 0.0833, "%", "No view: flat.", ""),
    ("cnyc_loss", "CNY corporate expected loss p.a.", 0.003, "%", "Same as RUB IG.", ""),
    ("__", "RUSSIAN EQUITIES — MCFTR (Grinold-Kroner)", None, None, None, None),
    ("pe0", "MOEX P/E today", 5.5, "x",
     "Between SberCIB's ~4.9x (Jan 2026) and T-Investments' 5.8x trailing (2026).",
     "https://www.tbank.ru/invest/research/strategy/2026/stocks/"),
    ("peT", "P/E at end of horizon", 7.0, "x", "Pre-2022 normal range 6–8.5x (MOEX factsheet 2017–2019: 7.8 / 8.5 / 6.2).", "https://fs.moex.com/files/19418"),
    ("payout", "Dividend payout ratio", 0.55, "ratio",
     "Implies a starting DY = 0.55/5.5 = 10%, in line with BCS's ~10% 12-month estimate.",
     "https://bcswm.ru/ajax/files/68847/2xiyc034ghlepnlny1lc66wwrss3wv01/"),
    ("g_real", "Real EPS growth", 0.010, "%", "≈ long-run real GDP growth.", ""),
    ("dilution", "Net dilution", 0.010, "%", "SPOs/IPOs; typical EM net-issuance drag.", ""),
    ("__", "GOLD", None, None, None, None),
    ("gold_real", "Gold real return (USD)", 0.015, "%", "Judgement; gold is priced globally in USD.", ""),
    ("__", "FX", None, None, None, None),
    ("w_ppp", "Weight on PPP in RUB depreciation", 0.5, "ratio", "d = w·PPP + (1−w)·UIP.", ""),
    ("cny_dep", "CNY depreciation vs USD p.a.", 0.0, "%", "No view (PPP alone would imply ~1% CNY appreciation; see the reference row).", ""),
]
A = {k: v for k, _, v, *_ in ASSUMPTIONS if k != "__"}

# ----------------------------------------------------------------------------- vol inputs
VOL_SPEC = [  # key, label, (sheet, selector), default (her 02.11.2026 run)
    ("fx_rub", "RUB vs USD vol (USD/RUB, CBR), full monthly", ("vw", "USD", "RUB cash (FX only)", "full monthly %"), 19.452),
    ("fx_cny", "CNY vs USD vol, full monthly", ("vw", "USD", "CNY cash (FX only)", "full monthly %"), 5.211),
    ("fx_cny22", "CNY vs USD vol, post-2022 monthly", ("vw", "USD", "CNY cash (FX only)", "post-2022 monthly %"), 7.131),
    ("fx_cnyrub", "CNY/RUB vol, full monthly", ("vw", "RUB", "CNYRUB", "full monthly %"), 19.495),
    ("rts_real", "RTS realised vol, full monthly", ("rvi", "RTS realized, full monthly"), 29.117),
    ("rvi_mean", "RVI full-sample mean", ("rvi", "RVI full-sample mean"), 32.810),
    ("vrp", "Mean variance risk premium (RVI − next-month realised)", ("rvi", "Mean VRP (RVI − next-month realized)"), 5.921),
    ("eq_rub", "MCFTR (RUB) realised, full monthly", ("vw", "RUB", "Russian equity", "full monthly %"), 20.721),
    ("eq_usd_real", "MCFTR in USD realised, full monthly", ("vw", "USD", "Russian equity", "full monthly %"), 30.228),
    ("ofz_sdy", "σ(Δy) OFZ, full sample (bp/yr)", ("ofz", "full", "σ(Δy) bp/yr"), 227.523),
    ("ofz_rub", "RGBITR (RUB) realised, full monthly", ("vw", "RUB", "OFZ (RGBITR)", "full monthly %"), 9.549),
    ("ofz_usd", "RGBITR in USD realised, full monthly", ("vw", "USD", "OFZ (RGBITR)", "full monthly %"), 26.131),
    ("corp_rub", "RUCBTRNS (RUB) realised, full monthly (from 2018)", ("vw", "RUB", "RUB corporates", "full monthly %"), 5.953),
    ("corp_usd", "RUCBTRNS in USD realised, full monthly", ("vw", "USD", "RUB corporates", "full monthly %"), 23.970),
    ("cny_loc", "RUCNYTR (CNY) realised, monthly (from 11.2022)", ("vw", "CNY", "CNY bonds (RUCNYTR)", "full monthly %"), 4.355),
    ("cny_usd", "RUCNYTR in USD realised, monthly", ("vw", "USD", "CNY bonds (RUCNYTR)", "full monthly %"), 7.742),
    ("cny_rub", "RUCNYTR in RUB realised, monthly", ("vw", "RUB", "CNY bonds (RUCNYTR)", "full monthly %"), 16.243),
    ("gold_usd", "Gold in USD realised, post-2022 monthly", ("vw", "USD", "Gold", "post-2022 monthly %"), 18.720),
    ("gold_rub", "Gold in RUB realised, post-2022 monthly", ("vw", "RUB", "Gold", "post-2022 monthly %"), 24.760),
]
VOL_JUDGE = [  # judgement inputs on the Volatility sheet
    ("w_rvi", "Weight on implied (RVI − VRP) in equity vol", 0.5, "ratio"),
    ("cny_Didx", "RUCNYTR index modified duration (assumed)", 2.0, "yrs"),
    ("cash_loc", "Cash vol in its own currency", 0.5, "%pt"),
    ("float_loc", "OFZ floater vol in RUB", 1.5, "%pt"),
    ("rho_float", "ρ(floater RUB return, RUB/USD)", 0.5, "ratio"),
]


def load_vol_inputs(path: str | None) -> tuple[dict, dict, pd.DataFrame | None]:
    v = {k: d for k, _, _, d in VOL_SPEC}
    live_ofz, corr = {}, None
    if not path:
        return v, live_ofz, corr
    x = pd.ExcelFile(path)
    vw = x.parse("vol_windows")
    rvi = x.parse("rvi_stats")
    ofz = x.parse("ofz_decomp")
    for k, _, sel, _ in VOL_SPEC:
        try:
            if sel[0] == "vw":
                val = vw[(vw.ccy == sel[1]) & (vw.series == sel[2])][sel[3]].iloc[0]
            elif sel[0] == "rvi":
                val = rvi[rvi.metric == sel[1]]["value_%"].iloc[0]
            else:
                val = ofz[ofz.window == sel[1]][sel[2]].iloc[0]
            if pd.notna(val):
                v[k] = float(val)
        except Exception:
            pass
    try:
        roll = x.parse("ofz_rolling").dropna()
        live_ofz = {"ofz_y0": float(roll["yield %"].iloc[-1]) / 100, "ofz_D": float(roll["mod dur"].iloc[-1]),
                    "date": str(roll.iloc[-1, 0])[:10]}
    except Exception:
        pass
    try:
        corr = x.parse("corr_USD", index_col=0)
    except Exception:
        corr = None
    return v, live_ofz, corr


def live_yields() -> dict:
    """Latest YIELD/DURATION of RGBI and RUCBTRNS from MOEX ISS (needs internet access to MOEX)."""
    import requests
    out = {}
    for code, (ky, kd) in {"RGBI": ("ofz_y0", "ofz_D"), "RUCBTRNS": ("corp_y0", "corp_D")}.items():
        url = f"https://iss.moex.com/iss/history/engines/stock/markets/index/securities/{code}.json"
        tot = requests.get(url, params={"iss.meta": "off", "from": "2026-01-01"}, timeout=30).json()
        cur = tot.get("history.cursor", {}).get("data") or [[0, 0, 100]]
        last = max(cur[0][1] - 1, 0)
        j = requests.get(url, params={"iss.meta": "off", "from": "2026-01-01", "start": last}, timeout=30).json()
        df = pd.DataFrame(j["history"]["data"], columns=j["history"]["columns"]).dropna(subset=["YIELD"])
        if df.empty:
            continue
        r = df.iloc[-1]
        y = float(r["YIELD"]) / 100
        D = float(r["DURATION"]) / (365 if float(r["DURATION"]) > 50 else 1) / (1 + y)
        out[ky], out[kd] = y, round(D, 3)
        print(f"  live {code} {r['TRADEDATE']}: yield {y:.2%}, mod. duration {D:.2f}")
    return out


# ----------------------------------------------------------------------------- the model (Python)
def geo(rs):
    return math.exp(sum(math.log(1 + r) for r in rs) / len(rs)) - 1


def yield_path(y0, ystar, N):
    return [y0 + (ystar - y0) * t / N if t <= N else ystar for t in range(1, H + 1)]


def bond_returns(y0, path, D, loss=0.0):
    out, prev = [], y0
    for y in path:
        dy = y - prev
        out.append(prev - D * dy + 0.5 * D ** 2 * dy ** 2 - loss)
        prev = y
    return out


def model(a: dict, v: dict) -> dict:
    r = {}
    neutral = a["pi_ru"] + a["r_star"]
    key = [a["key_y1"] if t == 1 else a["key_y2"] if t == 2 else neutral for t in range(1, H + 1)]
    cash = [k + a["mmf_spread"] for k in key]
    flo = [c + a["float_spread"] for c in cash]
    ofz_star = neutral + a["ofz_tp"]
    corp_star = ofz_star + a["corp_spread"]
    N = a["conv_N"]
    ofz_y = yield_path(a["ofz_y0"], ofz_star, N)
    corp_y = yield_path(a["corp_y0"], corp_star, N)
    ofzc_y = yield_path(a["ofzc_y0"], a["ofzc_ystar"], N)
    cnyc_y = yield_path(a["cnyc_y0"], a["cnyc_ystar"], N)
    g = (1 + a["pi_ru"]) * (1 + a["g_real"]) * (1 - a["dilution"]) - 1
    pe = [a["pe0"] * (a["peT"] / a["pe0"]) ** (t / H) for t in range(0, H + 1)]
    eq = [(1 + g) * pe[t] / pe[t - 1] - 1 + a["payout"] / pe[t - 1] * (1 + g) for t in range(1, H + 1)]
    loc = {
        "RUB cash": geo(cash), "OFZ floaters (ОФЗ-ПК)": geo(flo),
        "OFZ fixed (ОФЗ-ПД)": geo(bond_returns(a["ofz_y0"], ofz_y, a["ofz_D"])),
        "RUB corporate IG": geo(bond_returns(a["corp_y0"], corp_y, a["corp_D"], a["corp_loss"])),
        "CNY cash": a["cny_cash"],
        "OFZ in CNY": geo(bond_returns(a["ofzc_y0"], ofzc_y, a["ofzc_D"])),
        "CNY corporate IG": geo(bond_returns(a["cnyc_y0"], cnyc_y, a["cnyc_D"], a["cnyc_loss"])),
        "Russian equities": geo(eq),
        "Gold": (1 + a["pi_us"]) * (1 + a["gold_real"]) - 1,
    }
    cash_avg = loc["RUB cash"]
    ppp = (1 + a["pi_ru"]) / (1 + a["pi_us"]) - 1
    uip = (1 + cash_avg) / (1 + a["cash_usd"]) - 1
    d = a["w_ppp"] * ppp + (1 - a["w_ppp"]) * uip
    c = a["cny_dep"]
    usd, rub = {}, {}
    for k, x in loc.items():
        ccy = ASSET_CCY[k]
        if ccy == "RUB":
            rub[k], usd[k] = x, (1 + x) / (1 + d) - 1
        elif ccy == "CNY":
            usd[k] = (1 + x) / (1 + c) - 1
            rub[k] = (1 + usd[k]) * (1 + d) - 1
        else:
            usd[k], rub[k] = x, (1 + x) * (1 + d) - 1
    r.update(loc=loc, usd=usd, rub=rub, d=d, ppp=ppp, uip=uip, neutral=neutral, g=g,
             paths=dict(key=key, ofz_y=ofz_y, corp_y=corp_y, pe=pe[1:], eq=eq))

    # ---- volatility (all in % points)
    j = {k: val for k, _, val, _ in VOL_JUDGE}
    def comb(s_loc, s_fx, rho):
        return math.sqrt(s_loc ** 2 + s_fx ** 2 + 2 * rho * s_loc * s_fx)
    def rho_impl(s_usd, s_loc, s_fx):
        return (s_usd ** 2 - s_loc ** 2 - s_fx ** 2) / (2 * s_loc * s_fx)
    eq_usd = j["w_rvi"] * (v["rvi_mean"] - v["vrp"]) + (1 - j["w_rvi"]) * v["rts_real"]
    ofz_rub = a["ofz_D"] * v["ofz_sdy"] / 100
    rho_ofz = rho_impl(v["ofz_usd"], v["ofz_rub"], v["fx_rub"])
    sdy_cny = v["cny_loc"] / j["cny_Didx"]
    rho_cny = rho_impl(v["cny_usd"], v["cny_loc"], v["fx_cny22"])
    ofzc_loc, cnyc_loc = a["ofzc_D"] * sdy_cny, a["cnyc_D"] * sdy_cny
    vol_usd = {
        "RUB cash": comb(j["cash_loc"], v["fx_rub"], 0),
        "OFZ floaters (ОФЗ-ПК)": comb(j["float_loc"], v["fx_rub"], j["rho_float"]),
        "OFZ fixed (ОФЗ-ПД)": comb(ofz_rub, v["fx_rub"], rho_ofz),
        "RUB corporate IG": v["corp_usd"],
        "CNY cash": comb(j["cash_loc"], v["fx_cny"], 0),
        "OFZ in CNY": comb(ofzc_loc, v["fx_cny"], rho_cny),
        "CNY corporate IG": comb(cnyc_loc, v["fx_cny"], rho_cny),
        "Russian equities": eq_usd,
        "Gold": v["gold_usd"],
    }
    vol_rub = {
        "RUB cash": j["cash_loc"], "OFZ floaters (ОФЗ-ПК)": j["float_loc"], "OFZ fixed (ОФЗ-ПД)": ofz_rub,
        "RUB corporate IG": v["corp_rub"], "CNY cash": v["fx_cnyrub"],
        "OFZ in CNY": comb(ofzc_loc, v["fx_cnyrub"], 0), "CNY corporate IG": comb(cnyc_loc, v["fx_cnyrub"], 0),
        "Russian equities": v["eq_rub"], "Gold": v["gold_rub"],
    }
    r.update(vol_usd=vol_usd, vol_rub=vol_rub, rho_ofz=rho_ofz, rho_cny=rho_cny, sdy_cny=sdy_cny, eq_usd=eq_usd)
    return r


# ----------------------------------------------------------------------------- asset table
ASSETS = [  # name, sleeve, ccy, benchmark, liquidity
    ("RUB cash", "Cash", "RUB", "Key rate / RUSFAR money-market funds", "T0"),
    ("CNY cash", "Cash", "CNY", "RUSFARCNY / CNY repo", "T0"),
    ("OFZ floaters (ОФЗ-ПК)", "Fixed income RUB", "RUB", "OFZ 29xxx basket (RUONIA-linked)", "T1"),
    ("OFZ fixed (ОФЗ-ПД)", "Fixed income RUB", "RUB", "RGBITR (MOEX)", "T1"),
    ("RUB corporate IG", "Fixed income RUB", "RUB", "RUCBTRNS (MOEX)", "T1–T2"),
    ("OFZ in CNY", "FX bonds", "CNY", "ОФЗ 29/33/36 CNY basket", "T1"),
    ("CNY corporate IG", "FX bonds", "CNY", "RUCNYTR (MOEX) / AAA basket", "T2"),
    ("Russian equities", "Equity", "RUB", "MCFTR (RUB) / RTSI (USD); vol via RVI", "T1"),
    ("Gold", "Real assets", "USD", "GLDRUB_TOM / gold БПИФ", "T1"),
]
ASSET_CCY = {n: c for n, _, c, _, _ in ASSETS}
# factor in corr_USD that each line maps to
FACTOR = {"RUB cash": "RUB cash (FX only)", "CNY cash": "CNY cash (FX only)", "OFZ floaters (ОФЗ-ПК)": "RUB cash (FX only)",
          "OFZ fixed (ОФЗ-ПД)": "OFZ (RGBITR)", "RUB corporate IG": "RUB corporates", "OFZ in CNY": "CNY bonds (RUCNYTR)",
          "CNY corporate IG": "CNY bonds (RUCNYTR)", "Russian equities": "Russian equity", "Gold": "Gold"}
PAIR_OVERRIDE = {("OFZ in CNY", "CNY corporate IG"): 0.90, ("RUB cash", "OFZ floaters (ОФЗ-ПК)"): 0.98}
SCALE = {"OFZ floaters (ОФЗ-ПК)": 0.98}  # floater = RUB-cash factor, slightly diluted


def asset_corr(corr: pd.DataFrame | None) -> pd.DataFrame:
    names = [a[0] for a in ASSETS]
    if corr is None:
        corr = pd.DataFrame(np.eye(7), index=list(set(FACTOR.values())), columns=list(set(FACTOR.values())))
    R = np.eye(len(names))
    for i, a in enumerate(names):
        for k, b in enumerate(names):
            if i == k:
                continue
            if (a, b) in PAIR_OVERRIDE or (b, a) in PAIR_OVERRIDE:
                R[i, k] = PAIR_OVERRIDE.get((a, b), PAIR_OVERRIDE.get((b, a)))
                continue
            fa, fb = FACTOR[a], FACTOR[b]
            base = 1.0 if fa == fb else float(corr.loc[fa, fb])
            R[i, k] = base * SCALE.get(a, 1) * SCALE.get(b, 1)
    w, V = np.linalg.eigh((R + R.T) / 2)
    R2 = V @ np.diag(np.maximum(w, 1e-6)) @ V.T
    dd = np.sqrt(np.diag(R2))
    R2 = R2 / np.outer(dd, dd)
    return pd.DataFrame(np.round(R2, 4), index=names, columns=names)


# ----------------------------------------------------------------------------- Excel writer
def build_workbook(a: dict, v: dict, res: dict, R: pd.DataFrame, path: str, meta: dict):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter as L
    from openpyxl.comments import Comment

    F = "Arial"
    f_in = Font(name=F, color="0000FF")
    f_link = Font(name=F, color="008000")
    f_calc = Font(name=F, color="000000")
    f_b = Font(name=F, bold=True)
    f_h = Font(name=F, bold=True, color="FFFFFF")
    f_title = Font(name=F, bold=True, size=14)
    f_note = Font(name=F, italic=True, color="666666", size=9)
    fill_h = PatternFill("solid", start_color="1F3864")
    fill_sec = PatternFill("solid", start_color="D9E1F2")
    fill_key = PatternFill("solid", start_color="FFFF00")
    thin = Side(style="thin", color="BFBFBF")
    box = Border(top=thin, bottom=thin, left=thin, right=thin)
    PCT, PCT2, NUM2, X = "0.0%;(0.0%);-", "0.00%;(0.00%);-", "0.00", "0.0x"

    wb = Workbook()

    def header(ws, row, labels, widths=None):
        for i, t in enumerate(labels, 1):
            c = ws.cell(row, i, t)
            c.font, c.fill, c.alignment = f_h, fill_h, Alignment(wrap_text=True, vertical="center")
        if widths:
            for i, w in enumerate(widths, 1):
                ws.column_dimensions[L(i)].width = w

    # ---------------- Assumptions
    ws = wb.active
    ws.title = "Assumptions"
    ws["A1"] = "KEEPP Russia — capital market assumptions: inputs"
    ws["A1"].font = f_title
    ws["A2"] = ("Blue = input you can change · black = formula · yellow = key committee levers. "
                "All % stored as fractions. Horizon fixed at 10 years (2027–2036).")
    ws["A2"].font = f_note
    header(ws, 4, ["Key", "Assumption", "Value", "Unit", "Source / rationale", "Link"], [14, 44, 11, 7, 90, 50])
    ref = {}
    row = 5
    KEY_LEVERS = {"pi_ru", "key_y1", "r_star", "ofz_tp", "peT", "w_ppp", "conv_N"}
    for k, lab, val, unit, src, link in ASSUMPTIONS:
        if k == "__":
            ws.cell(row, 1, lab).font = f_b
            for col in range(1, 7):
                ws.cell(row, col).fill = fill_sec
            row += 1
            continue
        ws.cell(row, 1, k).font = f_note
        ws.cell(row, 2, lab).font = f_calc
        c = ws.cell(row, 3, val)
        c.font = f_in
        c.number_format = PCT2 if unit == "%" else (X if unit == "x" else NUM2 if unit in ("yrs", "ratio") else "General")
        if k in KEY_LEVERS:
            c.fill = fill_key
        ws.cell(row, 4, unit).font = f_calc
        ws.cell(row, 5, src).font = f_calc
        ws.cell(row, 5).alignment = Alignment(wrap_text=True)
        if link:
            ws.cell(row, 6, link).hyperlink = link
            ws.cell(row, 6).font = Font(name=F, color="0563C1", underline="single")
        ref[k] = f"Assumptions!$C${row}"
        row += 1
    # derived
    row += 1
    ws.cell(row, 1, "DERIVED (formulas)").font = f_b
    for col in range(1, 7):
        ws.cell(row, col).fill = fill_sec
    row += 1
    derived = [
        ("neutral", "Neutral nominal key rate = π_RU + r*", f"={ref['pi_ru']}+{ref['r_star']}", PCT2),
        ("ofz_star", "OFZ equilibrium yield = neutral + term premium", "=@neutral+" + ref["ofz_tp"], PCT2),
        ("corp_star", "Corporate equilibrium yield = OFZ* + spread", "=@ofz_star+" + ref["corp_spread"], PCT2),
        ("g_eps", "Nominal EPS growth per share = (1+π)(1+g_real)(1−dilution)−1",
         f"=(1+{ref['pi_ru']})*(1+{ref['g_real']})*(1-{ref['dilution']})-1", PCT2),
        ("gold_ret", "Gold return in USD = (1+π_US)(1+real)−1", f"=(1+{ref['pi_us']})*(1+{ref['gold_real']})-1", PCT2),
        ("cash_avg", "RUB cash 10y geometric average (from Paths)", "=Paths!$M$8", PCT2),
        ("ppp", "PPP depreciation = (1+π_RU)/(1+π_US)−1", f"=(1+{ref['pi_ru']})/(1+{ref['pi_us']})-1", PCT2),
        ("uip", "UIP depreciation = (1+cash_RUB)/(1+cash_USD)−1", "=(1+@cash_avg)/(1+" + ref["cash_usd"] + ")-1", PCT2),
        ("d", "RUB depreciation vs USD p.a. d = w·PPP+(1−w)·UIP", "=" + ref["w_ppp"] + "*@ppp+(1-" + ref["w_ppp"] + ")*@uip", PCT2),
        ("cny_ppp", "Reference: CNY PPP drift (negative = CNY appreciation)", f"=(1+{ref['pi_cn']})/(1+{ref['pi_us']})-1", PCT2),
    ]
    for k, lab, fml, fmt in derived:
        ref[k] = f"Assumptions!$C${row}"
        ws.cell(row, 1, k).font = f_note
        ws.cell(row, 2, lab).font = f_calc
        row += 1
    # second pass to resolve @names
    r0 = row - len(derived)
    for i, (k, lab, fml, fmt) in enumerate(derived):
        for kk in ("neutral", "ofz_star", "cash_avg", "ppp", "uip"):
            fml = fml.replace("@" + kk, ref[kk])
        c = ws.cell(r0 + i, 3, fml)
        c.font = f_link if "Paths!" in fml else f_calc
        c.number_format = fmt
        if k == "d":
            c.fill = fill_key
    ws.freeze_panes = "A5"

    # ---------------- Paths
    P = wb.create_sheet("Paths")
    P["A1"] = "Year-by-year building blocks (local currency)"
    P["A1"].font = f_title
    P["A2"] = "Column M = 10-year geometric average = EXP(Σ ln(1+r_t)/10) − 1 — the compound return the app expects."
    P["A2"].font = f_note
    P.column_dimensions["A"].width = 4
    P.column_dimensions["B"].width = 46
    for col in range(3, 14):
        P.column_dimensions[L(col)].width = 10
    P.cell(4, 2, "t").font = f_b
    P.cell(5, 2, "Year").font = f_b
    for t in range(1, H + 1):
        P.cell(4, 2 + t, t).font = f_b
        P.cell(5, 2 + t, str(START_YEAR + t - 1)).font = f_b
    P.cell(4, 13, "10y avg").font = f_b
    cols = [L(2 + t) for t in range(1, H + 1)]  # C..L
    N = ref["conv_N"]

    def put_row(r, label, fn, fmt=PCT2, avg=True, sec=None):
        if sec:
            P.cell(r - 1, 2, sec).font = f_b
            for col in range(2, 14):
                P.cell(r - 1, col).fill = fill_sec
        P.cell(r, 2, label).font = f_calc
        for t in range(1, H + 1):
            c = P.cell(r, 2 + t, fn(t, cols[t - 1], L(1 + t)))
            c.number_format, c.font = fmt, f_calc
        if avg:
            c = P.cell(r, 13, f"=EXP(SUMPRODUCT(LN(1+C{r}:L{r}))/10)-1")
            c.number_format, c.font = PCT2, f_b

    put_row(7, "Key rate (annual average)",
            lambda t, c, p: f"=IF({c}$4=1,{ref['key_y1']},IF({c}$4=2,{ref['key_y2']},{ref['neutral']}))", avg=False, sec="CASH")
    put_row(8, "RUB cash return = key rate + MMF spread", lambda t, c, p: f"={c}7+{ref['mmf_spread']}")
    put_row(9, "OFZ floater return = cash + spread", lambda t, c, p: f"={c}8+{ref['float_spread']}")
    put_row(10, "CNY cash return", lambda t, c, p: f"={ref['cny_cash']}")

    def bond_block(r, sec, y0, ystar, D, loss=None):
        put_row(r, "Yield y_t (linear to equilibrium over N years)",
                lambda t, c, p: f"=IF({c}$4<={N},{y0}+({ystar}-{y0})*{c}$4/{N},{ystar})", avg=False, sec=sec)
        def ret(t, c, p):
            prev = y0 if t == 1 else f"{p}{r}"
            s = f"={prev}-{D}*({c}{r}-{prev})+0.5*{D}^2*({c}{r}-{prev})^2"
            return s + (f"-{loss}" if loss else "")
        put_row(r + 1, "Return = y_{t−1} − D·Δy + ½D²Δy²" + (" − expected loss" if loss else ""), ret)

    bond_block(13, "OFZ FIXED (RGBI)", ref["ofz_y0"], ref["ofz_star"], ref["ofz_D"])
    bond_block(17, "RUB CORPORATE IG (RUCBTRNS)", ref["corp_y0"], ref["corp_star"], ref["corp_D"], ref["corp_loss"])
    bond_block(21, "OFZ IN CNY", ref["ofzc_y0"], ref["ofzc_ystar"], ref["ofzc_D"])
    bond_block(25, "CNY CORPORATE IG", ref["cnyc_y0"], ref["cnyc_ystar"], ref["cnyc_D"], ref["cnyc_loss"])
    put_row(29, "P/E at end of year = PE0·(PE_T/PE0)^(t/10)",
            lambda t, c, p: f"={ref['pe0']}*({ref['peT']}/{ref['pe0']})^({c}$4/10)", fmt="0.00x", avg=False,
            sec="RUSSIAN EQUITIES (Grinold-Kroner)")
    put_row(30, "Dividend yield on start-of-year price = payout / P/E_{t−1}",
            lambda t, c, p: f"={ref['payout']}/" + (ref["pe0"] if t == 1 else f"{p}29"), avg=False)
    put_row(31, "Return = (1+g)·PE_t/PE_{t−1} − 1 + DY·(1+g)",
            lambda t, c, p: f"=(1+{ref['g_eps']})*{c}29/" + (ref["pe0"] if t == 1 else f"{p}29")
            + f"-1+{c}30*(1+{ref['g_eps']})")
    put_row(34, "Gold return (USD)", lambda t, c, p: f"={ref['gold_ret']}", sec="GOLD")
    P.freeze_panes = "C6"

    # ---------------- Volatility
    V = wb.create_sheet("Volatility")
    V["A1"] = "Volatility: realised inputs → forward-looking USD σ (annualised, % points)"
    V["A1"].font = f_title
    V["A2"] = (f"Green = realised statistics read from {meta['vol_file']} (russia_vol.py run of {meta['vol_date']}); "
               "blue = judgement. Monthly windows: σ×√12 of month-end log returns.")
    V["A2"].font = f_note
    header(V, 4, ["Key", "Input", "Value", "Source"], [12, 58, 11, 70])
    vref, row = {}, 5
    src_map = {"vw": "sheet vol_windows", "rvi": "sheet rvi_stats", "ofz": "sheet ofz_decomp, window = full"}
    for k, lab, sel, _ in VOL_SPEC:
        V.cell(row, 1, k).font = f_note
        V.cell(row, 2, lab)
        c = V.cell(row, 3, round(v[k], 4))
        c.font, c.number_format = f_link, NUM2
        V.cell(row, 4, f"{meta['vol_file']} → {src_map[sel[0]]}").font = f_note
        vref[k] = f"Volatility!$C${row}"
        row += 1
    row += 1
    for k, lab, val, unit in VOL_JUDGE:
        V.cell(row, 1, k).font = f_note
        V.cell(row, 2, lab)
        c = V.cell(row, 3, val)
        c.font, c.number_format = f_in, NUM2
        V.cell(row, 4, "Judgement").font = f_note
        vref[k] = f"Volatility!$C${row}"
        row += 1
    row += 1
    V.cell(row, 1, "DERIVED").font = f_b
    for col in range(1, 5):
        V.cell(row, col).fill = fill_sec
    row += 1
    dv = [
        ("eq_usd", "Equity σ (USD) = w·(RVI mean − VRP) + (1−w)·RTS realised",
         f"={vref['w_rvi']}*({vref['rvi_mean']}-{vref['vrp']})+(1-{vref['w_rvi']})*{vref['rts_real']}"),
        ("ofz_fwd", "OFZ σ in RUB, forward = D_mod × σ(Δy)", f"={ref['ofz_D']}*{vref['ofz_sdy']}/100"),
        ("rho_ofz", "ρ(OFZ RUB return, RUB/USD) implied from monthly vols = (σ²_USD−σ²_RUB−σ²_FX)/(2σ_RUBσ_FX)",
         f"=({vref['ofz_usd']}^2-{vref['ofz_rub']}^2-{vref['fx_rub']}^2)/(2*{vref['ofz_rub']}*{vref['fx_rub']})"),
        ("sdy_cny", "CNY yield vol σ(Δy) ≈ σ(RUCNYTR in CNY) / D_index (% pts)", f"={vref['cny_loc']}/{vref['cny_Didx']}"),
        ("rho_cny", "ρ(CNY bond return, CNY/USD) implied (post-2022 FX vol, same sample as RUCNYTR)",
         f"=({vref['cny_usd']}^2-{vref['cny_loc']}^2-{vref['fx_cny22']}^2)/(2*{vref['cny_loc']}*{vref['fx_cny22']})"),
        ("ofzc_loc", "OFZ-CNY σ in CNY = D × σ(Δy)", f"={ref['ofzc_D']}*@sdy_cny"),
        ("cnyc_loc", "CNY corporate σ in CNY = D × σ(Δy)", f"={ref['cnyc_D']}*@sdy_cny"),
    ]
    for k, lab, fml in dv:
        vref[k] = f"Volatility!$C${row}"
        for kk in ("sdy_cny",):
            fml = fml.replace("@" + kk, vref.get(kk, "@" + kk))
        V.cell(row, 1, k).font = f_note
        V.cell(row, 2, lab)
        c = V.cell(row, 3, fml)
        c.number_format = "0.000" if k.startswith("rho") else NUM2
        c.font = f_calc
        row += 1
    V.cell(row + 1, 2, "USD σ of a local asset: σ_USD = √(σ_loc² + σ_FX² + 2ρ·σ_loc·σ_FX)").font = f_note

    def comb(sl, sf, rho):
        return f"=SQRT(({sl})^2+({sf})^2+2*({rho})*({sl})*({sf}))"

    vol_usd_f = {
        "RUB cash": comb(vref["cash_loc"], vref["fx_rub"], 0),
        "OFZ floaters (ОФЗ-ПК)": comb(vref["float_loc"], vref["fx_rub"], vref["rho_float"]),
        "OFZ fixed (ОФЗ-ПД)": comb(vref["ofz_fwd"], vref["fx_rub"], vref["rho_ofz"]),
        "RUB corporate IG": f"={vref['corp_usd']}",
        "CNY cash": comb(vref["cash_loc"], vref["fx_cny"], 0),
        "OFZ in CNY": comb(vref["ofzc_loc"], vref["fx_cny"], vref["rho_cny"]),
        "CNY corporate IG": comb(vref["cnyc_loc"], vref["fx_cny"], vref["rho_cny"]),
        "Russian equities": f"={vref['eq_usd']}",
        "Gold": f"={vref['gold_usd']}",
    }
    vol_rub_f = {
        "RUB cash": f"={vref['cash_loc']}", "OFZ floaters (ОФЗ-ПК)": f"={vref['float_loc']}",
        "OFZ fixed (ОФЗ-ПД)": f"={vref['ofz_fwd']}", "RUB corporate IG": f"={vref['corp_rub']}",
        "CNY cash": f"={vref['fx_cnyrub']}", "OFZ in CNY": comb(vref["ofzc_loc"], vref["fx_cnyrub"], 0),
        "CNY corporate IG": comb(vref["cnyc_loc"], vref["fx_cnyrub"], 0),
        "Russian equities": f"={vref['eq_rub']}", "Gold": f"={vref['gold_rub']}",
    }
    vol_method = {
        "RUB cash": "Cash vol in RUB (0.5) combined with RUB/USD vol, ρ = 0",
        "OFZ floaters (ОФЗ-ПК)": "1.5% in RUB (judgement) combined with RUB/USD vol, ρ = 0.5",
        "OFZ fixed (ОФЗ-ПД)": "D_mod × σ(Δy) in RUB; FX overlay with ρ implied from monthly data",
        "RUB corporate IG": "Realised RUCBTRNS in USD, monthly (since 2018; excludes 2014)",
        "CNY cash": "CNY/USD vol",
        "OFZ in CNY": "D × σ(Δy) from RUCNYTR; FX overlay with CNY/USD",
        "CNY corporate IG": "D × σ(Δy) from RUCNYTR; FX overlay with CNY/USD",
        "Russian equities": "½ (RVI mean − VRP) + ½ RTS realised (monthly)",
        "Gold": "Realised gold in USD, post-2022 monthly (full sample has gaps)",
    }

    # ---------------- Summary
    S = wb.create_sheet("Summary", 0)
    S["A1"] = "KEEPP Russia — expected returns and volatilities (10-year, compound)"
    S["A1"].font = f_title
    S["A2"] = ("Every figure is a formula on the Assumptions / Paths / Volatility sheets. Return = 10y geometric (compound) "
               "annual return, the convention the simulator's Return % column uses. App base currency: USD.")
    S["A2"].font = f_note
    heads = ["Asset", "Sleeve", "Local ccy", "Benchmark", "Return, local", "Return, RUB", "Return, USD",
             "Vol, USD", "Vol, RUB (reference)", "Arithmetic mean, USD = g + σ²/2", "Liquidity", "Return method", "Vol method"]
    header(S, 4, heads, [26, 17, 9, 36, 11, 11, 11, 10, 12, 15, 9, 60, 58])
    S.row_dimensions[4].height = 32
    ret_src = {"RUB cash": "Paths!$M$8", "OFZ floaters (ОФЗ-ПК)": "Paths!$M$9", "CNY cash": "Paths!$M$10",
               "OFZ fixed (ОФЗ-ПД)": "Paths!$M$14", "RUB corporate IG": "Paths!$M$18", "OFZ in CNY": "Paths!$M$22",
               "CNY corporate IG": "Paths!$M$26", "Russian equities": "Paths!$M$31", "Gold": "Paths!$M$34"}
    ret_method = {
        "RUB cash": "Average key-rate path (CBR 2027 forecast → neutral = π + r*) + MMF spread",
        "OFZ floaters (ОФЗ-ПК)": "Cash path + floater spread",
        "OFZ fixed (ОФЗ-ПД)": "Constant-duration fund: yield converges from RGBI YTM to neutral + term premium",
        "RUB corporate IG": "As OFZ, with an IG spread, less expected credit loss",
        "CNY cash": "Flat CNY money-market rate",
        "OFZ in CNY": "Constant-duration fund on the OFZ-CNY basket (flat curve)",
        "CNY corporate IG": "Constant-duration fund on the AAA basket, less loss",
        "Russian equities": "Grinold-Kroner: DY + nominal EPS growth − dilution + P/E re-rating",
        "Gold": "US inflation + real return (USD)",
    }
    d, c_ = ref["d"], ref["cny_dep"]
    srow = {}
    for i, (n, sl, ccy, bm, liq) in enumerate(ASSETS):
        r = 5 + i
        srow[n] = r
        S.cell(r, 1, n).font = f_b
        S.cell(r, 2, sl)
        S.cell(r, 3, ccy)
        S.cell(r, 4, bm)
        S.cell(r, 5, f"={ret_src[n]}").font = f_link
        if ccy == "RUB":
            S.cell(r, 6, f"=E{r}")
            S.cell(r, 7, f"=(1+E{r})/(1+{d})-1")
        elif ccy == "CNY":
            S.cell(r, 7, f"=(1+E{r})/(1+{c_})-1")
            S.cell(r, 6, f"=(1+G{r})*(1+{d})-1")
        else:
            S.cell(r, 7, f"=E{r}")
            S.cell(r, 6, f"=(1+G{r})*(1+{d})-1")
        S.cell(r, 8, f"={vol_usd_f[n][1:]}/100").font = f_link
        S.cell(r, 9, f"={vol_rub_f[n][1:]}/100").font = f_link
        S.cell(r, 10, f"=G{r}+H{r}^2/2")
        S.cell(r, 11, liq)
        S.cell(r, 12, ret_method[n]).font = f_note
        S.cell(r, 13, vol_method[n]).font = f_note
        for col in range(5, 11):
            S.cell(r, col).number_format = PCT
        for col in range(1, 14):
            S.cell(r, col).border = box
    last = 5 + len(ASSETS)
    S.cell(last + 1, 1, "RUB depreciation vs USD, d").font = f_b
    S.cell(last + 1, 7, f"={d}").number_format = PCT
    S.cell(last + 1, 7).font = f_link
    S.cell(last + 2, 1, "Not yet calibrated: USD replacement bonds (need the analyst's USD bond table), RUB high yield, "
                        "SMID equities, real-estate ЗПИФ, Personal PE/PC (manual in the app).").font = f_note
    S.freeze_panes = "B5"

    # ---------------- App inputs
    Ap = wb.create_sheet("App inputs", 1)
    Ap["A1"] = "Paste into the simulator (Custom mode / 'KEEPP Russia' provider) — USD base, % units"
    Ap["A1"].font = f_title
    header(Ap, 3, ["Asset", "Category", "Return %", "Vol %"], [28, 18, 11, 11])
    for i, (n, *_r) in enumerate(ASSETS):
        r = 4 + i
        Ap.cell(r, 1, f"=Summary!A{srow[n]}").font = f_link
        Ap.cell(r, 2, f"=Summary!B{srow[n]}").font = f_link
        Ap.cell(r, 3, f"=ROUND(Summary!G{srow[n]}*100,2)").font = f_link
        Ap.cell(r, 4, f"=ROUND(Summary!H{srow[n]}*100,2)").font = f_link

    # ---------------- Correlation
    C = wb.create_sheet("Correlation USD")
    C["A1"] = "Correlation matrix, USD monthly returns (positive-definite)"
    C["A1"].font = f_title
    C["A2"] = (f"From corr_USD in {meta['vol_file']} (Ledoit-Wolf shrunk, 48 months from 12.2022), mapped to asset lines: "
               "floaters = RUB-cash factor × 0.98; OFZ-CNY ↔ CNY corp = 0.90 (judgement); both CNY lines use the RUCNYTR factor. "
               "Eigenvalues floored at 1e-6, then rescaled to unit diagonal. Values (not formulas).")
    C["A2"].font = f_note
    names = list(R.index)
    C.column_dimensions["A"].width = 26
    for j, n in enumerate(names):
        hc = C.cell(4, 2 + j, n)
        hc.font, hc.fill, hc.alignment = f_h, fill_h, Alignment(wrap_text=True, text_rotation=0)
        C.column_dimensions[L(2 + j)].width = 13
        C.cell(5 + j, 1, n).font = f_b
        for k in range(len(names)):
            cc = C.cell(5 + j, 2 + k, float(R.iloc[j, k]))
            cc.number_format = "0.00"
            cc.font = f_calc
            cc.border = box
    C.row_dimensions[4].height = 45

    # ---------------- Sensitivity
    Sn = wb.create_sheet("Sensitivity")
    Sn["A1"] = "USD return sensitivity to RUB depreciation d (RUB-denominated lines)"
    Sn["A1"].font = f_title
    Sn["A2"] = "USD = (1 + RUB return)/(1 + d) − 1. Each 1pp of extra depreciation costs ≈ 1pp of USD return."
    Sn["A2"].font = f_note
    rub_assets = [n for n, _, ccy, *_ in ASSETS if ccy == "RUB"]
    header(Sn, 4, ["d (RUB depreciation p.a.)"] + rub_assets, [24] + [16] * len(rub_assets))
    for i, dd in enumerate([0.02, 0.03, 0.04, 0.05, 0.06, 0.07, 0.08]):
        r = 5 + i
        cd = Sn.cell(r, 1, dd)
        cd.font, cd.number_format = f_in, PCT
        for j, n in enumerate(rub_assets):
            cc = Sn.cell(r, 2 + j, f"=(1+Summary!$E${srow[n]})/(1+$A{r})-1")
            cc.number_format = PCT

    # ---------------- Sources
    So = wb.create_sheet("Sources")
    So["A1"] = "Data sources"
    So["A1"].font = f_title
    header(So, 3, ["Item", "Link / reference"], [55, 95])
    srcs = [
        ("MOEX ISS API (all index histories)", "https://iss.moex.com/iss/reference/"),
        ("RVI — Russian volatility index", "https://www.moex.com/ru/index/RVI"),
        ("RVI methodology", "https://www.moex.com/a5737"),
        ("RVI futures (e.g. VIZ6)", "https://www.moex.com/ru/derivatives/futures-on-indices/viz6"),
        ("RTSI — RTS index (USD)", "https://www.moex.com/ru/index/RTSI"),
        ("MCFTR — MOEX Russia total return", "https://www.moex.com/ru/index/MCFTR"),
        ("RGBI / RGBITR — OFZ index", "https://www.moex.com/ru/index/RGBITR"),
        ("RUCBTRNS — corporate bond index", "https://www.moex.com/ru/index/RUCBTRNS"),
        ("RUCNYTR — CNY bond index", "https://www.moex.com/ru/index/RUCNYTR"),
        ("GLDRUB_TOM — gold (ISS endpoint)",
         "https://iss.moex.com/iss/history/engines/currency/markets/selt/boards/CETS/securities/GLDRUB_TOM.json"),
        ("Bank of Russia FX fixings (XML_dynamic)", "https://www.cbr.ru/development/SXML/"),
        ("Bank of Russia key rate decisions / forecast", "https://www.cbr.ru/press/keypr/"),
        ("Bank of Russia macro survey", "https://www.cbr.ru/statistics/ddkp/mo_br/"),
        ("MOEX corporate index factsheet (duration)", "https://fs.moex.com/files/19822"),
        ("MOEX Russia index factsheet 2017–19 (P/E history)", "https://fs.moex.com/files/19418"),
        ("BCS: 12m dividend yield ~10% vs ~6.5% historical", "https://bcswm.ru/ajax/files/68847/2xiyc034ghlepnlny1lc66wwrss3wv01/"),
        ("T-Investments 2026 strategy (P/E 5.8x trailing)", "https://www.tbank.ru/invest/research/strategy/2026/stocks/"),
        ("CNY bond snapshot", "CNYbonds201026.xlsx from the Russian analyst (20.10.2026)"),
        ("Grinold & Kroner (2002), 'The Equity Risk Premium', Investment Insights", "Equity building block"),
        ("Leibowitz (1986) / Bogle: starting yield ≈ return over ~2×duration", "Bond building block"),
        ("Ledoit & Wolf (2004), 'Honey, I Shrunk the Sample Covariance Matrix'", "Correlation shrinkage"),
        ("Geltner (1993), desmoothing appraisal returns", "For future real-estate / private lines"),
    ]
    for i, (it, ln) in enumerate(srcs):
        So.cell(4 + i, 1, it)
        c = So.cell(4 + i, 2, ln)
        if ln.startswith("http"):
            c.hyperlink = ln
            c.font = Font(name=F, color="0563C1", underline="single")

    for wsx in wb.worksheets:
        for row_ in wsx.iter_rows():
            for cell in row_:
                if cell.font and cell.font.name != F:
                    cell.font = Font(name=F, bold=cell.font.bold, italic=cell.font.italic, color=cell.font.color,
                                     size=cell.font.size, underline=cell.font.underline)
    wb.save(path)


# ----------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--vol-file", default=str(_HERE / "russia_vol_inputs.xlsx"))
    ap.add_argument("--out", default=str(_HERE / "russia_cma.xlsx"))
    ap.add_argument("--live", action="store_true", help="refresh OFZ and corporate yields from MOEX ISS")
    args = ap.parse_args()

    try:
        v, ofz_last, corr = load_vol_inputs(args.vol_file)
        vol_file, vol_date = args.vol_file.split("/")[-1], ofz_last.get("date", "n/a")
    except FileNotFoundError:
        print(f"  ({args.vol_file} not found — using built-in defaults from the 02.11.2026 run)")
        v, ofz_last, corr = load_vol_inputs(None)
        vol_file, vol_date = "built-in defaults", "2026-11-02"
    a = dict(A)
    if ofz_last:
        a["ofz_y0"], a["ofz_D"] = round(ofz_last["ofz_y0"], 4), round(ofz_last["ofz_D"], 3)
    if args.live:
        try:
            a.update(live_yields())
        except Exception as e:
            print(f"  ! live refresh failed ({e}); keeping defaults")
    for i, row in enumerate(ASSUMPTIONS):  # push refreshed values into the table written to Excel
        if row[0] in a:
            ASSUMPTIONS[i] = (row[0], row[1], a[row[0]], *row[3:])

    res = model(a, v)
    R = asset_corr(corr)
    out = pd.DataFrame({
        "Return local": res["loc"], "Return RUB": res["rub"], "Return USD": res["usd"],
        "Vol USD": {k: x / 100 for k, x in res["vol_usd"].items()},
        "Vol RUB (ref)": {k: x / 100 for k, x in res["vol_rub"].items()},
    }).loc[[n for n, *_ in ASSETS]]
    pd.set_option("display.width", 200)
    print((out * 100).round(2).to_string())
    print(f"\nd (RUB depreciation) = {res['d']:.2%}  [PPP {res['ppp']:.2%}, UIP {res['uip']:.2%}];  "
          f"neutral key rate {res['neutral']:.2%};  EPS growth {res['g']:.2%}")
    build_workbook(a, v, res, R, args.out, {"vol_file": vol_file, "vol_date": vol_date})
    out.to_csv(args.out.replace(".xlsx", "_python.csv"))
    print(f"Wrote {args.out}")


if __name__ == "__main__":
    main()
