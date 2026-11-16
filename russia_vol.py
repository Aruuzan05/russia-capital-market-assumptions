"""
russia_vol.py — volatility & correlation inputs for the Russian version of the KEEPP simulator.

Pulls free public data (MOEX ISS + Bank of Russia official FX), then computes:
  1. Equity vol: realized (RUB and USD) vs RVI implied, variance risk premium, RVI futures curve
  2. Pavel #1 — realized vol of RGBI total return converted to USD (20/30/90d + long-run)
  3. Pavel #2 — vol of OFZ yield changes (MOVE-style, realized), and D x sigma(dy) cross-check
  4. Pavel #3 — duration-adjusted FX: how much of the USD vol of OFZ comes from USDRUB
  5. CNY bonds — RUCNYTR vol in CNY and USD, implied yield vol, applied to the OFZ-CNY curve
  6. Monthly USD (and RUB) correlation matrix, Ledoit-Wolf shrunk, for the app's Σ

Run on a machine with internet access (MOEX/CBR are not reachable from Claude's sandbox):
    pip install pandas numpy requests openpyxl
    python russia_vol.py                 # live data -> russia_vol_inputs.xlsx
    python russia_vol.py --demo          # synthetic data, just to see the output layout
Options: --start 2013-01-01  --out file.xlsx  --cny-basket cny_basket_summary.csv
"""
from __future__ import annotations
import argparse, sys, time, xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path
import numpy as np

_HERE = Path(__file__).parent
import pandas as pd

ISS = "https://iss.moex.com/iss"
TD = 252  # trading days / year

# Series used. Codes marked (verify) should be checked on moex.com/ru/indices — the script
# skips any code that returns no data and tells you, it never silently substitutes.
INDICES = {
    "EQ_RUB_TR": ["MCFTR", "IMOEX"],        # MOEX Russia total return (fallback: price index)
    "EQ_USD":    ["RTSI"],                  # RTS = MOEX Russia in USD; RVI is implied vol of THIS
    "RVI":       ["RVI"],
    "OFZ_TR":    ["RGBITR"],                # OFZ total return, duration > 1y
    "OFZ_PX":    ["RGBI"],                  # price index; history carries YIELD and DURATION
    "CORP_TR":   ["RUCBTRNS", "RUCBITR"],   # RUB corporate bonds TR (verify which code you want)
    "CNY_TR":    ["RUCNYTR"],               # Russian CNY-denominated bonds TR, since 14.11.2022
}
CBR_FX = {"USDRUB": "R01235", "CNYRUB": "R01375"}  # official CBR fixings (exchange USD trading stopped 06.2024)
GOLD = ("currency", "selt", "CETS", "GLDRUB_TOM")


# ----------------------------------------------------------------------------- data access
def _get(url: str, params: dict | None = None, tries: int = 3):
    import requests
    for k in range(tries):
        try:
            r = requests.get(url, params=params, timeout=30)
            r.raise_for_status()
            return r
        except Exception as e:  # noqa
            if k == tries - 1:
                raise
            time.sleep(1.5 * (k + 1))


def iss_history(path: str, start: str, cols=("TRADEDATE", "CLOSE")) -> pd.DataFrame:
    """Paged ISS history download. Returns DataFrame indexed by date with the requested cols present."""
    frames, pos = [], 0
    while True:
        j = _get(f"{ISS}/{path}.json", {"from": start, "start": pos, "iss.meta": "off"}).json()
        h = j["history"]
        df = pd.DataFrame(h["data"], columns=h["columns"])
        if df.empty:
            break
        frames.append(df)
        cur = j.get("history.cursor")
        if cur and cur["data"]:
            idx, total, size = cur["data"][0]
            pos = idx + size
            if pos >= total:
                break
        else:
            pos += len(df)
            if len(df) < 100:
                break
    if not frames:
        return pd.DataFrame()
    df = pd.concat(frames)
    df["TRADEDATE"] = pd.to_datetime(df["TRADEDATE"])
    keep = [c for c in cols if c in df.columns]
    return df[keep].drop_duplicates("TRADEDATE").set_index("TRADEDATE").sort_index()


def load_index(codes: list[str], start: str) -> tuple[str | None, pd.DataFrame]:
    for c in codes:
        try:
            df = iss_history(f"history/engines/stock/markets/index/securities/{c}", start,
                             cols=("TRADEDATE", "CLOSE", "YIELD", "DURATION"))
            if not df.empty and df["CLOSE"].notna().sum() > 50:
                return c, df
        except Exception as e:
            print(f"  ! {c}: {e}")
    return None, pd.DataFrame()


def load_cbr(code: str, start: str) -> pd.Series:
    d1 = pd.Timestamp(start).strftime("%d/%m/%Y")
    d2 = date.today().strftime("%d/%m/%Y")
    r = _get("https://www.cbr.ru/scripts/XML_dynamic.asp",
             {"date_req1": d1, "date_req2": d2, "VAL_NM_RQ": code})
    root = ET.fromstring(r.content)
    out = {}
    for rec in root.findall("Record"):
        nominal = float(rec.find("Nominal").text.replace(",", "."))
        val = float(rec.find("Value").text.replace(",", "."))
        out[pd.to_datetime(rec.get("Date"), dayfirst=True)] = val / nominal
    return pd.Series(out).sort_index()


def load_rvi_futures() -> pd.DataFrame:
    j = _get(f"{ISS}/engines/futures/markets/forts/securities.json", {"iss.meta": "off"}).json()
    sec = pd.DataFrame(j["securities"]["data"], columns=j["securities"]["columns"])
    md = pd.DataFrame(j["marketdata"]["data"], columns=j["marketdata"]["columns"])
    sec = sec[sec["ASSETCODE"].astype(str).str.upper() == "RVI"]
    if sec.empty:
        return pd.DataFrame()
    m = sec.merge(md, on="SECID", how="left", suffixes=("", "_md"))
    cols = [c for c in ["SECID", "SHORTNAME", "LASTTRADEDATE", "LAST", "SETTLEPRICE", "PREVSETTLEPRICE", "OPENPOSITION"]
            if c in m.columns]
    return m[cols].sort_values("LASTTRADEDATE")


def load_live(start: str) -> dict:
    data, used = {}, {}
    for key, codes in INDICES.items():
        print(f"- {key}: trying {codes}")
        code, df = load_index(codes, start)
        if code is None:
            print(f"  ! no data for {key}; it will be skipped")
            continue
        used[key] = code
        data[key] = df
    for k, c in CBR_FX.items():
        print(f"- {k}: CBR {c}")
        data[k] = load_cbr(c, start).to_frame("CLOSE")
    try:
        eng, mkt, brd, sec = GOLD
        g = iss_history(f"history/engines/{eng}/markets/{mkt}/boards/{brd}/securities/{sec}", start,
                        cols=("TRADEDATE", "CLOSE", "WAPRICE"))
        if not g.empty:
            g["CLOSE"] = g["CLOSE"].fillna(g.get("WAPRICE"))
            data["GOLD_RUB"] = g[["CLOSE"]]
            used["GOLD_RUB"] = sec
    except Exception as e:
        print(f"  ! gold: {e}")
    try:
        data["RVI_FUT"] = load_rvi_futures()
    except Exception as e:
        print(f"  ! RVI futures: {e}")
        data["RVI_FUT"] = pd.DataFrame()
    data["_used"] = used
    return data


# ----------------------------------------------------------------------------- demo data
def load_demo(start: str, seed: int = 7) -> dict:
    """Synthetic series with realistic-ish shape so the workbook layout can be inspected offline.
    NUMBERS FROM --demo ARE NOT ESTIMATES. Do not put them in the app."""
    rng = np.random.default_rng(seed)
    idx = pd.bdate_range(start, "2026-10-30")
    n = len(idx)
    regime = np.where(rng.random(n) < 0.04, 2.5, 1.0)
    regime = pd.Series(regime).rolling(20, min_periods=1).max().to_numpy()
    fx = rng.normal(0.0002, 0.008, n) * regime                    # dln USDRUB
    eq_usd = rng.normal(0.0002, 0.016, n) * regime - 0.5 * fx
    eq_rub = eq_usd + fx
    dy = rng.normal(0, 0.0009, n) * regime + 0.02 * fx            # OFZ yield change
    y = 0.09 + np.cumsum(dy)
    dur = np.full(n, 5.5)
    ofz = y / TD - dur * dy
    corp = 0.6 * ofz + rng.normal(0.0004, 0.002, n)
    cny = rng.normal(0.0003, 0.0018, n)
    cnyrub = 0.9 * fx + rng.normal(0, 0.003, n)
    gold = rng.normal(0.0003, 0.010, n) + fx
    rv = pd.Series(eq_usd).rolling(21).std().bfill().to_numpy() * np.sqrt(TD) * 100 + 4
    lvl = lambda r: 100 * np.exp(np.cumsum(r))
    f = lambda a: pd.DataFrame({"CLOSE": a}, index=idx)
    d = {
        "EQ_RUB_TR": f(lvl(eq_rub)), "EQ_USD": f(lvl(eq_usd)), "RVI": f(rv),
        "OFZ_TR": f(lvl(ofz)),
        "OFZ_PX": pd.DataFrame({"CLOSE": lvl(ofz - y / TD), "YIELD": y * 100, "DURATION": dur * 365}, index=idx),
        "CORP_TR": f(lvl(corp)),
        "CNY_TR": f(lvl(cny)).loc["2022-11-14":],
        "USDRUB": f(60 * np.exp(np.cumsum(fx))), "CNYRUB": f(9 * np.exp(np.cumsum(cnyrub))),
        "GOLD_RUB": f(lvl(gold)),
        "RVI_FUT": pd.DataFrame({"SECID": ["VIX6", "VIZ6", "VIF7"], "LASTTRADEDATE": ["2026-11-19", "2026-12-17", "2027-01-21"],
                                 "SETTLEPRICE": [28.5, 29.4, 30.1]}),
        "_used": {"(demo)": "synthetic"},
    }
    return d


# ----------------------------------------------------------------------------- maths
def ann_vol_daily(r: pd.Series) -> float:
    return float(r.std() * np.sqrt(TD))


def ann_vol_monthly(level: pd.Series) -> float:
    m = np.log(level.resample("ME").last()).diff().dropna()
    return float(m.std() * np.sqrt(12)) if len(m) > 6 else np.nan


def ewma_vol(r: pd.Series, lam: float = 0.94) -> float:
    v = (r ** 2).ewm(alpha=1 - lam, adjust=False).mean()
    return float(np.sqrt(v.iloc[-1] * TD))


def window_table(name: str, level: pd.Series, windows=(20, 30, 90, 250)) -> dict:
    level = level.dropna()
    r = np.log(level).diff().dropna()
    row = {"series": name, "from": level.index[0].date(), "to": level.index[-1].date()}
    for w in windows:
        row[f"{w}d %"] = ann_vol_daily(r.iloc[-w:]) * 100
    row["EWMA(0.94) %"] = ewma_vol(r) * 100
    row["full daily %"] = ann_vol_daily(r) * 100
    row["full monthly %"] = ann_vol_monthly(level) * 100
    row["post-2022 monthly %"] = ann_vol_monthly(level.loc["2022-04-01":]) * 100
    row["ex-2022 monthly %"] = ann_vol_monthly(level.drop(level.loc["2022-02-01":"2022-12-31"].index)) * 100
    return row


def to_usd(level_rub: pd.Series, usdrub: pd.Series) -> pd.Series:
    fx = usdrub.reindex(level_rub.index).ffill()
    return (level_rub / fx).dropna()


def ledoit_wolf(X: np.ndarray) -> tuple[np.ndarray, float]:
    """Ledoit-Wolf shrinkage of the sample covariance toward constant-correlation target."""
    X = X - X.mean(0)
    T, N = X.shape
    S = X.T @ X / T
    sd = np.sqrt(np.diag(S))
    R = S / np.outer(sd, sd)
    rbar = (R.sum() - N) / (N * (N - 1))
    F = rbar * np.outer(sd, sd)
    np.fill_diagonal(F, np.diag(S))
    Y = X ** 2
    pi_mat = Y.T @ Y / T - S ** 2
    pi_hat = pi_mat.sum()
    term1 = ((X ** 3).T @ X) / T
    help_ = np.diag(S)[:, None] * S
    theta_ii = term1 - help_
    theta_jj = theta_ii.T
    rho_off = rbar / 2 * ((np.sqrt(np.diag(S))[None, :] / sd[:, None]) * theta_ii
                          + (sd[:, None] / sd[None, :]) * theta_jj)
    np.fill_diagonal(rho_off, 0)
    rho_hat = np.trace(pi_mat) + rho_off.sum()
    gamma = ((F - S) ** 2).sum()
    k = (pi_hat - rho_hat) / gamma if gamma > 0 else 0
    delta = float(np.clip(k / T, 0, 1))
    return delta * F + (1 - delta) * S, delta


def corr_from_cov(C):
    sd = np.sqrt(np.diag(C))
    return C / np.outer(sd, sd)


def pos_def_fix(R: np.ndarray, eps=1e-8) -> np.ndarray:
    w, V = np.linalg.eigh((R + R.T) / 2)
    R2 = V @ np.diag(np.maximum(w, eps)) @ V.T
    d = np.sqrt(np.diag(R2))
    return R2 / np.outer(d, d)


# ----------------------------------------------------------------------------- analysis
def analyse(d: dict, cny_basket: pd.DataFrame | None) -> dict:
    out = {}
    usd = d["USDRUB"]["CLOSE"]
    cnyrub = d["CNYRUB"]["CLOSE"]
    usdcny = (usd / cnyrub.reindex(usd.index).ffill()).dropna()  # CNY per USD

    # ---- 0. assemble level series in RUB and USD
    lv_rub, lv_usd = {}, {}
    if "EQ_RUB_TR" in d:
        lv_rub["Russian equity"] = d["EQ_RUB_TR"]["CLOSE"]
        lv_usd["Russian equity"] = to_usd(lv_rub["Russian equity"], usd)
    if "OFZ_TR" in d:
        lv_rub["OFZ (RGBITR)"] = d["OFZ_TR"]["CLOSE"]
        lv_usd["OFZ (RGBITR)"] = to_usd(lv_rub["OFZ (RGBITR)"], usd)
    if "CORP_TR" in d:
        lv_rub["RUB corporates"] = d["CORP_TR"]["CLOSE"]
        lv_usd["RUB corporates"] = to_usd(lv_rub["RUB corporates"], usd)
    if "CNY_TR" in d:
        cny_cny = d["CNY_TR"]["CLOSE"]
        lv_usd["CNY bonds (RUCNYTR)"] = (cny_cny / usdcny.reindex(cny_cny.index).ffill()).dropna()
        lv_rub["CNY bonds (RUCNYTR)"] = (cny_cny * cnyrub.reindex(cny_cny.index).ffill()).dropna()
    if "GOLD_RUB" in d:
        lv_rub["Gold"] = d["GOLD_RUB"]["CLOSE"]
        lv_usd["Gold"] = to_usd(lv_rub["Gold"], usd)
    rub_cash_proxy = None
    lv_usd["RUB cash (FX only)"] = 1 / usd           # RUB cash vol in USD ≈ FX vol
    lv_usd["CNY cash (FX only)"] = 1 / usdcny
    lv_rub["USDRUB"] = usd
    lv_rub["CNYRUB"] = cnyrub

    rows = []
    for k, s in lv_rub.items():
        rows.append({"ccy": "RUB", **window_table(k, s)})
    for k, s in lv_usd.items():
        rows.append({"ccy": "USD", **window_table(k, s)})
    if "CNY_TR" in d:
        rows.append({"ccy": "CNY", **window_table("CNY bonds (RUCNYTR)", d["CNY_TR"]["CLOSE"])})
    out["vol_windows"] = pd.DataFrame(rows)

    # ---- 1. RVI vs realized RTS vol, variance risk premium
    if "RVI" in d and "EQ_USD" in d:
        rvi = d["RVI"]["CLOSE"].dropna()
        rts = d["EQ_USD"]["CLOSE"].dropna()
        r = np.log(rts).diff()
        fwd_rv = (r[::-1].rolling(21).std()[::-1].shift(-1) * np.sqrt(TD) * 100)   # next-21d realized
        both = pd.concat({"RVI": rvi, "fwd_realized": fwd_rv}, axis=1).dropna()
        both["VRP"] = both["RVI"] - both["fwd_realized"]
        q = rvi.quantile([0.1, 0.5, 0.9])
        rvi_stats = pd.DataFrame({
            "metric": ["RVI latest", "RVI 1y mean", "RVI 3y mean", "RVI full-sample mean", "RVI median",
                       "RVI 10th pct", "RVI 90th pct", "Mean VRP (RVI − next-month realized)",
                       "Median VRP", "RVI mean − mean VRP (implied-based LT vol)",
                       "RTS realized, full monthly", "RTS realized, post-2022 monthly"],
            "value_%": [rvi.iloc[-1], rvi.iloc[-TD:].mean(), rvi.iloc[-3 * TD:].mean(), rvi.mean(), q[0.5],
                        q[0.1], q[0.9], both["VRP"].mean(), both["VRP"].median(),
                        rvi.mean() - both["VRP"].mean(),
                        ann_vol_monthly(rts) * 100, ann_vol_monthly(rts.loc["2022-04-01":]) * 100],
        })
        out["rvi_stats"] = rvi_stats
        out["rvi_series"] = both.resample("W").last()
    if "RVI_FUT" in d and not d["RVI_FUT"].empty:
        out["rvi_futures"] = d["RVI_FUT"]

    # ---- 2+3. OFZ: yield vol, D×σ(Δy), FX decomposition
    if "OFZ_PX" in d and {"YIELD", "DURATION"} <= set(d["OFZ_PX"].columns) and "OFZ_TR" in d:
        px = d["OFZ_PX"].dropna(subset=["YIELD"])
        y = px["YIELD"] / 100
        # ISS publishes index Macaulay duration in days; guard in case a feed gives years
        D_mac = px["DURATION"] / (365 if px["DURATION"].median() > 50 else 1)
        D_mod = D_mac / (1 + y)
        dy = y.diff().dropna()
        tr = np.log(d["OFZ_TR"]["CLOSE"]).diff()
        fx = np.log(usd.reindex(tr.index).ffill()).diff()
        rows = []
        for lab, w in [("20d", 20), ("30d", 30), ("90d", 90), ("1y", 250), ("full", None)]:
            sl = slice(-w, None) if w else slice(None)
            s_dy = dy.iloc[sl].std() * np.sqrt(TD)
            D = D_mod.iloc[-w:].mean() if w else D_mod.mean()
            rl = tr.dropna().iloc[sl]
            fxw = fx.reindex(rl.index)
            ru = (rl - fxw).dropna()
            s_loc, s_fx, s_usd = rl.std() * np.sqrt(TD), fxw.std() * np.sqrt(TD), ru.std() * np.sqrt(TD)
            rho = np.corrcoef(rl.loc[ru.index], fxw.loc[ru.index])[0, 1]
            rho_dy_fx = pd.concat([dy, fx], axis=1, sort=True).dropna().iloc[sl].corr().iloc[0, 1]
            rows.append({
                "window": lab,
                "σ(Δy) bp/yr": s_dy * 1e4,
                "mod. duration": D,
                "D×σ(Δy) %": D * s_dy * 100,
                "realized RGBITR RUB %": s_loc * 100,
                "USDRUB vol %": s_fx * 100,
                "ρ(OFZ RUB ret, ΔlnUSDRUB)": rho,
                "ρ(Δy, ΔlnUSDRUB)": rho_dy_fx,
                "RGBITR in USD % (Pavel #1)": s_usd * 100,
                "duration-adj. FX formula %": np.sqrt(s_loc ** 2 + s_fx ** 2 - 2 * rho * s_loc * s_fx) * 100,
                "share of USD variance from FX": (s_fx ** 2 - rho * s_loc * s_fx) / s_usd ** 2,
            })
        out["ofz_decomp"] = pd.DataFrame(rows)
        out["ofz_rolling"] = pd.DataFrame({
            "σ(Δy) 30d bp": dy.rolling(30).std() * np.sqrt(TD) * 1e4,
            "RGBITR USD 30d %": (tr - fx).rolling(30).std() * np.sqrt(TD) * 100,
            "RGBITR USD 90d %": (tr - fx).rolling(90).std() * np.sqrt(TD) * 100,
            "yield %": y * 100, "mod dur": D_mod,
        }).dropna().resample("W").last()

    # ---- 5. CNY bonds: implied yield vol from RUCNYTR, applied to the sovereign CNY curve
    if "CNY_TR" in d and cny_basket is not None:
        r = np.log(d["CNY_TR"]["CLOSE"]).diff().dropna()
        s_idx = ann_vol_monthly(d["CNY_TR"]["CLOSE"])
        D_corp = float(cny_basket.loc["Corp AAA", "mod_dur"]) if "Corp AAA" in cny_basket.index else 1.8
        s_dy = s_idx / D_corp
        s_fx_cny = ann_vol_monthly(1 / usdcny.loc["2022-11-14":])
        s_cnyrub = ann_vol_monthly(cnyrub.loc["2022-11-14":])
        rows = []
        for g, rr in cny_basket.iterrows():
            s_loc = rr["mod_dur"] * s_dy
            rows.append({"bucket": g, "YTM %": rr["ytm"] * 100, "mod dur": rr["mod_dur"],
                         "σ in CNY %": s_loc * 100,
                         "σ in USD % (ρ=0)": np.sqrt(s_loc ** 2 + s_fx_cny ** 2) * 100,
                         "σ in RUB % (ρ=0)": np.sqrt(s_loc ** 2 + s_cnyrub ** 2) * 100})
        out["cny_curve"] = pd.DataFrame(rows)
        out["cny_note"] = pd.DataFrame({"item": ["RUCNYTR monthly vol (CNY) %", "assumed index mod. duration",
                                                 "implied σ(Δy) bp/yr", "USD/CNY vol since 11.2022 %",
                                                 "CNY/RUB vol since 11.2022 %"],
                                        "value": [s_idx * 100, D_corp, s_dy * 1e4, s_fx_cny * 100, s_cnyrub * 100]})

    # ---- 6. correlations (monthly log returns, USD and RUB), LW-shrunk, PD-repaired
    for ccy, lv in (("USD", lv_usd), ("RUB", lv_rub)):
        m = pd.DataFrame({k: np.log(v.resample("ME").last()).diff() for k, v in lv.items()
                          if k not in ("USDRUB", "CNYRUB")})
        full = m.dropna(axis=1, thresh=36).dropna()
        if full.shape[1] >= 2 and len(full) >= 24:
            C, delta = ledoit_wolf(full.to_numpy())
            R = pos_def_fix(corr_from_cov(C))
            out[f"corr_{ccy}"] = pd.DataFrame(R, index=full.columns, columns=full.columns)
            out[f"corr_{ccy}_meta"] = pd.DataFrame({"item": ["months", "from", "to", "LW shrinkage δ", "note"],
                                                    "value": [len(full), full.index[0].date(), full.index[-1].date(),
                                                              round(delta, 3),
                                                              "series with <36 months (e.g. RUCNYTR) excluded; set by judgement"]})
    return out


def write_xlsx(res: dict, used: dict, path: str, demo: bool):
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        readme = pd.DataFrame({"README": [
            "DEMO DATA — synthetic, not estimates" if demo else "Live data from MOEX ISS and Bank of Russia",
            f"Codes used: {used}",
            "All vols annualised, %. Daily: σ×√252. Monthly: σ×√12 of month-end log returns.",
            "For the app's long-horizon σ use the MONTHLY columns (they include crisis months and are robust to "
            "stale daily prices); daily 20/30/90d are for monitoring current regime.",
            "USD series = RUB level / CBR USDRUB; CNY bonds in USD use USD/CNY cross from CBR fixings.",
        ]})
        readme.to_excel(xw, "README", index=False)
        for k in ["vol_windows", "rvi_stats", "rvi_futures", "ofz_decomp", "cny_curve", "cny_note",
                  "corr_USD", "corr_USD_meta", "corr_RUB", "corr_RUB_meta", "ofz_rolling", "rvi_series"]:
            if k in res:
                df = res[k]
                df.to_excel(xw, k[:31], index=k.startswith("corr_") and not k.endswith("meta")
                            or k in ("ofz_rolling", "rvi_series"))
    print(f"\nWrote {path}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--start", default="2013-01-01")
    ap.add_argument("--out", default=str(_HERE / "russia_vol_inputs.xlsx"))
    ap.add_argument("--cny-basket", default=str(_HERE / "cny_basket_summary.csv"))
    ap.add_argument("--demo", action="store_true")
    a = ap.parse_args()
    d = load_demo(a.start) if a.demo else load_live(a.start)
    try:
        basket = pd.read_csv(a.cny_basket, index_col=0)
    except Exception:
        basket = None
        print("  (no CNY basket csv — CNY curve sheet skipped)")
    res = analyse(d, basket)
    pd.set_option("display.width", 220, "display.max_columns", 30, "display.precision", 3)
    for k in ("vol_windows", "rvi_stats", "ofz_decomp", "cny_curve"):
        if k in res:
            print(f"\n=== {k}\n", res[k].to_string(index=False))
    write_xlsx(res, d.get("_used", {}), a.out, a.demo)


if __name__ == "__main__":
    main()
