# KEEPP Russia — capital market assumptions: methodology (3 Nov 2026)

How `russia_returns.py` turns market data and committee assumptions into the expected returns, volatilities and correlations the simulator needs for the Russian portfolio. These replace J.P. Morgan and BlackRock, neither of which covers Russian assets.

**Files** (in `OneDrive\finance\Russian_MC`)

| File | What it is |
|---|---|
| `russia_vol.py` | Downloads MOEX/CBR history and writes the realised statistics to `russia_vol_inputs.xlsx` (run 02.11.2026). |
| `russia_returns.py` | Builds the return model, converts the realised statistics into forward volatilities, and writes `russia_cma.xlsx`. `--live` refreshes OFZ and corporate yields from MOEX. |
| `russia_cma.xlsx` | Every number is an Excel formula on the Assumptions sheet, so changing an input recalculates the whole workbook. Sheets: Summary, App inputs, Assumptions, Paths, Volatility, Correlation USD, Sensitivity, Sources. |

Checks run: the Excel results match the Python results to within 6×10⁻⁷ (the difference comes from rounding inputs to 4 decimals). There are 312 formulas and none returns an error. The correlation matrix is positive-definite (smallest eigenvalue 0.019).

---

## 1. Results (horizon 10 years, base currency USD)

| Asset | Return, local | Return, RUB | **Return, USD** | **Vol, USD** | Vol, RUB (ref.) |
|---|---|---|---|---|---|
| RUB cash | 8.65% | 8.65% | **5.01%** | **19.5%** | 0.5% |
| CNY cash | 4.00% | 7.60% | **4.00%** | **5.2%** | 19.5% |
| OFZ floaters (ОФЗ-ПК) | 8.95% | 8.95% | **5.30%** | **20.2%** | 1.5% |
| OFZ fixed (ОФЗ-ПД) | 13.91% | 13.91% | **10.10%** | **25.6%** | 8.9% |
| RUB corporate IG | 12.92% | 12.92% | **9.14%** | **24.0%** | 6.0% |
| OFZ in CNY | 7.87% | 11.60% | **7.87%** | **10.4%** | 21.8% |
| CNY corporate IG | 8.03% | 11.77% | **8.03%** | **6.0%** | 19.9% |
| Russian equities | 16.44% | 16.44% | **12.54%** | **28.0%** | 20.7% |
| Gold | 4.04% (USD) | 7.64% | **4.04%** | **18.7%** | 24.8% |

RUB depreciation against USD comes out at **d = 3.46% a year**. The neutral key rate is 8.5% and nominal EPS growth is 4.49%.

All returns are **compound** (geometric) annual returns, which is the convention the simulator's `Return %` column uses (the engine sets m = ln(1+g)/12). The workbook also shows the arithmetic mean, g + σ²/2.

---

## 2. Return building blocks

### 2.1 Notation and horizon averaging

Each asset gets a year-by-year local-currency return r₁ … r₁₀ (sheet **Paths**, 2027–2036). The 10-year figure is the geometric average:

```
R = exp( Σₜ ln(1 + rₜ) / 10 ) − 1
```

### 2.2 RUB cash and floaters

```
key₁ = 11.5%   (midpoint of CBR's 10.5–12.5% forecast for the 2027 average)
key₂ = 9.5%    (glide path)
keyₜ = neutral = π_RU + r* = 4.5% + 4.0% = 8.5%   for t ≥ 3
cashₜ    = keyₜ + MMF spread (−0.25%)        → R = 8.65%
floaterₜ = cashₜ + floater spread (+0.30%)   → R = 8.95%
```

**Why:** money-market funds track RUSFAR, which tracks the key rate. The key rate was held at 14% on 23 Oct 2026, with the next meeting on 18 Dec. The CBR's 2027 forecast range anchors year 1. After that the rate converges to the neutral rate: long-run inflation plus a real neutral rate. Both are committee levers.

### 2.3 OFZ, corporates, CNY bonds: constant-duration bond fund

The fund holds a constant modified duration D. Its yield moves linearly from today's level y₀ to an equilibrium y* over N = 5 years:

```
yₜ = y₀ + (y* − y₀)·t/N     for t ≤ N,   otherwise y*

rₜ = yₜ₋₁  −  D·(yₜ − yₜ₋₁)  +  ½·D²·(yₜ − yₜ₋₁)²  −  expected loss
     carry    price effect      convexity (C ≈ D²)
```

| Line | y₀ | D | y* | Loss |
|---|---|---|---|---|
| OFZ fixed (RGBI) | 15.95% (RGBI YIELD, 02.11.2026) | 3.90 | neutral 8.5% + term premium 1.0% = 9.5% | 0 |
| RUB corp IG (RUCBTRNS) | 16.0% | 2.0 | OFZ* + 1.0% spread = 10.5% | 0.30% |
| OFZ in CNY (ОФЗ 29/33/36) | 7.87% | 4.52 | flat (no view) | 0 |
| CNY corp AAA | 8.33% | 1.80 | flat | 0.30% |

Expected loss = default probability × loss given default ≈ 0.5% × 60%.

**Worked example, OFZ year 1:** the yield falls 1.29 pp, so r₁ = 15.95% + 3.90 × 1.29% + ½ × 3.90² × 1.29%² ≈ 21.1%. Years 1–5 run 21% → 16% as the falling yield adds price gains on top of the carry. Years 6–10 earn the 9.5% equilibrium yield. The average is 13.9%.

**This explains two results the committee may question:**

- **OFZ (13.9%) is well above cash (8.65%).** The model assumes the yield-curve normalisation the CBR forecast implies. A fund holding duration today locks in about 16% and gains when yields fall. If the committee believes rates stay high, raise y* or N.
- **OFZ beats corporates (12.9%).** The corporate index is shorter (D = 2.0 vs 3.9), so it gains less from falling yields. Its spread over OFZ is currently close to zero and is assumed to widen to 1.0%. With rates flat, the corporates would come out ahead. Run `--live` to replace the 16.0% placeholder with the index's actual yield.

This is the standard horizon logic: over about 2 × duration years, the starting yield is a good estimate of a bond fund's return, because price changes and reinvestment offset each other (Leibowitz).

**CNY bond baskets** come from the analyst's snapshot `CNYbonds201026.xlsx`. I kept one quote per bond (the CNY-settled line), measured duration to the put date where there is one (оферта), and used equal weights. I excluded ПолюсБ1P10, Полюс Б1P8 and НорНик1P11 as stale.

### 2.4 Russian equities: Grinold-Kroner, with P/E and dividend yield kept consistent

```
g      = (1+π_RU)(1+g_real)(1−dilution) − 1 = 1.045 × 1.01 × 0.99 − 1 = 4.49%
PEₜ    = PE₀ · (PE_T / PE₀)^(t/10),   PE₀ = 5.5x → PE_T = 7.0x   (+2.4% a year from re-rating)
DYₜ    = payout / PEₜ₋₁,   payout = 0.55   → DY falls from 10.0% to 8.0%
rₜ     = (1+g)·PEₜ/PEₜ₋₁ − 1  +  DYₜ·(1+g)
```

Result: **16.44% in RUB**, made up of roughly 9% dividends, 4.5% growth and 2.4% re-rating.

Deriving DY from payout / P/E keeps the two consistent: as the market re-rates, the dividend yield falls. Calibration:

- A starting DY of 10% matches BCS's 12-month estimate of about 10% (against a historical average of about 6.5%).
- PE₀ = 5.5x sits between SberCIB's ~4.9x and T-Investments' 5.8x trailing.
- PE_T = 7x is mid-range for 2017–2019 (7.8 / 8.5 / 6.2 per the MOEX factsheet).

The equity premium comes out at about 2.5% over OFZ and about 8% over cash.

### 2.5 Gold

The return is set in USD, because gold is priced globally: (1 + π_US)(1 + real) − 1 = 1.025 × 1.015 − 1 = **4.04%**.

### 2.6 Currency conversion

```
PPP:  (1+π_RU)/(1+π_US) − 1             = 1.95%
UIP:  (1+cash_RUB)/(1+cash_USD) − 1     = 4.97%    (cash_RUB = 8.65%, cash_USD = 3.5%)
d   = w·PPP + (1−w)·UIP,  w = 0.5        = 3.46%

RUB asset → USD:   (1 + r_RUB)/(1 + d) − 1
CNY asset → USD:   (1 + r_CNY)/(1 + c) − 1,   c = 0 (no view; PPP alone would give ~1% CNY appreciation)
USD/CNY asset → RUB: (1 + r_USD)(1 + d) − 1
```

**Why a blend:** pure UIP assumes a USD investor earns none of the ruble's interest-rate premium (carry). Pure PPP assumes they earn all of it. History lies between the two. **d is the most important input in USD terms:** each extra 1 pp of depreciation costs about 1 pp of USD return on every RUB line. The **Sensitivity** sheet tabulates d from 2% to 8%.

---

## 3. Volatility

The realised statistics come from `russia_vol_inputs.xlsx` (the `russia_vol.py` run of 02.11.2026). They use **monthly** log returns annualised by √12, because:

- monthly windows keep the crisis months (2014, 2022);
- they are robust to stale daily prices;
- they reflect the 2022 market closure.

The daily 20/30/90-day windows are for monitoring the current regime, not for long-run assumptions.

### 3.1 Equities: RVI as the Russian VIX

RVI is the 30-day implied volatility of RTS options, quoted in USD. It overstates the volatility that actually follows, because of the variance risk premium (VRP):

```
VRP_t = RVI_t − realised RTS vol over the next 21 trading days
Mean RVI 32.81,  mean VRP 5.92  →  implied-based long-run σ = 26.89
RTS realised, full sample monthly = 29.12

σ_eq,USD = ½·(RVI mean − VRP) + ½·RTS realised = ½·26.89 + ½·29.12 = 28.0%
```

The current RVI (31.7) and futures (VIX6 34.6, VIZ6 33.65) are near the long-run mean. The term structure only reaches about two months, so it informs short-term or stress settings, not a 10-year σ. The RVI 90th percentile (48.2) is a natural input for the crisis stress test.

### 3.2 OFZ: Pavel's three substitutes, combined

There is no implied volatility for OFZ, because there is no liquid options market. The forward estimate is built in two steps.

**Step 1 — volatility in RUB from yield volatility (Pavel #2):**

```
σ_RUB = D_now × σ(Δy) = 3.90 × 2.275% = 8.87%
```

The yield-volatility method uses **today's** duration. If RGBI's duration changes, the volatility follows; the index's own history (9.55%) reflects old durations.

**Step 2 — FX overlay (Pavel #3: duration-adjusted FX):**

```
r_USD = r_RUB + f,   f = Δln(1/USDRUB)
σ²_USD = σ²_RUB + σ²_FX + 2ρ·σ_RUB·σ_FX
```

ρ is backed out from the measured monthly volatilities (Pavel #1, RGBITR converted to USD):

```
ρ = (26.13² − 9.55² − 19.45²) / (2·9.55·19.45) = +0.574
```

So OFZ prices and the ruble fall together, and FX **adds** risk. Forward result: σ_USD = √(8.87² + 19.45² + 2·0.574·8.87·19.45) = **25.6%**. Of the USD variance, 84% comes from FX (sheet `ofz_decomp`).

The daily correlation is about 0, against +0.57 monthly. That gap comes from non-synchronous trading, and is another reason to use monthly data.

### 3.3 CNY bonds

RUCNYTR (MOEX's index of Russian CNY bonds, from 14.11.2022) gives a realised volatility of 4.36% in CNY.

```
σ(Δy)_CNY ≈ 4.36 / D_index(2.0) = 2.18%
OFZ-CNY:  σ_CNY = 4.52 × 2.18 = 9.84%      CNY corp: 1.80 × 2.18 = 3.92%
ρ(bond, CNY/USD) = (7.74² − 4.36² − 7.13²)/(2·4.36·7.13) = −0.16    (post-2022 FX, same sample)
σ_USD = √(σ² + 5.21² + 2ρσ·5.21)  →  10.4% and 6.0%
```

CNY/USD volatility is only about 5%, so CNY instruments are the low-volatility, diversifying part of a USD-based Russian portfolio. In RUB terms they are risky (CNY/RUB about 19.5%).

### 3.4 Other lines

| Line | Method |
|---|---|
| RUB cash | √(0.5² + σ_FX²) ≈ σ_FX = 19.5% |
| CNY cash | CNY/USD volatility, 5.2% |
| Floaters | 1.5% in RUB (judgement), FX overlay with ρ = 0.5 |
| RUB corporates | Realised RUCBTRNS in USD, 24.0% (history only from 2018, so it misses 2014) |
| Gold | Realised in USD after 2022, 18.7%. The full GLDRUB_TOM series has gaps, so the full-sample figure is missing; the global long-run gold volatility is about 15–16%. |

---

## 4. Correlations (sheet "Correlation USD")

- **Source:** `corr_USD` from `russia_vol.py`, built from monthly USD returns, Ledoit-Wolf shrinkage towards constant correlation (δ = 0.21), and a positive-definite repair.
- **Mapping to asset lines:**
  - floaters use the RUB-cash factor × 0.98;
  - both CNY lines use the RUCNYTR factor, with OFZ-CNY ↔ CNY corp set to 0.90 by judgement;
  - eigenvalues are floored at 10⁻⁶ and the matrix rescaled to a unit diagonal.
- **Caveat:** the common sample is only 48 months (12.2022–11.2026), because RUCNYTR is the shortest series. In USD, the RUB lines correlate 0.7–0.9 because they share the ruble factor. That's real, not an estimation error.
- **Next step:** estimate a second matrix over 2013–2026 without RUCNYTR, so the RUB lines' correlations include 2014 and 2022. Use the longer one for those lines.

---

## 5. Data sources

| Data | Link |
|---|---|
| MOEX ISS API (all histories) | https://iss.moex.com/iss/reference/ |
| RVI index / methodology / futures | https://www.moex.com/ru/index/RVI · https://www.moex.com/a5737 · https://www.moex.com/ru/derivatives/futures-on-indices/viz6 |
| RTSI, MCFTR | https://www.moex.com/ru/index/RTSI · https://www.moex.com/ru/index/MCFTR |
| RGBI / RGBITR | https://www.moex.com/ru/index/RGBITR |
| RUCBTRNS (corporate) | https://www.moex.com/ru/index/RUCBTRNS |
| RUCNYTR (CNY bonds) | https://www.moex.com/ru/index/RUCNYTR |
| Gold GLDRUB_TOM (ISS) | https://iss.moex.com/iss/history/engines/currency/markets/selt/boards/CETS/securities/GLDRUB_TOM.json |
| CBR FX fixings (XML) | https://www.cbr.ru/development/SXML/ |
| CBR key rate and forecast | https://www.cbr.ru/press/keypr/ |
| CBR macro survey | https://www.cbr.ru/statistics/ddkp/mo_br/ |
| MOEX corporate index factsheet (duration) | https://fs.moex.com/files/19822 |
| MOEX index factsheet (P/E history) | https://fs.moex.com/files/19418 |
| BCS: 12m dividend yield ~10% | https://bcswm.ru/ajax/files/68847/2xiyc034ghlepnlny1lc66wwrss3wv01/ |
| T-Investments 2026 strategy (P/E) | https://www.tbank.ru/invest/research/strategy/2026/stocks/ |
| CNY bond snapshot | `CNYbonds201026.xlsx` (Russian analyst, 20.10.2026) |

**Methods:** Grinold & Kroner (2002), equity building blocks · Leibowitz (1986), bond horizon return · Ledoit & Wolf (2004), constant-correlation shrinkage · Geltner (1993), desmoothing for future real-estate and private lines.

---

## 6. Open items

1. Run `python russia_returns.py --live` on your laptop. It replaces the corporate yield placeholder (16.0%) and duration with RUCBTRNS's live values.
2. Check CNY cash (4%) against RUSFARCNY.
3. Get the USD replacement-bond table from the analyst. Lines still not calibrated: RUB high yield, SMID, ЗПИФ.
4. Committee levers to agree (yellow cells):
   - π_RU
   - 2027 key rate
   - r*
   - term premium
   - P/E target
   - w (PPP vs UIP)
   - N (years for yields to converge)
5. Add a "KEEPP Russia" provider to the app. The sheet **App inputs** is in the app's format (USD, compound Return %, Vol %), and the correlations go in from "Correlation USD".
