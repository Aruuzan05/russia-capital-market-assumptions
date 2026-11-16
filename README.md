# Russia Capital Market Assumptions

10-year expected returns, volatilities, and correlations for Russian assets — built for the KEEPP Russia Monte Carlo simulator. These replace J.P. Morgan and BlackRock assumptions, neither of which covers Russian markets.

## Results (horizon 10 years, base currency USD)

| Asset | Return (USD) | Vol (USD) |
|---|---|---|
| RUB cash | 5.01% | 19.5% |
| CNY cash | 4.00% | 5.2% |
| OFZ floaters (ОФЗ-ПК) | 5.30% | 20.2% |
| OFZ fixed (ОФЗ-ПД) | 10.10% | 25.6% |
| RUB corporate IG | 9.14% | 24.0% |
| OFZ in CNY | 7.87% | 10.4% |
| CNY corporate IG | 8.03% | 6.0% |
| Russian equities | 12.54% | 28.0% |
| Gold | 4.04% | 18.7% |

Correlation matrix is positive-definite (smallest eigenvalue 0.019).

## Structure

```
russia-capital-market-assumptions/
├── russia_vol.py          # Step 1: downloads MOEX/CBR history, computes realised stats
├── russia_returns.py      # Step 2: builds return model, writes russia_cma.xlsx
├── data/
│   ├── russia_vol_inputs.xlsx     # realised vol/correlation inputs (output of russia_vol.py)
│   ├── russia_cma.xlsx            # full CMA workbook (output of russia_returns.py)
│   └── russia_volatility_report.xlsx
└── docs/
    └── russia-cma-methodology.md  # full methodology writeup
```

## Usage

```bash
# Step 1 — refresh realised statistics from MOEX/CBR
python russia_vol.py

# Step 2 — build return model (use --live to refresh OFZ/corporate yields)
python russia_returns.py
python russia_returns.py --live
```

## Methodology

See [`docs/russia-cma-methodology.md`](docs/russia-cma-methodology.md) for the full derivation: building blocks for each asset class (rate paths, yield paths, Grinold-Kroner equity model, FX drift), volatility methodology (RVI blend, duration × yield vol, FX overlay), and correlation estimation (Ledoit-Wolf shrinkage).

## Data sources

- MOEX ISS (equity prices, bond yields, RVI implied volatility)
- Bank of Russia (official FX rates, OFZ yield curves)
- CBR key rate path
