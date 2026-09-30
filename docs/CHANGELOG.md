# Change log

## 2026-09-30: Phase 1, company, chart of accounts, driver engine and synthetic actuals

- **Design:** `docs/DESIGN.md` covers the company, workbook, modules, web panel and phases.
- **Company:** `config/company.toml` defines Kestrel Row Lettings Ltd, a fictional letting agent (financial year April to March). It holds the FY2026/27 budget drivers:
  - portfolio and rent growth,
  - fee rates,
  - lettings activity with seasonality,
  - headcount by role, with a planned hire in October,
  - cost drivers.
- **Code:**
  - `fpa/company.py`: calendar helpers.
  - `fpa/accounts.py`: chart of accounts with 22 P&L accounts, groups and subtotals.
  - `fpa/drivers.py`: driver engine (monthly roll-forward, then P&L by account), shared by the budget, the forecast and the simulator.
  - `fpa/ledger.py`: reads and writes the monthly trial balance and KPI exports, with validation for unknown or duplicate accounts, wrong period, missing accounts and non-numeric values.
  - `fpa/simulate.py`: seeded synthetic actuals from April 2025 to September 2026.
- **The prior year** ran on a 11.5% management fee; the budget uses 12%.
- **FY2026/27 H1 events:**
  - rents rising faster than budget,
  - a £4k bad debt in May,
  - a property manager leaving in June and replaced in September, with agency cover,
  - a softer summer lettings market,
  - a portfolio landlord leaving in July (24 properties),
  - a portal price rise of 15% in August.
- **Results:**

  | | Revenue | EBITDA | EBITDA margin |
  |---|---|---|---|
  | FY2025/26 actual | £1.158m | £205k | 17.7% |
  | FY2026/27 budget | £1.285m | £239k | 18.6% |
  | H1 actual | £625k | £115k | |
  | H1 budget | £648k | £134k | |

- **Tests:** 10 passed.
