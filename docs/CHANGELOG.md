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

## 2026-09-30: Phase 2, driver-based budget in Excel with a Python mirror

- **`fpa/budget.py`:** the budget from the driver engine, by month and for the full year.
- **`fpa/workbook.py`:** generates the workbook with live formulas.
  - **Assumptions:** 35 named scalar inputs in blue, salaries by role with an NI-able helper column, and monthly seasonality and headcount (the October hire is included).
  - **Drivers:** portfolio roll-forward, rent, rent roll, lettings activity, headcount, salaries and NI.
  - **Budget:** management P&L by account with subtotals, EBITDA margin, depreciation, EBIT, tax and net income. Months run across columns D–O, with the full year in column P.
- **`fpa/recalc.py`:** reused from Lunoviq (sandbox-safe, runs in a child process with a timeout).
- **Budget FY2026/27:** revenue £1,284,541; EBITDA £239,320 (18.6%); net income £168,690.
- **Excel vs Python:** they match on every account in every month. The maximum difference is 2p, from rounding.
- **Tests:** 14 passed, including the Excel recalculation.

## 2026-09-30: Phase 3, budget vs actual

- **`fpa/variance.py`:** month and year-to-date variance by account and P&L line (positive = favourable).
  - **Materiality flags:** a line is flagged only when it passes both thresholds, £1k and 10% for the month, £3k and 5% for the year to date (config `[reporting]`).
  - **Driver bridges:**
    - management fees: volume (properties), rate (rent), other;
    - let-only fees: volume (lets), rate (rent), other;
    - set-up fees: volume (re-lets), other;
    - staff costs: headcount, other.
  - **Commentary:** generated text (the numbers) combined with analyst notes from `config/commentary.toml` (the reasons).
  - **Checks:** a reporting month must be closed and inside the budget year.
- **Workbook:**
  - **Actuals sheet:** trial-balance values, subtotal formulas and operating KPIs.
  - **BvA sheet:**
    - reporting month and materiality as named inputs;
    - month and YTD actual, budget and variance as `INDEX` / `SUMPRODUCT` formulas on the Actuals and Budget sheets, with a flag formula;
    - commentary, driver bridges and a KPI table.
- **September 2026:**
  - **Month:** EBITDA £16.9k vs £26.4k.
  - **YTD:** EBITDA £114.6k vs £134.3k (−£19.7k).
  - **YTD causes:**
    - let-only fees −£19.0k (81 lets vs 95),
    - set-up fees −£4.8k,
    - agency cover −£6.1k,
    - bad debt −£3.9k,
    - partly offset by staff savings of +£11.6k from the vacancy, and rents ahead of budget (+£4.0k).
- **Excel vs Python:** the BvA sheet matches to within 5p on every line, and the flags match.
- **Tests:** 19 passed.
