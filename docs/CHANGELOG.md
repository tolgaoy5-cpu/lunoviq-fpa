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

## 2026-09-30: Phase 4, rolling forecast and scenarios

- **`fpa/forecast.py`:** a reforecast after the last closed month (6+6 after September).
  - Closed months are taken from the actuals.
  - Open months run on the driver engine from the actual closing position (properties, rent).
  - The drivers are updated by the latest estimate in `config/forecast.toml`:
    - year-to-date lettings achievement carried forward (0.85),
    - rent growth of 0.5% a month,
    - the portal price rise continues.
  - The budget roster is used, including the October hire.
  - Scenarios (`upside`, `downside`) change lets, churn, new landlords and rent growth for the open months only.
- **Workbook:**
  - **Forecast sheet:** actual months are linked to Actuals; forecast months are shown in purple. It adds full-year, budget and variance columns, forecast operating drivers, and the latest-estimate assumptions with notes.
  - **Scenarios sheet:** budget vs base, upside and downside, with EBITDA vs budget, year-end portfolio and the scenario definitions.
- **FY2026/27 outturn (6+6):**

  | | EBITDA | vs budget |
  |---|---|---|
  | Budget | £239.3k | |
  | Base | £194.8k | −£44.6k |
  | Upside | £214.2k | |
  | Downside | £167.3k | |

- **Excel vs Python:** the forecast full year matches on every account.
- **Tests:** 23 passed.

## 2026-09-30: Phase 5, 13-week cash flow

- **`fpa/cash.py`:** 13-week cash forecast for the office account, built from the 6+6 monthly P&L with UK timing rules (config `[cash]`).
  - **Receipts:**
    - fees deducted from rent as it is collected (60% in the first week),
    - let-only and set-up fees on 14-day terms,
    - VAT at 20%, net of bad debts.
  - **Payments:**
    - net pay on the 28th,
    - PAYE/NI and pension on the 22nd of the next month,
    - suppliers on the 15th of the next month (VAT where charged),
    - office rent quarterly in advance on the English quarter days,
    - monthly business rates,
    - the VAT return one month and seven days after the quarter end,
    - the planned interim dividend (£120k on 18 December).
  - **Warning:** corporation tax for FY2025/26 (£47.7k), due on 1 January 2027, falls just after the window.
- **Cash13W sheet:** weekly receipts and payments, with formulas for totals, net flow, opening/closing roll, buffer, headroom and a below-buffer flag. Opening balance and buffer are named inputs.
- **Result:** the lowest closing balance is £102.3k (week 13). After the 1 January tax payment it would be £54.6k against a £50k buffer, so there is little headroom. This is the board discussion point on phasing the dividend.
- **Tests:** 27 passed, including Excel = Python on every weekly closing balance.

## 2026-09-30: Phase 6, management pack, checks and monthly pipeline

- **Sheets:**
  - **Cover:** contents with hyperlinks and the model-check status.
  - **Dashboard:** six KPI tiles and three charts:
    - revenue, actual/forecast vs budget;
    - EBITDA, actual/forecast vs budget;
    - 13-week closing cash vs the minimum balance.

    It also carries headline commentary.
  - **Checks:** nine reconciliations between the sheets, a cash buffer warning and an overall status.
- **Pack layout:** landscape, fit to width, with a footer. The Dashboard prints on one page.
- **Visual QA:** the sheets were exported to PDF through Excel and inspected. Fixes:
  - chart y-axes start at 0 (the revenue chart started at 85k and exaggerated the variances);
  - smaller chart titles;
  - headline rows no longer clipped;
  - chart data kept off the printed page;
  - section bands in Cash13W run the full width;
  - variances are rounded so an exact zero shows as "-".
- **`fpa/pipeline.py`** and `python -m fpa run --month 2026-09`:
  - builds the pack in `output/<month>_<stamp>/`,
  - recalculates it in Excel,
  - audits Excel against Python (budget, BvA, forecast and cash): 409 values, 0 mismatches, 0 Excel errors, model checks OK,
  - writes `summary.json` for the web panel.
- **Tests:** 30 passed.

## 2026-09-30: Phase 7, web panel

- **`fpa/web/server.py`** (`python -m fpa serve`, port 8766): standard-library server bound to localhost. It reuses the Lunoviq architecture:
  - job queue and a single worker (one Excel recalculation at a time),
  - friendly errors (Excel timeout, ledger errors),
  - traversal-safe run paths,
  - cache-busted assets,
  - Quit with a busy check.
- **Upload of month-end actuals:** `POST /api/upload`, trial balance and KPI CSVs. They are validated in a temporary folder with the ledger rules before saving. The server also:
  - rejects months outside the year, missing earlier months and missing KPIs;
  - asks before replacing an existing month.

  `GET /api/sample` serves synthetic files for the next month, for demos.
- **UI** (vanilla JS/CSS, the Lunoviq design language, light and dark, phone layout):
  - **Home:** the 12 months of the year (closed / open / forecast) with pack status, and the list of packs.
  - **Progress view** and **error view** with retry.
  - **Pack view,** four tabs:
    - Overview: tiles, revenue and EBITDA charts (actual/forecast bars, budget line), headlines.
    - Variance: month/YTD toggle, P&L table with materiality dots, commentary, driver-bridge bars, KPIs.
    - Forecast: budget/base/upside/downside cards, EBITDA chart, latest-estimate assumptions, monthly detail.
    - Cash: tiles, a warning callout, the 13-week chart with the buffer, the weekly receipts and payments table.
  - **Load actuals** page.
- **QA:**
  - Every page was captured at full size, and the phone layout checked at 375px with no horizontal overflow.
  - Fixed: a duplicate "No file chosen" (custom file picker), a stray focus outline on `main`, and a visible scrollbar on the tabs.
  - The web tests found a `log_message` crash on 404 responses; fixed.
  - A real pack build from the API passed the audit: 409 values, 0 mismatches, checks OK.
- **Tests:** 7 new web API tests (pipeline stubbed); 37 in total.

## 2026-09-30: Phase 8, documentation and packaging

- **README:**
  - what it does, and the September 2026 story with its numbers,
  - the Excel pack and the audit,
  - quick start and the monthly workflow,
  - tests, repository layout, limitations and license.
- **Screenshots:** web Overview, Variance and Cash tabs, and the Excel dashboard exported through Excel.
- **Packaging:** `LICENSE` (all rights reserved; viewing only) and `pyproject.toml`.
- **Mac app:** `tools/make_mac_app.py` builds `~/Applications/Lunoviq FPA.app`, with its own icon and a background server on port 8766. Verified: launched from the icon, the server started and the panel opened.
- **Fixed:** a race in the web worker, where the job status was set to "error" before the message was stored. The web test caught it intermittently. The same fix was applied to Lunoviq.
- **Tests:** 37 passed.

## 2026-09-30: Any company: importer, rule-based budget, flexed budget, second demo company

The user asked for the system to work for any company, with the example company kept.

- **Workspaces:** each company lives in `companies/<slug>/`:
  - `company.toml`,
  - `accounts.csv` (code, name, line, section, cash class),
  - `commentary.toml`,
  - `forecast.toml`,
  - `actuals/`.

  Kestrel Row moved to `companies/kestrel-row` with `model = "lettings"`.
- **Chart of accounts** (`fpa/accounts.py`): sections are revenue, cost of sales, opex, depreciation, interest, tax and balance sheet (ignored). Standard subtotals: gross profit, EBITDA, EBIT, profit before tax, net income. The P&L layout in every sheet is generated from the chart, with line subtotals for cost lines.
- **Cash:** timing classes per account replace the hardcoded account lists. The classes are rent collection, invoice, even, payroll, employer NI, pension, supplier with/without VAT, office rent and rates, bad debt, bank and non-cash. Empty payment lines are omitted.
- **`fpa/importer.py`:** reads exports from accounting systems.
  - It finds the header under report title lines and picks the delimiter (comma, semicolon or tab) by consistency.
  - It matches column names by synonyms (code/nominal, debit/Dr, balance, YTD) and understands £, thousands separators and (negatives).
  - Two layouts: trial balance (debit/credit or signed balance, monthly or year-to-date, which it converts to the month) and P&L by month. Headings and subtotals are skipped.
  - Accounts are matched by code, or by name when the export has no codes.
  - It stops on unknown accounts and names them, and it checks that the imported P&L reconciles to the file to the penny.
- **`fpa/rules.py`:** rule-based budget for any company.
  - Rules: growth (on the prior-year month), prior_year, pct_revenue, fixed, manual, tax. Defaults are set per section.
  - Excel sheets: Prior year, Assumptions (the rules, in blue) and Budget, all with live formulas.
  - The run-rate forecast scales the remaining budget by the year-to-date actual/budget of each section, plus adjustments. Scenarios apply multipliers.
- **Variance:**
  - A new rule `month_always` / `ytd_always` flags a variance on size alone.
  - Flexed budget for costs budgeted as a share of sales: the variance is split into sales volume and cost share.
  - The headline includes gross margin when there is cost of sales.
- **Second demo company, Brightwell Cleaning Services Ltd (fictional):**
  - Xero-style chart with balance-sheet accounts.
  - `fpa/demo_brightwell.py` builds the prior-year P&L by month and year-to-date trial balances with balance-sheet accounts (debits = credits), in report layout.
  - All of it is imported through the importer, exactly: 0 difference on every account and month.
  - Events: the April wage rise, a lost contract in July, a materials price rise in August, a new contract in September.
- **Results for Brightwell, September 2026:**
  - YTD EBITDA is £23.4k vs £38.5k budget.
  - Commentary finds the cleaners' cost share at 54.1% of sales vs 52.0% budgeted (−£9.1k), partly hidden by lower sales (+£6.8k), and the lost contract.
  - The pack audits clean first time: 463 values, 0 mismatches, checks OK.
- **Regression:** Kestrel budget, BvA, bridges, commentary, forecast and cash are unchanged (snapshot tests).
- **Tests:** 58 passed, including Excel.

## 2026-09-30: Web panel for any company: switcher, set-up wizard, export uploads

- **Company switcher** in the top bar. The choice is remembered per browser, and `?company=<slug>` opens a company directly.
- **Add company wizard** (`fpa/onboard.py`, `POST /api/companies/preview`, `POST /api/companies`):
  1. Upload last year's Profit and Loss by month.
  2. Every account gets a suggested section and report line. The suggestions come from:
     - unambiguous names first (depreciation, interest, corporation tax),
     - then the report's own section headings (e.g. "Cost of Sales"),
     - then keyword rules.

     The user can change any of them.
  3. Growth rates, opening cash and the minimum cash buffer are entered.
  4. The company is created with a rule-based budget, materiality thresholds scaled to its size, cash defaults and forecast scenarios, and its history is imported. Nothing half-made is left if the import fails.
- **Monthly upload for rule-based companies:** a single accounting-system export, with the basis auto-detected or chosen (month / year to date). The result reports the basis, whether debits equal credits, and the accounts matched. A synthetic next-month trial balance is available for the Brightwell demo.
- **Generic views:**
  - Variance: cost of sales, gross profit and margin; the flexed-budget card (sales volume / cost share); the company's own materiality thresholds.
  - Overview: a gross-margin tile.
  - Forecast: labelled assumptions with notes.
  - Cash: a warning with the company's own dividends.
- **Importer fixes found in browser QA:**
  - QuickBooks-style reports with an empty header over the account column are now read.
  - Report section headings guide the suggestions, so costs under "Cost of Sales" are classed as cost of sales.
- **Output:** packs are stored per company in `output/<slug>/`.
- **Tests:** 62 passed, including Excel, and 13 web tests (wizard, export upload, validation).

## 2026-09-30: Optional AI executive summary (OpenAI or Anthropic)

The user has an OpenAI account and asked to use it.

- **`fpa/ai.py`:** drafts a 5–6 sentence board summary from the pack's key figures:
  - totals and variances,
  - commentary,
  - full-year outturns,
  - cash.

  **Guardrail:** every £ amount and percentage in the draft is checked against the pack's figures (with rounding tolerance). A draft with an unknown figure is rejected; the model gets one retry with the reason, then the request fails with a clear message. The text is always labelled "AI draft · figures checked".
- **Settings:** the key goes in `local.toml` (git-ignored, mode 600) or in `OPENAI_API_KEY` / `ANTHROPIC_API_KEY`.
  - The provider and model are configurable. Defaults: OpenAI `gpt-4o-mini`, Anthropic `claude-sonnet-5`.
  - The API never returns the key.
  - Nothing is sent unless the user presses "Draft AI summary".
- **Web:**
  - Settings page: provider, model and key (password field), with a remove option.
  - Overview: an Executive summary card; the draft is saved with the pack as `ai_summary.json`.
- **Server:** it now always answers. An unexpected exception returns a JSON 500 instead of dropping the connection; the new web test caught the dropped connection.
- **Tests:** 63 fast. The AI tests mock the service: figure check, retry, give-up, settings privacy and request shape.
