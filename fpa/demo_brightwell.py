"""
Synthetic exports for Brightwell Cleaning Services Ltd (fictional), in the
shape an accounting system produces them, then imported through fpa/importer.py
exactly as a user's files would be:

  imports/pnl_by_month_FY2025-26.csv   prior year: Profit and Loss by month, with title
                                       rows, section headings and subtotals
  imports/trial_balance_<month>.csv    FY2026/27: year-to-date trial balance each month,
                                       balance-sheet accounts included (debits = credits)

Events in FY2026/27 H1 (for the variance analysis to find):
  - April: the National Living Wage rise lifts cleaners' pay to ~49% of sales (budget 47%)
  - July: a 12-site office contract ends (about £5.5k a month)
  - August: consumables prices rise ~20%
  - September: a new medical-centre contract starts (about £3k a month)

Usage:  python -m fpa.demo_brightwell      (writes imports/ and actuals/)
"""
import csv
import io
import random

from . import company, importer

SLUG = "brightwell-cleaning"
SEED = 20261001
TITLE = "Brightwell Cleaning Services Ltd"
MONTHS_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
# seasonality by financial month (Apr..Mar)
S_EOT = [1.0, 1.1, 1.35, 1.6, 1.6, 1.45, 1.0, 0.8, 0.55, 0.8, 0.9, 1.0]     # end-of-tenancy: summer moves
S_DOM = [1.0, 1.0, 0.97, 0.93, 0.9, 1.0, 1.02, 1.05, 0.9, 1.02, 1.03, 1.05]
S_CARPET = [1.25, 1.15, 1.0, 0.85, 0.8, 0.95, 1.0, 1.05, 0.7, 0.9, 1.1, 1.3]
S_UTIL = [0.85, 0.75, 0.7, 0.7, 0.7, 0.8, 1.0, 1.2, 1.35, 1.4, 1.3, 1.25]


def _true_pnl(co):
    """{month: {code: amount}} from April 2025 to the last closed month."""
    rng = random.Random(SEED)
    n = lambda s=0.03: rng.gauss(1.0, s)
    months = company.month_range(co.cfg["company"]["first_actual_month"], co.cfg["company"]["last_closed_month"])
    out = {}
    for i, m in enumerate(months):
        fm = co.month_index(m) - 1
        yr2 = m >= "2026-04"
        growth = 1.0 + 0.004 * i                                   # steady underlying growth
        commercial = 37500 * growth * n(0.015)
        if m >= "2026-04":
            commercial *= 1.03                                     # contract price rises in April
        if m >= "2026-07":
            commercial -= 5500                                     # lost 12-site contract
        if m >= "2026-09":
            commercial += 3000                                     # new medical-centre contract
        p = {"200": commercial,
             "201": 16800 * growth * S_DOM[fm] * n(),
             "202": 8200 * (1.06 if yr2 else 1.0) * S_EOT[fm] * n(0.06),
             "203": 4100 * S_CARPET[fm] * n(0.08)}
        rev = sum(p.values())
        wage_share = (0.49 if yr2 else 0.462) * n(0.01)
        wages = rev * wage_share
        p.update({"310": wages, "311": wages * 0.077, "312": wages * 0.03,
                  "320": rev * 0.05 * (1.2 if m >= "2026-08" else 1.0) * n(0.05),
                  "325": rev * 0.045 * (1.25 if fm in (2, 3, 4, 5) else 0.9) * n(0.08),
                  "477": 8900 if yr2 else 8550, "478": 880 if yr2 else 840, "479": 267 if yr2 else 256,
                  "469": 1650 if yr2 else 1575, "445": 340 * S_UTIL[fm] * n(0.05),
                  "449": 3150 * n(0.06), "450": 2400, "400": 1150 * n(0.1), "404": rev * 0.011,
                  "412": 600, "433": 560 if yr2 else 530, "463": 460, "489": 310, "416": 900, "437": 210})
        pbt = rev - sum(v for c, v in p.items() if c not in ("200", "201", "202", "203"))
        p["500"] = 0.25 * pbt
        out[m] = {c: round(v, 2) for c, v in p.items()}
    return out


def _csv(rows):
    buf = io.StringIO()
    csv.writer(buf, lineterminator="\n").writerows(rows)
    return buf.getvalue()


def _label(m):
    return "%s-%s" % (MONTHS_ABBR[int(m[5:]) - 1], m[2:4])


def pnl_by_month_export(co, pnl, months):
    """Profit and Loss by month, laid out like an accounting-system report."""
    chart = co.chart
    rows = [["Profit and Loss"], [TITLE], ["For the 12 months ended 31 March %s" % months[-1][:4]], [],
            ["Account Code", "Account"] + [_label(m) for m in months] + ["Total"]]
    fmt = lambda v: "{:,.2f}".format(v)
    for heading, sections in (("Trading Income", ("revenue",)), ("Cost of Sales", ("cost_of_sales",)),
                              ("Operating Expenses", ("opex", "depreciation", "interest")),
                              ("Taxation", ("tax",))):
        rows.append(["", heading])
        codes = chart.codes(*sections)
        for c in codes:
            vals = [pnl[m][c] for m in months]
            rows.append([c, chart.by_code[c].name] + [fmt(v) for v in vals] + [fmt(sum(vals))])
        tot = [sum(pnl[m][c] for c in codes) for m in months]
        rows.append(["", "Total " + heading] + [fmt(v) for v in tot] + [fmt(sum(tot))])
        if heading == "Cost of Sales":
            gp = [sum(pnl[m][c] for c in chart.codes("revenue")) - sum(pnl[m][c] for c in codes) for m in months]
            rows.append(["", "Gross Profit"] + [fmt(v) for v in gp] + [fmt(sum(gp))])
    net = [chart.totals(pnl[m])["net_income"] for m in months]
    rows.append(["", "Net Profit"] + [fmt(v) for v in net] + [fmt(sum(net))])
    return _csv(rows)


def trial_balance_export(co, pnl, month, rng):
    """Year-to-date trial balance at the end of `month`, with balance-sheet accounts (balances)."""
    chart = co.chart
    fy = [m for m in co.fy_months() if m <= month]
    ytd = {c: sum(pnl[m][c] for m in fy) for c in chart.by_code}
    lines = []
    for a in chart.accounts:
        v = ytd[a.code]
        if abs(v) < 0.005:
            continue
        credit = v if a.kind == "revenue" else -v              # revenue: credit balance; costs: debit balance
        dr, cr = (0.0, credit) if credit >= 0 else (-credit, 0.0)
        lines.append([a.code, a.name, {"revenue": "Revenue", "cost_of_sales": "Direct Costs"}.get(a.section, "Expense"),
                      dr, cr])
    rev_m = sum(pnl[month][c] for c in chart.codes("revenue"))
    bs = [("610", "Accounts receivable", "Current Asset", rev_m * 1.1 * rng.gauss(1, 0.05), 0.0),
          ("710", "Office equipment", "Fixed Asset", 6200.0, 0.0),
          ("720", "Vehicles", "Fixed Asset", 41000.0 - 450 * len(fy), 0.0),
          ("800", "Accounts payable", "Current Liability", 0.0, 9800 * rng.gauss(1, 0.1)),
          ("820", "VAT", "Current Liability", 0.0, 14500 * rng.gauss(1, 0.2)),
          ("825", "PAYE and NI payable", "Current Liability", 0.0, 7400 * rng.gauss(1, 0.05)),
          ("830", "Corporation tax payable", "Current Liability", 0.0, 21500.0 + ytd["500"]),
          ("900", "Van loan", "Non-current Liability", 0.0, 18000.0 - 400 * len(fy)),
          ("960", "Retained earnings", "Equity", 0.0, 64000.0)]
    lines += [list(b) for b in bs]
    dr = sum(l[3] for l in lines)
    cr = sum(l[4] for l in lines)
    bank = cr - dr                                             # the bank balances the trial balance
    lines.insert(0, ["090", "Business bank account", "Bank", bank if bank >= 0 else 0.0, -bank if bank < 0 else 0.0])
    lines.sort(key=lambda l: l[0])
    y, mm = int(month[:4]), int(month[5:])
    last_day = [31, 29 if y % 4 == 0 else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][mm - 1]
    fmt = lambda v: "" if abs(v) < 0.005 else "{:,.2f}".format(v)
    rows = [["Trial Balance"], [TITLE], ["As at %d %s %d" % (last_day, ["January", "February", "March", "April", "May",
            "June", "July", "August", "September", "October", "November", "December"][mm - 1], y)], [],
            ["Account Code", "Account", "Account Type", "Debit - Year to date", "Credit - Year to date"]]
    rows += [[l[0], l[1], l[2], fmt(l[3]), fmt(l[4])] for l in lines]
    rows.append(["", "Total", "", fmt(sum(l[3] for l in lines)), fmt(sum(l[4] for l in lines))])
    return _csv(rows)


def build(co=None, write_actuals=True):
    co = co or company.load(SLUG)
    pnl = _true_pnl(co)
    imports = co.dir / "imports"
    imports.mkdir(exist_ok=True)
    py = co.fy_months(co.budget_year - 1)
    text = pnl_by_month_export(co, pnl, py)
    (imports / "pnl_by_month_FY2025-26.csv").write_text(text)
    rng = random.Random(SEED + 7)
    tbs = {}
    for m in co.fy_months():
        if m in pnl:
            tbs[m] = trial_balance_export(co, pnl, m, rng)
            (imports / ("trial_balance_%s.csv" % m)).write_text(tbs[m])
    if write_actuals:
        importer.import_history(co, text)
        for m, t in tbs.items():
            importer.import_month(co, m, t, basis="auto")
    return pnl


def main():
    co = company.load(SLUG)
    pnl = build(co)
    print("Brightwell: %d months generated and imported through fpa.importer into %s" % (len(pnl), co.actuals_dir))


if __name__ == "__main__":
    main()
