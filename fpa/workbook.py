"""
Builds the FP&A workbook with live Excel formulas.

Scalar assumptions are named ranges (a_<driver>, s_<role>); monthly inputs sit
on the Assumptions sheet; the Drivers sheet rolls the operating drivers
forward; the Budget sheet turns them into a monthly P&L by account. Months run
across columns D..O, with the full year in column P. Inputs are blue.
"""
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

from . import accounts

FIRST_COL = 4
TOTAL = get_column_letter(FIRST_COL + 12)          # P
FONT = "Arial"

NAVY, GREY = "1F3A5F", "F2F4F7"
F_BASE = Font(name=FONT, size=9)
F_INPUT = Font(name=FONT, size=9, color="0000FF")
F_BOLD = Font(name=FONT, size=9, bold=True)
F_HEAD = Font(name=FONT, size=9, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=14, bold=True, color=NAVY)
F_NOTE = Font(name=FONT, size=8, italic=True, color="666666")
FILL_HEAD = PatternFill("solid", fgColor=NAVY)
FILL_SECTION = PatternFill("solid", fgColor=GREY)
TOP = Border(top=Side(style="thin", color="999999"))
TOP_BOTTOM = Border(top=Side(style="thin", color="999999"), bottom=Side(style="double", color="999999"))

GBP = '#,##0;(#,##0);"-"'
VAR = '#,##0;[Red](#,##0);"-"'
VARPCT = '0.0%;[Red](0.0%);"-"'
GBP2 = '#,##0.00;(#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
PCT2 = '0.00%;(0.00%);"-"'
NUM1 = '#,##0.0;(#,##0.0);"-"'
NUM = '#,##0;(#,##0);"-"'

SCALARS = [  # (section, key, label, unit, format)
    ("Portfolio", "opening_pum", "Properties under management at 1 April", "properties", NUM),
    ("Portfolio", "gained_per_month", "New managed properties per month (before seasonality)", "properties", NUM1),
    ("Portfolio", "lost_rate_annual", "Properties lost per year", "% of portfolio", PCT),
    ("Portfolio", "opening_avg_rent", "Average monthly rent at 1 April", "GBP", GBP),
    ("Portfolio", "rent_growth_monthly", "Rent growth per month", "%", PCT2),
    ("Portfolio", "collection_rate", "Rent collection rate", "%", PCT),
    ("Fees", "management_fee", "Management fee", "% of rent collected", PCT),
    ("Fees", "let_only_lets_per_month", "Let-only lets per month (before seasonality)", "lets", NUM1),
    ("Fees", "let_only_fee", "Let-only fee", "% of first year's rent", PCT),
    ("Fees", "relet_rate_annual", "Re-lets on managed stock per year", "% of portfolio", PCT),
    ("Fees", "setup_fee", "Tenancy set-up fee", "GBP per let", GBP),
    ("Fees", "renewal_rate_annual", "Renewals per year", "% of portfolio", PCT),
    ("Fees", "renewal_fee", "Renewal fee", "GBP per renewal", GBP),
    ("Fees", "certificates_per_property_year", "Certificates arranged per property per year", "number", NUM1),
    ("Fees", "certificate_margin", "Margin per certificate", "GBP", GBP),
    ("Fees", "maintenance_spend_per_property", "Contractor spend per property per month", "GBP", GBP),
    ("Fees", "maintenance_commission", "Maintenance commission", "% of contractor spend", PCT),
    ("Staff", "employer_ni_rate", "Employer NI rate", "%", PCT),
    ("Staff", "employer_ni_threshold", "Employer NI threshold", "GBP per employee per year", GBP),
    ("Staff", "pension_rate", "Employer pension", "% of salary", PCT),
    ("Staff", "commission_rate", "Commission", "% of new-business fees", PCT),
    ("Costs", "office_rent_rates", "Office rent and rates", "GBP per month", GBP),
    ("Costs", "utilities", "Utilities (before seasonality)", "GBP per month", GBP),
    ("Costs", "portals", "Property portals", "GBP per month", GBP),
    ("Costs", "marketing_base", "Marketing, fixed", "GBP per month", GBP),
    ("Costs", "marketing_rate", "Marketing, variable", "% of new-business fees", PCT),
    ("Costs", "software_base", "Software, fixed", "GBP per month", GBP),
    ("Costs", "software_per_user", "Software per user", "GBP per month", GBP),
    ("Costs", "insurance", "Insurance", "GBP per month", GBP),
    ("Costs", "professional_fees", "Professional fees", "GBP per month", GBP),
    ("Costs", "motor_per_property_manager", "Motor and travel per property manager", "GBP per month", GBP),
    ("Costs", "bad_debt_rate", "Bad debts", "% of revenue", PCT2),
    ("Costs", "other_overheads", "Other overheads", "GBP per month", GBP),
    ("Costs", "depreciation", "Depreciation", "GBP per month", GBP),
    ("Costs", "corporation_tax_rate", "Corporation tax", "% of EBIT", PCT),
]
ROLE_LABEL = {"director": "Directors", "branch_manager": "Branch managers", "property_manager": "Property managers",
              "negotiator": "Lettings negotiators", "accounts": "Accounts", "admin_compliance": "Admin and compliance"}


def mcol(i):
    return get_column_letter(FIRST_COL + i)


def _name(wb, name, sheet, ref):
    wb.defined_names[name] = DefinedName(name, attr_text="'%s'!%s" % (sheet, ref))


def _title(ws, title, subtitle):
    ws["A1"], ws["A2"] = title, subtitle
    ws["A1"].font, ws["A2"].font = F_TITLE, F_NOTE
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 12
    for i in range(13):
        ws.column_dimensions[get_column_letter(FIRST_COL + i)].width = 11


def _month_header(ws, row, months, co, label=""):
    ws.cell(row=row, column=1, value=label)
    for c in range(1, FIRST_COL + 13):
        ws.cell(row=row, column=c).fill = FILL_HEAD
        ws.cell(row=row, column=c).font = F_HEAD
    for i, m in enumerate(months):
        cell = ws.cell(row=row, column=FIRST_COL + i, value=_month_label(m))
        cell.alignment = Alignment(horizontal="right")
    ws.cell(row=row, column=FIRST_COL + 12, value=co.fy_label()).alignment = Alignment(horizontal="right")


def _index_row(ws, row):
    """Month numbers 1..12 above the header: used by YTD formulas (SUMPRODUCT on <= month)."""
    for i in range(12):
        c = ws.cell(row=row, column=FIRST_COL + i, value=i + 1)
        c.font, c.alignment = F_NOTE, Alignment(horizontal="right")


def _month_label(m):
    return ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][int(m[5:]) - 1] + "-" + m[2:4]


def _section(ws, row, text):
    ws.cell(row=row, column=1, value=text).font = F_BOLD
    for c in range(1, FIRST_COL + 13):
        ws.cell(row=row, column=c).fill = FILL_SECTION


def _row(ws, row, label, unit, formulas, fmt, font=F_BASE, total=None, border=None):
    """formulas: callable(i, col, prev) -> value for month i; total: 'sum', 'last', 'avg' or a formula."""
    ws.cell(row=row, column=1, value=label).font = font
    ws.cell(row=row, column=2, value=unit).font = F_NOTE
    for i in range(12):
        c = ws.cell(row=row, column=FIRST_COL + i, value=formulas(i, mcol(i), mcol(i - 1) if i else None))
        c.number_format, c.font = fmt, font
        if border:
            c.border = border
    if total:
        rng = "%s%d:%s%d" % (mcol(0), row, mcol(11), row)
        v = {"sum": "=SUM(%s)" % rng, "last": "=%s%d" % (mcol(11), row),
             "avg": "=AVERAGE(%s)" % rng}.get(total, total)
        c = ws.cell(row=row, column=FIRST_COL + 12, value=v)
        c.number_format, c.font = fmt, F_BOLD if font is F_BOLD else font
        if border:
            c.border = border
    return row


def assumptions_sheet(wb, co, months):
    ws = wb.create_sheet("Assumptions")
    _title(ws, "Budget assumptions: %s" % co.fy_label(),
           "%s. GBP, ex VAT. Blue = input; black = formula." % co.name)
    params = {}
    for sec in ("portfolio", "fees", "costs", "staff"):
        params.update({k: v for k, v in co.cfg[sec].items() if not isinstance(v, (dict, list))})
    row, section = 4, None
    for sec, key, label, unit, fmt in SCALARS:
        if sec != section:
            row += 1
            _section(ws, row, sec)
            section = sec
            row += 1
        ws.cell(row=row, column=1, value=label).font = F_BASE
        ws.cell(row=row, column=2, value=unit).font = F_NOTE
        c = ws.cell(row=row, column=3, value=params[key])
        c.font, c.number_format = F_INPUT, fmt
        _name(wb, "a_" + key, "Assumptions", "$C$%d" % row)
        row += 1

    row += 1
    _section(ws, row, "Salaries (annual, after the April pay review)")
    ws.cell(row=row, column=4, value="NI-able").font = F_BOLD
    row += 1
    roles = list(co.cfg["staff"]["roles"])
    sal_first = row
    for role in roles:
        ws.cell(row=row, column=1, value=ROLE_LABEL.get(role, role)).font = F_BASE
        ws.cell(row=row, column=2, value="GBP per year").font = F_NOTE
        c = ws.cell(row=row, column=3, value=co.cfg["staff"]["roles"][role][1])
        c.font, c.number_format = F_INPUT, GBP
        d = ws.cell(row=row, column=4, value="=MAX(0,C%d-a_employer_ni_threshold)" % row)
        d.font, d.number_format = F_BASE, GBP
        row += 1
    sal_last = row - 1

    row += 1
    _month_header(ws, row, months, co, "Monthly inputs")
    row += 1
    lets = _row(ws, row, "Lettings seasonality", "index",
                lambda i, c, p: co.cfg["seasonality"]["lets"][i], '0.00', F_INPUT, "avg")
    util = _row(ws, row + 1, "Utilities seasonality", "index",
                lambda i, c, p: co.cfg["seasonality"]["utilities"][i], '0.00', F_INPUT, "avg")
    row += 3
    _section(ws, row, "Headcount (planned hires included)")
    row += 1
    hires = {}
    for role, m in co.cfg["staff"].get("hires", []):
        hires.setdefault(role, []).append(m)
    hc_first = row
    for role in roles:
        n0 = co.cfg["staff"]["roles"][role][0]
        _row(ws, row, ROLE_LABEL.get(role, role), "people",
             lambda i, c, p, n0=n0, role=role: n0 + sum(1 for m in hires.get(role, []) if months[i] >= m),
             NUM, F_INPUT, "last")
        row += 1
    hc_last = row - 1
    ws.freeze_panes = "D4"
    return {"lets": lets, "util": util, "hc_first": hc_first, "hc_last": hc_last,
            "sal_first": sal_first, "sal_last": sal_last,
            "pm_row": hc_first + roles.index("property_manager")}


def drivers_sheet(wb, co, months, a):
    ws = wb.create_sheet("Drivers")
    _title(ws, "Operating drivers: %s budget" % co.fy_label(), "Calculated from Assumptions.")
    _month_header(ws, 4, months, co)
    A = "Assumptions!"
    hc = "%s{c}%d:{c}%d" % (A, a["hc_first"], a["hc_last"])
    sal = "%s$C$%d:$C$%d" % (A, a["sal_first"], a["sal_last"])
    nia = "%s$D$%d:$D$%d" % (A, a["sal_first"], a["sal_last"])
    r = {}
    rows = [
        ("open", "Properties at start of month", "properties",
         lambda i, c, p: "=a_opening_pum" if i == 0 else "=%s{close}" % p, NUM1, "=%s{open}" % mcol(0)),
        ("gained", "Properties gained", "properties",
         lambda i, c, p: "=a_gained_per_month*%s%s%d" % (A, c, a["lets"]), NUM1, "sum"),
        ("lost", "Properties lost", "properties",
         lambda i, c, p: "=%s{open}*a_lost_rate_annual/12" % c, NUM1, "sum"),
        ("close", "Properties at end of month", "properties",
         lambda i, c, p: "=%s{open}+%s{gained}-%s{lost}" % (c, c, c), NUM1, "last"),
        ("avg", "Average properties under management", "properties",
         lambda i, c, p: "=(%s{open}+%s{close})/2" % (c, c), NUM1, "avg"),
        ("rent", "Average monthly rent", "GBP",
         lambda i, c, p: "=a_opening_avg_rent" if i == 0 else "=%s{rent}*(1+a_rent_growth_monthly)" % p, GBP, "avg"),
        ("rent_roll", "Rent roll (rent due on managed stock)", "GBP",
         lambda i, c, p: "=%s{avg}*%s{rent}" % (c, c), GBP, "sum"),
        ("letonly", "Let-only lets", "lets",
         lambda i, c, p: "=a_let_only_lets_per_month*%s%s%d" % (A, c, a["lets"]), NUM1, "sum"),
        ("relets", "Re-lets on managed stock", "lets",
         lambda i, c, p: "=%s{avg}*a_relet_rate_annual/12*%s%s%d" % (c, A, c, a["lets"]), NUM1, "sum"),
        ("renewals", "Renewals", "tenancies",
         lambda i, c, p: "=%s{avg}*a_renewal_rate_annual/12" % c, NUM1, "sum"),
        ("certs", "Certificates arranged", "number",
         lambda i, c, p: "=%s{avg}*a_certificates_per_property_year/12" % c, NUM1, "sum"),
        ("maint", "Contractor spend on managed stock", "GBP",
         lambda i, c, p: "=%s{avg}*a_maintenance_spend_per_property" % c, GBP, "sum"),
        ("hc", "Headcount", "people", lambda i, c, p: "=SUM(%s)" % hc.format(c=c), NUM, "last"),
        ("salaries", "Salaries", "GBP", lambda i, c, p: "=SUMPRODUCT(%s,%s)/12" % (hc.format(c=c), sal), GBP, "sum"),
        ("ni", "Employer NI", "GBP",
         lambda i, c, p: "=SUMPRODUCT(%s,%s)*a_employer_ni_rate/12" % (hc.format(c=c), nia), GBP, "sum"),
    ]
    row = 5
    for key, *_ in rows:
        r[key] = row
        row += 1
    for key, label, unit, f, fmt, total in rows:
        fmt_f = (lambda f: lambda i, c, p: f(i, c, p).format(**r))(f)
        _row(ws, r[key], label, unit, fmt_f, fmt, total=total.format(**r) if "{" in total else total)
    ws.freeze_panes = "D5"
    return r


BUDGET_LINES = {
    "4000": "=Drivers!{c}{avg}*Drivers!{c}{rent}*a_collection_rate*a_management_fee",
    "4010": "=Drivers!{c}{letonly}*Drivers!{c}{rent}*12*a_let_only_fee",
    "4020": "=Drivers!{c}{relets}*a_setup_fee",
    "4030": "=Drivers!{c}{renewals}*a_renewal_fee",
    "4040": "=Drivers!{c}{certs}*a_certificate_margin",
    "4050": "=Drivers!{c}{maint}*a_maintenance_commission",
    "5000": "=Drivers!{c}{salaries}",
    "5010": "=Drivers!{c}{ni}",
    "5020": "=Drivers!{c}{salaries}*a_pension_rate",
    "5030": "=a_commission_rate*({c}{L4010}+{c}{L4020})",
    "6000": "=a_office_rent_rates",
    "6010": "=a_utilities*Assumptions!{c}{util}",
    "6100": "=a_portals",
    "6110": "=a_marketing_base+a_marketing_rate*({c}{L4010}+{c}{L4020})",
    "6200": "=a_software_base+a_software_per_user*Drivers!{c}{hc}",
    "6300": "=a_insurance",
    "6400": "=a_professional_fees",
    "6500": "=a_motor_per_property_manager*Assumptions!{c}{pm_row}",
    "6600": "=a_bad_debt_rate*{c}{revenue}",
    "6900": "=a_other_overheads",
    "7000": "=a_depreciation",
    "9000": "=a_corporation_tax_rate*{c}{ebit}",
}


def pnl_layout():
    """Row plan of the management P&L: (kind, key, label). kind: section/account/subtotal/ratio/blank."""
    g = lambda grp: [("account", a.code, a.name) for a in accounts.ACCOUNTS if a.group in grp]
    return ([("section", None, "Revenue")] + g(("Management fees", "New business fees", "Other fee income"))
            + [("subtotal", "revenue", "Total revenue"), ("blank", None, None),
               ("section", None, "Operating costs")] + g(("Staff costs",))
            + [("subtotal", "staff", "Staff costs")] + g(("Premises", "Marketing", "Overheads"))
            + [("subtotal", "opex", "Total operating costs"), ("blank", None, None),
               ("subtotal", "ebitda", "EBITDA"), ("ratio", "margin", "EBITDA margin"), ("blank", None, None)]
            + g(("Depreciation",)) + [("subtotal", "ebit", "EBIT")] + g(("Tax",))
            + [("subtotal", "net_income", "Net income")])


def pnl_rows(start):
    rows, row = {}, start
    for kind, key, _ in pnl_layout():
        if key:
            rows[("L" + key) if kind == "account" else key] = row
        row += 1
    return rows


SUBTOTALS = {
    "revenue": lambda c, r: "=SUM(%s%d:%s%d)" % (c, r["L4000"], c, r["L4050"]),
    "staff": lambda c, r: "=SUM(%s%d:%s%d)" % (c, r["L5000"], c, r["L5030"]),
    "opex": lambda c, r: "=%s%d+SUM(%s%d:%s%d)" % (c, r["staff"], c, r["L6000"], c, r["L6900"]),
    "ebitda": lambda c, r: "=%s%d-%s%d" % (c, r["revenue"], c, r["opex"]),
    "ebit": lambda c, r: "=%s%d-%s%d" % (c, r["ebitda"], c, r["L7000"]),
    "net_income": lambda c, r: "=%s%d-%s%d" % (c, r["ebit"], c, r["L9000"]),
}


def budget_sheet(wb, co, months, a, d):
    ws = wb.create_sheet("Budget")
    _title(ws, "Budget P&L: %s" % co.fy_label(), "%s. GBP, ex VAT. Calculated from Drivers and Assumptions." % co.name)
    _index_row(ws, 3)
    _month_header(ws, 4, months, co, "Management P&L")
    r = pnl_rows(5)
    ref = dict(d, **a, **r)
    for (kind, key, label), row in zip(pnl_layout(), range(5, 5 + len(pnl_layout()))):
        if kind == "section":
            _section(ws, row, label)
        elif kind == "account":
            f = BUDGET_LINES[key]
            _row(ws, row, label, key, lambda i, c, p, f=f: f.format(c=c, **ref), GBP, total="sum")
        elif kind == "subtotal":
            fn = SUBTOTALS[key]
            _row(ws, row, label, "", lambda i, c, p, fn=fn: fn(c, r), GBP, F_BOLD, total="sum",
                 border=TOP_BOTTOM if key in ("ebitda", "net_income") else TOP)
        elif kind == "ratio":
            _row(ws, row, label, "", lambda i, c, p: "=IF(%s%d=0,0,%s%d/%s%d)" % (c, r["revenue"], c, r["ebitda"], c, r["revenue"]),
                 PCT, F_NOTE, total="=IF(%s%d=0,0,%s%d/%s%d)" % (TOTAL, r["revenue"], TOTAL, r["ebitda"], TOTAL, r["revenue"]))
    ws.freeze_panes = "D5"
    return r


def build(co, path, month=None, root=None):
    """Budget sheets; with `month` also Actuals and BvA for that reporting month."""
    wb = Workbook()
    wb.remove(wb.active)
    months = co.fy_months()
    a = assumptions_sheet(wb, co, months)
    d = drivers_sheet(wb, co, months, a)
    b = budget_sheet(wb, co, months, a, d)
    out = {"assumptions": a, "drivers": d, "budget": b, "months": months}
    if month:
        from . import variance
        res = variance.analyse(co, month, root)
        out["actuals"] = actuals_sheet(wb, co, months, res)
        out["bva"] = bva_sheet(wb, co, months, res)
        out["variance"] = res
        from . import forecast
        sc = forecast.scenarios(co, month, root=root)
        out["forecast"] = forecast_sheet(wb, co, months, sc["base"], out["actuals"])
        out["scenarios"] = scenarios_sheet(wb, co, sc)
        out["scenario_results"] = sc
        from . import cash
        cf = cash.build(co, month, fc=sc["base"], root=root)
        out["cash"] = cash_sheet(wb, co, cf)
        out["cash_result"] = cf
        order = ["BvA", "Forecast", "Scenarios", "Cash13W", "Actuals", "Budget", "Drivers", "Assumptions"]
        wb._sheets = [wb[n] for n in order] + [ws for ws in wb._sheets if ws.title not in order]
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    return out



KPI_ROWS = [("pum_open", "Properties at start of month", NUM), ("gained", "Properties gained", NUM),
            ("lost", "Properties lost", NUM), ("pum_close", "Properties at end of month", NUM),
            ("avg_rent", "Average monthly rent", GBP), ("let_only_lets", "Let-only lets", NUM),
            ("relets", "Re-lets on managed stock", NUM), ("renewals", "Renewals", NUM),
            ("certificates", "Certificates arranged", NUM), ("headcount", "Headcount", NUM),
            ("arrears_pct", "Arrears (% of rent roll)", PCT)]
KPI_TOTAL = {"pum_open": "first", "pum_close": "last", "avg_rent": "avg", "headcount": "last", "arrears_pct": "avg"}


def actuals_sheet(wb, co, months, res):
    ws = wb.create_sheet("Actuals")
    last = res["month"]
    _title(ws, "Actuals: %s" % co.fy_label(),
           "%s. GBP, ex VAT. Monthly trial balance exports; closed months to %s." % (co.name, _month_label(last)))
    _index_row(ws, 3)
    _month_header(ws, 4, months, co, "Management P&L")
    r = pnl_rows(5)
    act = res["actuals"]
    for (kind, key, label), row in zip(pnl_layout(), range(5, 5 + len(pnl_layout()))):
        if kind == "section":
            _section(ws, row, label)
        elif kind == "account":
            _row(ws, row, label, key, lambda i, c, p, key=key: act[months[i]]["pnl"][key] if months[i] in act else None,
                 GBP, total="sum")
        elif kind == "subtotal":
            fn = SUBTOTALS[key]
            _row(ws, row, label, "", lambda i, c, p, fn=fn: fn(c, r) if months[i] in act else None, GBP, F_BOLD,
                 total="sum", border=TOP_BOTTOM if key in ("ebitda", "net_income") else TOP)
        elif kind == "ratio":
            _row(ws, row, label, "", lambda i, c, p: ("=IF(%s%d=0,0,%s%d/%s%d)" % (c, r["revenue"], c, r["ebitda"], c, r["revenue"]))
                 if months[i] in act else None, PCT, F_NOTE,
                 total="=IF(%s%d=0,0,%s%d/%s%d)" % (TOTAL, r["revenue"], TOTAL, r["ebitda"], TOTAL, r["revenue"]))
    row = 5 + len(pnl_layout()) + 1
    _section(ws, row, "Operating KPIs")
    kr = {}
    for key, label, fmt in KPI_ROWS:
        row += 1
        kr[key] = row
        tot = KPI_TOTAL.get(key, "sum")
        closed = sum(1 for m in months if m in act)
        total = {"first": "=%s%d" % (mcol(0), row), "last": "=%s%d" % (mcol(closed - 1), row),
                 "avg": "=AVERAGE(%s%d:%s%d)" % (mcol(0), row, mcol(closed - 1), row)}.get(tot, "sum")
        _row(ws, row, label, "", lambda i, c, p, key=key: act[months[i]]["kpis"][key] if months[i] in act else None,
             fmt, total=total)
    ws.freeze_panes = "D5"
    return {"pnl": r, "kpis": kr}


BVA_KIND = {"revenue": "revenue", "staff": "expense", "opex": "expense", "ebitda": "revenue", "ebit": "revenue",
            "net_income": "revenue"}


def bva_sheet(wb, co, months, res):
    ws = wb.create_sheet("BvA")
    sel = months.index(res["month"]) + 1
    _title(ws, "Budget vs actual: %s" % _month_label(res["month"]),
           "%s. GBP, ex VAT. Variance: positive = favourable, red = adverse." % co.name)
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 7
    ws.column_dimensions["C"].width = 3
    for col, w in zip("DEFGHIJKLM", (11, 11, 10, 8, 3, 11, 11, 10, 8, 7)):
        ws.column_dimensions[col].width = w
    rep = co.cfg["reporting"]
    inputs = [(4, "Reporting month (1 = April)", sel, "0", "r_sel", None, None),
              (5, "Materiality, month (GBP and %)", rep["month_abs"], GBP, "r_m_abs", rep["month_pct"], "r_m_pct"),
              (6, "Materiality, year to date (GBP and %)", rep["ytd_abs"], GBP, "r_y_abs", rep["ytd_pct"], "r_y_pct")]
    for row, label, v, fmt, name, v2, name2 in inputs:
        ws.cell(row=row, column=1, value=label).font = F_BASE
        c = ws.cell(row=row, column=4, value=v)
        c.font, c.number_format = F_INPUT, fmt
        _name(wb, name, "BvA", "$D$%d" % row)
        if name2:
            c = ws.cell(row=row, column=5, value=v2)
            c.font, c.number_format = F_INPUT, PCT
            _name(wb, name2, "BvA", "$E$%d" % row)
    ws.cell(row=4, column=5, value=_month_label(res["month"])).font = F_NOTE

    ws.cell(row=8, column=4, value="Month").font = F_BOLD
    ws.cell(row=8, column=9, value="Year to date").font = F_BOLD
    heads = {1: "", 2: "Code", 4: "Actual", 5: "Budget", 6: "Var", 7: "Var %", 9: "Actual", 10: "Budget",
             11: "Var", 12: "Var %", 13: "Flag"}
    for col in range(1, 14):
        cell = ws.cell(row=9, column=col, value=heads.get(col))
        cell.fill, cell.font = FILL_HEAD, F_HEAD
        if col >= 4:
            cell.alignment = Alignment(horizontal="right")
    src = pnl_rows(5)
    row = 10
    rows = {}
    for kind, key, label in pnl_layout():
        if kind == "blank":
            row += 1
            continue
        if kind == "section":
            ws.cell(row=row, column=1, value=label).font = F_BOLD
            for col in range(1, 14):
                ws.cell(row=row, column=col).fill = FILL_SECTION
            row += 1
            continue
        s = src[("L" + key) if kind == "account" else key]
        rows[key] = row
        bold = kind == "subtotal"
        font = F_BOLD if bold else (F_NOTE if kind == "ratio" else F_BASE)
        ws.cell(row=row, column=1, value=label).font = font
        if kind == "account":
            ws.cell(row=row, column=2, value=key).font = F_NOTE
        if kind == "ratio":
            r_rev, r_eb = rows["revenue"], rows["ebitda"]
            for col in "DEIJ":
                c = ws["%s%d" % (col, row)]
                c.value = "=IF(%s%d=0,0,%s%d/%s%d)" % (col, r_rev, col, r_eb, col, r_rev)
                c.number_format, c.font = PCT, font
            for col, a_, b_ in (("F", "D", "E"), ("K", "I", "J")):
                c = ws["%s%d" % (col, row)]
                c.value = "=%s%d-%s%d" % (a_, row, b_, row)
                c.number_format, c.font = VARPCT, font
            row += 1
            continue
        k = accounts.BY_CODE[key].kind if kind == "account" else BVA_KIND[key]
        rng = lambda sheet: "%s!$D$%d:$O$%d" % (sheet, s, s)
        ws["D%d" % row] = "=INDEX(%s,r_sel)" % rng("Actuals")
        ws["E%d" % row] = "=INDEX(%s,r_sel)" % rng("Budget")
        ws["I%d" % row] = "=SUMPRODUCT(%s*(Actuals!$D$3:$O$3<=r_sel))" % rng("Actuals")
        ws["J%d" % row] = "=SUMPRODUCT(%s*(Budget!$D$3:$O$3<=r_sel))" % rng("Budget")
        for v, a_, b_ in (("F", "D", "E"), ("K", "I", "J")):
            ws["%s%d" % (v, row)] = ("=%s%d-%s%d" if k == "revenue" else "=%s%d-%s%d") % (
                (a_, row, b_, row) if k == "revenue" else (b_, row, a_, row))
        ws["G%d" % row] = "=IF(E%d=0,0,F%d/ABS(E%d))" % (row, row, row)
        ws["L%d" % row] = "=IF(J%d=0,0,K%d/ABS(J%d))" % (row, row, row)
        ws["M%d" % row] = ('=TRIM(IF(AND(ABS(F{r})>=r_m_abs,ABS(G{r})>=r_m_pct),"M ","")&'
                           'IF(AND(ABS(K{r})>=r_y_abs,ABS(L{r})>=r_y_pct),"YTD",""))').format(r=row) if kind == "account" else None
        for col, fmt in zip("DEFGIJKLM", (GBP, GBP, VAR, VARPCT, GBP, GBP, VAR, VARPCT, "@")):
            c = ws["%s%d" % (col, row)]
            c.number_format, c.font = fmt, font
            if bold:
                c.border = TOP_BOTTOM if key in ("ebitda", "net_income") else TOP
        ws["M%d" % row].alignment = Alignment(horizontal="center")
        row += 1

    row += 1
    for period, title in (("month", "Commentary: %s" % _month_label(res["month"])), ("ytd", "Commentary: year to date")):
        _bva_section(ws, row, title)
        row += 1
        for label, text, fav in res["commentary"][period]:
            ws.cell(row=row, column=1, value=label).font = F_BOLD
            c = ws.cell(row=row, column=2, value=text)
            c.font = Font(name=FONT, size=9, color="1B5E20" if fav else "B71C1C") if label != "Headline" else F_BASE
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=13)
            ws.row_dimensions[row].height = 12 * (1 + len(text) // 120)
            row += 1
        row += 1

    _bva_section(ws, row, "Driver bridges (favourable +)")
    row += 1
    heads = ["Driver", "", "", "Volume", "Rate", "Other", "Total", "", "Volume", "Rate", "Other", "Total"]
    ws.cell(row=row, column=4, value="Month").font = F_BOLD
    ws.cell(row=row, column=9, value="Year to date").font = F_BOLD
    row += 1
    for col, h in enumerate(heads, start=1):
        c = ws.cell(row=row, column=col, value=h or None)
        c.fill, c.font = FILL_HEAD, F_HEAD
    ws.cell(row=row, column=13).fill = FILL_HEAD
    row += 1
    for name in res["bridges"]["month"]:
        ws.cell(row=row, column=1, value=name).font = F_BASE
        for base, period in ((4, "month"), (9, "ytd")):
            b = res["bridges"][period][name]
            for j, f in enumerate(("volume", "rate", "other", "var")):
                c = ws.cell(row=row, column=base + j, value=round(b[f], 2))
                c.number_format, c.font = VAR, F_BOLD if f == "var" else F_BASE
        row += 1
    ws.cell(row=row, column=1, value="Volume = properties, lets or headcount; rate = average rent; "
                                      "other = collection, fee mix and pay.").font = F_NOTE
    row += 2

    _bva_section(ws, row, "Operating KPIs: %s" % _month_label(res["month"]))
    row += 1
    for col, h in enumerate(["KPI", "", "", "Actual", "Budget", "Var"], start=1):
        c = ws.cell(row=row, column=col, value=h or None)
        c.fill, c.font = FILL_HEAD, F_HEAD
    row += 1
    for label, a_, b_, kind in res["kpis"]:
        ws.cell(row=row, column=1, value=label).font = F_BASE
        fmt = {"number": NUM, "gbp": GBP, "pct": PCT}[kind]
        for col, v in ((4, a_), (5, b_)):
            c = ws.cell(row=row, column=col, value=None if v is None else round(v, 4))
            c.number_format, c.font = fmt, F_BASE
        if b_ is not None:
            c = ws.cell(row=row, column=6, value="=D%d-E%d" % (row, row))
            c.number_format, c.font = VAR if kind != "pct" else VARPCT, F_BASE
        row += 1
    ws.freeze_panes = "D10"
    return rows


def _bva_section(ws, row, text):
    ws.cell(row=row, column=1, value=text).font = F_BOLD
    for col in range(1, 14):
        ws.cell(row=row, column=col).fill = FILL_SECTION


F_FCST = Font(name=FONT, size=9, color="5B2C83")


def forecast_sheet(wb, co, months, fc, act_rows):
    ws = wb.create_sheet("Forecast")
    _title(ws, "Forecast %s: %s" % (fc["label"], co.fy_label()),
           "%s. GBP, ex VAT. Actual months link to Actuals; forecast months (purple) come from the driver "
           "engine (fpa.forecast) and config/forecast.toml." % co.name)
    for col, w in (("Q", 3), ("R", 11), ("S", 10), ("T", 8)):
        ws.column_dimensions[col].width = w
    _index_row(ws, 3)
    for i, m in enumerate(months):
        c = ws.cell(row=3, column=FIRST_COL + i, value="Act" if fc["source"][m] == "A" else "Fcst")
        c.font, c.alignment = (F_NOTE if fc["source"][m] == "A" else Font(name=FONT, size=8, italic=True, color="5B2C83")), Alignment(horizontal="right")
    _month_header(ws, 4, months, co, "Management P&L")
    for col, h in (("R", "Budget"), ("S", "Var"), ("T", "Var %")):
        c = ws["%s4" % col]
        c.value, c.fill, c.font, c.alignment = h, FILL_HEAD, F_HEAD, Alignment(horizontal="right")
    r = pnl_rows(5)
    ap = act_rows["pnl"]
    for (kind, key, label), row in zip(pnl_layout(), range(5, 5 + len(pnl_layout()))):
        if kind == "section":
            _section(ws, row, label)
            continue
        if kind == "account":
            def f(i, c, p, key=key, row=row):
                m = months[i]
                return "=Actuals!%s%d" % (c, ap["L" + key]) if fc["source"][m] == "A" else round(fc["pnl"][m][key], 2)
            _row(ws, row, label, key, f, GBP, total="sum")
            for i, m in enumerate(months):
                if fc["source"][m] == "F":
                    ws.cell(row=row, column=FIRST_COL + i).font = F_FCST
            kind_rev = accounts.BY_CODE[key].kind == "revenue"
        elif kind == "subtotal":
            fn = SUBTOTALS[key]
            _row(ws, row, label, "", lambda i, c, p, fn=fn: fn(c, r), GBP, F_BOLD, total="sum",
                 border=TOP_BOTTOM if key in ("ebitda", "net_income") else TOP)
            kind_rev = BVA_KIND[key] == "revenue"
        elif kind == "ratio":
            _row(ws, row, label, "", lambda i, c, p: "=IF(%s%d=0,0,%s%d/%s%d)" % (c, r["revenue"], c, r["ebitda"], c, r["revenue"]),
                 PCT, F_NOTE, total="=IF(%s%d=0,0,%s%d/%s%d)" % (TOTAL, r["revenue"], TOTAL, r["ebitda"], TOTAL, r["revenue"]))
            ws["R%d" % row] = "=Budget!%s%d" % (TOTAL, row)
            ws["S%d" % row] = "=%s%d-R%d" % (TOTAL, row, row)
            for col, fmt in (("R", PCT), ("S", VARPCT)):
                ws["%s%d" % (col, row)].number_format, ws["%s%d" % (col, row)].font = fmt, F_NOTE
            continue
        else:
            continue
        ws["R%d" % row] = "=Budget!%s%d" % (TOTAL, row)
        ws["S%d" % row] = ("=%s%d-R%d" % (TOTAL, row, row)) if kind_rev else ("=R%d-%s%d" % (row, TOTAL, row))
        ws["T%d" % row] = "=IF(R%d=0,0,S%d/ABS(R%d))" % (row, row, row)
        bold = kind == "subtotal"
        for col, fmt in (("R", GBP), ("S", VAR), ("T", VARPCT)):
            c = ws["%s%d" % (col, row)]
            c.number_format, c.font = fmt, F_BOLD if bold else F_BASE

    row = 5 + len(pnl_layout()) + 1
    _section(ws, row, "Operating drivers")
    kp = act_rows["kpis"]
    drv = {d["month"]: d for d in fc["drivers"]}
    for key, label, fmt, dkey in (("pum_close", "Properties at end of month", NUM1, "pum_close"),
                                  ("avg_rent", "Average monthly rent", GBP, "avg_rent"),
                                  ("let_only_lets", "Let-only lets", NUM1, "let_only_lets"),
                                  ("relets", "Re-lets on managed stock", NUM1, "relets"),
                                  ("headcount", "Headcount", NUM, "headcount")):
        row += 1
        _row(ws, row, label, "", lambda i, c, p, key=key, dkey=dkey: ("=Actuals!%s%d" % (c, kp[key]))
             if fc["source"][months[i]] == "A" else round(drv[months[i]][dkey], 4), fmt,
             total={"pum_close": "last", "avg_rent": "avg", "headcount": "last"}.get(key, "sum"))
        for i, m in enumerate(months):
            if fc["source"][m] == "F":
                ws.cell(row=row, column=FIRST_COL + i).font = F_FCST

    row += 2
    _section(ws, row, "Latest estimate assumptions (config/forecast.toml)")
    from .forecast import load_le
    notes = load_le().get("latest_estimate", {}).get("notes", {})
    for k, v in fc["assumptions"].items():
        row += 1
        ws.cell(row=row, column=1, value=k.replace("_", " ").capitalize()).font = F_BASE
        c = ws.cell(row=row, column=3, value=round(v, 4))
        c.font, c.number_format = F_INPUT, (PCT2 if abs(v) < 0.05 else ('0.000' if v < 5 else GBP))
        ws.cell(row=row, column=4, value=notes.get(k, "")).font = F_NOTE
    ws.freeze_panes = "D5"
    return {"pnl": r}


def scenarios_sheet(wb, co, sc):
    ws = wb.create_sheet("Scenarios")
    base = sc["base"]
    _title(ws, "Scenarios: full-year outturn %s (%s)" % (co.fy_label(), base["label"]),
           "%s. GBP, ex VAT. Base = latest estimate; upside and downside change the open months only." % co.name)
    ws.column_dimensions["A"].width = 34
    names = ["budget"] + list(sc)
    heads = ["", "Budget"] + [n.capitalize() for n in sc]
    for col, h in enumerate(heads, start=1):
        c = ws.cell(row=4, column=col, value=h or None)
        c.fill, c.font = FILL_HEAD, F_HEAD
        if col > 1:
            c.alignment = Alignment(horizontal="right")
            ws.column_dimensions[get_column_letter(col)].width = 13
    groups = lambda p: accounts.by_group(p)
    rows = [("Total revenue", lambda r: r["revenue"], GBP, True), ("Staff costs", None, GBP, False),
            ("Total operating costs", lambda r: r["opex"], GBP, False), ("EBITDA", lambda r: r["ebitda"], GBP, True),
            ("EBITDA margin", lambda r: r["ebitda"] / r["revenue"], PCT, False),
            ("Net income", lambda r: r["net_income"], GBP, True)]
    row = 5
    out = {}
    for label, fn, fmt, bold in rows:
        ws.cell(row=row, column=1, value=label).font = F_BOLD if bold else F_BASE
        for j, n in enumerate(names):
            if n == "budget":
                t, p = base["budget_fy_totals"], base["budget_fy"]
            else:
                t, p = sc[n]["fy_totals"], sc[n]["fy"]
            v = groups(p)["Staff costs"] if fn is None else fn(t)
            c = ws.cell(row=row, column=2 + j, value=round(v, 4 if fmt == PCT else 2))
            c.number_format, c.font = fmt, F_BOLD if bold else F_BASE
        out[label] = row
        row += 1
    ws.cell(row=row, column=1, value="EBITDA vs budget").font = F_BASE
    for j in range(1, len(names)):
        col = get_column_letter(2 + j)
        c = ws.cell(row=row, column=2 + j, value="=%s%d-$B$%d" % (col, out["EBITDA"], out["EBITDA"]))
        c.number_format, c.font = VAR, F_BASE
    row += 1
    ws.cell(row=row, column=1, value="Properties under management at year end").font = F_BASE
    for j, n in enumerate(names):
        v = sc[n]["closing_pum"] if n != "budget" else base["budget_closing_pum"]
        c = ws.cell(row=row, column=2 + j, value=round(v, 1))
        c.number_format, c.font = NUM, F_BASE
    row += 2
    _bva_section(ws, row, "Scenario definitions")
    from .forecast import load_le
    le = load_le()
    for n, d in le.get("scenarios", {}).items():
        row += 1
        ws.cell(row=row, column=1, value=n.capitalize()).font = F_BOLD
        ws.cell(row=row, column=2, value=d.get("description", "")).font = F_BASE
    return out


def cash_sheet(wb, co, cf):
    from .cash import PAYMENT_LINES, RECEIPT_LINES
    ws = wb.create_sheet("Cash13W")
    _title(ws, "13-week cash flow: office account",
           "%s. GBP, incl. VAT. Built from the %s forecast with the timing rules in config [cash]; "
           "client money excluded." % (co.name, cf["forecast_label"]))
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 3
    ws.column_dimensions["C"].width = 3
    for i in range(14):
        ws.column_dimensions[get_column_letter(FIRST_COL + i)].width = 10
    wcol = lambda i: get_column_letter(FIRST_COL + i)
    last = wcol(12)
    ws.cell(row=4, column=1, value="Week commencing")
    ws.cell(row=5, column=1, value="Week")
    for c in range(1, FIRST_COL + 14):
        for r in (4, 5):
            ws.cell(row=r, column=c).fill, ws.cell(row=r, column=c).font = FILL_HEAD, F_HEAD
    for i, (w0, _) in enumerate(cf["weeks"]):
        for r, v in ((4, w0.strftime("%d-%b")), (5, i + 1)):
            ws.cell(row=r, column=FIRST_COL + i, value=v).alignment = Alignment(horizontal="right")
    ws.cell(row=4, column=FIRST_COL + 13, value="13 weeks").alignment = Alignment(horizontal="right")
    rows, row = {}, 6

    def line(label, values=None, formula=None, font=F_BASE, fmt=VAR, total=True, border=None):
        nonlocal row
        ws.cell(row=row, column=1, value=label).font = font
        for i in range(13):
            v = formula(i, wcol(i)) if formula else round(values[i], 2)
            c = ws.cell(row=row, column=FIRST_COL + i, value=v)
            c.number_format, c.font = fmt, font
            if border:
                c.border = border
        if total:
            c = ws.cell(row=row, column=FIRST_COL + 13, value="=SUM(%s%d:%s%d)" % (wcol(0), row, last, row))
            c.number_format, c.font = fmt, F_BOLD
        rows[label] = row
        row += 1
        return row - 1

    for sec, lines_ in (("Receipts", RECEIPT_LINES), ("Payments", PAYMENT_LINES)):
        _bva_section(ws, row, sec)
        row += 1
        first = row
        for l in lines_:
            line(l, cf["table"][(sec, l)])
        r0, r1 = first, row - 1
        line("Total " + sec.lower(), formula=lambda i, c, r0=r0, r1=r1: "=SUM(%s%d:%s%d)" % (c, r0, c, r1),
             font=F_BOLD, border=TOP)
        row += 1
    rr, pr = rows["Total receipts"], rows["Total payments"]
    net = line("Net cash flow", formula=lambda i, c: "=%s%d+%s%d" % (c, rr, c, pr), font=F_BOLD, border=TOP_BOTTOM)
    row += 1
    op = row
    cl = row + 1
    line("Opening balance", formula=lambda i, c: ("=cash_opening" if i == 0 else "=%s%d" % (wcol(i - 1), cl)),
         fmt=GBP, total=False)
    line("Closing balance", formula=lambda i, c: "=%s%d+%s%d" % (c, op, c, net), font=F_BOLD, fmt=GBP, total=False,
         border=TOP_BOTTOM)
    mn = line("Minimum balance (buffer)", formula=lambda i, c: "=cash_minimum", fmt=GBP, total=False)
    line("Headroom over the buffer", formula=lambda i, c: "=%s%d-%s%d" % (c, cl, c, mn), fmt=VAR, total=False)
    line("Below buffer?", formula=lambda i, c: '=IF(%s%d<%s%d,"YES","")' % (c, cl, c, mn), fmt="@", total=False)
    row += 1
    for label, v, name, fmt in (("Opening balance at %s" % cf["weeks"][0][0].strftime("%d %b %Y"), cf["opening"], "cash_opening", GBP),
                                ("Minimum balance (board buffer)", cf["minimum"], "cash_minimum", GBP)):
        ws.cell(row=row, column=1, value=label).font = F_BASE
        c = ws.cell(row=row, column=FIRST_COL, value=v)
        c.font, c.number_format = F_INPUT, fmt
        _name(wb, name, "Cash13W", "$%s$%d" % (wcol(0), row))
        row += 1
    ct = cf["corporation_tax"]
    row += 1
    _bva_section(ws, row, "Notes")
    for text in ("Corporation tax for FY2025/26 of £%s is due on %s (week 14, just after this window): "
                 "balance after payment £%s vs the £%s buffer."
                 % ("{:,.0f}".format(ct["amount"]), ct["due"].strftime("%d %b %Y"),
                    "{:,.0f}".format(ct["balance_after"]), "{:,.0f}".format(cf["minimum"])),
                 "Lowest closing balance in the window: £%s (week %d)." % ("{:,.0f}".format(cf["lowest"]), cf["lowest_week"]),
                 "Receipts are net of bad debts; VAT is collected on fees and paid quarterly."):
        row += 1
        ws.cell(row=row, column=1, value=text).font = F_BASE
    ws.freeze_panes = "D6"
    return {"rows": rows, "closing": cl}
