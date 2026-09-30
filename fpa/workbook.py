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


def build(co, path):
    wb = Workbook()
    wb.remove(wb.active)
    months = co.fy_months()
    a = assumptions_sheet(wb, co, months)
    d = drivers_sheet(wb, co, months, a)
    b = budget_sheet(wb, co, months, a, d)
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    return {"assumptions": a, "drivers": d, "budget": b, "months": months}

