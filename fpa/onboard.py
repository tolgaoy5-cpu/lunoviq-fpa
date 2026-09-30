"""
Adding a company from its own exports (used by the web wizard):

  preview(text)            parse a P&L-by-month (or trial balance) export and suggest, for every
                           account, a section and a management line
  create(...)              write companies/<slug>/ (company.toml, accounts.csv) with a rule-based
                           budget, then import the history

The suggestions are only a starting point: the user confirms or changes each account.
"""
import re
import tomllib

from . import company, importer
from .accounts import SECTIONS, Account, Chart

STAFF = ("wage", "salar", "payroll", "employer ni", "national insurance", "pension", "staff", "bonus", "commission")
PREMISES = ("rent", "rates", "light", "power", "heat", "utilit", "premises", "repairs", "service charge", "water")
VEHICLES = ("motor", "vehicle", "fuel", "travel", "mileage", "van", "parking")
MARKETING = ("advert", "marketing", "promotion", "website", "portal")


def suggest_line(section, name):
    n = name.lower()
    if section == "revenue":
        return re.sub(r"^(sales|revenue|income)\s*[-–:]\s*", "", name, flags=re.I).strip() or name
    if section == "cost_of_sales":
        if any(k in n for k in ("wage", "salar", "labour", "employer ni", "pension")):
            return "Direct labour"
        if any(k in n for k in ("material", "supplies", "purchases", "stock", "consumable")):
            return "Materials"
        if "subcontract" in n:
            return "Subcontractors"
        return "Direct costs"
    if section == "opex":
        for line, keys in (("Staff costs", STAFF), ("Premises", PREMISES), ("Vehicles and travel", VEHICLES),
                           ("Marketing", MARKETING)):
            if any(k in n for k in keys):
                return line
        return "Overheads"
    return {"depreciation": "Depreciation", "interest": "Interest", "tax": "Corporation tax",
            "balance_sheet": "Balance sheet"}[section]


def suggest_cash(section, name):
    n = name.lower()
    if section in ("cost_of_sales", "opex"):
        if any(k in n for k in ("employer ni", "national insurance")):
            return "employer_ni"
        if "pension" in n:
            return "pension"
        if any(k in n for k in ("wage", "salar", "payroll", "bonus", "commission")):
            return "payroll_gross"
        if any(k in n for k in ("insurance", "bank", "card fee", "mileage", "rates")):
            return "supplier_novat"
        if "bad debt" in n:
            return "bad_debt"
    return ""


def preview(text):
    parsed = importer.parse(text)
    accts = []
    seen = set()
    for i, r in enumerate(parsed["rows"]):
        code = r["code"] or "A%03d" % (i + 1)
        if code in seen:
            continue
        seen.add(code)
        sec = importer.suggest_section(r)
        total = sum(r.get("months", {}).values()) if parsed["layout"] == "pnl_by_month" else r["debit"] - r["credit"]
        accts.append({"code": code, "name": r["name"] or code, "section": sec, "line": suggest_line(sec, r["name"] or code),
                      "cash": suggest_cash(sec, r["name"] or ""), "total": round(total, 2), "has_code": bool(r["code"])})
    return {"layout": parsed["layout"], "months": parsed["months"], "accounts": accts, "sections": list(SECTIONS)}


def slugify(name):
    s = re.sub(r"\b(ltd|limited|plc|llp|inc)\b\.?", "", name.lower())
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return s[:40] or "company"


def _q(s):
    return '"%s"' % str(s).replace("\\", "\\\\").replace('"', '\\"')


def create(name, accounts, history_text, fy_start_month=4, revenue_growth=0.05, cost_growth=0.03,
           opening_balance=0.0, minimum_balance=0.0, description="", root=None):
    """Create the company, import its history and return the loaded Company.
    accounts: [{"code", "name", "section", "line", "cash"}] as confirmed by the user."""
    root = root or company.COMPANIES
    slug = slugify(name)
    d = root / slug
    if (d / "company.toml").exists():
        raise ValueError("A company called %s already exists." % name)
    chart = Chart([Account(str(a["code"]).strip(), a["name"].strip(), (a.get("line") or a["name"]).strip(),
                           a["section"], a.get("cash", "") or "") for a in accounts])
    parsed = importer.parse(history_text)
    if parsed["layout"] != "pnl_by_month":
        raise ValueError("The history must be a P&L by month (one column per month).")
    months = parsed["months"]
    if not months:
        raise ValueError("No months found in the file.")
    first, last = months[0], months[-1]
    fy_of = lambda m: int(m[:4]) if int(m[5:]) >= fy_start_month else int(m[:4]) - 1
    fy_end_month = (fy_start_month - 2) % 12 + 1
    budget_year = fy_of(last) + 1 if int(last[5:]) == fy_end_month else fy_of(last)
    end_year = budget_year if fy_start_month > 1 else budget_year - 1
    ct_due = company.add_months("%d-%02d" % (end_year, fy_end_month), 10) + "-01"   # 9 months and a day
    matched, _ = importer.resolve(chart, parsed["rows"])
    rev = sum(sum(i.get("months", {}).values()) for c, items in matched.items() if chart.all_by_code[c].section == "revenue"
              for i in items)
    per_month = rev / len(months)
    unit = max(250.0, round(per_month * 0.01 / 50) * 50)
    toml = "\n".join([
        "# %s: created with the Lunoviq FP&A wizard. Budgeted by rules (fpa/rules.py)." % name,
        "", "[company]", "name = %s" % _q(name), 'model = "rules"', "description = %s" % _q(description),
        'currency = "GBP"', "fy_start_month = %d" % fy_start_month, "budget_year = %d" % budget_year,
        "first_actual_month = %s" % _q(first), "last_closed_month = %s" % _q(last), "",
        "[budget.defaults]",
        'revenue = {method = "growth", rate = %.4f}' % revenue_growth,
        'cost_of_sales = {method = "growth", rate = %.4f}' % revenue_growth,
        'opex = {method = "growth", rate = %.4f}' % cost_growth,
        'depreciation = {method = "prior_year"}', 'interest = {method = "prior_year"}',
        'tax = {method = "tax", rate = 0.25}', "", "[budget.accounts]", "",
        "[reporting]", "month_abs = %.1f" % unit, "month_pct = 0.10", "ytd_abs = %.1f" % (unit * 3),
        "ytd_pct = 0.05", "month_always = %.1f" % (unit * 4), "ytd_always = %.1f" % (unit * 10), "",
        "[cash]", "opening_balance = %.2f" % opening_balance, "minimum_balance = %.2f" % minimum_balance,
        "vat_registered = true", "vat_rate = 0.20", "vat_quarter_end_months = [3, 6, 9, 12]",
        "vat_payment_lag_days = 37", "payroll_day = 28", "paye_day = 22", "supplier_day = 15",
        "employee_deductions = 0.15", "employee_pension = 0.05", "invoice_terms_days = 30",
        "corporation_tax_due = %s" % _q(ct_due),
        "corporation_tax_fy_prior = 0.0", "dividends = []", ""])
    tomllib.loads(toml)                                       # never write an invalid file
    d.mkdir(parents=True)
    (d / "company.toml").write_text(toml)
    chart.save(d / "accounts.csv")
    (d / "forecast.toml").write_text('[latest_estimate]\nrevenue = "ytd"\ncost_of_sales = "ytd"\nopex = "budget"\n\n'
                                     '[scenarios.upside]\nrevenue_mult = 1.05\ndescription = "Sales 5% above the '
                                     'run-rate."\n\n[scenarios.downside]\nrevenue_mult = 0.93\ncost_mult = 1.02\n'
                                     'description = "Sales 7% below the run-rate; costs 2% higher."\n')
    co = company.load(slug, root)
    try:
        importer.import_history(co, history_text)
    except Exception:
        import shutil
        shutil.rmtree(d)                                      # leave nothing half-made
        raise
    return co
