"""
Imports exports from accounting systems (Xero, QuickBooks, Sage and similar)
into the stored monthly format (fpa/ledger.py).

Two layouts are recognised:

  trial balance   one row per account, with a code and/or name column and either
                  debit and credit columns or one signed balance column (debit +);
                  balances may be the month's movement or year to date
  P&L by month    one row per account and one column per month (e.g. "Apr-26",
                  "April 2026", "2026-04"), amounts as presented in the report

The header row is found automatically (reports often start with title lines),
column names are matched by synonyms, and amounts such as "£1,234.56" or
"(1,234.56)" are understood. Accounts are matched to the company's chart by code,
or by name when the export has no codes. Nothing is stored unless:
  - every account in the file is in the chart (balance-sheet accounts may be listed
    in the chart with section "balance_sheet"; they are then ignored),
  - the P&L in the file reconciles to the imported P&L to the penny.
"""
import csv
import datetime as dt
import io
import re

from . import company, ledger

CODE = ("account code", "code", "account number", "account no", "acc code", "nominal code", "nominal", "a/c",
        "account #", "gl code", "ledger code")
NAME = ("account name", "account", "name", "description", "nominal name", "account description")
DEBIT = ("debit", "dr", "debits", "debit amount", "debit (gbp)", "debit gbp")
CREDIT = ("credit", "cr", "credits", "credit amount", "credit (gbp)", "credit gbp")
BALANCE = ("balance", "net", "amount", "net amount", "value", "closing balance", "net movement", "movement")
TYPE = ("account type", "type", "class", "account class")
MONTHS = {m: i for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov",
                                      "dec"], 1)}


class ImportError_(ValueError):
    pass


def _norm(h):
    h = re.sub(r"\s+", " ", (h or "").strip().lower())
    return re.sub(r"\s*-\s*(ytd|year to date|month|period)$", r" \1", h)


def amount(s):
    """'£1,234.56' -> 1234.56; '(1,234.56)' or '-1,234.56' -> -1234.56; '' or '-' -> 0."""
    t = (s or "").strip().replace("£", "").replace("$", "").replace("€", "").replace(",", "").replace(" ", "")
    if t in ("", "-", "–", "—"):
        return 0.0
    neg = t.startswith("(") and t.endswith(")")
    t = t.strip("()")
    if t.endswith("-"):
        neg, t = True, t[:-1]
    try:
        v = float(t)
    except ValueError:
        raise ImportError_("not an amount: %r" % s) from None
    return -v if neg else v


def month_of(header):
    """Column header -> 'YYYY-MM' or None ('Apr-26', 'April 2026', '2026-04', '30 Apr 2026', 'Apr 26')."""
    h = (header or "").strip().lower()
    m = re.fullmatch(r"(\d{4})-(\d{1,2})", h)
    if m:
        return "%s-%02d" % (m.group(1), int(m.group(2)))
    m = re.fullmatch(r"(?:\d{1,2}\s+)?([a-z]{3})[a-z]*[\s\-/']+(\d{2}|\d{4})", h)
    if m and m.group(1) in MONTHS:
        y = int(m.group(2))
        y = y + 2000 if y < 100 else y
        return "%d-%02d" % (y, MONTHS[m.group(1)])
    return None


def read_rows(text):
    """Rows of a CSV export; the delimiter (comma, semicolon or tab) is the one that splits the
    most lines into the same, largest number of fields (title lines above the table are ignored)."""
    text = text.lstrip("\ufeff")
    best, best_score = ",", -1
    for d in (",", ";", "\t"):
        rows = list(csv.reader(io.StringIO(text), delimiter=d))
        widths = [len(r) for r in rows if len(r) > 1]
        if not widths:
            continue
        mode = max(set(widths), key=lambda w: (widths.count(w), w))
        score = widths.count(mode) * mode
        if score > best_score:
            best, best_score = d, score
    return list(csv.reader(io.StringIO(text), delimiter=best))


def _find(header, names):
    for i, h in enumerate(header):
        if h in names:
            return i
    return None


def parse(text):
    """-> {"layout": "trial_balance"|"pnl_by_month", "rows": [{"code","name","type", values...}], ...}."""
    rows = read_rows(text)
    for hi, raw in enumerate(rows[:15]):
        header = [_norm(h) for h in raw]
        code_i, name_i = _find(header, CODE), _find(header, NAME)
        month_cols = [(i, month_of(raw[i])) for i in range(len(raw)) if month_of(raw[i])]
        if code_i is None and name_i is None:
            if len(month_cols) < 2:
                continue
            # P&L by month with an unnamed account column (e.g. QuickBooks): the first non-month column
            name_i = next((i for i in range(len(raw)) if i not in {c for c, _ in month_cols}), None)
            if name_i is None:
                continue
        sfx = ("", " ytd", " year to date", " month", " period", " this month")
        dr_i = _find(header, tuple(d + x for d in DEBIT for x in sfx))
        cr_i = _find(header, tuple(c + x for c in CREDIT for x in sfx))
        bal_i = _find(header, tuple(b + x for b in BALANCE for x in sfx))
        type_i = _find(header, TYPE)
        if month_cols:
            layout = "pnl_by_month"
        elif (dr_i is not None and cr_i is not None) or bal_i is not None:
            layout = "trial_balance"
        else:
            continue
        out = []
        heading = ""                               # the report section the rows sit under (e.g. "Cost of Sales")
        for n, r in enumerate(rows[hi + 1:], start=hi + 2):
            if not any(c.strip() for c in r):
                continue
            get = lambda i: r[i].strip() if i is not None and i < len(r) else ""
            code, name = get(code_i), get(name_i)
            if code_i is None:                     # name-only exports: split "200 - Sales" style names
                m = re.match(r"^(\d{2,6})\s*[-–:]\s*(.+)$", name)
                if m:
                    code, name = m.group(1), m.group(2)
            low = (name or code).lower()
            if not code and (not name or low.startswith(("total", "net ", "gross", "operating profit", "profit"))):
                continue                           # section titles and report subtotals
            item = {"line": n, "code": code, "name": name, "type": get(type_i) or heading}
            try:
                if layout == "pnl_by_month":
                    item["months"] = {mo: amount(get(i)) for i, mo in month_cols}
                elif dr_i is not None and cr_i is not None:
                    item["debit"], item["credit"] = amount(get(dr_i)), amount(get(cr_i))
                else:
                    b = amount(get(bal_i))
                    item["debit"], item["credit"] = (b, 0.0) if b >= 0 else (0.0, -b)
            except ImportError_ as e:
                raise ImportError_("line %d: %s" % (n, e)) from None
            vals = list(item.get("months", {}).values()) + [item.get("debit", 0.0), item.get("credit", 0.0)]
            if not code and not any(vals):
                heading = name                     # section headings (e.g. "Revenue") carry no amounts
                continue
            out.append(item)
        if not out:
            raise ImportError_("no account rows found under the header on line %d" % (hi + 1))
        return {"layout": layout, "rows": out, "header_line": hi + 1,
                "months": sorted({mo for _, mo in month_cols}) if layout == "pnl_by_month" else [],
                "ytd_columns": any("ytd" in _norm(h) or "year to date" in _norm(h) for h in rows[hi])}
    raise ImportError_("could not find a header row with an account column and debit/credit, balance or month "
                       "columns in the first 15 lines")


def _match(chart, item, by_name):
    if item["code"] and item["code"] in chart.all_by_code:
        return chart.all_by_code[item["code"]]
    if item["name"]:
        a = by_name.get(item["name"].strip().lower())
        if a:
            return a
    return None


def resolve(chart, rows):
    """-> ({code: row}, unmatched rows); several export rows for one account are added up."""
    by_name = {a.name.lower(): a for a in chart.all}
    matched, unmatched = {}, []
    for it in rows:
        a = _match(chart, it, by_name)
        if a is None:
            unmatched.append(it)
        else:
            matched.setdefault(a.code, []).append(it)
    return matched, unmatched


def _unmatched_error(unmatched):
    names = ["%s %s" % (u["code"], u["name"]) if u["code"] else u["name"] for u in unmatched[:12]]
    more = " and %d more" % (len(unmatched) - 12) if len(unmatched) > 12 else ""
    return ImportError_("accounts not in the chart: %s%s. Add them to accounts.csv (use section \"balance_sheet\" "
                        "for balance-sheet accounts)." % ("; ".join(names), more))


def trial_balance_to_pnl(co, parsed, basis="month", previous=None):
    """P&L amounts for the month from a parsed trial balance. basis="ytd": subtract `previous`
    (the year-to-date P&L at the end of the previous month, {} for the first month of the year)."""
    chart = co.chart
    matched, unmatched = resolve(chart, parsed["rows"])
    if unmatched:
        raise _unmatched_error(unmatched)
    pnl = {}
    for code, items in matched.items():
        a = chart.all_by_code[code]
        if a.section == "balance_sheet":
            continue
        dr, cr = sum(i["debit"] for i in items), sum(i["credit"] for i in items)
        pnl[code] = round(cr - dr if a.kind == "revenue" else dr - cr, 2)
    for code in chart.by_code:
        pnl.setdefault(code, 0.0)
    if basis == "ytd":
        prev = previous or {}
        pnl = {c: round(v - prev.get(c, 0.0), 2) for c, v in pnl.items()}
    # reconciliation: the file's P&L result (credits - debits on P&L accounts) must equal the imported result
    file_result = sum(sum(i["credit"] - i["debit"] for i in items) for code, items in matched.items()
                      if chart.all_by_code[code].section != "balance_sheet")
    if basis == "ytd":
        file_result -= sum(v if chart.by_code[c].kind == "revenue" else -v for c, v in (previous or {}).items())
    t = chart.totals(pnl)
    if abs(file_result - t["net_income"]) > 0.01:
        raise ImportError_("the imported P&L (%.2f) does not reconcile to the file (%.2f)" % (t["net_income"], file_result))
    dr_all = sum(i["debit"] for items in matched.values() for i in items)
    cr_all = sum(i["credit"] for items in matched.values() for i in items)
    has_bs = any(chart.all_by_code[c].section == "balance_sheet" for c in matched)
    return pnl, {"result": t["net_income"], "revenue": t["revenue"], "accounts": len(matched),
                 "balanced": abs(dr_all - cr_all) < 0.01 if has_bs else None}


def pnl_by_month_to_months(co, parsed):
    """{month: P&L amounts} from a P&L-by-month export (amounts as presented: revenue and costs positive)."""
    chart = co.chart
    matched, unmatched = resolve(chart, parsed["rows"])
    if unmatched:
        raise _unmatched_error(unmatched)
    out = {}
    for mo in parsed["months"]:
        p = {c: 0.0 for c in chart.by_code}
        for code, items in matched.items():
            if chart.all_by_code[code].section != "balance_sheet":
                p[code] = round(sum(i["months"].get(mo, 0.0) for i in items), 2)
        out[mo] = p
    return out


def ytd_before(co, month, root=None):
    """Year-to-date P&L at the end of the month before `month` (from stored actuals)."""
    y, m = int(month[:4]), int(month[5:])
    fy_start = "%04d-%02d" % (y if m >= co.fy_start_month else y - 1, co.fy_start_month)
    have = set(ledger.available_months(co, root))
    months = company.month_range(fy_start, company.add_months(month, -1)) if month > fy_start else []
    missing = [m for m in months if m not in have]
    if missing:
        raise ImportError_("year-to-date files need the earlier months of the year first: %s" % ", ".join(missing))
    total = {c: 0.0 for c in co.chart.by_code}
    for m in months:
        p, _ = ledger.read_month(co, m, root)
        for c, v in p.items():
            total[c] += v
    return total


def import_month(co, month, text, basis="auto", kpis=None, root=None, save=True):
    """Import one month's trial balance (or a single-month P&L) and store it."""
    parsed = parse(text)
    if parsed["layout"] == "pnl_by_month":
        months = pnl_by_month_to_months(co, parsed)
        if month not in months:
            raise ImportError_("the file has no column for %s (found %s)" % (month, ", ".join(parsed["months"])))
        pnl, info = months[month], {"layout": "pnl_by_month"}
    else:
        if basis == "auto":
            basis = "ytd" if parsed["ytd_columns"] else "month"
        prev = ytd_before(co, month, root) if basis == "ytd" else None
        pnl, info = trial_balance_to_pnl(co, parsed, basis, prev)
        info.update(layout="trial_balance", basis=basis)
    if save:
        ledger.write_month(co, month, pnl, kpis, root)
    return pnl, info


def import_history(co, text, root=None, save=True):
    """Import several months from a P&L-by-month export (e.g. the prior year for the budget)."""
    parsed = parse(text)
    if parsed["layout"] != "pnl_by_month":
        raise ImportError_("expected a P&L by month (one column per month)")
    months = pnl_by_month_to_months(co, parsed)
    if save:
        for m, p in months.items():
            ledger.write_month(co, m, p, None, root)
    return months


def suggest_section(item):
    """Best guess of the P&L section for an account in an export (used when building a chart)."""
    t = (item.get("type") or "").lower()
    name = (item.get("name") or "").lower()
    code = item.get("code") or ""
    # unambiguous names first: reports often list these under "Operating Expenses"
    if any(k in name for k in ("corporation tax", "income tax expense")):
        return "tax"
    if any(k in name for k in ("interest paid", "interest expense", "loan interest", "bank interest paid")):
        return "interest"
    if "depreciation" in name or "amortisation" in name or "amortization" in name:
        return "depreciation"
    by_type = [("cost_of_sales", ("direct cost", "cost of sales", "cogs", "direct costs", "cost of goods")),
               ("revenue", ("revenue", "sales", "income", "other income", "turnover")),
               ("opex", ("expense", "overhead", "overheads")),
               ("balance_sheet", ("asset", "liability", "liabilities", "equity", "bank", "current", "fixed",
                                  "non-current", "inventory", "prepayment"))]
    for sec, keys in by_type:
        if any(k in t for k in keys):
            return sec
    if any(k in name for k in ("sales", "revenue", "income", "fees received", "turnover")):
        return "revenue"
    if any(k in name for k in ("cost of sales", "materials", "subcontract", "purchases", "direct wages")):
        return "cost_of_sales"
    if code.isdigit() and len(code) <= 4 and code[:1] in "0189" and len(code) == 3:
        return "balance_sheet"
    return "opex"
