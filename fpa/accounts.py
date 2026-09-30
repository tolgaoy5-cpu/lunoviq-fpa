"""
Chart of accounts, per company (companies/<slug>/accounts.csv):

    code,name,line,section,cash

section   revenue | cost_of_sales | opex | depreciation | interest | tax | balance_sheet
line      the management P&L line the account rolls up into (e.g. "Staff costs")
cash      timing class for the 13-week cash flow (see fpa/cash.py); blank = default for the section

Balance-sheet accounts may be listed so that full trial balances import cleanly;
they are ignored by the P&L. All P&L amounts are positive (revenue as income,
costs as expense); subtotals follow the standard management P&L:

    revenue - cost of sales = gross profit - operating costs = EBITDA
    - depreciation = EBIT - interest = profit before tax - tax = net income
"""
import csv
from dataclasses import dataclass
from pathlib import Path

SECTIONS = ("revenue", "cost_of_sales", "opex", "depreciation", "interest", "tax", "balance_sheet")
PNL_SECTIONS = SECTIONS[:-1]
SECTION_LABEL = {"revenue": "Revenue", "cost_of_sales": "Cost of sales", "opex": "Operating costs",
                 "depreciation": "Depreciation", "interest": "Interest", "tax": "Tax"}


class ChartError(ValueError):
    pass


@dataclass(frozen=True)
class Account:
    code: str
    name: str
    line: str
    section: str
    cash: str = ""

    @property
    def kind(self):
        return "revenue" if self.section == "revenue" else "expense"

    @property
    def group(self):                       # the management line (kept for older callers)
        return self.line


class Chart:
    def __init__(self, accounts):
        codes = [a.code for a in accounts]
        dup = sorted({c for c in codes if codes.count(c) > 1})
        if dup:
            raise ChartError("duplicate account codes: %s" % ", ".join(dup))
        bad = [a.code for a in accounts if a.section not in SECTIONS]
        if bad:
            raise ChartError("unknown section for %s (use one of %s)" % (", ".join(bad), ", ".join(SECTIONS)))
        self.all = list(accounts)
        self.all_by_code = {a.code: a for a in accounts}
        self.accounts = [a for a in accounts if a.section != "balance_sheet"]
        self.by_code = {a.code: a for a in self.accounts}
        if not any(a.section == "revenue" for a in self.accounts):
            raise ChartError("the chart has no revenue accounts")

    @classmethod
    def load(cls, path):
        with open(path, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
        need = {"code", "name", "line", "section"}
        if not rows or not need <= set(rows[0]):
            raise ChartError("%s: expected columns code,name,line,section[,cash]" % Path(path).name)
        return cls([Account(r["code"].strip(), r["name"].strip(), r["line"].strip() or r["name"].strip(),
                            r["section"].strip().lower(), (r.get("cash") or "").strip()) for r in rows])

    def save(self, path):
        with open(path, "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(["code", "name", "line", "section", "cash"])
            for a in self.all:
                w.writerow([a.code, a.name, a.line, a.section, a.cash])

    def codes(self, *sections):
        return [a.code for a in self.accounts if a.section in sections]

    @property
    def has_cost_of_sales(self):
        return bool(self.codes("cost_of_sales"))

    @property
    def has_interest(self):
        return bool(self.codes("interest"))

    def lines(self, *sections):
        """Management lines in chart order: [(line, section)]."""
        out = []
        for a in self.accounts:
            if (a.line, a.section) not in out and (not sections or a.section in sections):
                out.append((a.line, a.section))
        return out

    def totals(self, pnl):
        s = lambda sec: sum(pnl.get(c, 0.0) for c in self.codes(sec))
        rev, cos, opex = s("revenue"), s("cost_of_sales"), s("opex")
        gp = rev - cos
        ebitda = gp - opex
        ebit = ebitda - s("depreciation")
        pbt = ebit - s("interest")
        return {"revenue": rev, "cost_of_sales": cos, "gross_profit": gp, "opex": opex, "ebitda": ebitda,
                "depreciation": s("depreciation"), "ebit": ebit, "interest": s("interest"), "pbt": pbt,
                "tax": s("tax"), "net_income": pbt - s("tax")}

    def by_line(self, pnl):
        out = {line: 0.0 for line, _ in self.lines()}
        for code, v in pnl.items():
            a = self.by_code.get(code)
            if a:
                out[a.line] += v
        return out
