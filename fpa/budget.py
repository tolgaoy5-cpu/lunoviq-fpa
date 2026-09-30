"""
Driver-based budget for the financial year: the Python mirror of the
Assumptions, Drivers and Budget sheets.
"""
from . import accounts
from .drivers import driver_path, pnl


def build(co, fy_start_year=None):
    """{"months", "drivers", "pnl": {month: {code: amount}}, "totals": {month: subtotals}}."""
    months = co.fy_months(fy_start_year)
    path = driver_path(co, months)
    by_month = {d["month"]: pnl(d, co) for d in path}
    return {"months": months, "drivers": path, "pnl": by_month,
            "totals": {m: accounts.totals(p) for m, p in by_month.items()}}


def full_year(result):
    """{code: FY amount} plus subtotals."""
    fy = {}
    for p in result["pnl"].values():
        for code, v in p.items():
            fy[code] = fy.get(code, 0.0) + v
    return fy, accounts.totals(fy)
