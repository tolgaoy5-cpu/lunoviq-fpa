"""
Budget for the financial year: the Python mirror of the Assumptions, Drivers
(or Prior year) and Budget sheets.
"""
from .drivers import driver_path, pnl


def build(co, fy_start_year=None, root=None):
    """{"months", "drivers", "pnl": {month: {code: amount}}, "totals": {month: subtotals}}.
    The lettings template budgets from operating drivers; any other company by rules."""
    months = co.fy_months(fy_start_year)
    if co.model == "lettings":
        path = driver_path(co, months)
        by_month = {d["month"]: pnl(d, co) for d in path}
    else:
        from . import rules
        path, by_month = [], rules.budget(co, months, root=root)
    return {"months": months, "drivers": path, "pnl": by_month, "chart": co.chart,
            "totals": {m: co.chart.totals(p) for m, p in by_month.items()}}


def full_year(result):
    """{code: FY amount} plus subtotals."""
    fy = {}
    for p in result["pnl"].values():
        for code, v in p.items():
            fy[code] = fy.get(code, 0.0) + v
    return fy, result["chart"].totals(fy)
