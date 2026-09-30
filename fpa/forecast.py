"""
Rolling forecast and scenarios.

A reforecast after month N (e.g. 6+6 after September) combines the closed months'
actuals with a driver-based forecast for the open months. The forecast starts
from the actual closing position of month N (properties under management and
average rent) and uses the budget drivers updated by the latest estimate in
companies/<slug>/forecast.toml. Scenarios apply further changes on top of it.
Companies budgeted by rules are reforecast by fpa/rules.py (run-rate method).
"""
import tomllib

from . import budget, ledger
from .drivers import driver_path, pnl


class ForecastError(ValueError):
    pass


def load_le(co):
    if not co.forecast_path.exists():
        return {}
    with open(co.forecast_path, "rb") as f:
        return tomllib.load(f)


def _ytd_lets_factor(act, bud, months):
    a = sum(act[m]["kpis"]["let_only_lets"] + act[m]["kpis"]["relets"] for m in months)
    b = sum(bud[m]["let_only_lets"] + bud[m]["relets"] for m in months)
    return min(1.2, max(0.6, a / b)) if b else 1.0


def reforecast(co, last_closed, le=None, scenario=None, root=None):
    """{"label", "months", "closed", "open", "pnl": {month: {code: amount}}, "source": {month: 'A'|'F'},
    "totals", "fy", "fy_totals", "budget_fy_totals", "assumptions"}."""
    le = load_le(co) if le is None else le
    fy = co.fy_months()
    if last_closed not in fy:
        raise ForecastError("%s is not in %s" % (last_closed, co.fy_label()))
    closed = [m for m in fy if m <= last_closed]
    open_ = [m for m in fy if m > last_closed]
    avail = set(ledger.available_months(co, root))
    if any(m not in avail for m in closed):
        raise ForecastError("actuals missing for the closed months")
    act = {}
    for m in closed:
        p, k = ledger.read_month(co, m, root)
        act[m] = {"pnl": p, "kpis": k}
    b = budget.build(co, root=root)
    if co.model != "lettings":
        from . import rules
        return rules.reforecast(co, b, act, closed, open_, le, scenario)
    bud = {d["month"]: d for d in b["drivers"]}

    est = dict(le.get("latest_estimate", {}))
    est.pop("notes", None)
    lets_factor = est.pop("lets_factor", 1.0)
    if lets_factor == "ytd":
        lets_factor = _ytd_lets_factor(act, bud, closed)
    params = dict(est)
    sc = (le.get("scenarios", {}).get(scenario) if scenario else None) or {}
    if scenario and not sc:
        raise ForecastError("unknown scenario %s" % scenario)
    lets_factor *= sc.get("lets_factor_mult", 1.0)
    base = {**co.cfg["portfolio"], **co.cfg["fees"]}
    if sc:
        params["lost_rate_annual"] = base["lost_rate_annual"] * sc.get("lost_rate_mult", 1.0)
        params["gained_per_month"] = base["gained_per_month"] * sc.get("gained_mult", 1.0)
        params["rent_growth_monthly"] = params.get("rent_growth_monthly", base["rent_growth_monthly"]) \
            + sc.get("rent_growth_add", 0.0)

    by_month = {m: act[m]["pnl"] for m in closed}
    drivers = []
    if open_:
        k = act[closed[-1]]["kpis"] if closed else None
        opening = {} if not closed else {
            "pum": k["pum_close"],
            "avg_rent": k["avg_rent"] * (1 + params.get("rent_growth_monthly", base["rent_growth_monthly"])),
        }
        events = {open_[0]: {"set": params}}
        for m in open_:
            events.setdefault(m, {})["lets_factor"] = lets_factor
        drivers = driver_path(co, open_, opening=opening, events=events)
        # the budget roster applies from the first open month (planned hires included)
        for d in drivers:
            bd = bud[d["month"]]
            d["roster"], d["headcount"] = bd["roster"], bd["headcount"]
            by_month[d["month"]] = pnl(d, co)
    fy_pnl = {c: sum(by_month[m][c] for m in fy) for c in co.chart.by_code}
    bfy, bfy_t = budget.full_year(b)
    return {
        "label": "%d+%d" % (len(closed), len(open_)), "scenario": scenario or "base",
        "months": fy, "closed": closed, "open": open_, "pnl": by_month,
        "source": {m: ("A" if m in act else "F") for m in fy},
        "totals": {m: co.chart.totals(by_month[m]) for m in fy},
        "fy": fy_pnl, "fy_totals": co.chart.totals(fy_pnl), "budget_fy": bfy, "budget_fy_totals": bfy_t,
        "drivers": drivers, "lets_factor": lets_factor,
        "closing_pum": drivers[-1]["pum_close"] if drivers else act[closed[-1]]["kpis"]["pum_close"],
        "budget_closing_pum": b["drivers"][-1]["pum_close"],
        "assumptions": {**params, "lets_factor": lets_factor},
    }


def scenarios(co, last_closed, le=None, root=None):
    le = load_le(co) if le is None else le
    out = {"base": reforecast(co, last_closed, le, None, root)}
    for name in le.get("scenarios", {}):
        out[name] = reforecast(co, last_closed, le, name, root)
    return out
