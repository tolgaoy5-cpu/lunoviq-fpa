"""
Synthetic actuals for Kestrel Row Lettings Ltd (fictional).

Produces the monthly trial balance and KPI exports from April 2025 (prior year)
to the last closed month. The same driver engine as the budget is used, with
random noise (seeded, so every run is identical) and a set of business events.
The events give the variance analysis real stories to explain:

  Prior year: management fee 11.5% (raised to 12% in April 2026, as budgeted).

  FY2026/27 H1
  - Market rents rise faster than budgeted (0.6% a month vs 0.4%)   -> favourable fee rate
  - May: a landlord goes into administration owing fees (GBP 4,000 written off)
  - June: a property manager resigns; replacement starts in September;
          agency cover in July and August                            -> salary saving, overhead cost
  - June-September: a softer lettings market (-12% new lets)         -> adverse new business
  - July: a portfolio landlord sells 24 properties                   -> adverse management fees
  - August: a property portal raises its price by 15%                -> adverse marketing

Usage:  python -m fpa.simulate            (writes companies/kestrel-row/actuals/<month>/)
"""
import copy
import random

from . import company, ledger
from .drivers import driver_path, pnl

SEED = 20260930
NOISE = {"gained": 0.25, "lost": 0.20, "let_only": 0.12, "relets": 0.10, "renewals": 0.10,
         "certs": 0.08, "maint": 0.06, "cost_utilities": 0.05, "cost_marketing": 0.06,
         "cost_other": 0.08, "cost_motor": 0.08}


def _noise(seed):
    rng = random.Random(seed)

    def f(key):
        return min(1.25, max(0.75, rng.gauss(1.0, NOISE.get(key, 0.0))))
    return f


def prior_year_roster(co):
    roster = copy.deepcopy(co.cfg["staff"]["roles"])
    roster["property_manager"][0] -= 1                       # sixth property manager joined April 2026
    for role in roster.values():
        role[1] = round(role[1] / 1.03 / 100) * 100          # before the April 2026 pay review
    return roster


def events(co):
    base = {k: co.cfg["costs"][k] for k in ("portals", "office_rent_rates", "insurance",
                                            "professional_fees", "software_per_user")}
    ev = {
        "2025-04": {"set": {"portals": 3000.0, "office_rent_rates": 5300.0, "insurance": 1250.0,
                            "professional_fees": 2000.0, "software_per_user": 80.0,
                            "management_fee": 0.115}},
        "2026-04": {"set": dict(base, rent_growth_monthly=0.006,
                                management_fee=co.cfg["fees"]["management_fee"]),
                    "roster": copy.deepcopy(co.cfg["staff"]["roles"])},
        "2026-05": {"add": {"6600": 4000.0}},
        "2026-06": {"leave": ["property_manager"], "lets_factor": 0.88},
        "2026-07": {"extra_lost": 24, "add": {"6900": 3000.0}, "lets_factor": 0.88},
        "2026-08": {"set": {"portals": round(base["portals"] * 1.15, 2)}, "add": {"6900": 3000.0},
                    "lets_factor": 0.88},
        "2026-09": {"join": ["property_manager"], "lets_factor": 0.88},
    }
    return ev


def simulate(co=None, last=None, seed=SEED):
    """{month: (pnl by account, kpis)} from April 2025 to the last closed month."""
    co = co or company.load()
    last = last or co.cfg["company"]["last_closed_month"]
    months = company.month_range(co.cfg["company"]["first_actual_month"], last)
    opening = {"pum": 397, "avg_rent": 1385.0, "roster": prior_year_roster(co)}
    path = driver_path(co, months, opening=opening, events=events(co), noise=_noise(seed),
                       whole_counts=True, planned_hires=False)
    rng = random.Random(seed + 1)
    out = {}
    for d in path:
        arrears = 0.020 + rng.gauss(0, 0.002) + (0.011 if d["month"] == "2026-05" else 0.0)
        kpis = {"pum_open": d["pum_open"], "gained": d["gained"], "lost": d["lost"],
                "pum_close": d["pum_close"], "avg_rent": round(d["avg_rent"], 2),
                "let_only_lets": d["let_only_lets"], "relets": d["relets"],
                "renewals": d["renewals"], "certificates": d["certificates"],
                "headcount": d["headcount"], "arrears_pct": round(arrears, 4)}
        out[d["month"]] = (pnl(d, co), kpis)
    return out


def main():
    co = company.load("kestrel-row")
    data = simulate(co)
    for month, (p, k) in data.items():
        ledger.write_month(co, month, p, k)
    print("Wrote %d months of actuals (%s to %s) to %s"
          % (len(data), min(data), max(data), co.actuals_dir))


if __name__ == "__main__":
    main()
