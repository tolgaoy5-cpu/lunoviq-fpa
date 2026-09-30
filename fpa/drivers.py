"""
Driver engine shared by the budget, the forecast and the synthetic actuals.

driver_path() rolls the operating drivers forward month by month (portfolio,
rent, lettings activity, headcount); pnl() turns one month of drivers into
amounts by GL account. The budget uses the configured drivers as expected
values; the simulator adds events and noise to produce "actual" months.

Events, keyed by month:
    set        {param: value}   permanent change from that month (e.g. a portal price rise)
    extra_lost int              properties lost on top of normal churn (a landlord leaving)
    lets_factor float           one-month factor on let-only and re-let activity
    leave / join [role, ...]    permanent headcount changes
    add        {code: amount}   one-off amounts (e.g. a bad-debt write-off, agency cover)
    roster     {role: [n, salary]}  replaces the whole roster (e.g. the April pay review)
"""
import copy


PARAM_SECTIONS = ("portfolio", "fees", "costs", "staff")


def _params(co):
    p = {}
    for sec in PARAM_SECTIONS:
        for k, v in co.cfg[sec].items():
            if not isinstance(v, dict):
                p[k] = v
    return p


def driver_path(co, months, opening=None, events=None, noise=None, whole_counts=False,
                planned_hires=True):
    """List of monthly driver dicts. `opening`: {"pum", "avg_rent", "roster"}; `noise(key)`
    returns a multiplicative factor (1.0 for the budget); whole_counts rounds activity;
    planned_hires applies the hires in config [staff] (the budget plan)."""
    p = _params(co)
    events = events or {}
    noise = noise or (lambda key: 1.0)
    opening = opening or {}
    pum = opening.get("pum", co.cfg["portfolio"]["opening_pum"])
    rent = opening.get("avg_rent", co.cfg["portfolio"]["opening_avg_rent"])
    roster = copy.deepcopy(opening.get("roster", co.cfg["staff"]["roles"]))
    hires = {m: r for r, m in co.cfg["staff"].get("hires", [])} if planned_hires else {}
    rnd = (lambda x: float(round(x))) if whole_counts else (lambda x: x)
    out = []
    first = True
    for m in months:
        ev = events.get(m, {})
        p.update(ev.get("set", {}))
        if "roster" in ev:
            roster = copy.deepcopy(ev["roster"])
        for role in ev.get("leave", []):
            roster[role][0] -= 1
        for role in ev.get("join", []) + ([hires[m]] if m in hires else []):
            roster[role][0] += 1
        if not first:
            rent *= 1 + p["rent_growth_monthly"]
        first = False
        s = co.season("lets", m)
        lets_f = ev.get("lets_factor", 1.0)
        gained = rnd(p["gained_per_month"] * s * noise("gained"))
        lost = rnd(pum * p["lost_rate_annual"] / 12 * noise("lost")) + ev.get("extra_lost", 0)
        pum_close = pum + gained - lost
        avg_pum = (pum + pum_close) / 2
        d = {
            "month": m, "pum_open": pum, "gained": gained, "lost": lost, "pum_close": pum_close,
            "avg_pum": avg_pum, "avg_rent": rent,
            "let_only_lets": rnd(p["let_only_lets_per_month"] * s * lets_f * noise("let_only")),
            "relets": rnd(avg_pum * p["relet_rate_annual"] / 12 * s * lets_f * noise("relets")),
            "renewals": rnd(avg_pum * p["renewal_rate_annual"] / 12 * noise("renewals")),
            "certificates": rnd(avg_pum * p["certificates_per_property_year"] / 12 * noise("certs")),
            "maintenance_spend": avg_pum * p["maintenance_spend_per_property"] * noise("maint"),
            "roster": copy.deepcopy(roster),
            "headcount": sum(n for n, _ in roster.values()),
            "params": dict(p),
            "add": dict(ev.get("add", {})),
            "cost_noise": {k: noise("cost_" + k) for k in ("utilities", "marketing", "other", "motor")},
        }
        out.append(d)
        pum = pum_close
    return out


def pnl(d, co):
    """{GL code: amount} for one month of drivers (all amounts positive)."""
    p, r = d["params"], {}
    r["4000"] = d["avg_pum"] * d["avg_rent"] * p["collection_rate"] * p["management_fee"]
    r["4010"] = d["let_only_lets"] * d["avg_rent"] * 12 * p["let_only_fee"]
    r["4020"] = d["relets"] * p["setup_fee"]
    r["4030"] = d["renewals"] * p["renewal_fee"]
    r["4040"] = d["certificates"] * p["certificate_margin"]
    r["4050"] = d["maintenance_spend"] * p["maintenance_commission"]
    new_business = r["4010"] + r["4020"]
    fee_revenue = sum(r.values())

    sal = ni = pen = 0.0
    for n, salary in d["roster"].values():
        sal += n * salary / 12
        ni += n * p["employer_ni_rate"] * max(0.0, salary - p["employer_ni_threshold"]) / 12
        pen += n * p["pension_rate"] * salary / 12
    r["5000"], r["5010"], r["5020"] = sal, ni, pen
    r["5030"] = p["commission_rate"] * new_business
    cn = d["cost_noise"]
    r["6000"] = p["office_rent_rates"]
    r["6010"] = p["utilities"] * co.season("utilities", d["month"]) * cn["utilities"]
    r["6100"] = p["portals"]
    r["6110"] = (p["marketing_base"] + p["marketing_rate"] * new_business) * cn["marketing"]
    r["6200"] = p["software_base"] + p["software_per_user"] * d["headcount"]
    r["6300"] = p["insurance"]
    r["6400"] = p["professional_fees"]
    r["6500"] = p["motor_per_property_manager"] * d["roster"]["property_manager"][0] * cn["motor"]
    r["6600"] = p["bad_debt_rate"] * fee_revenue
    r["6900"] = p["other_overheads"] * cn["other"]
    r["7000"] = p["depreciation"]
    for code, amount in d["add"].items():
        r[code] = r.get(code, 0.0) + amount
    ebit = sum(r[c] for c in r if c < "5000") - sum(r[c] for c in r if "5000" <= c < "9000")
    r["9000"] = p["corporation_tax_rate"] * ebit          # monthly accrual (a loss month accrues a credit)
    return {a.code: round(r.get(a.code, 0.0), 2) for a in co.chart.accounts}
