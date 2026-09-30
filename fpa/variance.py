"""
Budget vs actual for a reporting month: month and year to date, by account and
management P&L line, with materiality flags, driver bridges and commentary.

Sign convention: a variance is FAVOURABLE when positive (revenue above budget,
costs below budget).

For the lettings template, driver bridges split the largest variances into causes:
  management fees   volume (properties) / rate (average rent) / other (collection, fee)
  let-only fees     volume (lets) / rate (rent) / other
  set-up fees       volume (re-lets) / other
  staff costs       headcount / other (pay, NI, commission)
Commentary combines generated text (the numbers) with analyst notes from
companies/<slug>/commentary.toml (the business reasons).
"""
import tomllib

from . import budget, ledger

STAFF = ("5000", "5010", "5020", "5030")


class ReportError(ValueError):
    pass


def _fav(kind, actual, budget_):
    return actual - budget_ if kind == "revenue" else budget_ - actual


def _line(label, kind, a, b, abs_t, pct_t, always=None):
    """Flag when the variance passes both thresholds, or the absolute `always` threshold on its own."""
    var = _fav(kind, a, b)
    pct = var / abs(b) if b else 0.0
    return {"label": label, "actual": a, "budget": b, "var": var, "pct": pct,
            "flag": (abs(var) >= abs_t and abs(pct) >= pct_t) or (always is not None and abs(var) >= always)}


def load_notes(co):
    p = co.commentary_path
    if not p.exists():
        return {}
    with open(p, "rb") as f:
        return tomllib.load(f)


def _bridges(bm, am, co):
    """Driver bridge for one month: {name: {"volume", "rate", "other", "var"}} (favourable +)."""
    p = bm["params"]
    a_avg = (am["kpis"]["pum_open"] + am["kpis"]["pum_close"]) / 2
    a_rent = am["kpis"]["avg_rent"]
    out = {}
    var = am["pnl"]["4000"] - bm["pnl"]["4000"]
    vol = (a_avg - bm["avg_pum"]) * bm["avg_rent"] * p["collection_rate"] * p["management_fee"]
    rate = a_avg * (a_rent - bm["avg_rent"]) * p["collection_rate"] * p["management_fee"]
    out["Management fees"] = {"volume": vol, "rate": rate, "other": var - vol - rate, "var": var}
    var = am["pnl"]["4010"] - bm["pnl"]["4010"]
    a_lets = am["kpis"]["let_only_lets"]
    vol = (a_lets - bm["let_only_lets"]) * bm["avg_rent"] * 12 * p["let_only_fee"]
    rate = a_lets * (a_rent - bm["avg_rent"]) * 12 * p["let_only_fee"]
    out["Let-only fees"] = {"volume": vol, "rate": rate, "other": var - vol - rate, "var": var}
    var = am["pnl"]["4020"] - bm["pnl"]["4020"]
    vol = (am["kpis"]["relets"] - bm["relets"]) * p["setup_fee"]
    out["Tenancy set-up fees"] = {"volume": vol, "rate": 0.0, "other": var - vol, "var": var}
    b_staff = sum(bm["pnl"][c] for c in STAFF)
    a_staff = sum(am["pnl"][c] for c in STAFF)
    var = b_staff - a_staff
    hc = (bm["headcount"] - am["kpis"]["headcount"]) * (b_staff / bm["headcount"])
    out["Staff costs"] = {"volume": hc, "rate": 0.0, "other": var - hc, "var": var}
    return out


def _sum_bridges(items):
    out = {}
    for b in items:
        for k, v in b.items():
            o = out.setdefault(k, {"volume": 0.0, "rate": 0.0, "other": 0.0, "var": 0.0})
            for f in o:
                o[f] += v[f]
    return out


def _k(v):
    return "£%.1fk" % (abs(v) / 1000)


def _fa(v):
    return "favourable" if v >= 0 else "adverse"


def _signed(v):
    return ("+" if v >= 0 else "\u2212") + _k(v)


def _bridge_text(name, b, facts):
    parts = []
    if abs(b["volume"]) >= 300:
        parts.append("%s (%s)" % (facts["volume"], _signed(b["volume"])))
    if abs(b["rate"]) >= 300:
        parts.append("%s (%s)" % (facts["rate"], _signed(b["rate"])))
    if abs(b["other"]) >= 300:
        parts.append("other effects (%s)" % _signed(b["other"]))
    return "; ".join(parts)


def analyse(co, month, root=None, notes=None):
    """Full budget-vs-actual analysis for `month` (a closed month of the budget year)."""
    fy = co.fy_months()
    if month not in fy:
        raise ReportError("%s is not in %s" % (month, co.fy_label()))
    avail = set(ledger.available_months(co, root))
    ytd = [m for m in fy if m <= month]
    missing = [m for m in ytd if m not in avail]
    if missing:
        raise ReportError("actuals missing for %s" % ", ".join(missing))
    b = budget.build(co, root=root)
    lettings = co.model == "lettings"
    if lettings:
        bmonths = {d["month"]: dict(d, pnl=b["pnl"][d["month"]]) for d in b["drivers"]}
    else:
        bmonths = {m: {"pnl": b["pnl"][m]} for m in fy}
    act = {}
    for m in ytd:
        p, k = ledger.read_month(co, m, root)
        act[m] = {"pnl": p, "kpis": k}
    rep = co.cfg["reporting"]
    notes = load_notes(co) if notes is None else notes
    chart = co.chart

    def agg(src, months, code):
        return sum(src[m]["pnl"][code] for m in months)

    lines = {"month": {}, "ytd": {}}
    m_always, y_always = rep.get("month_always"), rep.get("ytd_always")
    for code, acc in chart.by_code.items():
        lines["month"][code] = _line(acc.name, acc.kind, act[month]["pnl"][code], bmonths[month]["pnl"][code],
                                     rep["month_abs"], rep["month_pct"], m_always)
        lines["ytd"][code] = _line(acc.name, acc.kind, agg(act, ytd, code), agg(bmonths, ytd, code),
                                   rep["ytd_abs"], rep["ytd_pct"], y_always)
    totals = {}
    for period, months in (("month", [month]), ("ytd", ytd)):
        a_p = {c: agg(act, months, c) for c in chart.by_code}
        b_p = {c: agg(bmonths, months, c) for c in chart.by_code}
        a_t, b_t = chart.totals(a_p), chart.totals(b_p)
        a_g, b_g = chart.by_line(a_p), chart.by_line(b_p)
        abs_t, pct_t = (rep["month_abs"], rep["month_pct"]) if period == "month" else (rep["ytd_abs"], rep["ytd_pct"])
        always = m_always if period == "month" else y_always
        t = {}
        for g, section in chart.lines("revenue", "cost_of_sales", "opex"):
            t[g] = _line(g, "revenue" if section == "revenue" else "expense", a_g[g], b_g[g], abs_t, pct_t, always)
        keys = [("revenue", "Total revenue", "revenue")]
        if chart.has_cost_of_sales:
            keys += [("cost_of_sales", "Total cost of sales", "expense"), ("gross_profit", "Gross profit", "revenue")]
        keys += [("opex", "Total operating costs", "expense"), ("ebitda", "EBITDA", "revenue"),
                 ("net_income", "Net income", "revenue")]
        for key, label, kind in keys:
            t[key] = _line(label, kind, a_t[key], b_t[key], abs_t, pct_t, always)
        t["margin"] = {"actual": a_t["ebitda"] / a_t["revenue"] if a_t["revenue"] else 0.0,
                       "budget": b_t["ebitda"] / b_t["revenue"] if b_t["revenue"] else 0.0}
        if chart.has_cost_of_sales:
            t["gp_margin"] = {"actual": a_t["gross_profit"] / a_t["revenue"] if a_t["revenue"] else 0.0,
                              "budget": b_t["gross_profit"] / b_t["revenue"] if b_t["revenue"] else 0.0}
        totals[period] = t

    if not lettings:
        kp = act[month]["kpis"]
        bridges = {"month": _flex_bridges(co, act, bmonths, [month]), "ytd": _flex_bridges(co, act, bmonths, ytd)}
        res = {"month": month, "ytd_months": ytd, "fy": co.fy_label(), "company": co.name, "model": co.model,
               "thresholds": {"month": rep["month_abs"], "ytd": rep["ytd_abs"]},
               "lines": lines, "totals": totals, "bridges": bridges,
               "kpis": [(k.replace("_", " ").capitalize(), v, None, "number") for k, v in kp.items()],
               "ytd_kpis": [], "actuals": act, "budget": bmonths, "chart": chart}
        res["commentary"] = commentary(res, notes)
        return res

    bridges = {"month": _bridges(bmonths[month], act[month], co),
               "ytd": _sum_bridges([_bridges(bmonths[m], act[m], co) for m in ytd])}

    bm, am = bmonths[month], act[month]["kpis"]
    kpis = [
        ("Properties under management (end)", am["pum_close"], bm["pum_close"], "number"),
        ("Average monthly rent", am["avg_rent"], bm["avg_rent"], "gbp"),
        ("Let-only lets", am["let_only_lets"], bm["let_only_lets"], "number"),
        ("Re-lets on managed stock", am["relets"], bm["relets"], "number"),
        ("Renewals", am["renewals"], bm["renewals"], "number"),
        ("Headcount", am["headcount"], bm["headcount"], "number"),
        ("Arrears (% of rent roll)", am["arrears_pct"], None, "pct"),
    ]
    ytd_kpis = [
        ("Let-only lets", sum(act[m]["kpis"]["let_only_lets"] for m in ytd), sum(bmonths[m]["let_only_lets"] for m in ytd)),
        ("Re-lets on managed stock", sum(act[m]["kpis"]["relets"] for m in ytd), sum(bmonths[m]["relets"] for m in ytd)),
        ("Properties gained", sum(act[m]["kpis"]["gained"] for m in ytd), sum(bmonths[m]["gained"] for m in ytd)),
        ("Properties lost", sum(act[m]["kpis"]["lost"] for m in ytd), sum(bmonths[m]["lost"] for m in ytd)),
    ]
    res = {"month": month, "ytd_months": ytd, "fy": co.fy_label(), "company": co.name, "model": co.model,
           "lines": lines, "totals": totals, "bridges": bridges, "kpis": kpis, "ytd_kpis": ytd_kpis,
           "actuals": act, "budget": bmonths, "chart": chart}
    res["commentary"] = commentary(res, notes)
    return res


def _flex_bridges(co, act, bmonths, months):
    """Flexed-budget split for costs budgeted as a share of revenue, by management line:
    volume = the cost change explained by sales being above/below budget (at the budgeted share);
    rate = the change in the cost share itself (actual share - budgeted share, on actual sales).
    Favourable +."""
    from . import rules as rules_mod
    chart = co.chart
    rs = rules_mod.rules(co)
    rev = chart.codes("revenue")
    a_rev = sum(act[m]["pnl"][c] for m in months for c in rev)
    b_rev = sum(bmonths[m]["pnl"][c] for m in months for c in rev)
    out = {}
    for a in chart.accounts:
        r = rs.get(a.code, {})
        if r.get("method") != "pct_revenue" or a.section not in ("cost_of_sales", "opex"):
            continue
        actual = sum(act[m]["pnl"][a.code] for m in months)
        budget_ = sum(bmonths[m]["pnl"][a.code] for m in months)
        b_share = budget_ / b_rev if b_rev else 0.0
        a_share = actual / a_rev if a_rev else 0.0
        o = out.setdefault(a.line, {"volume": 0.0, "rate": 0.0, "other": 0.0, "var": 0.0,
                                    "actual_share": 0.0, "budget_share": 0.0})
        o["volume"] += -b_share * (a_rev - b_rev)
        o["rate"] += (b_share - a_share) * a_rev
        o["var"] += budget_ - actual
        o["actual_share"] += a_share
        o["budget_share"] += b_share
    for o in out.values():
        o["other"] = o["var"] - o["volume"] - o["rate"]
    return out


def _facts(res, period):
    """Plain-language driver facts for the bridge text."""
    b, a = res["budget"], res["actuals"]
    months = [res["month"]] if period == "month" else res["ytd_months"]
    m = res["month"]
    pum_a, pum_b = a[m]["kpis"]["pum_close"], b[m]["pum_close"]
    rent_a, rent_b = a[m]["kpis"]["avg_rent"], b[m]["avg_rent"]
    lets_a = sum(a[x]["kpis"]["let_only_lets"] for x in months)
    lets_b = sum(b[x]["let_only_lets"] for x in months)
    rel_a = sum(a[x]["kpis"]["relets"] for x in months)
    rel_b = sum(b[x]["relets"] for x in months)
    hc_a, hc_b = a[m]["kpis"]["headcount"], b[m]["headcount"]
    return {
        "Management fees": {"volume": "portfolio of %.0f properties vs %.0f budgeted" % (pum_a, pum_b),
                            "rate": "average rent £%.0f vs £%.0f" % (rent_a, rent_b)},
        "Let-only fees": {"volume": "%.0f let-only lets vs %.0f budgeted" % (lets_a, lets_b),
                          "rate": "higher average rent" if rent_a >= rent_b else "lower average rent"},
        "Tenancy set-up fees": {"volume": "%.0f re-lets vs %.0f budgeted" % (rel_a, rel_b), "rate": ""},
        "Staff costs": {"volume": "headcount %.0f vs %.0f budgeted" % (hc_a, hc_b), "rate": ""},
    }


BRIDGE_ACCOUNTS = {"Management fees": ["4000"], "Let-only fees": ["4010"], "Tenancy set-up fees": ["4020"],
                   "Staff costs": list(STAFF)}


def commentary(res, notes):
    """{period: [(label, text, favourable)]}: headline first, then flagged lines, largest first."""
    out = {}
    for period in ("month", "ytd"):
        t = res["totals"][period]
        what = "In %s" % _month_name(res["month"]) if period == "month" else "Year to date"
        items = [("Headline", "%s, revenue was %s %s budget and EBITDA %s %s budget: EBITDA %s at a %.1f%% "
                  "margin (budget %.1f%%)."
                  % (what, _k(t["revenue"]["var"]), "above" if t["revenue"]["var"] >= 0 else "below",
                     _k(t["ebitda"]["var"]), "above" if t["ebitda"]["var"] >= 0 else "below",
                     _k(t["ebitda"]["actual"]), t["margin"]["actual"] * 100, t["margin"]["budget"] * 100),
                  t["ebitda"]["var"] >= 0)]
        lettings = res["model"] == "lettings"
        facts = _facts(res, period) if lettings else {}
        bridge_accounts = BRIDGE_ACCOUNTS if lettings else {}
        if "gp_margin" in t:
            items[0] = (items[0][0], items[0][1][:-1] + "; gross margin %.1f%% vs %.1f%%."
                        % (t["gp_margin"]["actual"] * 100, t["gp_margin"]["budget"] * 100), items[0][2])
        tax = set(res["chart"].codes("tax"))
        note_months = [res["month"]] if period == "month" else res["ytd_months"]
        flagged = sorted((c for c, l in res["lines"][period].items() if l["flag"] and c not in tax),
                         key=lambda c: -abs(res["lines"][period][c]["var"]))
        done_bridges = set()
        for code in flagged:
            l = res["lines"][period][code]
            bridge = next((n for n, codes in bridge_accounts.items() if code in codes), None)
            label = bridge if bridge in ("Staff costs",) else l["label"]
            if bridge and bridge in done_bridges:
                continue
            if bridge == "Staff costs":
                var = res["bridges"][period]["Staff costs"]["var"]
                text = "%s %s (%+.0f%%)" % (_k(var), _fa(var), 100 * var / abs(sum(res["lines"][period][c]["budget"] for c in STAFF)))
            else:
                text = "%s %s (%+.0f%%)" % (_k(l["var"]), _fa(l["var"]), l["pct"] * 100)
            if bridge:
                done_bridges.add(bridge)
                detail = _bridge_text(bridge, res["bridges"][period][bridge], facts[bridge])
                if detail:
                    text += ": " + detail
            codes = bridge_accounts.get(bridge, [code])
            reasons = []
            for m in note_months:
                for c in codes:
                    n = notes.get(m, {}).get(c)
                    if n and n not in reasons:
                        reasons.append(n)
            flex = None if lettings else res["bridges"][period].get(res["chart"].by_code[code].line)
            if flex and all(c in flagged for c in _line_codes(res, code)):
                text += "; cost share %.1f%% of sales vs %.1f%% budgeted" % (flex["actual_share"] * 100,
                                                                              flex["budget_share"] * 100)
            if reasons:
                text += ". " + " ".join(reasons)
            items.append((label, text.rstrip(".") + ".", (res["bridges"][period][bridge]["var"] if bridge == "Staff costs" else l["var"]) >= 0))
        if not lettings:
            items += _flex_commentary(res, period, notes, note_months, set(flagged))
        out[period] = items
    return out


def _line_codes(res, code):
    line = res["chart"].by_code[code].line
    return [a.code for a in res["chart"].accounts if a.line == line and a.section in ("cost_of_sales", "opex")]


def _flex_commentary(res, period, notes, note_months, flagged=()):
    """Cost lines whose share of sales moved materially (even when the total variance is small);
    lines whose accounts are all flagged already are covered by the account items."""
    rep_abs = res.get("thresholds", {}).get(period)
    out = []
    for line, b in sorted(res["bridges"][period].items(), key=lambda kv: -abs(kv[1]["rate"])):
        if rep_abs is not None and abs(b["rate"]) < rep_abs:
            continue
        codes = [a.code for a in res["chart"].accounts if a.line == line]
        if codes and all(c in flagged for c in codes):
            continue
        text = "%s %s overall; cost share %.1f%% of sales vs %.1f%% budgeted (%s), lower/higher sales (%s)" % (
            _k(b["var"]), _fa(b["var"]), b["actual_share"] * 100, b["budget_share"] * 100, _signed(b["rate"]),
            _signed(b["volume"]))
        text = text.replace("lower/higher sales", "lower sales" if b["volume"] >= 0 else "higher sales")
        reasons = []
        for m in note_months:
            for c in codes:
                n = notes.get(m, {}).get(c)
                if n and n not in reasons:
                    reasons.append(n)
        if reasons:
            text += ". " + " ".join(reasons)
        out.append((line + " (cost share)", text.rstrip(".") + ".", b["rate"] >= 0))
    return out


def _month_name(m):
    return ["January", "February", "March", "April", "May", "June", "July", "August", "September",
            "October", "November", "December"][int(m[5:]) - 1] + " " + m[:4]
