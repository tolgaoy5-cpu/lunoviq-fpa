"""
Budget vs actual for a reporting month: month and year to date, by account and
management P&L line, with materiality flags, driver bridges and commentary.

Sign convention: a variance is FAVOURABLE when positive (revenue above budget,
costs below budget).

Driver bridges split the largest variances into their causes:
  management fees   volume (properties) / rate (average rent) / other (collection, fee)
  let-only fees     volume (lets) / rate (rent) / other
  set-up fees       volume (re-lets) / other
  staff costs       headcount / other (pay, NI, commission)
Commentary combines generated text (the numbers) with analyst notes from
config/commentary.toml (the business reasons).
"""
import tomllib

from . import accounts, budget, company, ledger

STAFF = ("5000", "5010", "5020", "5030")


class ReportError(ValueError):
    pass


def _fav(code_or_kind, actual, budget_):
    kind = code_or_kind if code_or_kind in ("revenue", "expense") else accounts.BY_CODE[code_or_kind].kind
    return actual - budget_ if kind == "revenue" else budget_ - actual


def _line(label, kind, a, b, abs_t, pct_t):
    var = _fav(kind, a, b)
    pct = var / abs(b) if b else 0.0
    return {"label": label, "actual": a, "budget": b, "var": var, "pct": pct,
            "flag": abs(var) >= abs_t and abs(pct) >= pct_t}


def load_notes(path=None):
    p = path or (company.ROOT / "config" / "commentary.toml")
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
    avail = set(ledger.available_months(root))
    ytd = [m for m in fy if m <= month]
    missing = [m for m in ytd if m not in avail]
    if missing:
        raise ReportError("actuals missing for %s" % ", ".join(missing))
    b = budget.build(co)
    bmonths = {d["month"]: dict(d, pnl=b["pnl"][d["month"]]) for d in b["drivers"]}
    act = {}
    for m in ytd:
        p, k = ledger.read_month(m, root)
        act[m] = {"pnl": p, "kpis": k}
    rep = co.cfg["reporting"]
    notes = load_notes() if notes is None else notes

    def agg(src, months, code):
        return sum(src[m]["pnl"][code] for m in months)

    lines = {"month": {}, "ytd": {}}
    for code, acc in accounts.BY_CODE.items():
        lines["month"][code] = _line(acc.name, acc.kind, act[month]["pnl"][code], bmonths[month]["pnl"][code],
                                     rep["month_abs"], rep["month_pct"])
        lines["ytd"][code] = _line(acc.name, acc.kind, agg(act, ytd, code), agg(bmonths, ytd, code),
                                   rep["ytd_abs"], rep["ytd_pct"])
    totals = {}
    for period, months in (("month", [month]), ("ytd", ytd)):
        a_t = accounts.totals({c: agg(act, months, c) for c in accounts.BY_CODE})
        b_t = accounts.totals({c: agg(bmonths, months, c) for c in accounts.BY_CODE})
        a_g = accounts.by_group({c: agg(act, months, c) for c in accounts.BY_CODE})
        b_g = accounts.by_group({c: agg(bmonths, months, c) for c in accounts.BY_CODE})
        abs_t, pct_t = (rep["month_abs"], rep["month_pct"]) if period == "month" else (rep["ytd_abs"], rep["ytd_pct"])
        t = {}
        for g in accounts.GROUPS:
            kind = "revenue" if g in ("Management fees", "New business fees", "Other fee income") else "expense"
            t[g] = _line(g, kind, a_g[g], b_g[g], abs_t, pct_t)
        for key, label, kind in (("revenue", "Total revenue", "revenue"), ("opex", "Total operating costs", "expense"),
                                 ("ebitda", "EBITDA", "revenue"), ("net_income", "Net income", "revenue")):
            t[key] = _line(label, kind, a_t[key], b_t[key], abs_t, pct_t)
        t["margin"] = {"actual": a_t["ebitda"] / a_t["revenue"] if a_t["revenue"] else 0.0,
                       "budget": b_t["ebitda"] / b_t["revenue"] if b_t["revenue"] else 0.0}
        totals[period] = t

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
    res = {"month": month, "ytd_months": ytd, "fy": co.fy_label(), "company": co.name,
           "lines": lines, "totals": totals, "bridges": bridges, "kpis": kpis, "ytd_kpis": ytd_kpis,
           "actuals": act, "budget": bmonths}
    res["commentary"] = commentary(res, notes)
    return res


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
        facts = _facts(res, period)
        note_months = [res["month"]] if period == "month" else res["ytd_months"]
        flagged = sorted((c for c, l in res["lines"][period].items() if l["flag"] and c != "9000"),
                         key=lambda c: -abs(res["lines"][period][c]["var"]))
        done_bridges = set()
        for code in flagged:
            l = res["lines"][period][code]
            bridge = next((n for n, codes in BRIDGE_ACCOUNTS.items() if code in codes), None)
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
            codes = BRIDGE_ACCOUNTS.get(bridge, [code])
            reasons = []
            for m in note_months:
                for c in codes:
                    n = notes.get(m, {}).get(c)
                    if n and n not in reasons:
                        reasons.append(n)
            if reasons:
                text += ". " + " ".join(reasons)
            items.append((label, text.rstrip(".") + ".", (res["bridges"][period][bridge]["var"] if bridge == "Staff costs" else l["var"]) >= 0))
        out[period] = items
    return out


def _month_name(m):
    return ["January", "February", "March", "April", "May", "June", "July", "August", "September",
            "October", "November", "December"][int(m[5:]) - 1] + " " + m[:4]
