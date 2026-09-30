"""
13-week cash flow forecast for the office bank account.

The monthly P&L (actuals for closed months, the rolling forecast after that) is
turned into cash by timing rules (config [cash]); every flow is spread over
days and then summed into weeks:

  receipts   management, renewal, compliance and maintenance fees are deducted from
             rent as it is collected (mostly in the first week of the month);
             let-only and set-up fees are invoiced to landlords (14-day terms);
             VAT at 20% is collected on all fees; bad debts are never received
  payments   planned dividends (config [cash] dividends); net pay on the 28th; PAYE/NI and pension on the 22nd of the next month;
             suppliers on the 15th of the next month (with VAT where charged);
             office rent quarterly in advance on the English quarter days (+ VAT);
             business rates monthly; the VAT return one month and 7 days after the quarter

Client money (landlords' rent held in trust) is excluded.
"""
import datetime as dt

from . import company, forecast

# Cash timing classes (accounts.csv "cash" column; blank = the section default)
RECEIPT_CLASSES = ("rent_collection", "invoice", "even")
DEFAULT_CLASS = {"revenue": "invoice", "cost_of_sales": "supplier_vat", "opex": "supplier_vat",
                 "depreciation": "non_cash", "interest": "bank", "tax": "corporation_tax"}
PAYMENT_LINES = ["Net pay", "PAYE and NI", "Pension", "Suppliers", "Office rent", "Business rates", "Bank interest",
                 "VAT", "Dividend"]


def cash_class(a):
    return a.cash or DEFAULT_CLASS[a.section]


def _d(month, day):
    y, m = int(month[:4]), int(month[5:])
    last = (dt.date(y + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1)).day
    return dt.date(y, m, min(day, last))


def _days(month):
    return [_d(month, i) for i in range(1, _d(month, 31).day + 1)]


def _spread(amount, days, weights=None):
    weights = weights or [1.0] * len(days)
    tot = sum(weights)
    return [(d, amount * w / tot) for d, w in zip(days, weights)]


def _collection_weights(month, profile):
    out = []
    for d in _days(month):
        band = min(3, (d.day - 1) // 7)
        n_in_band = sum(1 for x in _days(month) if min(3, (x.day - 1) // 7) == band)
        out.append(profile[band] / n_in_band)
    return out


def build(co, last_closed, fc=None, root=None):
    c = co.cfg["cash"]
    fc = fc or forecast.reforecast(co, last_closed, root=root)
    start = _d(company.add_months(last_closed, 1), 1)
    weeks = [(start + dt.timedelta(days=7 * i), start + dt.timedelta(days=7 * i + 6)) for i in range(13)]
    end = weeks[-1][1]
    vat = c["vat_rate"]
    pnl = fc["pnl"]
    months = [m for m in fc["months"] if m in pnl]
    flows = []                                      # (date, section, line, signed amount)

    def add(items, section, line, sign):
        flows.extend((d, section, line, sign * a) for d, a in items)

    chart = co.chart
    cls = {a.code: cash_class(a) for a in chart.accounts}
    of = lambda *classes: [k for k, v in cls.items() if v in classes]
    revenue = chart.codes("revenue")
    receipt_lines = [line for line, _ in chart.lines("revenue")]
    vat_on = c.get("vat_registered", True)
    for m in months:
        p = pnl[m]
        fees = sum(p[k] for k in revenue)
        collect = 1 - (sum(p[k] for k in of("bad_debt")) / fees if fees else 0.0)   # bad debts are never received
        prof = _collection_weights(m, c["rent_collection_profile"]) if "rent_collection_profile" in c else None
        gross_up = (1 + vat) if vat_on else 1.0
        for line in receipt_lines:
            codes = [k for k in revenue if chart.by_code[k].line == line]
            for klass in RECEIPT_CLASSES:
                amt = sum(p[k] for k in codes if cls[k] == klass) * gross_up * collect
                if not amt:
                    continue
                if klass == "rent_collection":
                    add(_spread(amt, _days(m), prof), "Receipts", line, 1)
                elif klass == "invoice":
                    lag = dt.timedelta(days=c["invoice_terms_days"])
                    add([(d + lag, a) for d, a in _spread(amt, _days(m))], "Receipts", line, 1)
                else:
                    add(_spread(amt, _days(m)), "Receipts", line, 1)

        gross = sum(p[k] for k in of("payroll_gross"))
        nxt = company.add_months(m, 1)
        if gross or of("employer_ni", "pension"):
            add([(_d(m, c["payroll_day"]), gross * (1 - c["employee_deductions"] - c["employee_pension"]))],
                "Payments", "Net pay", -1)
            add([(_d(nxt, c["paye_day"]), gross * c["employee_deductions"] + sum(p[k] for k in of("employer_ni")))],
                "Payments", "PAYE and NI", -1)
            add([(_d(nxt, c["paye_day"]), gross * c["employee_pension"] + sum(p[k] for k in of("pension")))],
                "Payments", "Pension", -1)
        sup = sum(p[k] for k in of("supplier_vat")) * gross_up + sum(p[k] for k in of("supplier_novat"))
        add([(_d(nxt, c["supplier_day"]), sup)], "Payments", "Suppliers", -1)
        rent_rates = sum(p[k] for k in of("office_rent_rates"))
        rent = rent_rates * c.get("office_rent_share", 1.0)
        if rent_rates:
            add([(_d(m, 1), rent_rates - rent)], "Payments", "Business rates", -1)
            for qd in c.get("rent_quarter_days", []):
                if int(qd[:2]) == int(m[5:]):
                    add([(_d(m, int(qd[3:])), rent * 3 * gross_up)], "Payments", "Office rent", -1)
        bank = sum(p[k] for k in of("bank"))
        if bank:
            add([(_d(nxt, 1), bank)], "Payments", "Bank interest", -1)
        if vat_on and int(m[5:]) in c["vat_quarter_end_months"]:
            q = [company.add_months(m, -i) for i in (2, 1, 0)]
            if all(x in pnl for x in q):
                out_vat = sum(sum(pnl[x][k] for k in revenue) for x in q) * vat
                in_vat = sum(sum(pnl[x][k] for k in of("supplier_vat")) + sum(pnl[x][k] for k in of("office_rent_rates"))
                             * c.get("office_rent_share", 1.0) for x in q) * vat
                due = _d(m, 31) + dt.timedelta(days=c["vat_payment_lag_days"])
                add([(due, out_vat - in_vat)], "Payments", "VAT", -1)

    for day, amount in c.get("dividends", []):
        flows.append((dt.date.fromisoformat(day), "Payments", "Dividend", -amount))

    table = {("Receipts", l): [0.0] * 13 for l in receipt_lines}
    table.update({("Payments", l): [0.0] * 13 for l in PAYMENT_LINES})
    for d, sec, line, a in flows:
        for i, (w0, w1) in enumerate(weeks):
            if w0 <= d <= w1:
                table[(sec, line)][i] += a
    payment_lines = [l for l in PAYMENT_LINES if any(abs(v) > 0.005 for v in table[("Payments", l)])]
    for l in PAYMENT_LINES:
        if l not in payment_lines:
            del table[("Payments", l)]
    receipts = [sum(table[("Receipts", l)][i] for l in receipt_lines) for i in range(13)]
    payments = [sum(table[("Payments", l)][i] for l in payment_lines) for i in range(13)]
    net = [r + p for r, p in zip(receipts, payments)]
    bal, closing = c["opening_balance"], []
    for n in net:
        bal += n
        closing.append(bal)
    ct_due = dt.date.fromisoformat(c["corporation_tax_due"])
    prior_ct = c.get("corporation_tax_fy_prior") or _prior_year_tax(co, root)
    low = min(closing)
    return {
        "weeks": weeks, "table": table, "receipt_lines": receipt_lines, "payment_lines": payment_lines,
        "receipts": receipts, "payments": payments, "net": net,
        "opening": c["opening_balance"], "closing": closing, "minimum": c["minimum_balance"],
        "lowest": low, "lowest_week": closing.index(low) + 1,
        "below_minimum": [i + 1 for i, v in enumerate(closing) if v < c["minimum_balance"]],
        "corporation_tax": {"due": ct_due, "amount": prior_ct, "in_window": ct_due <= end,
                            "balance_after": closing[-1] - prior_ct},
        "forecast_label": fc["label"],
    }


def _prior_year_tax(co, root=None):
    from . import ledger
    months = co.fy_months(co.budget_year - 1)
    have = set(ledger.available_months(co, root))
    tax = co.chart.codes("tax")
    return sum(sum(ledger.read_month(co, m, root)[0][k] for k in tax) for m in months if m in have)
