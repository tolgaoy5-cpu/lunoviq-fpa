"""
Monthly actuals in the stored (native) format, per company:

    companies/<slug>/actuals/<YYYY-MM>/trial_balance.csv   period,account,description,debit,credit
    companies/<slug>/actuals/<YYYY-MM>/kpis.csv            period,metric,value

The trial balance holds the month's movements on every P&L account of the
company's chart. Revenue accounts carry credits, expense accounts debits;
read_month() returns positive amounts per account and rejects anything it
cannot trust. Exports from accounting systems are converted to this format by
fpa/importer.py before they are stored.
"""
import csv
from pathlib import Path

TB_HEADER = ["period", "account", "description", "debit", "credit"]
KPI_HEADER = ["period", "metric", "value"]


class LedgerError(ValueError):
    pass


def _num(s, where):
    s = (s or "").strip().replace(",", "")
    if s == "":
        return 0.0
    try:
        return float(s)
    except ValueError:
        raise LedgerError("%s: not a number: %r" % (where, s)) from None


def read_trial_balance(path, month, chart):
    path = Path(path)
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    if not rows or list(rows[0].keys())[:5] != TB_HEADER:
        raise LedgerError("%s: expected columns %s" % (path.name, ",".join(TB_HEADER)))
    out = {}
    for i, r in enumerate(rows, start=2):
        where = "%s line %d" % (path.name, i)
        if r["period"].strip() != month:
            raise LedgerError("%s: period %s, expected %s" % (where, r["period"], month))
        code = r["account"].strip()
        if code not in chart.by_code:
            raise LedgerError("%s: unknown account %s" % (where, code))
        if code in out:
            raise LedgerError("%s: account %s appears twice" % (where, code))
        dr, cr = _num(r["debit"], where), _num(r["credit"], where)
        out[code] = round(cr - dr if chart.by_code[code].kind == "revenue" else dr - cr, 2)
    missing = sorted(set(chart.by_code) - set(out))
    if missing:
        raise LedgerError("%s: missing accounts %s" % (path.name, ", ".join(missing)))
    return out


def read_kpis(path, month):
    path = Path(path)
    if not path.exists():
        return {}
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    out = {}
    for r in rows:
        if r["period"].strip() != month:
            raise LedgerError("%s: period %s, expected %s" % (path.name, r["period"], month))
        out[r["metric"].strip()] = _num(r["value"], path.name)
    return out


def read_month(co, month, root=None):
    d = Path(root or co.actuals_dir) / month
    return read_trial_balance(d / "trial_balance.csv", month, co.chart), read_kpis(d / "kpis.csv", month)


def available_months(co, root=None):
    d = Path(root or co.actuals_dir)
    return sorted(p.name for p in d.iterdir() if (p / "trial_balance.csv").exists()) if d.exists() else []


def write_month(co, month, pnl, kpis=None, root=None):
    d = Path(root or co.actuals_dir) / month
    d.mkdir(parents=True, exist_ok=True)
    with open(d / "trial_balance.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(TB_HEADER)
        for a in co.chart.accounts:
            amount = pnl.get(a.code, 0.0)
            credit_side = (a.kind == "revenue") == (amount >= 0)
            v = "%.2f" % abs(amount)
            w.writerow([month, a.code, a.name, "" if credit_side else v, v if credit_side else ""])
    kp = d / "kpis.csv"
    if kpis:
        with open(kp, "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(KPI_HEADER)
            for k, v in kpis.items():
                w.writerow([month, k, v])
    elif kp.exists():
        kp.unlink()
