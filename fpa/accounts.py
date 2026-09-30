"""
Chart of accounts (P&L). Revenue accounts are credits, expense accounts debits.
Groups roll the accounts up into the management P&L.
"""
from dataclasses import dataclass


@dataclass(frozen=True)
class Account:
    code: str
    name: str
    group: str          # management P&L line
    kind: str           # "revenue" or "expense"


ACCOUNTS = [
    Account("4000", "Management fees", "Management fees", "revenue"),
    Account("4010", "Let-only fees", "New business fees", "revenue"),
    Account("4020", "Tenancy set-up fees", "New business fees", "revenue"),
    Account("4030", "Renewal fees", "Other fee income", "revenue"),
    Account("4040", "Compliance income", "Other fee income", "revenue"),
    Account("4050", "Maintenance commission", "Other fee income", "revenue"),
    Account("5000", "Salaries", "Staff costs", "expense"),
    Account("5010", "Employer NI", "Staff costs", "expense"),
    Account("5020", "Pension", "Staff costs", "expense"),
    Account("5030", "Commission and bonus", "Staff costs", "expense"),
    Account("6000", "Office rent and rates", "Premises", "expense"),
    Account("6010", "Utilities", "Premises", "expense"),
    Account("6100", "Property portals", "Marketing", "expense"),
    Account("6110", "Marketing", "Marketing", "expense"),
    Account("6200", "Software and IT", "Overheads", "expense"),
    Account("6300", "Insurance", "Overheads", "expense"),
    Account("6400", "Professional fees", "Overheads", "expense"),
    Account("6500", "Motor and travel", "Overheads", "expense"),
    Account("6600", "Bad debts", "Overheads", "expense"),
    Account("6900", "Other overheads", "Overheads", "expense"),
    Account("7000", "Depreciation", "Depreciation", "expense"),
    Account("9000", "Corporation tax", "Tax", "expense"),
]
BY_CODE = {a.code: a for a in ACCOUNTS}
REVENUE = [a.code for a in ACCOUNTS if a.kind == "revenue"]
OPEX = [a.code for a in ACCOUNTS if a.kind == "expense" and a.group not in ("Depreciation", "Tax")]
GROUPS = ["Management fees", "New business fees", "Other fee income",
          "Staff costs", "Premises", "Marketing", "Overheads"]


def totals(pnl):
    """Management P&L subtotals from {code: amount} (all amounts positive)."""
    rev = sum(pnl.get(c, 0.0) for c in REVENUE)
    opex = sum(pnl.get(c, 0.0) for c in OPEX)
    ebitda = rev - opex
    ebit = ebitda - pnl.get("7000", 0.0)
    return {"revenue": rev, "opex": opex, "ebitda": ebitda, "ebit": ebit,
            "tax": pnl.get("9000", 0.0), "net_income": ebit - pnl.get("9000", 0.0)}


def by_group(pnl):
    out = {g: 0.0 for g in GROUPS}
    for code, v in pnl.items():
        g = BY_CODE[code].group
        if g in out:
            out[g] += v
    return out
