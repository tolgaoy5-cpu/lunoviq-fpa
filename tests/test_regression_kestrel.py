"""Regression guard: the Kestrel Row results must not change while the engine is generalised.
The snapshot was taken on 2026-09-30 from the committed single-company version."""
import json
from pathlib import Path

import pytest

from fpa import budget, cash, company, forecast, variance

SNAP = json.loads((Path(__file__).parent / "kestrel_snapshot.json").read_text())


def close(a, b):
    if isinstance(a, dict):
        assert set(a) == set(b), (sorted(a), sorted(b))
        for k in a:
            close(a[k], b[k])
    elif isinstance(a, list):
        assert len(a) == len(b)
        for x, y in zip(a, b):
            close(x, y)
    elif isinstance(a, float) or isinstance(b, float):
        assert float(a) == pytest.approx(float(b), abs=0.011)
    else:
        assert a == b


@pytest.fixture(scope="module")
def co():
    return company.load("kestrel-row") if "kestrel-row" in getattr(company, "available", lambda: [])() else company.load()


def test_budget_unchanged(co):
    close(SNAP["budget_pnl"], budget.build(co)["pnl"])


def test_bva_unchanged(co):
    res = variance.analyse(co, "2026-09")
    close(SNAP["bva_lines"], {p: {c: {k: l[k] for k in ("actual", "budget", "var", "flag")}
                                  for c, l in res["lines"][p].items()} for p in ("month", "ytd")})
    close(SNAP["bva_totals"], {p: {k: v["var"] for k, v in res["totals"][p].items() if "var" in v}
                               for p in ("month", "ytd")})
    close(SNAP["bridges"], res["bridges"])
    close(SNAP["commentary"], json.loads(json.dumps(res["commentary"])))


def test_forecast_and_cash_unchanged(co):
    sc = forecast.scenarios(co, "2026-09")
    close(SNAP["forecast"], {n: {"fy": r["fy"], "pnl": r["pnl"]} for n, r in sc.items()})
    cf = cash.build(co, "2026-09", fc=sc["base"])
    close(SNAP["cash"]["closing"], cf["closing"])
    close(SNAP["cash"]["table"], {"|".join(k): v for k, v in cf["table"].items()})
    close(SNAP["cash"]["ct"], cf["corporation_tax"]["amount"])
