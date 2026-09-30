"""Phase 1: company calendar, chart of accounts, driver engine, ledger I/O, synthetic actuals."""
import pytest

from fpa import accounts, company, ledger, simulate
from fpa.drivers import driver_path, pnl

CO = company.load()


def test_financial_year_calendar():
    assert CO.fy_months() == ["2026-%02d" % m for m in range(4, 13)] + ["2027-01", "2027-02", "2027-03"]
    assert CO.fy_label() == "FY2026/27"
    assert CO.month_index("2026-04") == 1 and CO.month_index("2027-03") == 12
    assert company.add_months("2026-11", 3) == "2027-02"
    assert sum(CO.cfg["seasonality"]["lets"]) == pytest.approx(12.0)


def test_budget_drivers_follow_the_plan():
    path = driver_path(CO, CO.fy_months())
    assert path[0]["pum_open"] == CO.cfg["portfolio"]["opening_pum"]
    assert [d["headcount"] for d in path[:6]] == [18] * 6          # planned hire from October
    assert [d["headcount"] for d in path[6:]] == [19] * 6
    for a, b in zip(path, path[1:]):
        assert b["pum_open"] == a["pum_close"]
        assert b["avg_rent"] == pytest.approx(a["avg_rent"] * (1 + CO.cfg["portfolio"]["rent_growth_monthly"]))


def test_pnl_from_drivers():
    d = driver_path(CO, CO.fy_months())[0]
    p = pnl(d, CO)
    assert set(p) == set(accounts.BY_CODE)
    fees = CO.cfg["fees"]
    assert p["4000"] == pytest.approx(d["avg_pum"] * d["avg_rent"] * CO.cfg["portfolio"]["collection_rate"]
                                      * fees["management_fee"], abs=0.01)
    assert p["5030"] == pytest.approx(0.10 * (p["4010"] + p["4020"]), abs=0.01)
    t = accounts.totals(p)
    assert p["9000"] == pytest.approx(0.25 * t["ebit"], abs=0.01)
    assert t["net_income"] == pytest.approx(t["ebit"] * 0.75, abs=0.02)


def test_simulation_is_reproducible_and_tells_the_stories():
    a, b = simulate.simulate(CO), simulate.simulate(CO)
    assert a == b
    assert list(a) == company.month_range("2025-04", "2026-09")
    k = {m: v[1] for m, v in a.items()}
    p = {m: v[0] for m, v in a.items()}
    assert k["2026-07"]["lost"] >= 24                            # portfolio landlord left
    assert p["2026-08"]["6100"] == pytest.approx(3200 * 1.15)     # portal price rise
    assert p["2026-05"]["6600"] > 4000                            # bad-debt write-off
    assert k["2026-06"]["headcount"] == 17 and k["2026-09"]["headcount"] == 18
    assert k["2026-09"]["avg_rent"] > driver_path(CO, CO.fy_months())[5]["avg_rent"]   # rents ahead of budget
    for m, (pm, km) in a.items():
        assert km["pum_close"] == km["pum_open"] + km["gained"] - km["lost"]


def test_ledger_round_trip(tmp_path):
    data = simulate.simulate(CO)
    ledger.write_month("2026-09", *data["2026-09"], root=tmp_path)
    tb, kpis = ledger.read_month("2026-09", root=tmp_path)
    assert tb == data["2026-09"][0]
    assert kpis["pum_close"] == data["2026-09"][1]["pum_close"]
    assert ledger.available_months(tmp_path) == ["2026-09"]


@pytest.mark.parametrize("edit, message", [
    (lambda rows: rows.__setitem__(1, rows[1].replace("4000", "4999")), "unknown account"),
    (lambda rows: rows.__setitem__(1, rows[1].replace("2026-09", "2026-08")), "expected 2026-09"),
    (lambda rows: rows.append(rows[1]), "appears twice"),
    (lambda rows: rows.pop(1), "missing accounts"),
    (lambda rows: rows.__setitem__(1, rows[1].rstrip() + "x\n"), "not a number"),
])
def test_ledger_rejects_bad_exports(tmp_path, edit, message):
    ledger.write_month("2026-09", *simulate.simulate(CO)["2026-09"], root=tmp_path)
    f = tmp_path / "2026-09" / "trial_balance.csv"
    rows = f.read_text().splitlines(keepends=True)
    edit(rows)
    f.write_text("".join(rows))
    with pytest.raises(ledger.LedgerError, match=message):
        ledger.read_month("2026-09", root=tmp_path)
