"""Phase 5: 13-week cash flow."""
import datetime as dt

import openpyxl
import pytest

from fpa import cash, company, workbook

CO = company.load()


@pytest.fixture(scope="module")
def cf():
    return cash.build(CO, "2026-09")


def test_weeks_and_balance_roll_forward(cf):
    assert cf["weeks"][0][0] == dt.date(2026, 10, 1) and len(cf["weeks"]) == 13
    bal = cf["opening"]
    for n, c in zip(cf["net"], cf["closing"]):
        bal += n
        assert c == pytest.approx(bal)
    assert cf["net"] == pytest.approx([r + p for r, p in zip(cf["receipts"], cf["payments"])])


def test_timing_rules(cf):
    t = cf["table"]
    week = lambda d: next(i for i, (a, b) in enumerate(cf["weeks"]) if a <= d <= b)
    assert t[("Payments", "VAT")][week(dt.date(2026, 11, 7))] < 0          # Jul-Sep VAT return
    assert sum(t[("Payments", "VAT")]) == t[("Payments", "VAT")][week(dt.date(2026, 11, 7))]
    assert t[("Payments", "Office rent")][week(dt.date(2026, 12, 25))] < 0  # quarter day
    assert t[("Payments", "Dividend")][week(dt.date(2026, 12, 18))] == -120000
    mf = t[("Receipts", "Management fees")]
    assert mf[0] > mf[1] > mf[2]                                            # rent collected early in the month


def test_corporation_tax_warning(cf):
    ct = cf["corporation_tax"]
    assert ct["due"] == dt.date(2027, 1, 1) and not ct["in_window"]
    assert ct["amount"] > 0 and ct["balance_after"] == pytest.approx(cf["closing"][-1] - ct["amount"])


@pytest.mark.excel
def test_excel_cash_matches_python(tmp_path):
    from fpa.recalc import recalc
    path = tmp_path / "pack.xlsx"
    lay = workbook.build(CO, path, month="2026-09")
    recalc(path)
    ws = openpyxl.load_workbook(path, data_only=True)["Cash13W"]
    cf = lay["cash_result"]
    row = lay["cash"]["closing"]
    for i in range(13):
        assert ws.cell(row=row, column=workbook.FIRST_COL + i).value == pytest.approx(cf["closing"][i], abs=0.5)
