"""Web panel API: state, runs (pipeline stubbed), uploads with validation, downloads, quit."""
import json
import shutil
import threading
import time
import urllib.error
import urllib.request

import pytest

from fpa import company, ledger, simulate
from fpa.web import server

CO = company.load()
RUN = "2026-09_20260930-120000"


@pytest.fixture(scope="module")
def base(tmp_path_factory):
    root = tmp_path_factory.mktemp("web")
    companies, out = root / "companies", root / "output"
    shutil.copytree(company.COMPANIES, companies)
    actuals = companies / "kestrel-row" / "actuals"
    run_dir = out / "kestrel-row" / RUN
    run_dir.mkdir(parents=True)
    summary = {"month": "2026-09", "generated": "2026-09-30T12:00:00",
               "totals": {"ytd": {"ebitda": {"actual": 114555.0, "var": -19748.0}}},
               "full_year": {"base": {"ebitda": 194752.0}},
               "audit": {"mismatches": [], "excel_errors": [], "model_checks": "OK", "values_checked": 409}}
    (run_dir / "summary.json").write_text(json.dumps(summary))
    (run_dir / "Kestrel_Row_Lettings_FPA_2026-09.xlsx").write_bytes(b"PK-fake")
    server.state.update(output=str(out), companies=str(companies))

    def fake_run(month, out_root=None, progress=None, slug=None, **kw):
        for s in ("actuals", "workbook", "recalc", "audit"):
            progress(s)
        if month == "2026-04":
            raise ValueError("no actuals for 2026-04")
        return {"dir": str(run_dir), "summary": dict(summary, month=month)}

    import fpa.pipeline
    orig = fpa.pipeline.run
    fpa.pipeline.run = fake_run
    threading.Thread(target=server.worker, daemon=True).start()
    httpd = server.ThreadingHTTPServer(("127.0.0.1", 0), server.Handler)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    yield "http://127.0.0.1:%d" % httpd.server_address[1], actuals
    httpd.shutdown()
    fpa.pipeline.run = orig
    server.state.update(output=None, companies=None)


def call(url, body=None):
    req = urllib.request.Request(url, data=None if body is None else json.dumps(body).encode(),
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=10) as r:
            raw = r.read()
            return r.status, (json.loads(raw) if r.headers.get_content_type() == "application/json" else raw)
    except urllib.error.HTTPError as e:
        raw = e.read()
        try:
            return e.code, json.loads(raw)
        except ValueError:
            return e.code, raw


def wait(url, job):
    for _ in range(50):
        _, j = call(url + "/api/runs/" + job)
        if j["status"] in ("done", "error"):
            return j
        time.sleep(0.1)
    raise AssertionError("job did not finish")


def october():
    data = simulate.simulate(CO, last="2026-10")
    import tempfile
    from pathlib import Path
    with tempfile.TemporaryDirectory() as tmp:
        ledger.write_month(CO, "2026-10", *data["2026-10"], root=tmp)
        return ((Path(tmp) / "2026-10" / "trial_balance.csv").read_text(),
                (Path(tmp) / "2026-10" / "kpis.csv").read_text())


def test_index_and_state(base):
    url, _ = base
    code, html = call(url + "/")
    assert code == 200 and b"Lunoviq FP&amp;A" in html and b"/app.js?v=" in html
    code, s = call(url + "/api/state")
    assert s["slug"] == "kestrel-row" and s["closed"][-1] == "2026-09" and s["next_month"] == "2026-10"
    assert s["runs"][0]["run"] == RUN and s["runs"][0]["healthy"]
    assert "kestrel-row" in [c["slug"] for c in s["companies"]]
    assert call(url + "/api/state?company=../etc")[0] == 404
    assert call(url + "/api/state?company=nope")[0] == 404


def test_run_job_reports_steps(base):
    url, _ = base
    code, res = call(url + "/api/runs", {"month": "2026-09"})
    assert code == 202
    j = wait(url, res["job"])
    assert j["status"] == "done" and j["step"] == "audit" and j["run"] == RUN


def test_run_errors_are_friendly(base):
    url, _ = base
    assert call(url + "/api/runs", {"month": "2026-12"})[0] == 400
    code, res = call(url + "/api/runs", {"month": "2026-04"})
    j = wait(url, res["job"])
    assert j["status"] == "error" and j["error"].startswith("No actuals") and j["retryable"] is False


def test_upload_validation(base):
    url, _ = base
    tb, kp = october()
    cases = [({"month": "2026-13", "trial_balance": tb, "kpis": kp}, 400, "Choose the month"),
             ({"month": "2027-06", "trial_balance": tb, "kpis": kp}, 400, "outside"),
             ({"month": "2026-10", "trial_balance": tb, "kpis": ""}, 400, "Both files"),
             ({"month": "2026-10", "trial_balance": tb.replace("4000,", "4999,"), "kpis": kp}, 400, "unknown account"),
             ({"month": "2026-10", "trial_balance": tb.replace("2026-10", "2026-11"), "kpis": kp}, 400, "expected 2026-10"),
             ({"month": "2026-11", "trial_balance": tb, "kpis": kp}, 400, "earlier months"),
             ({"month": "2026-10", "trial_balance": tb, "kpis": "period,metric,value\n2026-10,pum_open,1\n"}, 400, "missing")]
    for body, want, msg in cases:
        code, res = call(url + "/api/upload", body)
        assert code == want and msg in res["error"], (body.get("month"), res)


def test_upload_then_replace(base):
    url, actuals = base
    tb, kp = october()
    code, res = call(url + "/api/upload", {"month": "2026-10", "trial_balance": tb, "kpis": kp})
    assert code == 200 and res["revenue"] > 90000
    assert (actuals / "2026-10" / "trial_balance.csv").exists()
    assert call(url + "/api/state")[1]["next_month"] == "2026-11"
    code, res = call(url + "/api/upload", {"month": "2026-10", "trial_balance": tb, "kpis": kp})
    assert code == 409 and res["exists"]
    assert call(url + "/api/upload", {"month": "2026-10", "trial_balance": tb, "kpis": kp, "replace": True})[0] == 200


def test_sample_download_and_files(base):
    url, _ = base
    code, raw = call(url + "/api/sample?company=kestrel-row&month=2026-10&file=trial_balance")
    assert code == 200 and raw.startswith(b"period,account")
    assert call(url + "/api/sample?month=2030-01&file=kpis")[0] == 404
    code, raw = call(url + "/api/download?company=kestrel-row&run=" + RUN)
    assert code == 200 and raw == b"PK-fake"
    assert call(url + "/api/download?company=kestrel-row&run=../../etc")[0] == 404
    assert call(url + "/api/download?company=..&run=" + RUN)[0] == 404
    assert call(url + "/api/summary?company=kestrel-row&run=nope")[0] == 404
    assert call(url + "/../fpa/company.py")[0] == 404


def test_quit_refuses_while_busy(base):
    url, _ = base
    server.JOBS["busy"] = {"id": "busy", "month": "2026-09", "status": "running"}
    code, res = call(url + "/api/quit", {})
    assert code == 409 and res["busy"]
    del server.JOBS["busy"]


# ---------------------------------------------------------------- any company
def test_rules_company_state_and_sample(base):
    url, _ = base
    code, s = call(url + "/api/state?company=brightwell-cleaning")
    assert code == 200 and s["model"] == "rules" and s["upload"] == "export" and s["next_month"] == "2026-10"
    code, raw = call(url + "/api/sample?company=brightwell-cleaning&month=2026-10&file=trial_balance")
    assert code == 200 and raw.startswith(b"Trial Balance")


def test_export_upload_for_a_rules_company(base):
    url, _ = base
    _, tb = call(url + "/api/sample?company=brightwell-cleaning&month=2026-10&file=trial_balance")
    tb = tb.decode()
    bad = tb.replace("200,Sales - Commercial contracts", "299,Sales - Unknown line")
    code, res = call(url + "/api/upload", {"company": "brightwell-cleaning", "month": "2026-10", "file": bad})
    assert code == 400 and "299 Sales - Unknown line" in res["error"]
    code, res = call(url + "/api/upload", {"company": "brightwell-cleaning", "month": "2026-10", "file": tb})
    assert code == 200 and res["info"]["basis"] == "ytd" and res["info"]["balanced"] is True
    from fpa import demo_brightwell
    true = demo_brightwell._true_pnl(company.load("brightwell-cleaning"), last="2026-10")["2026-10"]
    assert res["revenue"] == pytest.approx(sum(true[c] for c in ("200", "201", "202", "203")), abs=0.01)
    assert call(url + "/api/upload", {"company": "brightwell-cleaning", "month": "2026-10", "file": tb})[0] == 409


def test_add_company_wizard(base):
    url, _ = base
    hist = (company.COMPANIES / "brightwell-cleaning" / "imports" / "pnl_by_month_FY2025-26.csv").read_text()
    code, pv = call(url + "/api/companies/preview", {"text": hist})
    assert code == 200 and pv["layout"] == "pnl_by_month" and len(pv["months"]) == 12
    sections = {a["code"]: a["section"] for a in pv["accounts"]}
    assert sections["200"] == "revenue" and sections["310"] == "cost_of_sales" and sections["416"] == "depreciation"
    body = {"name": "Harbour Test Cleaning Ltd", "fy_start_month": 4, "revenue_growth": 0.04, "cost_growth": 0.03,
            "opening_balance": 20000, "accounts": pv["accounts"], "history": hist}
    code, res = call(url + "/api/companies", body)
    assert code == 201 and res["slug"] == "harbour-test-cleaning"
    code, s = call(url + "/api/state?company=harbour-test-cleaning")
    assert s["closed"] == [] and s["next_month"] == "2026-04" and s["fy"] == "FY2026/27"
    assert call(url + "/api/companies", body)[0] == 400                          # already exists
    assert call(url + "/api/companies/preview", {"text": "no,table,here"})[0] == 400


def test_ai_settings_and_summary(base, tmp_path, monkeypatch):
    url, _ = base
    from fpa import ai
    monkeypatch.setattr(ai, "SETTINGS", tmp_path / "local.toml")
    for v in ("OPENAI_API_KEY", "ANTHROPIC_API_KEY"):
        monkeypatch.delenv(v, raising=False)
    assert call(url + "/api/ai")[1]["configured"] is False
    assert call(url + "/api/ai/settings", {"provider": "openai", "api_key": "bad"})[0] == 400
    key = "sk-proj-" + "b" * 40
    assert call(url + "/api/ai/settings", {"provider": "openai", "api_key": key})[0] == 200
    code, st = call(url + "/api/ai")
    assert st["configured"] and st["provider"] == "openai" and key not in json.dumps(st)
    # the fake pack summary has no figures to draft from -> a clear error, nothing saved
    code, res = call(url + "/api/ai/summary", {"company": "kestrel-row", "run": RUN})
    assert code in (400, 500) or "error" in res
    assert call(url + "/api/ai/summary", {"company": "kestrel-row", "run": "nope"})[0] == 404
    assert call(url + "/api/ai/settings", {"clear": True})[0] == 200
    assert call(url + "/api/ai")[1]["configured"] is False
