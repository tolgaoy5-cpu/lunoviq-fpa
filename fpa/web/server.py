"""
Local web panel: `python -m fpa serve` -> http://127.0.0.1:8766

Standard library only; binds to localhost. Packs are built one at a time on a
worker thread (Excel recalculates one workbook at a time); the browser polls.

API (company = folder name under companies/, default kestrel-row)
    GET  /api/state?company=<c>       companies, closed months, packs of the company
    POST /api/runs {company, month}   build the pack for a month -> {job}
    GET  /api/runs/<job>              status, current step, summary when done
    GET  /api/summary?company=<c>&run=<dir>    summary of a previous pack
    GET  /api/download?company=<c>&run=<dir>   the Excel pack
    POST /api/open {company, run}     open the pack in Excel (macOS)
    POST /api/upload {company, month, trial_balance, kpis, replace}   add a month of actuals (CSV text);
                                      for rule-based companies: {company, month, file, basis, replace},
                                      any accounting-system export (fpa/importer.py)
    POST /api/companies/preview {text}                  suggested chart from a P&L-by-month export
    POST /api/companies {name, accounts, history, ...}  create a company (fpa/onboard.py)
    GET  /api/ai                      AI summary settings (provider, model; never the key)
    POST /api/ai/settings {provider, api_key, model} | {clear: true}
    POST /api/ai/summary {company, run}   draft the executive summary of a pack (checked figures)
    GET  /api/sample?month=<m>&file=trial_balance|kpis        synthetic export for a demo upload
    POST /api/quit                    stop the server
"""
import json
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import traceback
import uuid
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from .. import company, ledger

STATIC = Path(__file__).resolve().parent / "static"
MONTH_RE = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
RUN_RE = re.compile(r"^\d{4}-\d{2}_\d{8}-\d{6}$")
TYPES = {".html": "text/html; charset=utf-8", ".css": "text/css; charset=utf-8",
         ".js": "text/javascript; charset=utf-8", ".svg": "image/svg+xml", ".png": "image/png",
         ".csv": "text/csv; charset=utf-8"}
MAX_UPLOAD = 2_000_000

JOBS = {}
QUEUE = queue.Queue()
SLUG_RE = re.compile(r"^[a-z0-9][a-z0-9-]{0,40}$")
state = {"output": None, "companies": None}        # overridable by tests


def output_root(slug):
    from ..pipeline import OUTPUT
    return Path(state["output"] or OUTPUT) / slug


def get_company(slug):
    slug = slug or company.DEFAULT
    if not SLUG_RE.match(slug) or slug not in company.available(state["companies"]):
        raise LookupError("Unknown company.")
    return company.load(slug, state["companies"])


def friendly_error(exc):
    """-> (message for the user, technical detail, retry makes sense?)"""
    msg = str(exc)
    from ..recalc import RecalcTimeout
    if isinstance(exc, RecalcTimeout):
        return ("Excel did not respond. A dialog may be open in Excel (for example document recovery "
                "or a file warning): close it, then try again.", msg, True)
    if isinstance(exc, ledger.LedgerError):
        return "The actuals for this month could not be read: %s" % msg, msg, False
    if isinstance(exc, ValueError):
        return msg[:1].upper() + msg[1:], msg, False
    return "An unexpected error occurred while building the pack.", msg, True


def worker():
    from .. import pipeline
    while True:
        job_id = QUEUE.get()
        job = JOBS[job_id]
        job["status"] = "running"
        try:
            from ..pipeline import OUTPUT
            res = pipeline.run(job["month"], out_root=state["output"] or OUTPUT, slug=job["company"],
                               companies_root=state["companies"], progress=lambda s: job.__setitem__("step", s))
            job["run"] = Path(res["dir"]).name
            job["summary"] = res["summary"]
            job["status"] = "done"
        except Exception as e:                              # noqa: BLE001 - reported to the UI
            job["error"], job["detail"], job["retryable"] = friendly_error(e)
            job["status"] = "error"                          # last: pollers must see the message with it
            traceback.print_exc()
        finally:
            QUEUE.task_done()


def run_path(slug, name):
    if not name or not RUN_RE.match(name) or not slug or not SLUG_RE.match(slug):
        return None
    p = output_root(slug) / name
    return p if p.is_dir() and p.resolve().parent == output_root(slug).resolve() else None


def runs(slug):
    root = output_root(slug)
    out = []
    if root.exists():
        for p in sorted(root.iterdir(), reverse=True):
            if RUN_RE.match(p.name) and (p / "summary.json").exists():
                s = json.loads((p / "summary.json").read_text())
                au = s.get("audit") or {}
                out.append({"run": p.name, "month": s["month"], "generated": s["generated"],
                            "ebitda_ytd": s["totals"]["ytd"]["ebitda"]["actual"],
                            "ebitda_ytd_var": s["totals"]["ytd"]["ebitda"]["var"],
                            "fy_ebitda": s["full_year"]["base"]["ebitda"],
                            "healthy": bool(au) and not au.get("mismatches") and not au.get("excel_errors")
                            and au.get("model_checks") == "OK"})
    return out


def app_state(slug=None):
    co = get_company(slug)
    fy = co.fy_months()
    have = set(ledger.available_months(co))
    closed = [m for m in fy if m in have]
    nxt = company.add_months(closed[-1], 1) if closed else fy[0]
    companies = []
    for s in company.available(state["companies"]):
        c = company.load(s, state["companies"])
        companies.append({"slug": s, "name": c.name, "model": c.model, "demo": bool(c.cfg["company"].get("demo")),
                          "description": c.cfg["company"].get("description", "")})
    return {"slug": co.slug, "company": co.name, "model": co.model, "demo": bool(co.cfg["company"].get("demo")),
            "description": co.cfg["company"].get("description", ""), "fy": co.fy_label(), "fy_months": fy,
            "closed": closed, "next_month": nxt if nxt in fy else None, "runs": runs(co.slug),
            "companies": companies, "sample": co.model == "lettings" or co.slug == "brightwell-cleaning",
            "upload": "native" if co.model == "lettings" else "export"}


def upload(body):
    month = str(body.get("month", ""))
    if not MONTH_RE.match(month):
        return {"error": "Choose the month the files belong to."}, 400
    try:
        co = get_company(body.get("company"))
    except LookupError as e:
        return {"error": str(e)}, 404
    if month not in co.fy_months():
        return {"error": "%s is outside %s." % (month, co.fy_label())}, 400
    have = ledger.available_months(co)
    if co.model != "lettings":
        return upload_export(co, month, body, have)
    tb, kp = body.get("trial_balance") or "", body.get("kpis") or ""
    if not tb.strip() or not kp.strip():
        return {"error": "Both files are needed: the trial balance and the KPIs."}, 400
    if len(tb) + len(kp) > MAX_UPLOAD:
        return {"error": "The files are too large."}, 413
    if month in have and not body.get("replace"):
        return {"error": "Actuals for %s already exist. Replace them?" % month, "exists": True}, 409
    earlier = [m for m in co.fy_months() if m < month and m not in have]
    if earlier:
        return {"error": "Load the earlier months first: %s." % ", ".join(earlier)}, 400
    with tempfile.TemporaryDirectory() as tmp:
        d = Path(tmp) / month
        d.mkdir()
        (d / "trial_balance.csv").write_text(tb, encoding="utf-8")
        (d / "kpis.csv").write_text(kp, encoding="utf-8")
        try:
            p, k = ledger.read_month(co, month, tmp)
        except (ledger.LedgerError, KeyError, ValueError) as e:
            return {"error": "The files were not accepted: %s" % e}, 400
        if co.model == "lettings":
            need = {"pum_open", "pum_close", "gained", "lost", "avg_rent", "let_only_lets", "relets",
                    "renewals", "headcount", "arrears_pct"}
            missing = sorted(need - set(k))
            if missing:
                return {"error": "The KPI file is missing: %s." % ", ".join(missing)}, 400
        dest = co.actuals_dir / month
        if dest.exists():
            shutil.rmtree(dest)
        shutil.copytree(d, dest)
    t = co.chart.totals(p)
    return {"ok": True, "month": month, "revenue": t["revenue"], "ebitda": t["ebitda"]}, 200


def upload_export(co, month, body, have):
    """A month from an accounting-system export (trial balance or P&L), checked before it is stored."""
    from .. import importer
    text = body.get("file") or ""
    if not text.strip():
        return {"error": "Choose the export file."}, 400
    if len(text) > MAX_UPLOAD:
        return {"error": "The file is too large."}, 413
    if month in have and not body.get("replace"):
        return {"error": "Actuals for %s already exist. Replace them?" % month, "exists": True}, 409
    earlier = [m for m in co.fy_months() if m < month and m not in have]
    if earlier:
        return {"error": "Load the earlier months first: %s." % ", ".join(earlier)}, 400
    basis = body.get("basis") or "auto"
    if basis not in ("auto", "month", "ytd"):
        return {"error": "Unknown basis."}, 400
    try:
        pnl, info = importer.import_month(co, month, text, basis=basis, save=False)
    except (importer.ImportError_, ledger.LedgerError, ValueError) as e:
        return {"error": "The file was not accepted: %s" % e}, 400
    ledger.write_month(co, month, pnl)
    t = co.chart.totals(pnl)
    return {"ok": True, "month": month, "revenue": t["revenue"], "ebitda": t["ebitda"], "info": info}, 200


def create_company(body):
    from .. import onboard
    name = str(body.get("name", "")).strip()
    if not 2 <= len(name) <= 80:
        return {"error": "Enter the company name."}, 400
    try:
        fy = int(body.get("fy_start_month") or 4)
        rg, cg = float(body.get("revenue_growth", 0.05)), float(body.get("cost_growth", 0.03))
        ob, mb = float(body.get("opening_balance") or 0), float(body.get("minimum_balance") or 0)
    except (TypeError, ValueError):
        return {"error": "Check the numbers in the form."}, 400
    if not 1 <= fy <= 12 or not -0.5 <= rg <= 1 or not -0.5 <= cg <= 1:
        return {"error": "Check the financial year and the growth rates."}, 400
    accounts = body.get("accounts") or []
    if not accounts:
        return {"error": "Confirm the accounts first."}, 400
    try:
        co = onboard.create(name, accounts, body.get("history") or "", fy_start_month=fy, revenue_growth=rg,
                            cost_growth=cg, opening_balance=ob, minimum_balance=mb,
                            description=str(body.get("description", ""))[:200], root=state["companies"] and Path(state["companies"]))
    except (ValueError, KeyError) as e:
        return {"error": str(e)}, 400
    return {"ok": True, "slug": co.slug}, 201


class Handler(BaseHTTPRequestHandler):
    server_version = "LunoviqFPA"

    def log_message(self, fmt, *args):
        if "/api/runs/" not in str(args[0] if args else ""):
            sys.stderr.write("  %s\n" % (fmt % args))

    def _json(self, obj, code=200):
        body = json.dumps(obj, default=str).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _body(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n > MAX_UPLOAD * 2:
            return None
        try:
            return json.loads(self.rfile.read(n) or b"{}")
        except ValueError:
            return {}

    def do_GET(self):
        try:
            return self._get()
        except Exception:                                   # noqa: BLE001 - always answer the browser
            traceback.print_exc()
            return self._json({"error": "Something went wrong on the server; see the log."}, 500)

    def do_POST(self):
        try:
            return self._post()
        except Exception:                                   # noqa: BLE001 - always answer the browser
            traceback.print_exc()
            return self._json({"error": "Something went wrong on the server; see the log."}, 500)

    def _get(self):
        u = urlparse(self.path)
        q = {k: v[0] for k, v in parse_qs(u.query).items()}
        if u.path == "/api/ai":
            from .. import ai
            cfg = ai.settings()
            return self._json({"configured": bool(cfg), "provider": cfg and cfg["provider"], "model": cfg and cfg["model"],
                               "source": "local.toml" if ai.SETTINGS.exists() else ("environment" if cfg else None)})
        if u.path == "/api/state":
            try:
                return self._json(app_state(q.get("company")))
            except LookupError as e:
                return self._json({"error": str(e)}, 404)
        if u.path.startswith("/api/runs/"):
            job = JOBS.get(u.path.rsplit("/", 1)[-1])
            if not job:
                return self._json({"error": "Job not found."}, 404)
            return self._json(job)
        if u.path == "/api/summary":
            p = run_path(q.get("company"), q.get("run"))
            if not p or not (p / "summary.json").exists():
                return self._json({"error": "Run not found."}, 404)
            extra = {"run": p.name, "slug": p.parent.name}
            if (p / "ai_summary.json").exists():
                extra["ai_summary"] = json.loads((p / "ai_summary.json").read_text())
            return self._json(json.loads((p / "summary.json").read_text()) | extra)
        if u.path == "/api/sample":
            month, which = q.get("month", ""), q.get("file", "")
            try:
                co = get_company(q.get("company"))
            except LookupError:
                return self._json({"error": "Unknown sample."}, 404)
            if month not in co.fy_months() or which not in ("trial_balance", "kpis"):
                return self._json({"error": "Unknown sample."}, 404)
            if co.model == "lettings":
                from ..simulate import simulate
                with tempfile.TemporaryDirectory() as tmp:
                    data = simulate(co, last=month)
                    ledger.write_month(co, month, *data[month], root=tmp)
                    body = (Path(tmp) / month / ("%s.csv" % which)).read_bytes()
            elif co.slug == "brightwell-cleaning" and which == "trial_balance":
                from .. import demo_brightwell
                body = demo_brightwell.sample_trial_balance(co, month).encode()
            else:
                return self._json({"error": "Unknown sample."}, 404)
            self.send_response(200)
            self.send_header("Content-Type", "text/csv; charset=utf-8")
            self.send_header("Content-Disposition", 'attachment; filename="%s_%s.csv"' % (month, which))
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return None
        if u.path == "/api/download":
            p = run_path(q.get("company"), q.get("run"))
            files = list(p.glob("*_FPA_*.xlsx")) if p else []
            if not files:
                return self._json({"error": "File not found."}, 404)
            data = files[0].read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="%s"' % files[0].name)
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return None
        name = "index.html" if u.path in ("/", "") else u.path.lstrip("/")
        f = (STATIC / name).resolve()
        if STATIC not in f.parents or not f.is_file():
            self.send_error(404)
            return None
        data = f.read_bytes()
        if f.name == "index.html":
            for asset in ("app.css", "app.js"):
                ver = int((STATIC / asset).stat().st_mtime)
                data = data.replace(b'"/%s"' % asset.encode(), b'"/%s?v=%d"' % (asset.encode(), ver))
        self.send_response(200)
        self.send_header("Content-Type", TYPES.get(f.suffix, "application/octet-stream"))
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()
        self.wfile.write(data)
        return None

    def _post(self):
        u = urlparse(self.path)
        body = self._body()
        if body is None:
            return self._json({"error": "The request is too large."}, 413)
        if u.path == "/api/runs":
            month = str(body.get("month", ""))
            try:
                st = app_state(body.get("company"))
            except LookupError as e:
                return self._json({"error": str(e)}, 404)
            if month not in st["closed"]:
                return self._json({"error": "There are no actuals for that month yet."}, 400)
            job_id = uuid.uuid4().hex[:12]
            JOBS[job_id] = {"id": job_id, "company": st["slug"], "month": month, "status": "queued", "step": None,
                            "error": None}
            QUEUE.put(job_id)
            return self._json({"job": job_id}, 202)
        if u.path == "/api/upload":
            res, code = upload(body)
            return self._json(res, code)
        if u.path == "/api/companies/preview":
            from .. import onboard, importer
            text = body.get("text") or ""
            if not text.strip() or len(text) > MAX_UPLOAD:
                return self._json({"error": "Choose the export file."}, 400)
            try:
                return self._json(onboard.preview(text))
            except importer.ImportError_ as e:
                return self._json({"error": "The file could not be read: %s" % e}, 400)
        if u.path == "/api/companies":
            res, code = create_company(body)
            return self._json(res, code)
        if u.path == "/api/ai/settings":
            from .. import ai
            if body.get("clear"):
                ai.clear_settings()
                return self._json({"ok": True})
            try:
                ai.save_settings(str(body.get("provider", "openai")), str(body.get("api_key", "")), str(body.get("model", "")))
            except ai.AIError as e:
                return self._json({"error": str(e)[:1].upper() + str(e)[1:] + "."}, 400)
            return self._json({"ok": True})
        if u.path == "/api/ai/summary":
            from .. import ai
            p = run_path(body.get("company"), body.get("run"))
            if not p or not (p / "summary.json").exists():
                return self._json({"error": "Pack not found."}, 404)
            try:
                res = ai.draft(json.loads((p / "summary.json").read_text()))
            except ai.AIError as e:
                return self._json({"error": "No summary: %s." % str(e).rstrip(".")}, 400)
            except (KeyError, TypeError, ValueError):
                return self._json({"error": "No summary: this pack has no figures to summarise; rebuild it."}, 400)
            import datetime as _dt
            res["created"] = _dt.datetime.now().isoformat(timespec="seconds")
            (p / "ai_summary.json").write_text(json.dumps(res, indent=1))
            return self._json(res)
        if u.path == "/api/open":
            p = run_path(body.get("company"), body.get("run"))
            files = list(p.glob("*_FPA_*.xlsx")) if p else []
            if not files:
                return self._json({"error": "File not found."}, 404)
            if sys.platform == "darwin":
                subprocess.run(["open", str(files[0])], check=False)
                return self._json({"ok": True})
            return self._json({"error": "Opening in Excel is supported on macOS only; download the file instead."}, 501)
        if u.path == "/api/quit":
            busy = any(j["status"] in ("queued", "running") for j in JOBS.values())
            if busy and not body.get("force"):
                return self._json({"error": "A pack is still being built. Quit anyway?", "busy": True}, 409)
            self._json({"ok": True})
            threading.Thread(target=self.server.shutdown, daemon=True).start()
            return None
        return self._json({"error": "Unknown request."}, 404)


def serve(port=8766, open_browser=True):
    threading.Thread(target=worker, daemon=True).start()
    httpd = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = "http://127.0.0.1:%d" % port
    print("Lunoviq FP&A is running at %s   (Ctrl+C to stop)" % url)
    if open_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()
