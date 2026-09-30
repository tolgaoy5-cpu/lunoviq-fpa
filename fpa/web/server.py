"""
Local web panel: `python -m fpa serve` -> http://127.0.0.1:8766

Standard library only; binds to localhost. Packs are built one at a time on a
worker thread (Excel recalculates one workbook at a time); the browser polls.

API
    GET  /api/state                   company, closed months, runs per month
    POST /api/runs {month}            build the pack for a month -> {job}
    GET  /api/runs/<job>              status, current step, summary when done
    GET  /api/summary?run=<dir>       summary of a previous run
    GET  /api/download?run=<dir>      the Excel pack
    POST /api/open {run}              open the pack in Excel (macOS)
    POST /api/upload {month, trial_balance, kpis, replace}   add a month of actuals (CSV text)
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
state = {"output": None, "actuals": None}          # overridable by tests


def output_root():
    from ..pipeline import OUTPUT
    return Path(state["output"] or OUTPUT)


def actuals_root():
    return Path(state["actuals"] or ledger.ACTUALS)


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
            res = pipeline.run(job["month"], out_root=output_root(), root=actuals_root(),
                               progress=lambda s: job.__setitem__("step", s))
            job["run"] = Path(res["dir"]).name
            job["summary"] = res["summary"]
            job["status"] = "done"
        except Exception as e:                              # noqa: BLE001 - reported to the UI
            job["status"] = "error"
            job["error"], job["detail"], job["retryable"] = friendly_error(e)
            traceback.print_exc()
        finally:
            QUEUE.task_done()


def run_path(name):
    if not name or not RUN_RE.match(name):
        return None
    p = output_root() / name
    return p if p.is_dir() and p.resolve().parent == output_root().resolve() else None


def runs():
    root = output_root()
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


def app_state():
    co = company.load()
    fy = co.fy_months()
    have = set(ledger.available_months(actuals_root()))
    closed = [m for m in fy if m in have]
    nxt = company.add_months(closed[-1], 1) if closed else fy[0]
    return {"company": co.name, "fy": co.fy_label(), "fy_months": fy, "closed": closed,
            "next_month": nxt if nxt in fy else None, "runs": runs()}


def upload(body):
    month = str(body.get("month", ""))
    if not MONTH_RE.match(month):
        return {"error": "Choose the month the files belong to."}, 400
    co = company.load()
    if month not in co.fy_months():
        return {"error": "%s is outside %s." % (month, co.fy_label())}, 400
    tb, kp = body.get("trial_balance") or "", body.get("kpis") or ""
    if not tb.strip() or not kp.strip():
        return {"error": "Both files are needed: the trial balance and the KPIs."}, 400
    if len(tb) + len(kp) > MAX_UPLOAD:
        return {"error": "The files are too large."}, 413
    have = ledger.available_months(actuals_root())
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
            p, k = ledger.read_month(month, tmp)
        except (ledger.LedgerError, KeyError, ValueError) as e:
            return {"error": "The files were not accepted: %s" % e}, 400
        need = {"pum_open", "pum_close", "gained", "lost", "avg_rent", "let_only_lets", "relets",
                "renewals", "headcount", "arrears_pct"}
        missing = sorted(need - set(k))
        if missing:
            return {"error": "The KPI file is missing: %s." % ", ".join(missing)}, 400
        dest = actuals_root() / month
        if dest.exists():
            shutil.rmtree(dest)
        shutil.copytree(d, dest)
    from ..accounts import totals
    t = totals(p)
    return {"ok": True, "month": month, "revenue": t["revenue"], "ebitda": t["ebitda"]}, 200


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
        u = urlparse(self.path)
        q = {k: v[0] for k, v in parse_qs(u.query).items()}
        if u.path == "/api/state":
            return self._json(app_state())
        if u.path.startswith("/api/runs/"):
            job = JOBS.get(u.path.rsplit("/", 1)[-1])
            if not job:
                return self._json({"error": "Job not found."}, 404)
            return self._json(job)
        if u.path == "/api/summary":
            p = run_path(q.get("run"))
            if not p or not (p / "summary.json").exists():
                return self._json({"error": "Run not found."}, 404)
            return self._json(json.loads((p / "summary.json").read_text()) | {"run": p.name})
        if u.path == "/api/sample":
            month, which = q.get("month", ""), q.get("file", "")
            co = company.load()
            if month not in co.fy_months() or which not in ("trial_balance", "kpis"):
                return self._json({"error": "Unknown sample."}, 404)
            from ..simulate import simulate
            with tempfile.TemporaryDirectory() as tmp:
                data = simulate(co, last=month)
                ledger.write_month(month, *data[month], root=tmp)
                body = (Path(tmp) / month / ("%s.csv" % which)).read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "text/csv; charset=utf-8")
            self.send_header("Content-Disposition", 'attachment; filename="%s_%s.csv"' % (month, which))
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return None
        if u.path == "/api/download":
            p = run_path(q.get("run"))
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

    def do_POST(self):
        u = urlparse(self.path)
        body = self._body()
        if body is None:
            return self._json({"error": "The request is too large."}, 413)
        if u.path == "/api/runs":
            month = str(body.get("month", ""))
            if month not in app_state()["closed"]:
                return self._json({"error": "There are no actuals for that month yet."}, 400)
            job_id = uuid.uuid4().hex[:12]
            JOBS[job_id] = {"id": job_id, "month": month, "status": "queued", "step": None, "error": None}
            QUEUE.put(job_id)
            return self._json({"job": job_id}, 202)
        if u.path == "/api/upload":
            res, code = upload(body)
            return self._json(res, code)
        if u.path == "/api/open":
            p = run_path(body.get("run"))
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
