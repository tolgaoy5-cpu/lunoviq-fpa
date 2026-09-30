"""
Command line:

    python -m fpa run [--company kestrel-row] [--month 2026-09] [--no-recalc]
                                                        build, recalculate and audit the monthly pack
    python -m fpa simulate                              regenerate the synthetic actuals
    python -m fpa serve [--port 8766]                   start the local web panel
"""
import argparse
import sys


def main(argv=None):
    ap = argparse.ArgumentParser(prog="fpa")
    sub = ap.add_subparsers(dest="cmd", required=True)
    r = sub.add_parser("run")
    r.add_argument("--month")
    r.add_argument("--company", help="company folder under companies/ (default: kestrel-row)")
    r.add_argument("--no-recalc", action="store_true")
    sub.add_parser("simulate")
    s = sub.add_parser("serve")
    s.add_argument("--port", type=int, default=8766)
    s.add_argument("--no-browser", action="store_true")
    a = ap.parse_args(argv)
    if a.cmd == "simulate":
        from .simulate import main as sim
        sim()
    elif a.cmd == "run":
        from .pipeline import run
        res = run(a.month, recalc=not a.no_recalc, progress=lambda s: print("..", s, flush=True), slug=a.company)
        s = res["summary"]
        t = s["totals"]["ytd"]
        print("Pack: %s" % res["workbook"])
        print("YTD revenue %.0f vs %.0f budget; EBITDA %.0f vs %.0f" % (
            t["revenue"]["actual"], t["revenue"]["budget"], t["ebitda"]["actual"], t["ebitda"]["budget"]))
        print("Full-year EBITDA (%s): %.0f vs budget %.0f" % (
            s["forecast_label"], s["full_year"]["base"]["ebitda"], s["full_year"]["budget"]["ebitda"]))
        if s["audit"]:
            au = s["audit"]
            print("Audit: %d values checked, %d mismatches, %d Excel errors; model checks %s" % (
                au["values_checked"], len(au["mismatches"]), len(au["excel_errors"]), au["model_checks"]))
            if au["mismatches"] or au["excel_errors"] or au["model_checks"] != "OK":
                return 1
    elif a.cmd == "serve":
        from .web.server import serve
        serve(a.port, open_browser=not a.no_browser)
    return 0


if __name__ == "__main__":
    sys.exit(main())
