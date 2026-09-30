"""
AI-drafted executive summary of a management pack (optional).

The model gets the pack's key figures (totals, variances, commentary, forecast,
cash) and writes a short summary for the board. Every number in its answer is
then checked against those figures: if any amount or percentage cannot be found
in the pack, the draft is rejected. The text is always labelled as an AI draft.

Settings (never committed; local.toml is git-ignored):

    local.toml   [ai] provider = "openai" | "anthropic", api_key = "...", model = "..."
    or the environment variables OPENAI_API_KEY / ANTHROPIC_API_KEY

Nothing is sent unless the user asks for a summary. Standard library only.
"""
import json
import os
import re
import tomllib
import urllib.error
import urllib.request

from .company import ROOT

SETTINGS = ROOT / "local.toml"
DEFAULT_MODEL = {"openai": "gpt-4o-mini", "anthropic": "claude-sonnet-5"}


class AIError(RuntimeError):
    pass


def settings():
    """{"provider", "api_key", "model"} or None when no key is configured."""
    cfg = {}
    if SETTINGS.exists():
        with open(SETTINGS, "rb") as f:
            cfg = tomllib.load(f).get("ai", {})
    provider = cfg.get("provider") or ("anthropic" if os.environ.get("ANTHROPIC_API_KEY") and
                                       not os.environ.get("OPENAI_API_KEY") else "openai")
    key = cfg.get("api_key") or os.environ.get("%s_API_KEY" % provider.upper())
    if not key:
        return None
    return {"provider": provider, "api_key": key, "model": cfg.get("model") or DEFAULT_MODEL[provider]}


def save_settings(provider, api_key, model=""):
    if provider not in DEFAULT_MODEL:
        raise AIError("unknown provider %s" % provider)
    key = (api_key or "").strip()
    if not re.fullmatch(r"[A-Za-z0-9_\-]{20,300}", key):
        raise AIError("that does not look like an API key")
    q = lambda s: '"%s"' % s.replace("\\", "\\\\").replace('"', '\\"')
    text = "# Local settings: never committed (see .gitignore).\n[ai]\nprovider = %s\napi_key = %s\n" % (q(provider), q(key))
    if model.strip():
        text += "model = %s\n" % q(model.strip())
    SETTINGS.write_text(text)
    os.chmod(SETTINGS, 0o600)


def clear_settings():
    if SETTINGS.exists():
        SETTINGS.unlink()


def facts(s):
    """The figures the model may use, from summary.json, rounded as they would be written."""
    t = s["totals"]
    f = {"company": s["company"], "month": s["month"], "financial_year": s["fy"], "forecast": s["forecast_label"]}
    for p in ("month", "ytd"):
        f[p] = {k: {"actual": round(v["actual"]), "budget": round(v["budget"]), "variance": round(v["var"]),
                    "variance_pct": round(v["pct"] * 100, 1)}
                for k, v in t[p].items() if isinstance(v, dict) and "var" in v}
        f[p]["ebitda_margin_pct"] = {"actual": round(t[p]["margin"]["actual"] * 100, 1),
                                     "budget": round(t[p]["margin"]["budget"] * 100, 1)}
        if "gp_margin" in t[p]:
            f[p]["gross_margin_pct"] = {"actual": round(t[p]["gp_margin"]["actual"] * 100, 1),
                                        "budget": round(t[p]["gp_margin"]["budget"] * 100, 1)}
        f[p]["commentary"] = [c["label"] + ": " + c["text"] for c in s["commentary"][p]]
    fy = s["full_year"]
    f["full_year_ebitda"] = {n: round(fy[n]["ebitda"]) for n in fy if isinstance(fy[n], dict) and "ebitda" in fy[n]}
    f["full_year_revenue"] = {n: round(fy[n]["revenue"]) for n in fy if isinstance(fy[n], dict) and "revenue" in fy[n]}
    c = s["cash"]
    f["cash"] = {"opening": round(c["opening"]), "lowest": round(c["lowest"]), "lowest_week": c["lowest_week"],
                 "minimum_buffer": round(c["minimum"]),
                 "corporation_tax": round(c["corporation_tax"]["amount"]), "corporation_tax_due": c["corporation_tax"]["due"],
                 "balance_after_corporation_tax": round(c["corporation_tax"]["balance_after"]),
                 "dividends": c.get("dividends", [])}
    return f


PROMPT = """You are the head of FP&A writing the executive summary at the top of a monthly management pack
for the board of {company}. Use ONLY the figures in the JSON below; do not calculate new numbers
except simple differences already given as "variance". Write 5 or 6 short sentences in British English:
1) the headline for the year to date, 2) the two or three biggest drivers and their business reasons,
3) the full-year outlook against budget, 4) cash and anything the board must decide.
Write money as £ with one decimal in thousands (e.g. £19.7k) and percentages with one decimal.
No bullet points, no headings, no markdown.

JSON:
{facts}"""


def _numbers(obj):
    out = []
    if isinstance(obj, dict):
        for v in obj.values():
            out += _numbers(v)
    elif isinstance(obj, list):
        for v in obj:
            out += _numbers(v)
    elif isinstance(obj, bool):
        pass
    elif isinstance(obj, (int, float)):
        out.append(float(obj))
    elif isinstance(obj, str):
        for m in re.finditer(r"(−|-)?£?(\d[\d,]*\.?\d*)\s*(k|m)?%?", obj):
            v = float(m.group(2).replace(",", "")) if m.group(2).replace(",", "").replace(".", "").isdigit() else None
            if v is None:
                continue
            mult = {"k": 1000, "m": 1e6}.get(m.group(3) or "", 1)
            out.append(v * mult)
    return out


def verify(text, f):
    """Every amount and percentage in `text` must be one of the pack's figures (as written, rounded)."""
    known = [abs(x) for x in _numbers(f)]
    bad = []
    for m in re.finditer(r"(£\s?\d[\d,]*(?:\.\d+)?\s?(?:k|m|bn)?|\d[\d,]*(?:\.\d+)?\s?%)", text):
        tok = m.group(1)
        raw = tok.replace("£", "").replace(",", "").replace(" ", "")
        if raw.endswith("%"):
            v, tol = float(raw[:-1]), 0.051
        elif raw.endswith("k"):
            v, tol = float(raw[:-1]) * 1000, 51.0 if "." in raw else 501.0
        elif raw.endswith("m"):
            v, tol = float(raw[:-1]) * 1e6, 5100.0
        else:
            v, tol = float(raw), 1.01
        if not any(abs(v - k) <= tol for k in known):
            bad.append(tok)
    if bad:
        raise AIError("the draft used figures that are not in the pack: %s" % ", ".join(sorted(set(bad))))
    return True


def _post(url, headers, body, timeout=60):
    req = urllib.request.Request(url, data=json.dumps(body).encode(), headers={"Content-Type": "application/json",
                                                                               **headers})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read())
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", "replace")[:300]
        if e.code == 401:
            raise AIError("the API key was rejected (401). Check it in Settings.") from None
        if e.code == 429:
            raise AIError("the API account has no credit or hit a rate limit (429).") from None
        raise AIError("the AI service returned %d: %s" % (e.code, detail)) from None
    except (urllib.error.URLError, TimeoutError) as e:
        raise AIError("the AI service could not be reached: %s" % e) from None


def call(cfg, prompt):
    if cfg["provider"] == "openai":
        r = _post("https://api.openai.com/v1/chat/completions", {"Authorization": "Bearer " + cfg["api_key"]},
                  {"model": cfg["model"], "temperature": 0.2, "messages": [{"role": "user", "content": prompt}]})
        return r["choices"][0]["message"]["content"].strip()
    r = _post("https://api.anthropic.com/v1/messages", {"x-api-key": cfg["api_key"], "anthropic-version": "2023-06-01"},
              {"model": cfg["model"], "max_tokens": 600, "messages": [{"role": "user", "content": prompt}]})
    return "".join(b.get("text", "") for b in r["content"]).strip()


def draft(summary, cfg=None, attempts=2):
    """-> {"text", "provider", "model", "checked": True}; retries once if the check fails."""
    cfg = cfg or settings()
    if not cfg:
        raise AIError("no API key is set. Add one in Settings.")
    f = facts(summary)
    prompt = PROMPT.format(company=summary["company"], facts=json.dumps(f, indent=1))
    last = None
    for _ in range(attempts):
        text = call(cfg, prompt)
        try:
            verify(text, f)
            return {"text": text, "provider": cfg["provider"], "model": cfg["model"], "checked": True}
        except AIError as e:
            last = e
            prompt += "\n\nYour previous draft was rejected because %s. Use only the figures in the JSON." % e
    raise last
