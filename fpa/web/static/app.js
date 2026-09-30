"use strict";
/* Lunoviq FP&A — local web panel. Vanilla JS; data from the local API. */

const $ = (s, el = document) => el.querySelector(s);
const isNum = (v) => typeof v === "number" && isFinite(v);
const nf = (d = 0) => new Intl.NumberFormat("en-GB", { minimumFractionDigits: d, maximumFractionDigits: d });
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const gbp = (v) => (isNum(v) ? (v < 0 ? "−£" : "£") + nf(0).format(Math.abs(v)) : "—");
const k = (v, d = 1) => (isNum(v) ? (v < 0 ? "−£" : "£") + (Math.abs(v) >= 1e6 ? nf(2).format(Math.abs(v) / 1e6) + "m" : nf(d).format(Math.abs(v) / 1000) + "k") : "—");
const sk = (v) => (isNum(v) ? (v >= 0 ? "+" : "") + k(v) : "—");
const pct = (v, d = 1) => (isNum(v) ? nf(d).format(v * 100) + "%" : "—");
const spct = (v, d = 1) => (isNum(v) ? (v >= 0 ? "+" : "−") + nf(d).format(Math.abs(v) * 100) + "%" : "—");
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const mlabel = (m) => MONTHS[+m.slice(5) - 1] + "-" + m.slice(2, 4);
const mlong = (m) => ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"][+m.slice(5) - 1] + " " + m.slice(0, 4);
const STEPS = [["actuals", "Reading the ledger exports"], ["workbook", "Building the budget, variance, forecast and cash sheets"],
  ["recalc", "Recalculating in Excel"], ["audit", "Checking every figure against the Python engine"]];

let S = null;              // app state
let SLUG = new URLSearchParams(location.search).get("company")
  || (() => { try { return localStorage.getItem("fpa.company"); } catch { return null; } })();   // current company
const cq = () => "company=" + encodeURIComponent(SLUG || "");
let lastJob = null, pollTimer = null, current = null;

async function api(path, opts = {}) {
  const r = await fetch(path, { headers: { "Content-Type": "application/json" }, ...opts });
  let body = null;
  try { body = await r.json(); } catch { body = {}; }
  if (!r.ok) { const e = new Error(body.error || "Request failed (" + r.status + ")"); e.status = r.status; e.body = body; throw e; }
  return body;
}

function show(id) {
  document.querySelectorAll(".view").forEach((v) => (v.hidden = v.id !== id));
  $("#app").focus({ preventScroll: true });
}

function toast(msg) {
  const t = $("#toast");
  t.textContent = msg; t.hidden = false;
  clearTimeout(t._h); t._h = setTimeout(() => (t.hidden = true), 3200);
}

async function loadState() {
  try { S = await api("/api/state" + (SLUG ? "?" + cq() : "")); }
  catch (e) { if (e.status !== 404) throw e; S = await api("/api/state"); }        // stored company was removed
  SLUG = S.slug;
  try { localStorage.setItem("fpa.company", SLUG); } catch { /* private mode */ }
  const sel = $("#co-sel");
  sel.innerHTML = S.companies.map((c) => `<option value="${esc(c.slug)}"${c.slug === SLUG ? " selected" : ""}>${esc(c.name)}${c.demo ? " (demo)" : ""}</option>`).join("");
  sel.onchange = () => { SLUG = sel.value; current = null; try { localStorage.setItem("fpa.company", SLUG); } catch {} location.hash = "#/"; route(); };
  return S;
}

/* ---------------------------------------------------------------- router */
async function route() {
  clearTimeout(pollTimer);
  const [, kind, a, b] = (location.hash || "#/").split("/");
  try {
    await loadState();
    if (kind === "run" && a) return showRun(a, b || "overview");
    if (kind === "job" && a) return poll(a);
    if (kind === "upload") return renderUpload();
    if (kind === "add") return renderAdd();
    return renderHome();
  } catch (e) {
    toast(e.message);
  }
}
window.addEventListener("hashchange", route);

/* ---------------------------------------------------------------- home */
function latestRun(month) {
  return S.runs.find((r) => r.month === month);
}

function renderHome() {
  show("view-home");
  document.title = "Lunoviq FP&A";
  const last = S.closed[S.closed.length - 1];
  const cells = S.fy_months.map((m) => {
    const closed = S.closed.includes(m), run = latestRun(m);
    if (closed) {
      return `<div class="mcell closed">
        <div class="mname">${mlabel(m)}</div>
        ${run ? `<div class="mstat"><span class="pill ${run.healthy ? "ok" : "warn"}">${run.healthy ? "Pack ready" : "Check"}</span></div>
          <div class="mfig">YTD EBITDA <b>${k(run.ebitda_ytd)}</b><span class="${run.ebitda_ytd_var >= 0 ? "up" : "down"}"> ${sk(run.ebitda_ytd_var)}</span></div>
          <div class="mact"><a class="btn btn-sm" href="#/run/${run.run}">Open</a><button class="btn btn-sm linkish" data-build="${m}">Rebuild</button></div>`
        : `<div class="mstat"><span class="pill neutral">Actuals loaded</span></div><div class="mfig muted">No pack yet</div>
          <div class="mact"><button class="btn btn-sm btn-primary" data-build="${m}">Build pack</button></div>`}
      </div>`;
    }
    if (m === S.next_month) {
      return `<div class="mcell next"><div class="mname">${mlabel(m)}</div><div class="mstat"><span class="pill warn">Open</span></div>
        <div class="mfig muted">Waiting for the ledger export</div><div class="mact"><a class="btn btn-sm" href="#/upload">Load actuals</a></div></div>`;
    }
    return `<div class="mcell future"><div class="mname">${mlabel(m)}</div><div class="mfig muted">Forecast</div></div>`;
  }).join("");
  const rows = S.runs.map((r) => `<tr class="link" data-run="${r.run}"><td>${mlabel(r.month)}</td>
    <td class="num">${k(r.ebitda_ytd)}</td><td class="num ${r.ebitda_ytd_var >= 0 ? "up" : "down"}">${sk(r.ebitda_ytd_var)}</td>
    <td class="num">${k(r.fy_ebitda)}</td><td class="muted small">${esc(new Date(r.generated).toLocaleString("en-GB", { dateStyle: "medium", timeStyle: "short" }))}</td>
    <td>${r.healthy ? '<span class="pill ok">Audit OK</span>' : '<span class="pill warn">Not audited</span>'}</td></tr>`).join("");
  $("#view-home").innerHTML = `
    <div class="hero">
      <p class="eyebrow">Budget · actuals · forecast · cash</p>
      <h1>Monthly management pack</h1>
      <p class="lede">${esc(S.company)}${S.demo ? " (fictional demo company)" : ""}${S.description ? ": " + esc(S.description) : ""}.
        ${S.model === "lettings" ? "Driver-based budget (lettings template)." : "Budget built by rules from last year's actuals."}
        Pick a closed month to build its pack: budget vs actual with commentary, the rolling forecast and scenarios,
        and a 13-week cash flow. The workbook is recalculated in Excel and every figure is checked.</p>
      ${last ? `<div class="actions" style="justify-content:flex-start;margin-top:18px">
        ${latestRun(last) ? `<a class="btn btn-primary" href="#/run/${latestRun(last).run}">Open ${mlabel(last)} pack</a>` : `<button class="btn btn-primary" data-build="${last}">Build ${mlabel(last)} pack</button>`}
        ${S.next_month ? `<a class="btn" href="#/upload">Load ${mlabel(S.next_month)} actuals</a>` : ""}</div>` : ""}
    </div>
    <section class="recent" aria-labelledby="m-h">
      <div class="section-head"><h2 id="m-h">${esc(S.fy)}</h2><span class="muted small">${S.closed.length} of 12 months closed</span></div>
      <div class="mgrid">${cells}</div>
    </section>
    <section class="recent" aria-labelledby="r-h">
      <div class="section-head"><h2 id="r-h">Packs</h2><span class="muted small">${S.runs.length || ""}</span></div>
      ${S.runs.length ? `<div class="table-wrap"><table class="tbl"><thead><tr><th>Month</th><th class="num">YTD EBITDA</th><th class="num">vs budget</th><th class="num">FY forecast EBITDA</th><th>Built</th><th>Status</th></tr></thead><tbody>${rows}</tbody></table></div>`
        : '<p class="empty">No packs yet. Build the first one above.</p>'}
    </section>`;
  $("#view-home").querySelectorAll("[data-build]").forEach((b) => (b.onclick = () => startRun(b.dataset.build)));
  $("#view-home").querySelectorAll("tr[data-run]").forEach((tr) => (tr.onclick = () => (location.hash = "#/run/" + tr.dataset.run)));
}

/* ---------------------------------------------------------------- runs */
async function startRun(month) {
  try {
    const { job } = await api("/api/runs", { method: "POST", body: JSON.stringify({ company: SLUG, month }) });
    lastJob = month;
    location.hash = "#/job/" + job;
  } catch (e) { toast(e.message); }
}

async function poll(id) {
  let job;
  try { job = await api("/api/runs/" + id); } catch (e) { toast(e.message); return renderHome(); }
  if (job.status === "done") {
    history.replaceState(null, "", "#/run/" + job.run);
    return renderResult({ ...job.summary, run: job.run }, "overview");
  }
  if (job.status === "error") return renderError(job);
  show("view-run");
  $("#run-title").textContent = mlong(job.month);
  document.title = "Building " + mlabel(job.month) + " · Lunoviq FP&A";
  const idx = STEPS.findIndex(([s]) => s === job.step);
  $("#run-steps").innerHTML = STEPS.map(([s, t], i) => `<li class="${i < idx ? "done" : i === idx ? "active" : ""}"><span class="ic" aria-hidden="true"></span><span class="t">${t}</span></li>`).join("");
  pollTimer = setTimeout(() => poll(id), 900);
}

function renderError(job) {
  show("view-error");
  $("#err-title").textContent = mlong(job.month);
  $("#err-msg").textContent = job.error;
  $("#err-detail").textContent = job.detail || "";
  const retry = $("#err-retry");
  retry.hidden = !job.retryable;
  retry.onclick = () => startRun(job.month);
}

async function showRun(run, tab) {
  if (current && current.run === run && current.slug === SLUG) return renderResult(current, tab);
  try {
    const d = await api("/api/summary?" + cq() + "&run=" + encodeURIComponent(run));
    renderResult(d, tab);
  } catch (e) { toast(e.message); renderHome(); }
}

/* ---------------------------------------------------------------- charts */
const chartWidth = (max) => Math.max(300, Math.min(max, (window.innerWidth || max) - 72));

function niceTicks(lo, hi, n = 4) {
  const span = hi - lo || 1, step0 = span / n, mag = 10 ** Math.floor(Math.log10(step0));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((s) => span / s <= n) || 10 * mag;
  const out = [];
  for (let v = Math.floor(lo / step) * step; v <= hi + 1e-9; v += step) out.push(v);
  while (out[out.length - 1] < hi) out.push(out[out.length - 1] + step);
  return out;
}

/* bars (actual solid, forecast hatched) + budget line; values in GBP, axis in £k */
function monthChart(rows, key, bkey, title) {
  const W = chartWidth(560), H = 230, L = 44, B = 26, T = 12, n = rows.length;
  const vals = rows.flatMap((r) => [r[key], r[bkey]]).filter(isNum);
  const lo = Math.min(0, ...vals), hi = Math.max(...vals) * 1.08;
  const ticks = niceTicks(lo, hi, 4), top = ticks[ticks.length - 1], bot = ticks[0];
  const y = (v) => T + (1 - (v - bot) / (top - bot)) * (H - T - B);
  const slot = (W - L) / n, bw = Math.min(22, slot * 0.55);
  let s = `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(title)}">`;
  ticks.forEach((t) => (s += `<line class="gridl" x1="${L}" x2="${W}" y1="${y(t)}" y2="${y(t)}"/><text x="${L - 6}" y="${y(t) + 4}" text-anchor="end">${nf(0).format(t / 1000)}</text>`));
  let path = "";
  rows.forEach((r, i) => {
    const cx = L + slot * i + slot / 2, f = r.source === "F";
    const v = r[key], y0 = y(Math.max(0, v)), h = Math.abs(y(v) - y(0));
    s += `<rect class="${f ? "bar-f" : "bar-h"}" x="${cx - bw / 2}" y="${y0}" width="${bw}" height="${h}" rx="2"><title>${mlabel(r.month)}${f ? " forecast" : " actual"} ${k(v)}; budget ${k(r[bkey])}</title></rect>`;
    path += (i ? "L" : "M") + cx + "," + y(r[bkey]);
    s += `<text x="${cx}" y="${H - 8}" text-anchor="middle">${W < 520 ? MONTHS[+r.month.slice(5) - 1][0] : mlabel(r.month).slice(0, 3)}</text>`;
  });
  s += `<path class="budget-line" d="${path}"/>`;
  rows.forEach((r, i) => (s += `<circle class="budget-dot" cx="${L + slot * i + slot / 2}" cy="${y(r[bkey])}" r="2.5"/>`));
  const fi = rows.findIndex((r) => r.source === "F");
  if (fi > 0) { const dx = L + slot * fi; s += `<line class="div" x1="${dx}" x2="${dx}" y1="${T}" y2="${H - B}"/><text x="${dx + 5}" y="${T + 10}">forecast →</text>`; }
  return s + "</svg>";
}

function cashChart(c) {
  const W = chartWidth(1060), H = 240, L = 50, B = 26, T = 12, n = c.closing.length;
  const vals = c.closing.concat([c.minimum, c.opening]);
  const ticks = niceTicks(0, Math.max(...vals) * 1.08, 4), top = ticks[ticks.length - 1];
  const y = (v) => T + (1 - v / top) * (H - T - B);
  const x = (i) => L + ((W - L - 12) / (n - 1)) * i;
  let s = `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="13-week closing cash balance">`;
  ticks.forEach((t) => (s += `<line class="gridl" x1="${L}" x2="${W}" y1="${y(t)}" y2="${y(t)}"/><text x="${L - 6}" y="${y(t) + 4}" text-anchor="end">${nf(0).format(t / 1000)}</text>`));
  const pts = c.closing.map((v, i) => `${x(i)},${y(v)}`).join(" ");
  s += `<polygon class="cash-area" points="${x(0)},${y(0)} ${pts} ${x(n - 1)},${y(0)}"/>`;
  s += `<polyline class="cash-line" points="${pts}"/>`;
  s += `<line class="buffer" x1="${L}" x2="${W - 12}" y1="${y(c.minimum)}" y2="${y(c.minimum)}"/><text class="buffer-lab" x="${W - 14}" y="${y(c.minimum) - 6}" text-anchor="end">buffer ${k(c.minimum, 0)}</text>`;
  c.closing.forEach((v, i) => {
    s += `<circle class="${v < c.minimum ? "cash-dot low" : "cash-dot"}" cx="${x(i)}" cy="${y(v)}" r="3"><title>Week ${i + 1} (${c.weeks[i].slice(5)}): ${gbp(v)}</title></circle>`;
    const d = new Date(c.weeks[i]);
    s += `<text x="${x(i)}" y="${H - 8}" text-anchor="middle">${W < 600 ? i + 1 : d.getDate() + " " + MONTHS[d.getMonth()]}</text>`;
  });
  return s + "</svg>";
}

function bridgeBars(b, flex = false) {
  const parts = flex ? [["volume", "Sales volume"], ["rate", "Cost share"], ["other", "Other"]] : [["volume", "Volume"], ["rate", "Rate"], ["other", "Other"]];
  const max = Math.max(1, ...Object.values(b).flatMap((x) => parts.map(([p]) => Math.abs(x[p])).concat(Math.abs(x.var))));
  return Object.entries(b).map(([name, x]) => `<div class="bgroup"><div class="bname">${esc(name)} <span class="${x.var >= 0 ? "up" : "down"} num">${sk(x.var)}</span></div>
    ${parts.filter(([p]) => Math.abs(x[p]) >= 1).map(([p, lab]) => `<div class="brow"><span class="blab">${!flex && name === "Staff costs" && p === "volume" ? "Headcount" : lab}</span>
      <span class="btrack"><span class="bzero"></span><span class="bseg ${x[p] >= 0 ? "pos" : "neg"}" style="${x[p] >= 0 ? "left:50%" : "right:50%"};width:${(Math.abs(x[p]) / max) * 50}%"></span></span>
      <span class="num bval">${sk(x[p])}</span></div>`).join("")}</div>`).join("");
}

/* ---------------------------------------------------------------- result */
let resizeT = null;
window.addEventListener("resize", () => { clearTimeout(resizeT); resizeT = setTimeout(() => { if (current && !$("#view-result").hidden) renderResult(current, current._tab, true); }, 200); });

function renderResult(d, tab = "overview", keep = false) {
  current = d; d._tab = tab;
  show("view-result");
  document.title = mlabel(d.month) + " pack · Lunoviq FP&A";
  const au = d.audit || {};
  const healthy = au.model_checks === "OK" && !(au.mismatches || []).length && !(au.excel_errors || []).length;
  const tabs = [["overview", "Overview"], ["variance", "Variance"], ["forecast", "Forecast"], ["cash", "Cash"]];
  const body = { overview, variance, forecast, cash }[tab] || overview;
  $("#view-result").innerHTML = `
    <div class="res-head">
      <div>
        <p class="eyebrow">${esc(d.fy)} · reporting month ${mlabel(d.month)} · forecast ${esc(d.forecast_label)}</p>
        <h1>${mlong(d.month)} management pack</h1>
        <div class="meta"><span>${esc(d.company)}</span>
          ${au.values_checked ? `<span class="pill ${healthy ? "ok" : "crit"}">${healthy ? "Excel = Python · " + nf(0).format(au.values_checked) + " values" : "Checks need attention"}</span>` : '<span class="pill warn">Not recalculated</span>'}
          ${au.model_checks ? `<span class="pill ${au.model_checks === "OK" ? "ok" : "crit"}">Model checks ${esc(au.model_checks)}</span>` : ""}</div>
      </div>
      <div class="res-actions">
        <button class="btn btn-primary" id="open-xl">Open in Excel</button>
        <a class="btn" href="/api/download?${cq()}&run=${encodeURIComponent(d.run)}">Download pack</a>
      </div>
    </div>
    <nav class="tabs" role="tablist">${tabs.map(([id, lab]) => `<a role="tab" aria-selected="${id === tab}" class="tab${id === tab ? " on" : ""}" href="#/run/${d.run}/${id}">${lab}</a>`).join("")}</nav>
    <div class="tabbody">${body(d)}</div>`;
  $("#open-xl").onclick = async () => {
    try { await api("/api/open", { method: "POST", body: JSON.stringify({ company: SLUG, run: d.run }) }); toast("Opening in Excel…"); } catch (e) { toast(e.message); }
  };
  const tog = $("#vtoggle");
  if (tog) tog.querySelectorAll("button").forEach((b) => (b.onclick = () => { d._period = b.dataset.p; renderResult(d, "variance", true); }));
  if (!keep) window.scrollTo(0, 0);
}

function tile(label, value, sub, good, note) {
  return `<div class="m"><div class="lbl">${label}</div><div class="val">${value}</div><div class="vs ${good ? "up" : "down"}">${sub}</div>${note ? `<div class="muted small">${note}</div>` : ""}</div>`;
}

function overview(d) {
  const tm = d.totals.month, ty = d.totals.ytd, fy = d.full_year, c = d.cash;
  const fyVar = fy.base.ebitda - fy.budget.ebitda;
  const pum = d.kpis[0] || {};
  const heads = [d.commentary.month[0], d.commentary.ytd[0], ...d.commentary.ytd.slice(1, 4)];
  const ct = c.corporation_tax;
  return `
    <div class="tiles">
      ${tile("Revenue, " + mlabel(d.month), k(tm.revenue.actual), sk(tm.revenue.var) + " vs budget", tm.revenue.var >= 0)}
      ${tile("Revenue, year to date", k(ty.revenue.actual), sk(ty.revenue.var) + " vs budget", ty.revenue.var >= 0)}
      ${tile("EBITDA, year to date", k(ty.ebitda.actual), sk(ty.ebitda.var) + " · margin " + pct(ty.margin.actual), ty.ebitda.var >= 0)}
      ${tile("EBITDA, full-year forecast", k(fy.base.ebitda), sk(fyVar) + " vs budget", fyVar >= 0)}
      ${tile("Lowest cash, next 13 weeks", k(c.lowest), sk(c.lowest - c.minimum) + " over the buffer", c.lowest >= c.minimum)}
      ${d.model === "lettings" ? tile("Properties under management", nf(0).format(pum.actual), (pum.actual - pum.budget >= 0 ? "+" : "−") + nf(0).format(Math.abs(pum.actual - pum.budget)) + " vs budget", pum.actual >= pum.budget)
        : ty.gp_margin ? tile("Gross margin, year to date", pct(ty.gp_margin.actual), spct(ty.gp_margin.actual - ty.gp_margin.budget).replace("%", " pts") + " vs budget", ty.gp_margin.actual >= ty.gp_margin.budget)
        : tile("EBITDA margin, year to date", pct(ty.margin.actual), spct(ty.margin.actual - ty.margin.budget).replace("%", " pts") + " vs budget", ty.margin.actual >= ty.margin.budget)}
    </div>
    <div class="grid g-2">
      <section class="card"><div class="card-head"><h2>Revenue</h2><span class="muted small">£k · bars actual/forecast, line budget</span></div>${monthChart(d.monthly, "revenue", "budget_revenue", "Monthly revenue vs budget")}</section>
      <section class="card"><div class="card-head"><h2>EBITDA</h2><span class="muted small">£k · bars actual/forecast, line budget</span></div>${monthChart(d.monthly, "ebitda", "budget_ebitda", "Monthly EBITDA vs budget")}</section>
    </div>
    <section class="card" style="margin-top:18px"><div class="card-head"><h2>Headlines</h2></div>
      <ul class="review heads">${heads.map((h) => `<li class="${h.label === "Headline" ? "info" : h.favourable ? "fav" : ""}"><span>${h.label === "Headline" ? "" : `<b>${esc(h.label)}:</b> `}${esc(h.text)}</span></li>`).join("")}
        <li class="${fyVar >= 0 ? "info" : ""}"><span><b>Full year (${esc(d.forecast_label)}):</b> EBITDA ${k(fy.base.ebitda)} vs budget ${k(fy.budget.ebitda)} (${sk(fyVar)}); range ${k(fy.downside.ebitda)} – ${k(fy.upside.ebitda)} across the scenarios.</span></li>
        <li class="${ct.balance_after >= c.minimum + 10000 ? "info" : ""}"><span><b>Cash:</b> lowest ${k(c.lowest)} in week ${c.lowest_week}${ct.amount > 0 ? `; corporation tax of ${k(ct.amount)} on ${new Date(ct.due).toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" })} would leave ${k(ct.balance_after)}` : ""} against the ${k(c.minimum, 0)} buffer.</span></li>
      </ul></section>`;
}

function materiality(r, p) {
  const abs = p === "month" ? r.month_abs : r.ytd_abs, pc = p === "month" ? r.month_pct : r.ytd_pct;
  const always = p === "month" ? r.month_always : r.ytd_always;
  if (!isNum(abs)) return "material variance";
  return `at least ${k(abs, abs % 1000 ? 1 : 0)} and ${pct(pc, 0)}` + (isNum(always) ? `, or ${k(always, always % 1000 ? 1 : 0)} on its own` : "");
}

function variance(d) {
  const p = d._period || "ytd";
  const t = d.totals[p];
  const lines = d.lines[p];
  const row = (l) => `<tr class="${l.flag ? "flagged" : ""}"><td>${esc(l.label)}<span class="code">${l.code}</span></td>
    <td class="num">${gbp(l.actual)}</td><td class="num">${gbp(l.budget)}</td>
    <td class="num ${Math.abs(l.var) < 0.5 ? "" : l.var >= 0 ? "up" : "down"}">${Math.abs(l.var) < 0.5 ? "–" : (l.var >= 0 ? "+" : "−") + nf(0).format(Math.abs(l.var))}</td>
    <td class="num ${Math.abs(l.var) < 0.5 ? "" : l.var >= 0 ? "up" : "down"}">${Math.abs(l.var) < 0.5 ? "" : spct(l.pct)}</td>
    <td class="flagc">${l.flag ? '<span class="flag" title="Material variance">●</span>' : ""}</td></tr>`;
  const tot = (lab, x, cls = "") => `<tr class="tot ${cls}"><td>${lab}</td><td class="num">${gbp(x.actual)}</td><td class="num">${gbp(x.budget)}</td>
    <td class="num ${x.var >= 0 ? "up" : "down"}">${(x.var >= 0 ? "+" : "−") + nf(0).format(Math.abs(x.var))}</td><td class="num ${x.var >= 0 ? "up" : "down"}">${spct(x.pct)}</td><td></td></tr>`;
  const sec = (s) => lines.filter((l) => l.section === s).map(row).join("");
  let body = `<tr class="sec"><td colspan="6">Revenue</td></tr>` + sec("revenue") + tot("Total revenue", t.revenue);
  if (t.cost_of_sales) {
    body += `<tr class="sec"><td colspan="6">Cost of sales</td></tr>` + sec("cost_of_sales") + tot("Total cost of sales", t.cost_of_sales)
      + tot("Gross profit", t.gross_profit, "strong")
      + `<tr class="muted"><td>Gross margin</td><td class="num">${pct(t.gp_margin.actual)}</td><td class="num">${pct(t.gp_margin.budget)}</td><td class="num">${spct(t.gp_margin.actual - t.gp_margin.budget)}</td><td></td><td></td></tr>`;
  }
  body += `<tr class="sec"><td colspan="6">Operating costs</td></tr>` + sec("opex") + tot("Total operating costs", t.opex);
  body += tot("EBITDA", t.ebitda, "strong");
  body += `<tr class="muted"><td>EBITDA margin</td><td class="num">${pct(t.margin.actual)}</td><td class="num">${pct(t.margin.budget)}</td><td class="num">${spct(t.margin.actual - t.margin.budget)}</td><td></td><td></td></tr>`;
  const com = d.commentary[p];
  const kp = p === "month" ? d.kpis.filter((x) => isNum(x.budget)).map((x) => `<tr><td>${esc(x.label)}</td><td class="num">${x.kind === "gbp" ? gbp(x.actual) : nf(x.kind === "number" && x.actual % 1 ? 1 : 0).format(x.actual)}</td><td class="num">${x.kind === "gbp" ? gbp(x.budget) : nf(1).format(x.budget)}</td></tr>`).join("")
    : d.ytd_kpis.map((x) => `<tr><td>${esc(x.label)}</td><td class="num">${nf(0).format(x.actual)}</td><td class="num">${nf(1).format(x.budget)}</td></tr>`).join("");
  const hasBridges = Object.keys(d.bridges[p] || {}).length > 0;
  const flex = d.model !== "lettings";
  return `
    <div class="vbar"><div class="seg" id="vtoggle" role="group" aria-label="Period">
      <button type="button" data-p="month" aria-pressed="${p === "month"}">${mlabel(d.month)}</button><button type="button" data-p="ytd" aria-pressed="${p === "ytd"}">Year to date</button></div>
      <span class="muted small">Positive = favourable. ● = material: ${materiality(d.reporting || {}, p)}.</span></div>
    <div class="grid g-var">
      <section class="card"><div class="card-head"><h2>Budget vs actual</h2><span class="muted small">${p === "month" ? mlong(d.month) : "April – " + mlong(d.month)}</span></div>
        <div class="table-wrap"><table class="tbl vtbl"><thead><tr><th></th><th class="num">Actual</th><th class="num">Budget</th><th class="num">Var £</th><th class="num">Var %</th><th></th></tr></thead><tbody>${body}</tbody></table></div></section>
      <div>
        <section class="card"><div class="card-head"><h2>Commentary</h2></div>
          <ul class="review heads">${com.map((h) => `<li class="${h.label === "Headline" ? "info" : h.favourable ? "fav" : ""}"><span>${h.label === "Headline" ? "" : `<b>${esc(h.label)}:</b> `}${esc(h.text)}</span></li>`).join("")}</ul>
          <p class="muted small" style="margin:10px 0 0">Numbers are generated from the ledger and the budget; reasons come from the analyst notes (commentary.toml).</p></section>
        ${hasBridges ? `<section class="card" style="margin-top:18px"><div class="card-head"><h2>${flex ? "Flexed budget" : "What drove it"}</h2><span class="muted small">favourable +</span></div>
          <p class="how">${flex ? "Costs budgeted as a share of sales: sales volume = the effect of sales above or below budget; cost share = the change in the cost as a share of sales."
            : "Volume = properties, lets or headcount; rate = average rent; other = collection, fee mix, pay."}</p>${bridgeBars(d.bridges[p], flex)}</section>` : ""}
        ${kp ? `<section class="card" style="margin-top:18px"><div class="card-head"><h2>Operating KPIs</h2></div>
          <div class="table-wrap"><table class="tbl"><thead><tr><th></th><th class="num">Actual</th><th class="num">Budget</th></tr></thead><tbody>${kp}</tbody></table></div></section>` : ""}
      </div>
    </div>`;
}

function forecast(d) {
  const fy = d.full_year, names = ["budget", "base", "upside", "downside"];
  const lab = { budget: "Budget", base: "Base (latest estimate)", upside: "Upside", downside: "Downside" };
  const cards = names.map((n) => {
    const t = fy[n], v = t.ebitda - fy.budget.ebitda;
    const pum = n === "budget" ? fy.budget_closing_pum : (fy.closing_pum || {})[n];
    return `<div class="m ${n === "base" ? "base" : ""}"><div class="lbl">${lab[n]}</div><div class="val">${k(t.ebitda)}</div>
      <div class="vs ${n === "budget" ? "" : v >= 0 ? "up" : "down"}">${n === "budget" ? "EBITDA · margin " + pct(t.ebitda / t.revenue) : sk(v) + " vs budget · " + pct(t.ebitda / t.revenue)}</div>
      <div class="muted small">Revenue ${k(t.revenue)}${isNum(pum) ? " · " + nf(0).format(pum) + " properties at year end" : t.gross_profit !== t.revenue ? " · gross margin " + pct(t.gross_profit / t.revenue) : ""}</div></div>`;
  }).join("");
  const a = d.assumptions, notes = d.assumption_notes || {};
  const AL = { lets_factor: ["Let-only and re-let volumes vs budget", (v) => pct(v, 0)], rent_growth_monthly: ["Rent growth per month", (v) => pct(v, 2)],
    portals: ["Property portals", (v) => gbp(v) + " a month"], revenue_factor: ["Revenue vs budget (run-rate)", (v) => pct(v, 1)],
    cost_of_sales_factor: ["Direct costs vs budget (run-rate)", (v) => pct(v, 1)], opex_factor: ["Overheads vs budget", (v) => pct(v, 1)] };
  const al = Object.entries(a).map(([key, v]) => {
    const [lab, f] = AL[key] || (key.startsWith("adjustment_") ? ["Adjustment, account " + key.slice(11), (x) => sk(x) + " a month"] : [key.replace(/_/g, " "), (x) => nf(3).format(x)]);
    return `<dt>${esc(lab)}</dt><dd>${f(v)}</dd>`;
  }).join("");
  const noteText = [...new Set(Object.values(notes))].map(esc).join(" ");
  const rows = d.monthly.map((r) => `<tr><td>${mlabel(r.month)} <span class="pill ${r.source === "A" ? "neutral" : "fc"}">${r.source === "A" ? "Actual" : "Forecast"}</span></td>
    <td class="num">${gbp(r.revenue)}</td><td class="num">${gbp(r.budget_revenue)}</td><td class="num">${gbp(r.ebitda)}</td><td class="num">${gbp(r.budget_ebitda)}</td>
    <td class="num ${r.ebitda - r.budget_ebitda >= 0 ? "up" : "down"}">${(r.ebitda - r.budget_ebitda >= 0 ? "+" : "−") + nf(0).format(Math.abs(r.ebitda - r.budget_ebitda))}</td></tr>`).join("");
  return `
    <section class="card"><div class="card-head"><h2>Full-year outturn: ${esc(d.fy)} (${esc(d.forecast_label)})</h2></div>
      <p class="how">${d.model === "lettings" ? "Closed months are actuals; open months run on the budget drivers from the actual closing position, updated by the latest estimate. Scenarios change lets, churn, new landlords and rent growth for the open months only."
        : "Closed months are actuals; open months are the remaining budget scaled by the year-to-date run-rate of each section, plus adjustments. Scenarios change sales and costs for the open months only."}</p>
      <div class="methods four">${cards}</div></section>
    <div class="grid g-2" style="margin-top:18px">
      <section class="card"><div class="card-head"><h2>EBITDA by month</h2><span class="muted small">£k</span></div>${monthChart(d.monthly, "ebitda", "budget_ebitda", "EBITDA by month")}</section>
      <section class="card"><div class="card-head"><h2>Latest estimate</h2><span class="muted small">forecast.toml</span></div>
        <dl class="kv">${al}</dl><p class="muted small">${noteText}</p></section>
    </div>
    <section class="card" style="margin-top:18px"><div class="card-head"><h2>Monthly detail</h2></div>
      <div class="table-wrap"><table class="tbl"><thead><tr><th>Month</th><th class="num">Revenue</th><th class="num">Budget</th><th class="num">EBITDA</th><th class="num">Budget</th><th class="num">Var</th></tr></thead><tbody>${rows}</tbody></table></div></section>`;
}

function cash(d) {
  const c = d.cash, ct = c.corporation_tax;
  const lines = Object.entries(c.lines);
  const wk = c.weeks.map((w, i) => `<th class="num">${i + 1}</th>`).join("");
  const tr = (lab, arr, cls = "") => `<tr class="${cls}"><td>${esc(lab)}</td>${arr.map((v) => `<td class="num">${Math.abs(v) < 0.5 ? "–" : nf(0).format(v / 1000)}</td>`).join("")}</tr>`;
  const rec = lines.filter(([kk]) => kk.startsWith("Receipts")), pay = lines.filter(([kk]) => kk.startsWith("Payments"));
  const tight = ct.balance_after < c.minimum + 10000;
  return `
    <div class="tiles three">
      ${tile("Opening balance", k(c.opening), "office account, " + new Date(c.weeks[0]).toLocaleDateString("en-GB", { day: "numeric", month: "short" }), true)}
      ${tile("Lowest balance", k(c.lowest), "week " + c.lowest_week + " · " + sk(c.lowest - c.minimum) + " over the buffer", c.lowest >= c.minimum)}
      ${tile("After corporation tax", k(ct.balance_after), "due " + new Date(ct.due).toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" }) + " · " + k(ct.amount), !tight)}
    </div>
    ${tight ? `<div class="callout warn"><b>Little headroom after corporation tax.</b> ${(c.dividends || []).length ? "Planned dividends of " + c.dividends.map(([day, amt]) => k(amt, 0) + " (" + new Date(day).toLocaleDateString("en-GB", { day: "numeric", month: "short" }) + ")").join(", ") + " and c" : "C"}orporation tax of ${k(ct.amount)} (${new Date(ct.due).toLocaleDateString("en-GB", { day: "numeric", month: "short" })}) take the balance to ${k(ct.balance_after)}, ${ct.balance_after >= c.minimum ? "just " + k(ct.balance_after - c.minimum) + " above" : k(c.minimum - ct.balance_after) + " below"} the ${k(c.minimum, 0)} buffer.${(c.dividends || []).length ? " Phasing the dividend would restore headroom." : ""}</div>` : ""}
    <section class="card" style="margin-top:18px"><div class="card-head"><h2>Closing balance by week</h2><span class="muted small">£k, incl. VAT</span></div>${cashChart(c)}</section>
    <section class="card" style="margin-top:18px"><div class="card-head"><h2>Receipts and payments</h2><span class="muted small">£k by week (week 1 = ${new Date(c.weeks[0]).toLocaleDateString("en-GB", { day: "numeric", month: "short" })})</span></div>
      <div class="table-wrap"><table class="tbl ctbl"><thead><tr><th></th>${wk}</tr></thead><tbody>
        ${rec.map(([kk, v]) => tr(kk.split("|")[1], v)).join("")}${tr("Total receipts", c.receipts, "tot")}
        ${pay.map(([kk, v]) => tr(kk.split("|")[1], v)).join("")}${tr("Total payments", c.payments, "tot")}
        ${tr("Closing balance", c.closing, "tot strong")}
      </tbody></table></div>
      <p class="muted small" style="margin:10px 0 0">Timing: rent is mostly collected in the first week and fees are deducted from it; landlord invoices on 14-day terms; net pay on the 28th, PAYE/NI and pension on the 22nd of the next month; suppliers on the 15th of the next month; office rent quarterly in advance; VAT one month and 7 days after the quarter. Client money is excluded.</p></section>`;
}

/* ---------------------------------------------------------------- upload */
const readFile = (inp) => new Promise((res, rej) => { const f = inp.files[0]; if (!f) return res(""); const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsText(f); });
const picker = (id, label) => `<label class="drop"><span class="dl">${label}</span><span class="btn btn-sm pick">Choose file</span><input type="file" accept=".csv,text/csv,.txt" id="f-${id}"><span class="fn muted small" id="n-${id}">No file chosen</span></label>`;
const bindPicker = (id) => ($("#f-" + id).onchange = (e) => ($("#n-" + id).textContent = e.target.files[0] ? e.target.files[0].name : "No file chosen"));

function renderUpload() {
  show("view-upload");
  document.title = "Load actuals · Lunoviq FP&A";
  const m = S.next_month, native = S.upload === "native";
  const sample = S.sample && m ? (native
    ? `<p class="muted small">No export at hand? Download the synthetic ${mlabel(m)} files: <a href="/api/sample?${cq()}&month=${m}&file=trial_balance">trial balance</a> · <a href="/api/sample?${cq()}&month=${m}&file=kpis">KPIs</a>.</p>`
    : `<p class="muted small">No export at hand? Download a synthetic <a href="/api/sample?${cq()}&month=${m}&file=trial_balance">${mlabel(m)} trial balance</a> (year to date, accounting-system layout).</p>`) : "";
  $("#view-upload").innerHTML = `
    <div class="run-card upload-card">
      <p class="eyebrow">Month-end close · ${esc(S.company)}</p>
      <h1>${m ? "Load " + mlong(m) + " actuals" : "All months of " + esc(S.fy) + " are loaded"}</h1>
      ${!m ? `<div class="actions"><a class="btn" href="#/">Back to months</a></div>` : native ? `
      <p class="lede">Upload the two exports: the <b>trial balance</b> (period, account, description, debit, credit) and the <b>KPIs</b> (period, metric, value). They are checked before anything is saved.</p>
      <form class="uform" id="uform">${picker("tb", "Trial balance (CSV)")}${picker("kp", "KPIs (CSV)")}
        <div class="form-actions"><button class="btn btn-primary" id="u-go" type="submit">Check and load</button><span class="form-err" id="u-err" role="alert"></span></div></form>` : `
      <p class="lede">Upload the month's <b>trial balance</b> (or a P&amp;L with a ${mlabel(m)} column) exported from your accounting system (Xero, QuickBooks, Sage…). Title lines, codes or names, debit/credit or balance columns and year-to-date figures are all understood. Every account must be in the chart and the file must reconcile, or nothing is saved.</p>
      <form class="uform" id="uform">${picker("tb", "Export (CSV)")}
        <label class="drop"><span class="dl">Figures are</span><select id="u-basis"><option value="auto">Detect from the column names</option><option value="month">For the month</option><option value="ytd">Year to date</option></select></label>
        <div class="form-actions"><button class="btn btn-primary" id="u-go" type="submit">Check and load</button><span class="form-err" id="u-err" role="alert"></span></div></form>`}
      ${sample}
      <div id="u-ok" hidden></div>
    </div>`;
  if (!m) return;
  bindPicker("tb");
  if (native) bindPicker("kp");
  $("#uform").onsubmit = async (e) => {
    e.preventDefault();
    const err = $("#u-err"); err.textContent = "";
    const tb = await readFile($("#f-tb")), kp = native ? await readFile($("#f-kp")) : "";
    if (!tb || (native && !kp)) return (err.textContent = native ? "Choose both files." : "Choose the export file.");
    const payload = (replace) => native ? { company: SLUG, month: m, trial_balance: tb, kpis: kp, replace }
      : { company: SLUG, month: m, file: tb, basis: $("#u-basis").value, replace };
    const send = (replace) => api("/api/upload", { method: "POST", body: JSON.stringify(payload(replace)) });
    $("#u-go").disabled = true;
    try {
      let r;
      try { r = await send(false); } catch (x) {
        if (x.status === 409 && x.body.exists && confirm(x.message)) r = await send(true); else throw x;
      }
      $("#uform").hidden = true;
      const ok = $("#u-ok"); ok.hidden = false;
      const how = r.info && r.info.basis ? ` Read as a ${r.info.basis === "ytd" ? "year-to-date" : "monthly"} trial balance${r.info.balanced ? " (debits = credits)" : ""}; ${r.info.accounts} accounts matched.` : "";
      ok.innerHTML = `<div class="callout okc"><b>${mlong(r.month)} loaded.</b> Revenue ${k(r.revenue)}, EBITDA ${k(r.ebitda)}.${how}</div>
        <div class="actions" style="justify-content:flex-start"><button class="btn btn-primary" id="u-build">Build ${mlabel(r.month)} pack</button><a class="btn" href="#/">Back to months</a></div>`;
      $("#u-build").onclick = () => startRun(r.month);
      await loadState();
    } catch (x) { err.textContent = x.message; }
    finally { $("#u-go") && ($("#u-go").disabled = false); }
  };
}

/* ---------------------------------------------------------------- add a company */
const SECTION_LABEL = { revenue: "Revenue", cost_of_sales: "Cost of sales", opex: "Operating costs", depreciation: "Depreciation",
  interest: "Interest", tax: "Tax", balance_sheet: "Balance sheet (ignored)" };

function renderAdd() {
  show("view-add");
  document.title = "Add a company · Lunoviq FP&A";
  const months = MONTHS.map((m, i) => `<option value="${i + 1}"${i === 3 ? " selected" : ""}>${m}</option>`).join("");
  $("#view-add").innerHTML = `
    <div class="run-card add-card">
      <p class="eyebrow">Add a company</p>
      <h1>Set up from last year's P&amp;L</h1>
      <p class="lede">Export last year's <b>Profit and Loss by month</b> from your accounting system (one column per month) and upload it here. Every account is given a suggested section and report line: check them, then create the company. The budget grows from these actuals by simple rules you can change later.</p>
      <form class="form add-form" id="aform" novalidate>
        <div class="field"><label for="a-name">Company name</label><div class="inp"><input id="a-name" placeholder="e.g. Harbour Dental Ltd"></div></div>
        <div class="field"><label for="a-desc">What it does (optional)</label><div class="inp"><input id="a-desc" placeholder="e.g. Two dental practices"></div></div>
        <div class="field"><label for="a-fy">Financial year starts in</label><div class="inp"><select id="a-fy">${months}</select></div></div>
        <div class="field"><label for="a-rg">Sales growth for the budget</label><div class="inp"><input id="a-rg" inputmode="decimal" value="5"><span class="unit">%</span></div></div>
        <div class="field"><label for="a-cg">Overhead inflation</label><div class="inp"><input id="a-cg" inputmode="decimal" value="3"><span class="unit">%</span></div></div>
        <div class="field"><label for="a-ob">Bank balance today</label><div class="inp"><span class="unit">£</span><input id="a-ob" inputmode="decimal" placeholder="0"></div></div>
        <div class="field"><label for="a-mb">Minimum cash to keep</label><div class="inp"><span class="unit">£</span><input id="a-mb" inputmode="decimal" placeholder="0"></div></div>
      </form>
      <div class="uform">${picker("hist", "P&amp;L by month (CSV)")}</div>
      <div class="form-actions"><button class="btn btn-primary" id="a-prev">Read the file</button><span class="form-err" id="a-err" role="alert"></span></div>
      <div id="a-map"></div>
    </div>`;
  bindPicker("hist");
  let history = "", preview = null;
  $("#a-prev").onclick = async () => {
    const err = $("#a-err"); err.textContent = "";
    history = await readFile($("#f-hist"));
    if (!history) return (err.textContent = "Choose the P&L by month export.");
    try { preview = await api("/api/companies/preview", { method: "POST", body: JSON.stringify({ text: history }) }); }
    catch (x) { return (err.textContent = x.message); }
    if (preview.layout !== "pnl_by_month") return (err.textContent = "This looks like a trial balance; the set-up needs a P&L by month (one column per month).");
    const opts = (sel) => preview.sections.map((s) => `<option value="${s}"${s === sel ? " selected" : ""}>${SECTION_LABEL[s]}</option>`).join("");
    $("#a-map").innerHTML = `
      <h2 style="margin:22px 0 6px">Check the accounts</h2>
      <p class="muted small">${preview.accounts.length} accounts, ${preview.months.length} months (${mlabel(preview.months[0])} – ${mlabel(preview.months[preview.months.length - 1])}). Accounts with the same report line are added together in the reports.</p>
      <div class="table-wrap"><table class="tbl maptbl"><thead><tr><th>Code</th><th>Account</th><th class="num">Year total</th><th>Section</th><th>Report line</th></tr></thead><tbody>
        ${preview.accounts.map((a, i) => `<tr><td class="mono small">${esc(a.code)}</td><td>${esc(a.name)}</td><td class="num">${gbp(a.total)}</td>
          <td><select data-i="${i}" data-f="section">${opts(a.section)}</select></td><td><input data-i="${i}" data-f="line" value="${esc(a.line)}"></td></tr>`).join("")}
      </tbody></table></div>
      <div class="form-actions"><button class="btn btn-primary" id="a-go">Create company</button><span class="form-err" id="a-err2" role="alert"></span></div>`;
    $("#a-map").querySelectorAll("[data-f]").forEach((el) => (el.onchange = () => { preview.accounts[+el.dataset.i][el.dataset.f] = el.value; }));
    $("#a-go").onclick = async () => {
      const e2 = $("#a-err2"); e2.textContent = "";
      const num = (id, d = 0) => { const v = $(id).value.trim().replace(/[£,%\s]/g, ""); return v === "" ? d : Number(v); };
      const body = { name: $("#a-name").value.trim(), description: $("#a-desc").value.trim(), fy_start_month: +$("#a-fy").value,
        revenue_growth: num("#a-rg", 5) / 100, cost_growth: num("#a-cg", 3) / 100, opening_balance: num("#a-ob"), minimum_balance: num("#a-mb"),
        accounts: preview.accounts, history };
      if (!body.name) return (e2.textContent = "Enter the company name.");
      if ([body.revenue_growth, body.cost_growth, body.opening_balance, body.minimum_balance].some((v) => !isFinite(v))) return (e2.textContent = "Check the numbers in the form.");
      if (!preview.accounts.some((a) => a.section === "revenue")) return (e2.textContent = "At least one account must be revenue.");
      $("#a-go").disabled = true;
      try {
        const r = await api("/api/companies", { method: "POST", body: JSON.stringify(body) });
        SLUG = r.slug; current = null;
        try { localStorage.setItem("fpa.company", SLUG); } catch {}
        toast(body.name + " created");
        location.hash = "#/";
      } catch (x) { e2.textContent = x.message; $("#a-go").disabled = false; }
    };
  };
}

/* ---------------------------------------------------------------- quit */
function bindQuit() {
  const b = $("#quit");
  let armed = null;
  const reset = () => { b.textContent = "Quit"; b.classList.remove("arm"); armed = null; };
  const stop = async (force) => {
    try { await api("/api/quit", { method: "POST", body: JSON.stringify({ force }) }); }
    catch (e) { if (e.status === 409) { b.textContent = "Pack building — click to quit anyway"; b.classList.add("arm"); armed = "force"; return; } }
    clearTimeout(pollTimer);
    show("view-stopped");
    document.title = "Lunoviq FP&A (stopped)";
    b.hidden = true;
  };
  b.addEventListener("click", () => {
    if (armed === "force") return stop(true);
    if (armed) { clearTimeout(armed); return stop(false); }
    b.textContent = "Click again to quit"; b.classList.add("arm");
    armed = setTimeout(reset, 3000);
  });
}

bindQuit();
route();
