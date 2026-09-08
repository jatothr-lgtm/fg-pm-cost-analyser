/* ============================================================================
   FG / PM Cost Analyser - UI layer
   Charts are hand-rolled SVG so the mark specs hold exactly:
   <=24px bars with a 4px rounded data-end square at the baseline, 2px lines,
   >=8px markers carrying a 2px surface ring, solid hairline grid.
   ========================================================================== */
(() => {
'use strict';

const $ = s => document.querySelector(s);
const el = (t, a = {}, kids = []) => {
  const n = document.createElementNS(t === 'div' || t === 'span' ? 'http://www.w3.org/1999/xhtml' : 'http://www.w3.org/2000/svg', t);
  for (const k in a) n.setAttribute(k, a[k]);
  kids.forEach(c => n.appendChild(c));
  return n;
};
const svgEl = (t, a = {}) => {
  const n = document.createElementNS('http://www.w3.org/2000/svg', t);
  for (const k in a) if (a[k] !== null && a[k] !== undefined) n.setAttribute(k, a[k]);
  return n;
};

const nf0 = new Intl.NumberFormat('en-IN', { maximumFractionDigits: 0 });
const nf2 = new Intl.NumberFormat('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmt = (v, d = 0) => !isFinite(v) ? '0' : (d ? nf2 : nf0).format(v);
const compact = v => {
  const a = Math.abs(v);
  if (a >= 1e7) return (v / 1e7).toFixed(a >= 1e8 ? 0 : 1) + ' Cr';
  if (a >= 1e5) return (v / 1e5).toFixed(a >= 1e6 ? 0 : 1) + ' L';
  if (a >= 1e3) return (v / 1e3).toFixed(a >= 1e4 ? 0 : 1) + 'k';
  return fmt(v);
};
const css = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim();
const SERIES = () => [1, 2, 3, 4, 5].map(i => css('--series-' + i));

/* ------------------------------------------------------------------ state */
let worker = null, buffer = null, fileName = '', result = null;
const opts = { mode: 'firstRow', denom: 'pmQty' };
const filters = { month: '', wh: '', group: '', whAxis: 'targetWh' };
let tab = 'target', page = 0, query = '';
const PAGE = 200;

function ensureWorker() {
  if (worker) return worker;
  worker = new Worker('assets/worker.js');
  worker.onmessage = e => {
    const m = e.data;
    if (m.type === 'progress') { $('#stage').textContent = m.stage + '…'; $('#fill').style.width = m.pct + '%'; }
    else if (m.type === 'result') { result = m.result; busy(false); render(); }
    else if (m.type === 'export') { saveBook(m.buffer); busy(false); }
    else if (m.type === 'error') { busy(false); showError(m.message); }
  };
  worker.onerror = e => { busy(false); showError(e.message || 'Worker failed to start.'); };
  return worker;
}

const busy = on => { $('#overlay').classList.toggle('on', on); if (on) $('#fill').style.width = '4%'; };
function showError(msg) {
  $('#err').classList.remove('hide');
  $('#errmsg').textContent = msg;
  $('#upload').hidden = false;
  $('#dash').hidden = true;
  $('#download').hidden = true;
  $('#reset').hidden = true;
}

/* ------------------------------------------------------------------ input */
function accept(file) {
  if (!file) return;
  if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) { showError('Please choose an .xlsx, .xlsm or .xls file.'); return; }
  fileName = file.name;
  $('#err').classList.add('hide');
  busy(true);
  const fr = new FileReader();
  fr.onerror = () => { busy(false); showError('The file could not be read from disk.'); };
  fr.onload = () => { buffer = fr.result; process(); };
  fr.readAsArrayBuffer(file);
}
function process() {
  if (!buffer) return;
  busy(true);
  const copy = buffer.slice(0);
  ensureWorker().postMessage({ cmd: 'process', buffer: copy, opts: { ...opts } }, [copy]);
}

const drop = $('#drop');
['dragenter', 'dragover'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => accept(e.dataTransfer.files[0]));
$('#pick').addEventListener('click', () => $('#file').click());
$('#file').addEventListener('change', e => accept(e.target.files[0]));
$('#reset').addEventListener('click', () => {
  result = null; buffer = null; $('#file').value = '';
  $('#upload').hidden = false; $('#dash').hidden = true;
  $('#download').hidden = true; $('#reset').hidden = true;
  $('#err').classList.add('hide');
});

/* ------------------------------------------------------------------ theme */
const savedTheme = (() => { try { return localStorage.getItem('fgpm-theme'); } catch { return null; } })();
if (savedTheme) document.documentElement.setAttribute('data-theme', savedTheme);
else document.documentElement.removeAttribute('data-theme');
$('#theme').addEventListener('click', () => {
  const now = document.documentElement.getAttribute('data-theme');
  const isDark = now ? now === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches;
  const next = isDark ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', next);
  try { localStorage.setItem('fgpm-theme', next); } catch {}
  if (result) drawCharts();
});

/* ---------------------------------------------------------------- segments */
function seg(id, key, after) {
  $(id).addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b) return;
    [...e.currentTarget.querySelectorAll('button')].forEach(x => x.setAttribute('aria-pressed', String(x === b)));
    const v = b.dataset.v;
    if (key in opts) opts[key] = v; else filters[key] = v;
    after();
  });
}
seg('#fMode', 'mode', () => process());
seg('#fDenom', 'denom', () => process());
seg('#fWhAxis', 'whAxis', () => { page = 0; render(false); });

['fMonth', 'fWh', 'fGroup'].forEach(id => {
  $('#' + id).addEventListener('change', e => {
    filters[{ fMonth: 'month', fWh: 'wh', fGroup: 'group' }[id]] = e.target.value;
    page = 0; render(false);
  });
});
$('#tabs').addEventListener('click', e => {
  const b = e.target.closest('button'); if (!b) return;
  [...e.currentTarget.querySelectorAll('button')].forEach(x => x.setAttribute('aria-selected', String(x === b)));
  tab = b.dataset.t; page = 0; query = ''; $('#search').value = ''; drawTable();
});
$('#search').addEventListener('input', e => { query = e.target.value.toLowerCase(); page = 0; drawTable(); });
$('#prev').addEventListener('click', () => { if (page > 0) { page--; drawTable(); } });
$('#next').addEventListener('click', () => { page++; drawTable(); });

/* ------------------------------------------------------------------ derive */
const filtered = () => result.long.filter(r =>
  (!filters.month || r.month === filters.month) &&
  (!filters.wh || r[filters.whAxis] === filters.wh) &&
  (!filters.group || r.group === filters.group));

function byMonth(rows) {
  const m = new Map(result.months.map(x => [x, { month: x, fgQty: 0, pmQty: 0, pmValue: 0 }]));
  rows.forEach(r => { const c = m.get(r.month); if (c) { c.fgQty += r.fgQty; c.pmQty += r.pmQty; c.pmValue += r.pmValue; } });
  return [...m.values()];
}
function byGroup(rows) {
  const m = new Map();
  rows.forEach(r => {
    let c = m.get(r.group);
    if (!c) m.set(r.group, c = { group: r.group, fgQty: 0, pmQty: 0, pmValue: 0 });
    c.fgQty += r.fgQty; c.pmQty += r.pmQty; c.pmValue += r.pmValue;
  });
  return [...m.values()].sort((a, b) => b.pmValue - a.pmValue);
}
const denomOf = c => opts.denom === 'fgQty' ? c.fgQty : c.pmQty;

/* ------------------------------------------------------------------ render */
function render(refill = true) {
  $('#upload').hidden = true;
  $('#dash').hidden = false;
  $('#download').hidden = false;
  $('#reset').hidden = false;
  if (refill) fillFilters();
  drawStatus(); drawTiles(); drawCharts(); drawTable();
}

function fillFilters() {
  const set = (id, values, keep) => {
    const s = $('#' + id), prev = keep && values.includes(filters[keep]) ? filters[keep] : '';
    s.length = 1;
    values.forEach(v => s.appendChild(Object.assign(document.createElement('option'), { value: v, textContent: v })));
    s.value = prev; if (keep) filters[keep] = prev;
  };
  set('fMonth', result.months, 'month');
  const whs = [...new Set(result.long.map(r => r[filters.whAxis]))].sort();
  set('fWh', whs, 'wh');
  set('fGroup', [...new Set(result.long.map(r => r.group))].sort(), 'group');
}

function drawStatus() {
  const s = result.stats, ok = Math.abs(s.difference) < 0.5;
  const unmapped = s.unmappedValue;
  const host = $('#statusbar');
  host.innerHTML = '';
  const cls = ok && !unmapped ? 'good' : (unmapped ? 'warning' : 'critical');
  const bar = document.createElement('div');
  bar.className = 'bar ' + cls;
  const modeTxt = opts.mode === 'firstRow'
    ? 'bucketed on each work order&rsquo;s first row'
    : 'bucketed on the FG row item group';
  bar.innerHTML =
    `<span class="dot"></span><div>` +
    `<b>${ok ? 'Control total matches' : 'Control total does not match'}</b> &mdash; ` +
    `source PKG <b class="tnum">${fmt(s.sourcePkg)}</b>, allocated <b class="tnum">${fmt(s.allocated)}</b>` +
    (ok ? '' : `, difference <b class="tnum">${fmt(s.difference)}</b>`) +
    `<div class="muted" style="margin-top:3px">` +
    `${fmt(result.rowCount)} rows &middot; ${fmt(s.workorderCount)} work orders &middot; ${modeTxt}` +
    (unmapped ? ` &middot; <b>${fmt(unmapped)}</b> unmapped across ${fmt(s.unmappedCount)} work orders` : '') +
    `</div></div>`;
  host.appendChild(bar);
}

function drawTiles() {
  const rows = filtered();
  const pmValue = rows.reduce((a, r) => a + r.pmValue, 0);
  const pmQty = rows.reduce((a, r) => a + r.pmQty, 0);
  const fgQty = rows.reduce((a, r) => a + r.fgQty, 0);
  const den = opts.denom === 'fgQty' ? fgQty : pmQty;
  const s = result.stats;
  const tiles = [
    ['PM Value', fmt(pmValue), 'From PKG rows only'],
    ['FG Qty', fmt(fgQty, 2), 'Sum of Qty on FG rows'],
    ['PM Qty', fmt(pmQty, 2), 'Value In FG &times; PKG%'],
    ['Blended PM Cost/kg', den ? fmt(pmValue / den, 2) : '—', 'PM Value &divide; ' + (opts.denom === 'fgQty' ? 'FG Qty' : 'PM Qty')],
    ['Work orders', fmt(s.workorderCount), `${fmt(s.multiFgCount)} multi-group &middot; ${fmt(s.noFgCount)} no FG line`]
  ];
  $('#tiles').innerHTML = tiles.map(([l, v, n]) =>
    `<div class="tile"><div class="label">${l}</div><div class="value tnum">${v}</div><div class="note">${n}</div></div>`).join('');
}

/* ------------------------------------------------------------- chart utils */
const tip = $('#tip');
function showTip(evt, title, lines) {
  tip.innerHTML = `<div class="th">${title}</div>` +
    lines.map(([k, v]) => `<div class="tr"><span>${k}</span><b class="tnum">${v}</b></div>`).join('');
  tip.classList.add('on');
  const r = tip.getBoundingClientRect();
  let x = evt.clientX + 14, y = evt.clientY + 14;
  if (x + r.width > innerWidth - 8) x = evt.clientX - r.width - 14;
  if (y + r.height > innerHeight - 8) y = evt.clientY - r.height - 14;
  tip.style.left = x + 'px'; tip.style.top = y + 'px';
}
const hideTip = () => tip.classList.remove('on');

function niceTicks(max, count = 4) {
  if (max <= 0) return [0, 1];
  const raw = max / count, mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const step = [1, 2, 2.5, 5, 10].map(m => m * mag).find(s => s >= raw) || mag * 10;
  const out = []; for (let v = 0; v <= max * 1.0001 + step; v += step) out.push(v);
  return out;
}
// rounded data-end at the top, square at the baseline
const colPath = (x, y, w, h, r) => {
  r = Math.max(0, Math.min(r, w / 2, h));
  return `M${x},${y + h}V${y + r}a${r},${r} 0 0 1 ${r},-${r}h${w - 2 * r}a${r},${r} 0 0 1 ${r},${r}V${y + h}Z`;
};
// rounded data-end at the right, square at the baseline
const barPath = (x, y, w, h, r) => {
  r = Math.max(0, Math.min(r, h / 2, w));
  return `M${x},${y}h${w - r}a${r},${r} 0 0 1 ${r},${r}v${h - 2 * r}a${r},${r} 0 0 1 ${-r},${r}H${x}Z`;
};
const clear = n => { while (n.firstChild) n.removeChild(n.firstChild); };

/* ------------------------------------------------------------- the charts */
function drawCharts() {
  if (!result) return;
  const rows = filtered();
  chartValueByMonth(byMonth(rows));
  chartTrend(rows);
  chartGroups(byGroup(rows));
}

function chartValueByMonth(data) {
  const svg = $('#cValue'), W = svg.clientWidth || 520, H = 260;
  const M = { t: 18, r: 14, b: 30, l: 58 };
  clear(svg); svg.setAttribute('viewBox', `0 0 ${W} ${H}`); svg.setAttribute('height', H);
  if (!data.length) return;

  const max = Math.max(...data.map(d => d.pmValue), 1);
  const ticks = niceTicks(max);
  const top = ticks[ticks.length - 1] || 1;
  const pw = W - M.l - M.r, ph = H - M.t - M.b;
  const y = v => M.t + ph - (v / top) * ph;
  const band = pw / data.length;
  const bw = Math.min(24, band - 12);

  ticks.forEach(t => {
    svg.appendChild(svgEl('line', { x1: M.l, x2: W - M.r, y1: y(t), y2: y(t), stroke: css('--grid'), 'stroke-width': 1 }));
    const lb = svgEl('text', { x: M.l - 9, y: y(t) + 4, 'text-anchor': 'end', class: 'axis-txt tnum' });
    lb.textContent = compact(t); svg.appendChild(lb);
  });

  data.forEach((d, i) => {
    const cx = M.l + band * i + band / 2, h = Math.max(0, y(0) - y(d.pmValue));
    if (h > 0) {
      const p = svgEl('path', { d: colPath(cx - bw / 2, y(d.pmValue), bw, h, 4), fill: css('--series-1') });
      svg.appendChild(p);
    }
    const hit = svgEl('rect', { x: cx - band / 2, y: M.t, width: band, height: ph, fill: 'transparent' });
    hit.addEventListener('mousemove', e => showTip(e, d.month, [
      ['PM Value', fmt(d.pmValue)], ['FG Qty', fmt(d.fgQty, 2)], ['PM Qty', fmt(d.pmQty, 2)],
      ['PM Cost/kg', denomOf(d) ? fmt(d.pmValue / denomOf(d), 2) : '—']
    ]));
    hit.addEventListener('mouseleave', hideTip);
    svg.appendChild(hit);

    const cap = svgEl('text', { x: cx, y: y(d.pmValue) - 7, 'text-anchor': 'middle', class: 'val-txt tnum' });
    cap.textContent = compact(d.pmValue); svg.appendChild(cap);
    const mx = svgEl('text', { x: cx, y: H - 9, 'text-anchor': 'middle', class: 'axis-txt' });
    mx.textContent = d.month; svg.appendChild(mx);
  });
  svg.appendChild(svgEl('line', { x1: M.l, x2: W - M.r, y1: y(0), y2: y(0), stroke: css('--axis'), 'stroke-width': 1 }));
}

function chartTrend(rows) {
  const svg = $('#cTrend'), W = svg.clientWidth || 520, H = 260;
  const M = { t: 18, r: 62, b: 30, l: 54 };
  clear(svg); svg.setAttribute('viewBox', `0 0 ${W} ${H}`); svg.setAttribute('height', H);
  const legend = $('#lTrend'); legend.innerHTML = '';

  const top5 = byGroup(rows).slice(0, 5).map(g => g.group);
  if (!top5.length) return;
  const months = result.months;
  const colors = SERIES();

  const series = top5.map((g, i) => ({
    name: g, color: colors[i],
    points: months.map(m => {
      const cs = rows.filter(r => r.group === g && r.month === m)
        .reduce((a, r) => (a.fgQty += r.fgQty, a.pmQty += r.pmQty, a.pmValue += r.pmValue, a), { fgQty: 0, pmQty: 0, pmValue: 0 });
      const d = denomOf(cs);
      return { month: m, v: d ? cs.pmValue / d : null, cs };
    })
  }));

  const vals = series.flatMap(s => s.points.map(p => p.v)).filter(v => v !== null && isFinite(v));
  if (!vals.length) return;
  const ticks = niceTicks(Math.max(...vals));
  const top = ticks[ticks.length - 1] || 1;
  const pw = W - M.l - M.r, ph = H - M.t - M.b;
  const x = i => months.length === 1 ? M.l + pw / 2 : M.l + (pw * i) / (months.length - 1);
  const y = v => M.t + ph - (v / top) * ph;

  // one decimal precision for the whole axis, taken from the tick step
  const step = ticks.length > 1 ? ticks[1] - ticks[0] : top;
  const axisDec = step >= 10 ? 0 : step >= 1 ? 1 : 2;
  ticks.forEach(t => {
    svg.appendChild(svgEl('line', { x1: M.l, x2: W - M.r, y1: y(t), y2: y(t), stroke: css('--grid'), 'stroke-width': 1 }));
    const lb = svgEl('text', { x: M.l - 9, y: y(t) + 4, 'text-anchor': 'end', class: 'axis-txt tnum' });
    lb.textContent = fmt(t, axisDec); svg.appendChild(lb);
  });
  months.forEach((m, i) => {
    const lb = svgEl('text', { x: x(i), y: H - 9, 'text-anchor': 'middle', class: 'axis-txt' });
    lb.textContent = m; svg.appendChild(lb);
  });

  const endLabels = [];
  series.forEach(s => {
    const pts = s.points.map((p, i) => ({ ...p, x: x(i), y: p.v === null ? null : y(p.v) })).filter(p => p.y !== null);
    if (pts.length > 1) {
      svg.appendChild(svgEl('path', {
        d: pts.map((p, i) => (i ? 'L' : 'M') + p.x + ',' + p.y).join(''),
        fill: 'none', stroke: s.color, 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round'
      }));
    }
    pts.forEach(p => {
      const dot = svgEl('circle', { cx: p.x, cy: p.y, r: 4, fill: s.color, stroke: css('--surface-1'), 'stroke-width': 2 });
      svg.appendChild(dot);
      const hit = svgEl('circle', { cx: p.x, cy: p.y, r: 11, fill: 'transparent' });
      hit.addEventListener('mousemove', e => showTip(e, `${s.name} · ${p.month}`, [
        ['PM Cost/kg', fmt(p.v, 2)], ['PM Value', fmt(p.cs.pmValue)],
        ['FG Qty', fmt(p.cs.fgQty, 2)], ['PM Qty', fmt(p.cs.pmQty, 2)]
      ]));
      hit.addEventListener('mouseleave', hideTip);
      svg.appendChild(hit);
    });
    const last = pts[pts.length - 1];
    if (last) endLabels.push({ x: last.x, y: last.y, v: last.v });

    const key = document.createElement('span');
    key.innerHTML = `<i style="background:${s.color}"></i>${s.name}`;
    legend.appendChild(key);
  });

  // Direct end-labels only where they do not collide. Converging lines get no
  // label rather than a stacked pile detached from its line - the legend and
  // the hover tooltip carry those values instead.
  endLabels.sort((a, b) => a.y - b.y);
  let lastY = -Infinity;
  endLabels.forEach(p => {
    if (p.y - lastY < 15) return;
    lastY = p.y;
    const lb = svgEl('text', { x: p.x + 9, y: p.y + 4, class: 'val-txt tnum' });
    lb.textContent = fmt(p.v, 2);
    svg.appendChild(lb);
  });
  svg.appendChild(svgEl('line', { x1: M.l, x2: W - M.r, y1: y(0), y2: y(0), stroke: css('--axis'), 'stroke-width': 1 }));
}

function chartGroups(groups) {
  const data = groups.slice(0, 12);
  const svg = $('#cGroups'), W = svg.clientWidth || 900;
  const rowH = 30, M = { t: 8, r: 96, b: 26, l: 166 };
  const H = M.t + M.b + Math.max(1, data.length) * rowH;
  clear(svg); svg.setAttribute('viewBox', `0 0 ${W} ${H}`); svg.setAttribute('height', H);
  if (!data.length) return;

  const max = Math.max(...data.map(d => d.pmValue), 1);
  const ticks = niceTicks(max, 5);
  const top = ticks[ticks.length - 1] || 1;
  const pw = W - M.l - M.r;
  const x = v => M.l + (v / top) * pw;
  const bh = Math.min(24, rowH - 10);

  ticks.forEach(t => {
    svg.appendChild(svgEl('line', { x1: x(t), x2: x(t), y1: M.t, y2: H - M.b, stroke: css('--grid'), 'stroke-width': 1 }));
    const lb = svgEl('text', { x: x(t), y: H - 8, 'text-anchor': 'middle', class: 'axis-txt tnum' });
    lb.textContent = compact(t); svg.appendChild(lb);
  });

  data.forEach((d, i) => {
    const cy = M.t + i * rowH + rowH / 2, w = Math.max(0, x(d.pmValue) - M.l);
    if (w > 0) svg.appendChild(svgEl('path', { d: barPath(M.l, cy - bh / 2, w, bh, 4), fill: css('--series-1') }));

    const name = svgEl('text', { x: M.l - 11, y: cy + 4, 'text-anchor': 'end', class: 'axis-txt' });
    name.textContent = d.group.length > 24 ? d.group.slice(0, 23) + '…' : d.group;
    svg.appendChild(name);

    const den = denomOf(d);
    const v = svgEl('text', { x: x(d.pmValue) + 9, y: cy + 4, class: 'val-txt tnum' });
    v.textContent = compact(d.pmValue) + (den ? `  ·  ${fmt(d.pmValue / den, 2)}/kg` : '');
    svg.appendChild(v);

    const hit = svgEl('rect', { x: M.l, y: cy - rowH / 2, width: pw + M.r - 8, height: rowH, fill: 'transparent' });
    hit.addEventListener('mousemove', e => showTip(e, d.group, [
      ['PM Value', fmt(d.pmValue)], ['FG Qty', fmt(d.fgQty, 2)], ['PM Qty', fmt(d.pmQty, 2)],
      ['PM Cost/kg', den ? fmt(d.pmValue / den, 2) : '—']
    ]));
    hit.addEventListener('mouseleave', hideTip);
    svg.appendChild(hit);
  });
  svg.appendChild(svgEl('line', { x1: M.l, x2: M.l, y1: M.t, y2: H - M.b, stroke: css('--axis'), 'stroke-width': 1 }));
}

let rt; addEventListener('resize', () => { clearTimeout(rt); rt = setTimeout(() => result && drawCharts(), 160); });

/* ------------------------------------------------------------------ tables */
function tableData() {
  const rows = filtered();
  if (tab === 'target' || tab === 'source') {
    const whKey = tab === 'target' ? 'targetWh' : 'sourceWh';
    const label = tab === 'target' ? 'Target Warehouse' : 'Source Warehouse';
    const map = new Map();
    rows.forEach(r => {
      const k = r[whKey] + '' + r.group;
      let o = map.get(k);
      if (!o) map.set(k, o = { wh: r[whKey], group: r.group, c: {} });
      const c = o.c[r.month] || (o.c[r.month] = { fgQty: 0, pmQty: 0, pmValue: 0 });
      c.fgQty += r.fgQty; c.pmQty += r.pmQty; c.pmValue += r.pmValue;
    });
    const months = result.months;
    const header = [label, 'FG Item Group'];
    months.forEach(m => header.push(`${m} (FG Qty)`, `${m} (PM Qty)`, `${m} (PM Value)`, `${m} (PM Cost/KG)`));
    header.push('Total (FG Qty)', 'Total (PM Qty)', 'Total (PM Value)', 'Total (PM Cost/KG)');
    const body = [...map.values()].sort((a, b) => (a.wh + a.group).localeCompare(b.wh + b.group)).map(o => {
      const line = [o.wh, o.group]; let tq = 0, tp = 0, tv = 0;
      months.forEach(m => {
        const c = o.c[m] || { fgQty: 0, pmQty: 0, pmValue: 0 };
        const d = denomOf(c);
        line.push(c.fgQty, c.pmQty, c.pmValue, d ? c.pmValue / d : 0);
        tq += c.fgQty; tp += c.pmQty; tv += c.pmValue;
      });
      const td = opts.denom === 'fgQty' ? tq : tp;
      line.push(tq, tp, tv, td ? tv / td : 0);
      return line;
    });
    if (body.length) {
      const g = ['Grand Total', ''];
      for (let c = 2; c < header.length; c++) g.push(body.reduce((s, r) => s + r[c], 0));
      for (let i = 0; i < months.length; i++) {
        const b = 2 + i * 4, d = opts.denom === 'fgQty' ? g[b] : g[b + 1];
        g[b + 3] = d ? g[b + 2] / d : 0;
      }
      const tb = 2 + months.length * 4, td = opts.denom === 'fgQty' ? g[tb] : g[tb + 1];
      g[tb + 3] = td ? g[tb + 2] / td : 0;
      body.push(g);
    }
    return { header, body, textCols: 2, totalLast: true };
  }
  if (tab === 'wo') {
    return {
      header: ['Workorder', 'Bucket Item Group', 'First Row Type', 'Month', 'FG Item Groups', 'FG Groups',
               'Target Warehouse', 'Source Warehouse', 'FG Qty', 'FG Value', 'PM Qty', 'PM Value', 'PKG Lines', 'FG Lines'],
      body: result.workorders.map(w => [w.wo, w.firstGroup, w.firstType, w.month, w.fgGroups, w.fgGroupCount,
        w.targetWh, w.sourceWh, w.fgQty, w.fgValue, w.pmQty, w.pmValue, w.pkgLines, w.fgLines]),
      textCols: 8
    };
  }
  if (tab === 'audit') return null;   // rendered by renderAudit()
  return {
    header: ['Workorder', 'Month', 'Bucket Item Group', 'FG Item Groups', 'Count', 'PM Value', 'FG Qty'],
    body: result.multiFg.map(r => [r.wo, r.month, r.firstGroup, r.fgGroups, r.count, r.pmValue, r.fgQty]),
    textCols: 4
  };
}

/* --------------------------------------------------- logic & audit panel */
function auditChecks() {
  const a = result.audit, s = result.stats, L = result.long;
  const near = (x, y) => Math.abs(x - y) < 0.5;

  const carriers = L.filter(r => r.pmValue !== 0);
  const distinctCarriers = new Set(carriers.map(r => r.wo)).size;

  const monthSum = L.reduce((m, r) => (m[r.month] = (m[r.month] || 0) + r.pmValue, m), {});
  const monthTotal = Object.values(monthSum).reduce((x, y) => x + y, 0);

  const byT = L.reduce((m, r) => (m[r.targetWh] = (m[r.targetWh] || 0) + r.pmValue, m), {});
  const byS = L.reduce((m, r) => (m[r.sourceWh] = (m[r.sourceWh] || 0) + r.pmValue, m), {});
  const tSum = Object.values(byT).reduce((x, y) => x + y, 0);
  const sSum = Object.values(byS).reduce((x, y) => x + y, 0);

  const typeTotal = Object.values(a.typeCounts).reduce((x, y) => x + y, 0);

  return [
    {
      ok: near(s.allocated + s.unmappedValue, s.sourcePkg),
      what: 'Every rupee of packaging spend is accounted for',
      why: 'Allocated + unmapped must equal the PKG total read straight off the source rows. A gap here would mean value silently vanished between reading and reporting.',
      fig: `${fmt(s.allocated)} + ${fmt(s.unmappedValue)} vs ${fmt(s.sourcePkg)}`
    },
    {
      ok: near(s.unmappedValue, 0),
      what: 'Nothing is left unallocated',
      why: opts.mode === 'firstRow'
        ? 'Every work order resolves to exactly one item group, so no packaging value is stranded.'
        : 'FG-row bucketing cannot resolve work orders with zero or several FG item groups, so their value is reported as unmapped rather than guessed at. Switch the bucket to “Work order first row” to clear this.',
      fig: fmt(s.unmappedValue)
    },
    {
      ok: near(a.pkgAmountRaw - a.pkgValueNoWorkorder, s.sourcePkg),
      what: 'The source total re-adds from the raw rows',
      why: 'Summing Total Amount over every PKG row independently of the work order build reproduces the same control figure — the aggregation is not inventing or dropping value.',
      fig: `${fmt(a.pkgAmountRaw - a.pkgValueNoWorkorder)} vs ${fmt(s.sourcePkg)}`
    },
    {
      ok: distinctCarriers === carriers.length,
      what: 'No work order is counted twice',
      why: 'Each work order attaches its PM Value to exactly one output cell. If a work order appeared as a value carrier twice, its packaging cost would be double counted in every roll-up above.',
      fig: `${fmt(carriers.length)} cells · ${fmt(distinctCarriers)} work orders`
    },
    {
      ok: near(monthTotal, s.allocated),
      what: 'The monthly figures re-add to the total',
      why: 'Summing PM Value across the month columns returns the allocated total, so no value falls outside the reported months.',
      fig: `${fmt(monthTotal)} vs ${fmt(s.allocated)}`
    },
    {
      ok: near(tSum, sSum) && near(tSum, s.allocated),
      what: 'Both warehouse views agree',
      why: 'The Target and Source warehouse summaries are two cuts of the same numbers. They must total identically; a difference would mean one axis is dropping or duplicating rows.',
      fig: `${fmt(tSum)} vs ${fmt(sSum)}`
    },
    {
      ok: near(s.totalFgQty, a.fgQtyRaw),
      what: 'FG Qty ties back to the FG rows',
      why: 'The quantity denominator equals the plain sum of Qty over every FG row in the file — nothing is filtered out on the way to the report.',
      fig: `${fmt(s.totalFgQty, 2)} vs ${fmt(a.fgQtyRaw, 2)}`
    },
    {
      ok: a.pkgRowsNoWorkorder === 0,
      what: 'Every PKG row carries a work order',
      why: 'A packaging row with a blank work order cannot be attributed to any item group, and its value would be dropped from the report entirely.',
      fig: `${fmt(a.pkgRowsNoWorkorder)} orphan rows` + (a.pkgValueNoWorkorder ? ` · ${fmt(a.pkgValueNoWorkorder)}` : '')
    },
    {
      ok: a.woNegativePm === 0,
      what: 'No negative packaging value',
      why: 'A negative PM Value points at a reversal or credit note booked into the extract, which would understate the cost per kg for its item group.',
      fig: `${fmt(a.woNegativePm)} work orders`
    },
    {
      ok: typeTotal === result.rowCount,
      what: 'Every row is classified',
      why: 'Each of the rows read carries an Item Type that was counted. An unclassified row is a row whose value went nowhere.',
      fig: `${fmt(typeTotal)} of ${fmt(result.rowCount)}`
    },
    {
      ok: a.unknownMonthRows === 0,
      what: 'Every row lands in a month',
      why: 'Rows with an unreadable Date and Month fall into an “Unknown” column instead of the calendar. Their value is still counted, but not in the month it belongs to.',
      fig: `${fmt(a.unknownMonthRows)} rows`
    }
  ];
}

function renderAudit() {
  const a = result.audit, s = result.stats;
  const checks = auditChecks();
  const passed = checks.filter(c => c.ok).length;

  const pill = ok => `<span class="pill ${ok ? 'pass' : 'fail'}">${ok ? '✓ Pass' : '✕ Check'}</span>`;
  const census = Object.entries(a.typeCounts).sort((x, y) => y[1] - x[1])
    .map(([k, v]) => `<tr><td class="what">${k}</td><td class="fig">${fmt(v)} rows</td></tr>`).join('');
  const prov = Object.entries(a.bucketSource).sort((x, y) => y[1] - x[1])
    .map(([k, v]) => `<tr><td class="what">First row is ${k}</td><td class="fig">${fmt(v)} work orders</td></tr>`).join('');

  const modeName = opts.mode === 'firstRow' ? 'Work order first row' : 'FG row Item Group';
  const denName = opts.denom === 'fgQty' ? 'FG Qty' : 'PM Qty';

  return `<div class="doc">

<h4>Reconciliation &mdash; ${passed} of ${checks.length} checks pass</h4>
<p>Recomputed from the file you loaded, every time the settings change. These are the reasons
to trust the numbers above; each one states what would be wrong if it failed.</p>
<table class="checks"><tbody>
${checks.map(c => `<tr>
  <td class="st">${pill(c.ok)}</td>
  <td><div class="what">${c.what}</div><div class="why">${c.why}</div></td>
  <td class="fig">${c.fig}</td>
</tr>`).join('')}
</tbody></table>

<div class="split">
  <div>
    <h4>Row census</h4>
    <p>Where the ${fmt(result.rowCount)} rows of <code>${result.sheetName}</code> went.</p>
    <table class="checks"><tbody>${census}
      <tr><td class="what">No work order</td><td class="fig">${fmt(a.rowsNoWorkorder)} rows</td></tr>
    </tbody></table>
  </div>
  <div>
    <h4>Bucket provenance</h4>
    <p>Which line each work order took its item group from.</p>
    <table class="checks"><tbody>${prov}</tbody></table>
  </div>
</div>

<h4>Current settings</h4>
<table class="checks"><tbody>
  <tr><td class="what">PM Value bucket</td><td class="fig">${modeName}</td></tr>
  <tr><td class="what">PM Cost/kg denominator</td><td class="fig">${denName}</td></tr>
  <tr><td class="what">Work orders</td><td class="fig">${fmt(s.workorderCount)}</td></tr>
  <tr><td class="what">Work orders with &gt;1 FG item group</td><td class="fig">${fmt(s.multiFgCount)} &middot; ${fmt(s.multiFgValue)}</td></tr>
  <tr><td class="what">Work orders with no FG line</td><td class="fig">${fmt(s.noFgCount)} &middot; ${fmt(s.noFgValue)}</td></tr>
</tbody></table>

<h4>The logic applied, in order</h4>

<div class="rule"><div class="n">1</div><div class="b">
  <div class="t">Read the first sheet in file order</div>
  <p>Row order is load-bearing &mdash; rule 3 depends on it. Header whitespace is stripped;
  numbers are coerced from text, and anything unparseable becomes 0 rather than breaking the run.</p>
</div></div>

<div class="rule"><div class="n">2</div><div class="b">
  <div class="t">PM Value &mdash; packaging rows only</div>
  <pre>PM Value(work order) = SUM(Total Amount) WHERE Item Type = 'PKG'</pre>
  <p><code>RM</code>, <code>FG</code> and <code>BiProduct</code> rows contribute nothing.</p>
</div></div>

<div class="rule"><div class="n">3</div><div class="b">
  <div class="t">Bucketing &mdash; which item group that value is reported under</div>
  <p><b>Work order first row</b> (in use${opts.mode === 'firstRow' ? '' : ' &mdash; currently off'}): the item group and month
  of the work order&rsquo;s <b>first row in file order</b>, which is its primary input line &mdash; the RM line for
  most work orders, the PKG line for those with no RM at all. Equivalent to
  <code>VLOOKUP(Workorder, &lt;data&gt;, Item Group)</code>.</p>
  <p><b>FG row</b>: the item group on the work order&rsquo;s FG rows. Work orders with zero or several
  distinct FG groups cannot be resolved and their value is reported unmapped.</p>
  <p>The two differ because a work order routinely consumes one item group and produces another,
  and because many work orders have no FG line to read at all.</p>
</div></div>

<div class="rule"><div class="n">4</div><div class="b">
  <div class="t">One value carrier per work order</div>
  <p>Where a work order spans several months or warehouses on its FG rows, its PM Value is attached
  to the single largest FG cell rather than repeated against each. Repeating it is the ordinary way
  this calculation gets inflated; check 4 above is what proves it did not happen here.</p>
</div></div>

<div class="rule"><div class="n">5</div><div class="b">
  <div class="t">Quantities</div>
  <pre>PM Qty = SUM over FG rows of ( Value In FG &times; PKG / 100 )
FG Qty = SUM over FG rows of Qty</pre>
  <p><code>PKG</code> is a percentage held as a plain number, so <code>2.02</code> means 2.02%.</p>
</div></div>

<div class="rule"><div class="n">6</div><div class="b">
  <div class="t">Cost per kg</div>
  <pre>PM Cost/kg = PM Value &divide; ${denName}</pre>
  <p>Totals re-derive the ratio from the summed numerator over the summed denominator &mdash;
  never an average of the monthly ratios, which would weight a small month the same as a large one.</p>
</div></div>

<div class="rule"><div class="n">7</div><div class="b">
  <div class="t">Warehouses</div>
  <p>Target Warehouse is read from the FG rows. <b>Source Warehouse is read from the input
  (RM / PKG) rows</b>, because FG rows carry it blank &mdash; taking it from the FG line makes the
  entire source view read <code>Unknown</code>.</p>
</div></div>

<div class="rule"><div class="n">8</div><div class="b">
  <div class="t">Months</div>
  <p>Taken from <code>Date</code> where it parses, otherwise from a text <code>Month</code> column.
  Columns appear in calendar order, and only for months actually present in the data.</p>
</div></div>

</div>`;
}

function drawTable() {
  const host = $('#tablehost');

  if (tab === 'audit') {
    $('#search').hidden = true;
    $('#rowcount').textContent = '';
    $('#pager').classList.add('hide');
    host.innerHTML = renderAudit();
    return;
  }
  $('#search').hidden = false;

  const { header, body, textCols, totalLast } = tableData();
  const grand = totalLast && body.length ? body[body.length - 1] : null;
  let rows = grand ? body.slice(0, -1) : body;

  if (query) rows = rows.filter(r => r.some(c => String(c).toLowerCase().includes(query)));
  const total = rows.length;
  const pages = Math.max(1, Math.ceil(total / PAGE));
  page = Math.min(page, pages - 1);
  const slice = rows.slice(page * PAGE, page * PAGE + PAGE);

  $('#rowcount').textContent = total ? `${fmt(total)} row${total === 1 ? '' : 's'}` : '';
  $('#pager').classList.toggle('hide', pages <= 1);
  $('#pageinfo').textContent = `Page ${page + 1} of ${pages}`;
  $('#prev').disabled = page === 0;
  $('#next').disabled = page >= pages - 1;

  if (!total) { host.innerHTML = '<div class="empty">Nothing to show for the current filters.</div>'; return; }

  const cell = (v, i) => {
    if (typeof v === 'number') {
      return `<td class="num">${fmt(v, Number.isInteger(v) ? 0 : 2)}</td>`;
    }
    return `<td class="txt">${String(v ?? '')}</td>`;
  };
  const tr = (r, cls = '') => `<tr class="${cls}">${r.map((v, i) => cell(v, i)).join('')}</tr>`;

  host.innerHTML =
    `<table><thead><tr>${header.map((h, i) => `<th class="${i < textCols ? 'txt' : ''}">${h}</th>`).join('')}</tr></thead>` +
    `<tbody>${slice.map(r => tr(r)).join('')}${grand && !query ? tr(grand, 'total') : ''}</tbody></table>`;
}

/* ------------------------------------------------------------------ export */
$('#download').addEventListener('click', () => {
  if (!result) return;
  busy(true); $('#stage').textContent = 'Building workbook…'; $('#fill').style.width = '55%';
  const checks = auditChecks().map(c => ({ what: c.what, ok: c.ok, fig: c.fig }));
  ensureWorker().postMessage({ cmd: 'export', result, opts: { ...opts }, checks });
});
function saveBook(buf) {
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = (fileName.replace(/\.[^.]+$/, '') || 'FG_PM') + '_Analysis.xlsx';
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
}
})();
