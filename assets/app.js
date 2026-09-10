/* ============================================================================
   FG Packaging Cost Analyser - UI layer
   Charts are hand-rolled SVG so the mark specs hold exactly:
   <=24px bars with a 4px rounded data-end square at the baseline, 2px lines,
   >=8px markers carrying a 2px surface ring, solid hairline grid.
   ========================================================================== */
(() => {
'use strict';

// bumped whenever worker.js changes, so browsers never run a cached worker
const BUILD = '11';
self.__BUILD = BUILD;

const $ = s => document.querySelector(s);
const svgEl = (t, a = {}) => {
  const n = document.createElementNS('http://www.w3.org/2000/svg', t);
  for (const k in a) if (a[k] !== null && a[k] !== undefined) n.setAttribute(k, a[k]);
  return n;
};

const nf0 = new Intl.NumberFormat('en-IN', { maximumFractionDigits: 0 });
const nfD = d => new Intl.NumberFormat('en-IN', { minimumFractionDigits: d, maximumFractionDigits: d });
const fmt = (v, d = 0) => !isFinite(v) ? '0' : (d ? nfD(d) : nf0).format(v);
const compact = v => {
  const a = Math.abs(v);
  if (a >= 1e7) return (v / 1e7).toFixed(a >= 1e8 ? 0 : 1) + ' Cr';
  if (a >= 1e5) return (v / 1e5).toFixed(a >= 1e6 ? 0 : 1) + ' L';
  if (a >= 1e3) return (v / 1e3).toFixed(a >= 1e4 ? 0 : 1) + 'k';
  return fmt(v, a < 100 && a % 1 ? 2 : 0);
};
const css = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim();
const SERIES = () => [1, 2, 3, 4, 5].map(i => css('--series-' + i));

const METRICS = ['Sum of Qty', 'Sum of PKg Cost', 'Qty / PKg Cost'];
const ratio = (q, c) => c ? q / c : 0;

/* ------------------------------------------------------------------ state */
let worker = null, buffer = null, fileName = '', result = null;
const filters = { month: '', group: '' };
let sheet = '';                    // '' means let the worker choose
let tab = 'trend', page = 0, query = '';
const PAGE = 200;

function ensureWorker() {
  if (worker) return worker;
  worker = new Worker('assets/worker.js?v=' + BUILD);
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
  ensureWorker().postMessage({ cmd: 'process', buffer: copy, opts: { sheet } }, [copy]);
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

/* ---------------------------------------------------------------- controls */
$('#fSheet').addEventListener('change', e => {
  sheet = e.target.value;
  filters.month = ''; filters.group = ''; page = 0;
  process();                        // a different sheet is a different dataset
});
['fMonth', 'fGroup'].forEach(id => {
  $('#' + id).addEventListener('change', e => {
    filters[{ fMonth: 'month', fGroup: 'group' }[id]] = e.target.value;
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
  (!filters.group || r.group === filters.group));

const monthsIn = rows => result.months.filter(m => rows.some(r => r.month === m));
const groupsIn = rows => [...new Set(rows.map(r => r.group))].sort((a, b) => a.localeCompare(b));

function byMonth(rows) {
  const m = new Map(monthsIn(rows).map(x => [x, { month: x, qty: 0, cost: 0 }]));
  rows.forEach(r => { const c = m.get(r.month); if (c) { c.qty += r.qty; c.cost += r.cost; } });
  const out = [...m.values()]; out.forEach(c => c.ratio = ratio(c.qty, c.cost));
  return out;
}
function byGroup(rows) {
  const m = new Map();
  rows.forEach(r => {
    let c = m.get(r.group);
    if (!c) m.set(r.group, c = { group: r.group, qty: 0, cost: 0 });
    c.qty += r.qty; c.cost += r.cost;
  });
  const out = [...m.values()]; out.forEach(c => c.ratio = ratio(c.qty, c.cost));
  return out.sort((a, b) => b.cost - a.cost);
}
const totalsOf = rows => {
  const qty = rows.reduce((a, r) => a + r.qty, 0), cost = rows.reduce((a, r) => a + r.cost, 0);
  return { qty, cost, ratio: ratio(qty, cost) };
};

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
  const set = (id, values, key) => {
    const s = $('#' + id), prev = values.includes(filters[key]) ? filters[key] : '';
    s.length = 1;
    values.forEach(v => s.appendChild(Object.assign(document.createElement('option'), { value: v, textContent: v })));
    s.value = prev; filters[key] = prev;
  };
  set('fMonth', result.months, 'month');
  set('fGroup', result.groups, 'group');

  // the workbook may hold several plausible tabs, so name the one in use
  const sel = $('#fSheet');
  sel.innerHTML = '';
  (result.sheets || []).filter(s => s.usable).forEach(s => {
    sel.appendChild(Object.assign(document.createElement('option'),
      { value: s.name, textContent: `${s.name} (${fmt(s.rows)} rows)` }));
  });
  if (![...sel.options].some(o => o.value === result.sheetName)) {
    sel.appendChild(Object.assign(document.createElement('option'),
      { value: result.sheetName, textContent: result.sheetName }));
  }
  sel.value = result.sheetName;
  sheet = result.sheetName;
}

function drawStatus() {
  const a = result.audit;
  const recomputeOk = !a.hasSourcePkgCost || a.recomputeMaxDiff < 1e-6;
  const host = $('#statusbar');
  host.innerHTML = '';
  const bar = document.createElement('div');
  bar.className = 'bar ' + (recomputeOk ? 'good' : 'critical');
  bar.innerHTML =
    `<span class="dot"></span><div>` +
    `<b>PKg Cost = Total Cost &times; PKG / 100</b>` +
    (a.hasSourcePkgCost
      ? ` &mdash; reproduces the <code>PKg Cost</code> column in your file` +
        ` (max difference <b class="tnum">${a.recomputeMaxDiff.toExponential(2)}</b>)`
      : ` &mdash; computed from <code>Total Cost</code> and <code>PKG %</code>`) +
    `<div class="muted" style="margin-top:3px">` +
    `sheet <b>${result.sheetName}</b> &middot; ` +
    `${fmt(result.rowCount)} rows read &middot; <b>${fmt(a.fgRows)}</b> FG rows analysed &middot; ` +
    `${fmt(result.rowCount - a.fgRows)} non-FG rows excluded &middot; ` +
    `${result.groups.length} item groups &middot; ${result.months.join(', ')}` +
    (a.unknownMonthRows ? ` &middot; <b>${fmt(a.unknownMonthRows)}</b> rows with no readable month` : '') +
    `</div></div>`;
  host.appendChild(bar);
}

function drawTiles() {
  const rows = filtered();
  const t = totalsOf(rows);
  const scope = filters.month || filters.group ? 'current filters' : 'all FG rows';
  const tiles = [
    ['Sum of Qty', fmt(t.qty, 2), `Kg produced &middot; ${scope}`],
    ['Sum of PKg Cost', fmt(t.cost, 2), 'Total Cost &times; PKG / 100'],
    ['Qty / PKg Cost', t.cost ? fmt(t.ratio, 4) : '—', 'Re-derived from both totals'],
    ['FG rows', fmt(result.audit.fgRows), `of ${fmt(result.rowCount)} rows read`],
    ['Item groups', fmt(groupsIn(rows).length), `${result.months.length} month${result.months.length === 1 ? '' : 's'} in file`]
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
const axisTxt = (x, y, s, anchor = 'middle') => {
  const t = svgEl('text', { x, y, 'text-anchor': anchor, class: 'axis-txt tnum' });
  t.textContent = s; return t;
};

/* ------------------------------------------------------------- the charts */
function drawCharts() {
  if (!result) return;
  const rows = filtered();
  chartCostByMonth(byMonth(rows));
  chartRatioTrend(rows);
  chartGroups(byGroup(rows));
}

/* One series, so one colour and no legend - the card title names it. */
function chartCostByMonth(data) {
  const svg = $('#cCost'), W = svg.clientWidth || 520, H = 260;
  const M = { t: 18, r: 14, b: 30, l: 62 };
  clear(svg); svg.setAttribute('viewBox', `0 0 ${W} ${H}`); svg.setAttribute('height', H);
  if (!data.length) return;

  const ticks = niceTicks(Math.max(...data.map(d => d.cost), 1));
  const top = ticks[ticks.length - 1] || 1;
  const pw = W - M.l - M.r, ph = H - M.t - M.b;
  const y = v => M.t + ph - (v / top) * ph;
  const band = pw / data.length;
  const bw = Math.min(24, band - 12);

  ticks.forEach(t => {
    svg.appendChild(svgEl('line', { x1: M.l, x2: W - M.r, y1: y(t), y2: y(t), stroke: css('--grid'), 'stroke-width': 1 }));
    svg.appendChild(axisTxt(M.l - 9, y(t) + 4, compact(t), 'end'));
  });

  data.forEach((d, i) => {
    const cx = M.l + band * i + band / 2, h = Math.max(0, y(0) - y(d.cost));
    if (h > 0) svg.appendChild(svgEl('path', {
      d: colPath(cx - bw / 2, y(d.cost), bw, h, 4), fill: css('--series-1')
    }));
    svg.appendChild(axisTxt(cx, H - 10, d.month));
    const hit = svgEl('rect', { x: cx - band / 2, y: M.t, width: band, height: ph, fill: 'transparent' });
    hit.addEventListener('mousemove', e => showTip(e, d.month, [
      ['Sum of PKg Cost', fmt(d.cost, 2)],
      ['Sum of Qty', fmt(d.qty, 2)],
      ['Qty / PKg Cost', fmt(d.ratio, 4)]
    ]));
    hit.addEventListener('mouseleave', hideTip);
    svg.appendChild(hit);
  });
  svg.appendChild(svgEl('line', { x1: M.l, x2: W - M.r, y1: y(0), y2: y(0), stroke: css('--axis'), 'stroke-width': 1 }));
}

/* Up to five groups. A legend is always present; only the top series carries
   a direct end-label, so converging lines never collide. */
function chartRatioTrend(rows) {
  const svg = $('#cTrend'), legend = $('#lTrend');
  const W = svg.clientWidth || 520, H = 260;
  const M = { t: 18, r: 74, b: 30, l: 56 };
  clear(svg); legend.innerHTML = '';
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`); svg.setAttribute('height', H);

  const months = monthsIn(rows);
  const top5 = byGroup(rows).slice(0, 5);
  if (!months.length || !top5.length) return;

  const at = new Map(rows.map(r => [r.group + '||' + r.month, r]));
  const series = top5.map((g, i) => ({
    name: g.group, colour: SERIES()[i], total: g.cost,
    pts: months.map((m, x) => {
      const c = at.get(g.group + '||' + m);
      return { x, month: m, v: c ? ratio(c.qty, c.cost) : null, qty: c ? c.qty : 0, cost: c ? c.cost : 0 };
    })
  }));

  const vals = series.flatMap(s => s.pts.map(p => p.v)).filter(v => v !== null);
  const ticks = niceTicks(Math.max(...vals, 1));
  const top = ticks[ticks.length - 1] || 1;
  const pw = W - M.l - M.r, ph = H - M.t - M.b;
  const y = v => M.t + ph - (v / top) * ph;
  const x = i => months.length === 1 ? M.l + pw / 2 : M.l + (i / (months.length - 1)) * pw;

  ticks.forEach(t => {
    svg.appendChild(svgEl('line', { x1: M.l, x2: W - M.r, y1: y(t), y2: y(t), stroke: css('--grid'), 'stroke-width': 1 }));
    svg.appendChild(axisTxt(M.l - 9, y(t) + 4, fmt(t, t < 10 ? 1 : 0), 'end'));
  });
  months.forEach((m, i) => svg.appendChild(axisTxt(x(i), H - 10, m)));

  series.forEach(s => {
    const segs = [];
    let cur = [];
    s.pts.forEach(p => { if (p.v === null) { if (cur.length) segs.push(cur); cur = []; } else cur.push(p); });
    if (cur.length) segs.push(cur);
    segs.forEach(seg => {
      if (seg.length === 1) return;
      svg.appendChild(svgEl('path', {
        d: 'M' + seg.map(p => `${x(p.x)},${y(p.v)}`).join('L'),
        fill: 'none', stroke: s.colour, 'stroke-width': 2,
        'stroke-linejoin': 'round', 'stroke-linecap': 'round'
      }));
    });
    // a 2px surface ring keeps markers legible where lines cross
    for (const p of s.pts) {
      if (p.v === null) continue;
      svg.appendChild(svgEl('circle', {
        cx: x(p.x), cy: y(p.v), r: 4, fill: s.colour,
        stroke: css('--surface-1'), 'stroke-width': 2
      }));
    }
  });

  // Direct-label exactly one series: the one sitting highest at the right
  // edge, which is the one furthest clear of the others. Labelling the lowest
  // line instead would drop the text into the converging bundle.
  const ends = series.map(s => ({ s, p: [...s.pts].reverse().find(q => q.v !== null) }))
    .filter(e => e.p);
  const lead = ends.sort((a, b) => b.p.v - a.p.v)[0];
  if (lead) {
    const t = svgEl('text', {
      x: Math.min(x(lead.p.x) + 10, W - 4), y: y(lead.p.v) + 4, class: 'axis-txt',
      fill: css('--text-secondary')
    });
    t.textContent = lead.s.name; svg.appendChild(t);
  }

  // hover column: one tooltip listing every series at that month
  months.forEach((m, i) => {
    const half = months.length === 1 ? pw / 2 : pw / (months.length - 1) / 2;
    const hit = svgEl('rect', {
      x: Math.max(M.l, x(i) - half), y: M.t,
      width: Math.min(half * 2, pw), height: ph, fill: 'transparent'
    });
    hit.addEventListener('mousemove', e => showTip(e, m,
      series.map(s => {
        const p = s.pts[i];
        return [s.name, p.v === null ? '—' : fmt(p.v, 4)];
      })));
    hit.addEventListener('mouseleave', hideTip);
    svg.appendChild(hit);
  });

  legend.innerHTML = series.map(s =>
    `<span class="lg"><i style="background:${s.colour}"></i>${s.name}</span>`).join('');
  svg.appendChild(svgEl('line', { x1: M.l, x2: M.l, y1: M.t, y2: H - M.b, stroke: css('--axis'), 'stroke-width': 1 }));
}

/* One series again: nominal categories all wear slot 1, never a value ramp. */
function chartGroups(data) {
  const svg = $('#cGroups');
  const rows = data.slice(0, 15);
  const W = svg.clientWidth || 900, rowH = 26, H = Math.max(120, rows.length * rowH + 44);
  const M = { t: 10, r: 78, b: 30, l: 132 };
  clear(svg); svg.setAttribute('viewBox', `0 0 ${W} ${H}`); svg.setAttribute('height', H);
  if (!rows.length) return;

  const ticks = niceTicks(Math.max(...rows.map(d => d.cost), 1));
  const top = ticks[ticks.length - 1] || 1;
  const pw = W - M.l - M.r;
  const x = v => M.l + (v / top) * pw;
  const bh = Math.min(24, rowH - 8);

  ticks.forEach(t => {
    svg.appendChild(svgEl('line', { x1: x(t), x2: x(t), y1: M.t, y2: H - M.b, stroke: css('--grid'), 'stroke-width': 1 }));
    svg.appendChild(axisTxt(x(t), H - 10, compact(t)));
  });

  rows.forEach((d, i) => {
    const yy = M.t + i * rowH + (rowH - bh) / 2;
    const w = Math.max(0, x(d.cost) - M.l);
    if (w > 0) svg.appendChild(svgEl('path', {
      d: barPath(M.l, yy, w, bh, 4), fill: css('--series-1')
    }));
    const lb = svgEl('text', { x: M.l - 10, y: yy + bh / 2 + 4, 'text-anchor': 'end', class: 'axis-txt' });
    lb.textContent = d.group.length > 17 ? d.group.slice(0, 16) + '…' : d.group;
    svg.appendChild(lb);
    // value at the tip, outside the bar so it can never be clipped
    svg.appendChild(axisTxt(Math.min(x(d.cost) + 8, W - 4), yy + bh / 2 + 4, compact(d.cost), 'start'));

    const hit = svgEl('rect', { x: M.l, y: M.t + i * rowH, width: pw, height: rowH, fill: 'transparent' });
    hit.addEventListener('mousemove', e => showTip(e, d.group, [
      ['Sum of PKg Cost', fmt(d.cost, 2)],
      ['Sum of Qty', fmt(d.qty, 2)],
      ['Qty / PKg Cost', fmt(d.ratio, 4)]
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

  if (tab === 'trend') {
    const months = monthsIn(rows), groups = groupsIn(rows);
    const at = new Map(rows.map(r => [r.group + '||' + r.month, r]));
    const header = ['Item Group'];
    months.forEach(m => METRICS.forEach(k => header.push(`${m} (${k})`)));
    METRICS.forEach(k => header.push(`Total (${k})`));

    const body = groups.map(g => {
      const line = [g]; let tq = 0, tc = 0;
      months.forEach(m => {
        const c = at.get(g + '||' + m);
        const q = c ? c.qty : 0, k = c ? c.cost : 0;
        tq += q; tc += k;
        line.push(q, k, ratio(q, k));
      });
      line.push(tq, tc, ratio(tq, tc));
      return line;
    });
    const grand = ['Grand Total'];
    for (let c = 1; c < header.length; c++) grand.push(body.reduce((s, r) => s + r[c], 0));
    for (let c = 1; c < header.length; c += 3) grand[c + 2] = ratio(grand[c], grand[c + 1]);
    body.push(grand);
    return { header, body, textCols: 1, totalLast: true, dec: c => (c % 3 === 1 ? 2 : 4) };
  }

  if (tab === 'months') {
    const header = ['Month', 'Month No.', ...METRICS];
    const body = byMonth(rows).map(m => [m.month, result.months.indexOf(m.month) + 1, m.qty, m.cost, m.ratio]);
    const t = totalsOf(rows);
    body.push(['Grand Total', '', t.qty, t.cost, t.ratio]);
    return { header, body, textCols: 2, totalLast: true, dec: c => (c === 2 ? 2 : 4) };
  }

  if (tab === 'groups') {
    const t = totalsOf(rows);
    const header = ['Item Group', ...METRICS, 'Share of PKg Cost'];
    const body = byGroup(rows).map(g => [g.group, g.qty, g.cost, g.ratio, t.cost ? g.cost / t.cost : 0]);
    body.push(['Grand Total', t.qty, t.cost, t.ratio, t.cost ? 1 : 0]);
    return { header, body, textCols: 1, totalLast: true, dec: c => (c === 1 ? 2 : 4), pct: 4 };
  }

  return { header: [], body: [], textCols: 0 };
}

function renderAudit() {
  const a = result.audit;
  const checks = result.checks;
  const passed = checks.filter(c => c.ok).length;
  const pill = ok => `<span class="pill ${ok ? 'pass' : 'fail'}">${ok ? '✓ Pass' : '✕ Check'}</span>`;
  const census = Object.entries(a.typeCounts).sort((x, y) => y[1] - x[1])
    .map(([k, v]) => `<tr><td class="what">${k}${k.toUpperCase() === 'FG' ? ' <b>(analysed)</b>' : ' (excluded)'}</td><td class="fig">${fmt(v)} rows</td></tr>`).join('');
  const t = totalsOf(result.long);

  return `<div class="doc">

<h4>The conditions, as applied</h4>
<p>Every number in this dashboard follows these five rules and nothing else. The exported
workbook writes each of them into the cells as live formulas.</p>
<div class="cond">
  <div class="c"><div class="n">1</div><div>
    <div class="ct">PKg Cost = Total Cost &times; PKG / 100</div>
    <div class="cd"><code>PKG</code> is a percentage held as a plain number, so <code>10.55</code>
    means 10.55%. <code>Total Cost</code> is
    ${result.hasTotalCost ? 'taken from the file' : 'derived as <code>Value In FG + Additional Cost</code>'}.
    This is <b>not</b> <code>Value In FG &times; PKG / 100</code> &mdash; that variant misses the
    reference extract by up to 25.75.</div></div></div>
  <div class="c"><div class="n">2</div><div>
    <div class="ct">Item Type = FG only</div>
    <div class="cd">Only finished-goods rows enter the analysis.
    <b>${fmt(a.fgRows)}</b> of ${fmt(result.rowCount)} rows qualify;
    ${fmt(result.rowCount - a.fgRows)} were excluded.</div></div></div>
  <div class="c"><div class="n">3</div><div>
    <div class="ct">Packaging cost per kg = SUM(Qty) &divide; SUM(PKg Cost)</div>
    <div class="cd">Kept in that order because it is what the reference pivot shows. Totals
    re-derive the ratio from the summed numerator and denominator &mdash; never an average of
    the monthly ratios.</div></div></div>
  <div class="c"><div class="n">4</div><div>
    <div class="ct">Date 1 is derived from Date</div>
    <div class="cd">Time of day stripped.</div></div></div>
  <div class="c"><div class="n">5</div><div>
    <div class="ct">Month is derived from Date, as a month name</div>
    <div class="cd">Written <code>Jan</code>, <code>Feb</code>, <code>Mar</code> &mdash; not
    <code>1</code>&ndash;<code>12</code> &mdash; and ordered by calendar position rather than
    alphabetically. Months present: <b>${result.months.join(', ')}</b>.</div></div></div>
</div>

<h4>Reconciliation &mdash; ${passed} of ${checks.length} checks pass</h4>
<table class="checks"><tbody>
${checks.map(c => `<tr>
  <td class="st">${pill(c.ok)}</td>
  <td><div class="what">${c.what}</div></td>
  <td class="fig">${c.fig}</td>
</tr>`).join('')}
</tbody></table>

<div class="two">
<div><h4>Row types seen</h4>
<table class="checks"><tbody>${census}</tbody></table></div>
<div><h4>Totals across all FG rows</h4>
<table class="checks"><tbody>
<tr><td class="what">Sum of Qty</td><td class="fig">${fmt(t.qty, 2)}</td></tr>
<tr><td class="what">Sum of PKg Cost</td><td class="fig">${fmt(t.cost, 2)}</td></tr>
<tr><td class="what">Qty / PKg Cost</td><td class="fig">${fmt(t.ratio, 4)}</td></tr>
<tr><td class="what">Sheet analysed</td><td class="fig">${result.sheetName}</td></tr>
</tbody></table></div>
</div>

<h4>What the export contains</h4>
<p><b>Raw Data</b> is the first tab &mdash; your source rows, with <code>Date 1</code>,
<code>Month</code> and <code>PKg Cost</code> added as formulas. <b>Monthly Trend</b> holds one row
per item group and three columns per month, every cell a <code>SUMIFS</code> back into Raw Data
carrying all three conditions. <b>Month Totals</b> and <b>Item Group Summary</b> read the same data
the other two ways, and <b>Logic &amp; Audit</b> restates these rules beside the control totals.
Filters on this page do not narrow the export &mdash; it always covers the whole file.</p>
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

  const { header, body, textCols, totalLast, dec, pct } = tableData();
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
      if (pct === i) return `<td class="num">${(v * 100).toFixed(1)}%</td>`;
      if (i < textCols) return `<td class="num">${fmt(v)}</td>`;
      return `<td class="num">${fmt(v, dec ? dec(i) : 2)}</td>`;
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
  ensureWorker().postMessage({ cmd: 'export', result, opts: {} });
});
function saveBook(buf) {
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = (fileName.replace(/\.[^.]+$/, '') || 'FG_Packaging') + '_Analysis.xlsx';
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
}
})();
