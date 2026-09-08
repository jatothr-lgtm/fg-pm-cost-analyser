/* ============================================================================
   FG / PM CALCULATION ENGINE  (Web Worker)
   ---------------------------------------------------------------------------
   Everything heavy happens here so the UI never blocks: xlsx parse, the
   single-pass work order build, the pivots, and the output workbook.

   PM VALUE
     PM Value(workorder) = SUM(Total Amount) WHERE Item Type = 'PKG'
     RM / FG / BiProduct rows contribute nothing.

   BUCKETING  (which Item Group + Month that PM Value is reported under)
     mode 'firstRow'  - the Workorder's FIRST row in file order. That row is the
                        primary input line (RM, or PKG when the work order has
                        no RM). Ties to the reference pivot exactly.
     mode 'fgRow'     - the Item Group on the Workorder's FG rows. Work orders
                        with zero or several distinct FG groups go unmapped.

   PM QTY
     PM Qty = SUM over FG rows of ( Value In FG x PKG / 100 )
   ========================================================================== */

importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const MONTH_NUM = MONTHS.reduce((a, m, i) => (a[m] = i + 1, a), {});
const MONTH_LOOKUP = {
  jan: 'Jan', january: 'Jan', feb: 'Feb', february: 'Feb', mar: 'Mar', march: 'Mar',
  apr: 'Apr', april: 'Apr', may: 'May', jun: 'Jun', june: 'Jun', jul: 'Jul', july: 'Jul',
  aug: 'Aug', august: 'Aug', sep: 'Sep', sept: 'Sep', september: 'Sep',
  oct: 'Oct', october: 'Oct', nov: 'Nov', november: 'Nov', dec: 'Dec', december: 'Dec'
};

const REQUIRED = ['Workorder', 'Item Group', 'Item Type', 'Qty', 'Total Amount'];
const UNKNOWN = 'Unknown';

function post(stage, pct) { self.postMessage({ type: 'progress', stage, pct }); }

const num = v => {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const n = parseFloat(String(v).replace(/,/g, '').trim());
  return isFinite(n) ? n : 0;
};
const txt = v => (v === null || v === undefined) ? '' : String(v).trim();

function monthOf(row) {
  const d = row['Date'];
  if (d instanceof Date && !isNaN(d)) return MONTHS[d.getMonth()];
  if (typeof d === 'number' && d > 0) {                     // excel serial
    const p = XLSX.SSF.parse_date_code(d);
    if (p && p.m >= 1 && p.m <= 12) return MONTHS[p.m - 1];
  }
  if (typeof d === 'string' && d.trim()) {
    const parsed = new Date(d);
    if (!isNaN(parsed)) return MONTHS[parsed.getMonth()];
  }
  const m = txt(row['Month']).toLowerCase();
  if (m && MONTH_LOOKUP[m]) return MONTH_LOOKUP[m];
  return UNKNOWN;
}

/* -------------------------------------------------------------------------
   Single pass. One record per work order, built in file order so that the
   first row seen is genuinely the work order's first row.
   ------------------------------------------------------------------------- */
function buildWorkorders(rows) {
  const map = new Map();
  const order = [];
  const audit = {
    typeCounts: {}, rowsNoWorkorder: 0, pkgRowsNoWorkorder: 0, pkgValueNoWorkorder: 0,
    unknownMonthRows: 0, fgQtyRaw: 0, pkgAmountRaw: 0, fgAmountRaw: 0
  };

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const wo = txt(r['Workorder']);

    const t = txt(r['Item Type']) || '(blank)';
    audit.typeCounts[t] = (audit.typeCounts[t] || 0) + 1;
    if (monthOf(r) === UNKNOWN) audit.unknownMonthRows++;
    const tU = t.toUpperCase();
    if (tU === 'PKG') audit.pkgAmountRaw += num(r['Total Amount']);
    if (tU === 'FG') { audit.fgQtyRaw += num(r['Qty']); audit.fgAmountRaw += num(r['Total Amount']); }

    if (!wo) {
      audit.rowsNoWorkorder++;
      if (tU === 'PKG') { audit.pkgRowsNoWorkorder++; audit.pkgValueNoWorkorder += num(r['Total Amount']); }
      continue;
    }

    let rec = map.get(wo);
    if (!rec) {
      rec = {
        wo,
        firstGroup: txt(r['Item Group']) || UNKNOWN,   // <- first row wins
        firstMonth: monthOf(r),                        // <- first row wins
        firstType: txt(r['Item Type']),
        targetWh: '', sourceWh: '', anyWh: '',
        pmValue: 0, pmQty: 0, fgQty: 0, fgValue: 0,
        pkgLines: 0, fgLines: 0,
        fgGroups: [],
        fgCells: new Map(),   // group|month|targetWh|sourceWh -> {fgQty, pmQty}
        firstRowIndex: i
      };
      map.set(wo, rec);
      order.push(wo);
    }

    const type = txt(r['Item Type']).toUpperCase();
    const qty = num(r['Qty']);
    const amt = num(r['Total Amount']);
    const tWh = txt(r['Target Warehouse']);
    const sWh = txt(r['Source Warehouse']);
    if (!rec.anyWh && (tWh || sWh)) rec.anyWh = tWh || sWh;

    if (type === 'PKG') {
      rec.pmValue += amt;                       // <- PM Value: PKG rows only
      rec.pkgLines++;
      // Source Warehouse lives on the input lines, not the FG line.
      if (!rec.sourceWh && sWh) rec.sourceWh = sWh;
    } else if (type === 'FG') {
      rec.fgQty += qty;
      rec.fgValue += amt;
      rec.fgLines++;
      const g = txt(r['Item Group']) || UNKNOWN;
      if (!rec.fgGroups.includes(g)) rec.fgGroups.push(g);
      if (!rec.targetWh && tWh) rec.targetWh = tWh;

      const pmQty = num(r['Value In FG']) * num(r['PKG']) / 100;
      const key = g + '' + monthOf(r);
      const cell = rec.fgCells.get(key);
      if (cell) { cell.fgQty += qty; cell.pmQty += pmQty; }
      else rec.fgCells.set(key, { group: g, month: monthOf(r), fgQty: qty, pmQty });
    } else {
      if (!rec.sourceWh && sWh) rec.sourceWh = sWh;   // RM lines carry it too
    }
  }

  const recs = order.map(w => {
    const rec = map.get(w);
    rec.targetWh = rec.targetWh || rec.anyWh || UNKNOWN;
    rec.sourceWh = rec.sourceWh || rec.anyWh || UNKNOWN;
    rec.fgGroupCount = rec.fgGroups.length;
    rec.pmQtyTotal = 0;
    rec.fgCells.forEach(c => { rec.pmQtyTotal += c.pmQty; });
    return rec;
  });

  // where each work order's bucket came from
  audit.bucketSource = {};
  recs.forEach(r => {
    const k = (r.firstType || '(blank)').toUpperCase();
    audit.bucketSource[k] = (audit.bucketSource[k] || 0) + 1;
  });
  audit.woWithPmValue = recs.filter(r => r.pmValue > 0).length;
  audit.woNegativePm = recs.filter(r => r.pmValue < 0).length;

  return { recs, audit };
}

/* -------------------------------------------------------------------------
   Long-form rows: one per (bucket group, month, target WH, source WH).
   PM Value is attached to exactly ONE cell per work order, so summing the
   long form along any axis can never duplicate it.
   ------------------------------------------------------------------------- */
function toLong(recs, mode) {
  const out = [];
  let unmappedValue = 0, unmappedCount = 0;

  for (const rec of recs) {
    if (mode === 'firstRow') {
      out.push({
        wo: rec.wo, group: rec.firstGroup, month: rec.firstMonth,
        targetWh: rec.targetWh, sourceWh: rec.sourceWh,
        fgQty: rec.fgQty, pmQty: rec.pmQtyTotal, pmValue: rec.pmValue,
        carriesValue: true
      });
      continue;
    }

    // mode === 'fgRow'
    const mappable = rec.fgGroupCount === 1;
    if (!mappable) { unmappedValue += rec.pmValue; if (rec.pmValue) unmappedCount++; }

    const cells = [...rec.fgCells.values()];
    if (!cells.length) continue;                    // no FG rows -> nothing to bucket

    // the largest FG cell carries the whole PM Value; the rest carry zero
    let carrier = 0;
    for (let i = 1; i < cells.length; i++) if (cells[i].fgQty > cells[carrier].fgQty) carrier = i;

    cells.forEach((c, i) => out.push({
      wo: rec.wo, group: c.group, month: c.month,
      targetWh: rec.targetWh, sourceWh: rec.sourceWh,
      fgQty: c.fgQty, pmQty: c.pmQty,
      pmValue: (mappable && i === carrier) ? rec.pmValue : 0,
      carriesValue: mappable && i === carrier
    }));
  }
  return { long: out, unmappedValue, unmappedCount };
}

/* ------------------------------- pivots ---------------------------------- */
function monthsIn(long) {
  const seen = new Set(long.map(r => r.month));
  const known = MONTHS.filter(m => seen.has(m));
  return seen.has(UNKNOWN) ? known.concat(UNKNOWN) : known;
}

function pivot(long, whKey, whLabel, months, denom) {
  const rows = new Map();
  for (const r of long) {
    const key = r[whKey] + '' + r.group;
    let row = rows.get(key);
    if (!row) { row = { wh: r[whKey], group: r.group, cells: {} }; rows.set(key, row); }
    const c = row.cells[r.month] || (row.cells[r.month] = { fgQty: 0, pmQty: 0, pmValue: 0 });
    c.fgQty += r.fgQty; c.pmQty += r.pmQty; c.pmValue += r.pmValue;
  }

  const header = [whLabel, 'FG Item Group'];
  months.forEach(m => header.push(`${m} (FG Qty)`, `${m} (PM Qty)`, `${m} (PM Value)`, `${m} (PM Cost/KG)`));
  header.push('Total (FG Qty)', 'Total (PM Qty)', 'Total (PM Value)', 'Total (PM Cost/KG)');

  const body = [...rows.values()]
    .sort((a, b) => (a.wh + a.group).localeCompare(b.wh + b.group))
    .map(row => {
      const line = [row.wh, row.group];
      let tq = 0, tp = 0, tv = 0;
      months.forEach(m => {
        const c = row.cells[m] || { fgQty: 0, pmQty: 0, pmValue: 0 };
        const d = denom === 'fgQty' ? c.fgQty : c.pmQty;
        line.push(c.fgQty, c.pmQty, c.pmValue, d ? c.pmValue / d : 0);
        tq += c.fgQty; tp += c.pmQty; tv += c.pmValue;
      });
      const td = denom === 'fgQty' ? tq : tp;
      line.push(tq, tp, tv, td ? tv / td : 0);
      return line;
    });

  // grand total row, ratio re-derived from the summed numerator/denominator
  if (body.length) {
    const g = ['Grand Total', ''];
    for (let c = 2; c < header.length; c++) g.push(body.reduce((s, r) => s + r[c], 0));
    const stride = 4;
    for (let i = 0; i < months.length; i++) {
      const base = 2 + i * stride;
      const d = denom === 'fgQty' ? g[base] : g[base + 1];
      g[base + 3] = d ? g[base + 2] / d : 0;
    }
    const tBase = 2 + months.length * stride;
    const td = denom === 'fgQty' ? g[tBase] : g[tBase + 1];
    g[tBase + 3] = td ? g[tBase + 2] / td : 0;
    body.push(g);
  }
  return { header, body };
}

/* ------------------------------ main run --------------------------------- */
function run(buffer, opts) {
  post('Reading workbook', 8);
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true, dense: true });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];

  post('Extracting rows', 22);
  let rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });

  // trim header whitespace
  if (rows.length) {
    const keys = Object.keys(rows[0]);
    const dirty = keys.filter(k => k !== k.trim());
    if (dirty.length) {
      rows = rows.map(r => {
        const o = {};
        for (const k in r) o[k.trim()] = r[k];
        return o;
      });
    }
  }
  if (!rows.length) throw new Error('The first sheet has no data rows.');

  const columns = Object.keys(rows[0]);
  const missing = REQUIRED.filter(c => !columns.includes(c));
  if (missing.length) {
    throw new Error(`Missing required column(s): ${missing.join(', ')}.\n\nColumns found: ${columns.join(', ')}`);
  }

  post('Building work orders', 45);
  const { recs, audit } = buildWorkorders(rows);

  post('Bucketing PM Value', 62);
  const { long, unmappedValue, unmappedCount } = toLong(recs, opts.mode);
  const months = monthsIn(long);
  audit.cellsCarryingValue = long.filter(r => r.carriesValue && r.pmValue !== 0).length;
  audit.longRows = long.length;

  post('Pivoting', 78);
  const target = pivot(long, 'targetWh', 'Target Warehouse', months, opts.denom);
  const source = pivot(long, 'sourceWh', 'Source Warehouse', months, opts.denom);

  const sourcePkg = recs.reduce((s, r) => s + r.pmValue, 0);
  const allocated = long.reduce((s, r) => s + r.pmValue, 0);

  const multiFg = recs.filter(r => r.fgGroupCount > 1).map(r => ({
    wo: r.wo, month: r.firstMonth, firstGroup: r.firstGroup,
    fgGroups: r.fgGroups.join(' | '), count: r.fgGroupCount,
    pmValue: r.pmValue, fgQty: r.fgQty
  }));
  const noFg = recs.filter(r => r.fgGroupCount === 0);

  post('Done', 95);
  return {
    sheetName, rowCount: rows.length, columns, audit,
    months,
    workorders: recs.map(r => ({
      wo: r.wo, firstGroup: r.firstGroup, firstType: r.firstType, month: r.firstMonth,
      fgGroups: r.fgGroups.join(' | '), fgGroupCount: r.fgGroupCount,
      targetWh: r.targetWh, sourceWh: r.sourceWh,
      fgQty: r.fgQty, fgValue: r.fgValue, pmQty: r.pmQtyTotal, pmValue: r.pmValue,
      pkgLines: r.pkgLines, fgLines: r.fgLines
    })),
    long, target, source, multiFg,
    stats: {
      sourcePkg, allocated, difference: sourcePkg - allocated,
      unmappedValue, unmappedCount,
      workorderCount: recs.length,
      multiFgCount: multiFg.length,
      multiFgValue: multiFg.reduce((s, r) => s + r.pmValue, 0),
      noFgCount: noFg.length,
      noFgValue: noFg.reduce((s, r) => s + r.pmValue, 0),
      totalFgQty: long.reduce((s, r) => s + r.fgQty, 0),
      totalPmQty: long.reduce((s, r) => s + r.pmQty, 0)
    }
  };
}

/* --------------------------- output workbook -----------------------------
   Derived cells are written as live Excel formulas with a cached value, so the
   workbook recalculates and every figure can be traced back to its inputs by
   clicking the cell. Only the raw sums (FG Qty, PM Qty, PM Value per month) are
   literal numbers - everything downstream of them is a formula.
   ------------------------------------------------------------------------ */
const A1 = (r, c) => XLSX.utils.encode_cell({ r, c });           // 0-based
const COL = c => XLSX.utils.encode_col(c);

function setFormula(ws, r, c, f, cached, z) {
  ws[A1(r, c)] = { t: 'n', f, v: isFinite(cached) ? cached : 0, z: z || '#,##0.00' };
}
function setFmt(ws, r, c, z) { const cell = ws[A1(r, c)]; if (cell && cell.t === 'n') cell.z = z; }

function buildWorkbook(res, opts, checks) {
  const wb = XLSX.utils.book_new();
  const widths = n => Array.from({ length: n }, () => ({ wch: 16 }));

  const addAoa = (name, header, body, firstWide) => {
    const ws = XLSX.utils.aoa_to_sheet([header, ...body]);
    const cols = widths(header.length);
    if (firstWide) { cols[0] = { wch: 42 }; if (cols[1]) cols[1] = { wch: 26 }; }
    ws['!cols'] = cols;
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: body.length, c: header.length - 1 } }) };
    // Column widths, number formats and autofilter are what the browser build of
    // SheetJS writes; frozen panes and bold headers need the pro build.
    XLSX.utils.book_append_sheet(wb, ws, name);
    return ws;
  };

  /* Monthly pivot: cost-per-kg, the Total block and the Grand Total row all
     become formulas over the month cells beside them. */
  const formulaiseMonthly = (ws, months, body) => {
    const n = body.length;                 // includes the Grand Total row
    const dataLast = n;                    // sheet row (1-based) of last data row
    const grandR = n;                      // 0-based row index of Grand Total
    const dOff = opts.denom === 'fgQty' ? 0 : 1;   // FG Qty or PM Qty column
    const base = m => 2 + m * 4;
    const tBase = 2 + months.length * 4;

    for (let i = 0; i < n; i++) {
      const r = i + 1, row = r + 1;        // 0-based row index, 1-based sheet row
      const isGrand = i === n - 1;         // the Grand Total row sums the rows above

      months.forEach((m, mi) => {
        const b = base(mi);
        if (isGrand) {
          for (let k = 0; k < 3; k++) {
            const L = COL(b + k);
            setFormula(ws, r, b + k, `SUM(${L}2:${L}${dataLast})`, body[i][b + k], k === 2 ? '#,##0' : '#,##0.00');
          }
        } else {
          setFmt(ws, r, b, '#,##0.00'); setFmt(ws, r, b + 1, '#,##0.00'); setFmt(ws, r, b + 2, '#,##0');
        }
        // cost per kg = PM Value / denominator, guarded against divide-by-zero
        const vC = COL(b + 2), dC = COL(b + dOff);
        setFormula(ws, r, b + 3, `IF(${dC}${row}=0,0,${vC}${row}/${dC}${row})`, body[i][b + 3]);
      });

      // Total (FG Qty | PM Qty | PM Value) = sum across the month columns
      for (let k = 0; k < 3; k++) {
        const refs = months.map((_, mi) => `${COL(base(mi) + k)}${row}`).join(',');
        setFormula(ws, r, tBase + k, `SUM(${refs})`, body[i][tBase + k], k === 2 ? '#,##0' : '#,##0.00');
      }
      const tV = COL(tBase + 2), tD = COL(tBase + dOff);
      setFormula(ws, r, tBase + 3, `IF(${tD}${row}=0,0,${tV}${row}/${tD}${row})`, body[i][tBase + 3]);
    }
  };

  const wsT = addAoa('Monthly Summary', res.target.header, res.target.body, true);
  if (res.target.body.length) formulaiseMonthly(wsT, res.months, res.target.body);
  const wsS = addAoa('Monthly Summary - Source WH', res.source.header, res.source.body, true);
  if (res.source.body.length) formulaiseMonthly(wsS, res.months, res.source.body);

  /* Workorder Summary - PM Cost/KG as a formula on every row */
  const woHeader = ['Workorder', 'Bucket Item Group', 'First Row Item Type', 'Month',
    'FG Item Groups', 'FG Item Group Count', 'Target Warehouse', 'Source Warehouse',
    'FG Qty', 'FG Value', 'PM Qty', 'PM Value', 'PM Cost/KG', 'PKG Lines', 'FG Lines'];
  const woBody = res.workorders.map(w => {
    const d = opts.denom === 'fgQty' ? w.fgQty : w.pmQty;
    return [w.wo, w.firstGroup, w.firstType, w.month, w.fgGroups, w.fgGroupCount,
      w.targetWh, w.sourceWh, w.fgQty, w.fgValue, w.pmQty, w.pmValue,
      d ? w.pmValue / d : 0, w.pkgLines, w.fgLines];
  });
  const wsW = addAoa('Workorder Summary', woHeader, woBody, false);
  {
    const dC = opts.denom === 'fgQty' ? COL(8) : COL(10);   // I = FG Qty, K = PM Qty
    for (let i = 0; i < woBody.length; i++) {
      const r = i + 1, row = r + 1;
      setFmt(wsW, r, 8, '#,##0.00'); setFmt(wsW, r, 9, '#,##0');
      setFmt(wsW, r, 10, '#,##0.00'); setFmt(wsW, r, 11, '#,##0');
      setFormula(wsW, r, 12, `IF(${dC}${row}=0,0,${COL(11)}${row}/${dC}${row})`, woBody[i][12]);
    }
  }

  /* Logic & Audit - the reconciliation, with the control total as a formula */
  const s = res.stats, a = res.audit;
  const auditRows = [
    ['SETTINGS', '', ''],
    ['PM Value bucket', opts.mode === 'firstRow' ? 'Work order first row' : 'FG row Item Group', ''],
    ['PM Cost/KG denominator', opts.denom === 'fgQty' ? 'FG Qty' : 'PM Qty', ''],
    ['Source sheet', res.sheetName, ''],
    ['Rows read', res.rowCount, ''],
    ['Work orders', s.workorderCount, ''],
    ['', '', ''],
    ['CONTROL TOTAL', '', ''],
    ['Source PKG Total Amount', s.sourcePkg, 'SUM(Total Amount) WHERE Item Type = PKG'],
    ['PM Value allocated to buckets', s.allocated, 'Summed over the Monthly Summary'],
    ['Unmapped PM Value', s.unmappedValue, 'Work orders whose item group could not be resolved'],
    ['Difference (must be 0)', 0, 'Source - allocated - unmapped'],
    ['', '', ''],
    ['POPULATION', '', ''],
    ['Work orders with >1 FG item group', s.multiFgCount, `PM Value ${Math.round(s.multiFgValue)}`],
    ['Work orders with no FG line', s.noFgCount, `PM Value ${Math.round(s.noFgValue)}`],
    ['PKG rows with no work order', a.pkgRowsNoWorkorder, 'Cannot be attributed to any item group'],
    ['Rows with an unreadable month', a.unknownMonthRows, 'Reported under "Unknown"'],
    ['', '', ''],
    ['ROW CENSUS', '', ''],
    ...Object.entries(a.typeCounts).sort((x, y) => y[1] - x[1]).map(([k, v]) => [k, v, 'rows']),
    ['', '', ''],
    ['BUCKET PROVENANCE', '', ''],
    ...Object.entries(a.bucketSource).sort((x, y) => y[1] - x[1])
      .map(([k, v]) => [`First row is ${k}`, v, 'work orders']),
    ['', '', ''],
    ['CHECKS', '', ''],
    ...(checks || []).map(c => [c.what, c.ok ? 'PASS' : 'CHECK', c.fig]),
    ['', '', ''],
    ['LOGIC APPLIED', '', ''],
    ['1. Read first sheet in file order', 'Row order is load-bearing for step 3', ''],
    ['2. PM Value', "SUM(Total Amount) WHERE Item Type = 'PKG', by Workorder", 'RM / FG / BiProduct excluded'],
    ['3. Bucket', opts.mode === 'firstRow'
      ? "Item Group + Month of the work order's FIRST row in file order (its input line)"
      : "Item Group on the work order's FG rows; unresolved when 0 or >1 distinct groups",
      'VLOOKUP(Workorder, data, Item Group)'],
    ['4. One carrier per work order', 'PM Value attached to the single largest FG cell', 'Prevents double counting'],
    ['5. PM Qty', 'SUM over FG rows of ( Value In FG x PKG / 100 )', 'PKG is a percent held as a number'],
    ['5. FG Qty', 'SUM over FG rows of Qty', ''],
    ['6. PM Cost/KG', `PM Value / ${opts.denom === 'fgQty' ? 'FG Qty' : 'PM Qty'}`, 'Totals re-derive the ratio, never average it'],
    ['7. Target Warehouse', 'From the FG rows', ''],
    ['7. Source Warehouse', 'From the input (RM / PKG) rows', 'FG rows carry it blank'],
    ['8. Months', 'From Date, else a text Month column', 'Calendar order, months present only']
  ];
  const wsA = addAoa('Logic & Audit', ['Item', 'Value', 'Note'], auditRows, false);
  wsA['!cols'] = [{ wch: 46 }, { wch: 62 }, { wch: 52 }];
  {
    // difference stays a live formula so the workbook proves itself on open
    const r = auditRows.findIndex(x => x[0] === 'Difference (must be 0)') + 1;
    const srcR = auditRows.findIndex(x => x[0] === 'Source PKG Total Amount') + 2;
    const allR = auditRows.findIndex(x => x[0] === 'PM Value allocated to buckets') + 2;
    const unmR = auditRows.findIndex(x => x[0] === 'Unmapped PM Value') + 2;
    setFormula(wsA, r, 1, `B${srcR}-B${allR}-B${unmR}`, 0, '#,##0');
  }

  if (res.multiFg.length) {
    addAoa('Multiple FG Groups',
      ['Workorder', 'Month', 'Bucket Item Group', 'FG Item Groups', 'Count', 'PM Value', 'FG Qty'],
      res.multiFg.map(r => [r.wo, r.month, r.firstGroup, r.fgGroups, r.count, r.pmValue, r.fgQty]),
      false);
  }
  return XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
}

self.onmessage = e => {
  const { cmd } = e.data;
  try {
    if (cmd === 'process') {
      const res = run(e.data.buffer, e.data.opts);
      self.postMessage({ type: 'result', result: res });
    } else if (cmd === 'export') {
      const buf = buildWorkbook(e.data.result, e.data.opts, e.data.checks);
      self.postMessage({ type: 'export', buffer: buf }, [buf]);
    }
  } catch (err) {
    self.postMessage({ type: 'error', message: err && err.message ? err.message : String(err) });
  }
};
