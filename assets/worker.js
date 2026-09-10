/* ============================================================================
   FG PACKAGING COST ENGINE  (Web Worker)
   ---------------------------------------------------------------------------
   Everything heavy happens here so the UI never blocks: xlsx parse, the FG
   aggregation, the month-wise pivot, and the output workbook.

   THE CONDITIONS  (all three are applied exactly as written)

     1. PKg Cost        = Total Amount x PKG / 100
                          Total Amount (or Total Cost) is multiplied by PKG %,
                          where PKG is a percentage held as a plain number, so 10.55
                          means 10.55%. Formulated as =N2*U2/100.

     2. Item Type = FG  Only finished-goods rows enter the analysis. RM, PKG,
                        BiProduct and anything else are excluded outright.

     3. Packaging cost per kg
                        = SUM(PKg Cost) / SUM(Qty)
                          Calculated as total PKg Cost divided by total Qty.
                          Totals re-derive the ratio from the summed numerator
                          and denominator - never an average of the monthly ratios.

   DERIVED COLUMNS      Date 1 and Month both come from Date. Month is written
                        as a real month name (Jan, Feb, Mar ...), never 1-12,
                        and is ordered by calendar position rather than
                        alphabetically.
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

const UNKNOWN = 'Unknown';
const FG = 'FG';

// what a sheet must have before it can be analysed
const NEEDED = ['Item Group', 'Item Type', 'Qty', 'PKG'];

function post(stage, pct) { self.postMessage({ type: 'progress', stage, pct }); }

let lastRows = null, lastColumns = null, lastMeta = null;   // reused by the exporter

// signals that the source will not fit alongside the rest of the workbook
class RawTooBig extends Error {}

const num = v => {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const n = parseFloat(String(v).replace(/,/g, '').trim());
  return isFinite(n) ? n : 0;
};
const txt = v => (v === null || v === undefined) ? '' : String(v).trim();

/* ------------------------------------------------------------------ months */
function monthOf(row) {
  const d = row['Date'];
  if (d instanceof Date && !isNaN(d)) return MONTHS[d.getMonth()];
  if (typeof d === 'number' && d > 0) {                       // excel serial
    const p = XLSX.SSF.parse_date_code(d);
    if (p && p.m >= 1 && p.m <= 12) return MONTHS[p.m - 1];
  }
  if (typeof d === 'string' && d.trim()) {
    const parsed = new Date(d);
    if (!isNaN(parsed)) return MONTHS[parsed.getMonth()];
  }
  // fall back to a Month column, which may hold 1-12 or a name
  const m = row['Month'];
  if (typeof m === 'number' && m >= 1 && m <= 12) return MONTHS[m - 1];
  const s = txt(m).toLowerCase();
  if (s && MONTH_LOOKUP[s]) return MONTH_LOOKUP[s];
  const asNum = parseInt(s, 10);
  if (asNum >= 1 && asNum <= 12) return MONTHS[asNum - 1];
  return UNKNOWN;
}

function dateOnly(row) {
  const d = row['Date'];
  if (d instanceof Date && !isNaN(d)) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  if (typeof d === 'number' && d > 0) {
    const p = XLSX.SSF.parse_date_code(d);
    if (p) return new Date(p.y, p.m - 1, p.d);
  }
  if (typeof d === 'string' && d.trim()) {
    const p = new Date(d);
    if (!isNaN(p)) return new Date(p.getFullYear(), p.getMonth(), p.getDate());
  }
  return null;
}

/* ------------------------------------------------------------------- costs */
// Total Amount is used as given, and reconstructed when omitted.
const totalAmountOf = (row, hasTotalAmount) => {
  if (hasTotalAmount) return num(row['Total Amount']);
  if (row['Amount'] !== undefined && row['Amount'] !== null) return num(row['Amount']);
  if (row['Total Cost'] !== undefined && row['Total Cost'] !== null) return num(row['Total Cost']);
  return num(row['Value In FG']) + num(row['Additional Cost']);
};

// CONDITION 1
const pkgCostOf = (row, hasTotalAmount) => totalAmountOf(row, hasTotalAmount) * num(row['PKG']) / 100;

/* -------------------------------------------------------------------------
   Aggregate. One pass over the rows, FG only, keyed by Item Group + Month +
   Target Warehouse. The warehouse only makes the grain finer - every total
   downstream sums the same cells, so no existing figure moves.
   ------------------------------------------------------------------------- */
function aggregate(rows, hasTotalAmount) {
  const cells = new Map();                 // "group|month|warehouse" -> cell
  const biCells = new Map();               // BiProduct: "group|month|warehouse" -> cell
  const audit = {
    typeCounts: {}, fgRows: 0, biRows: 0, unknownMonthRows: 0, blankPkgRows: 0,
    negativeCostRows: 0, qtyTotal: 0, costTotal: 0,
    sourcePkgCostTotal: 0, recomputeMaxDiff: 0, hasSourcePkgCost: false
  };

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const rawType = txt(r['Item Type']) || '(blank)';
    const type = rawType.toUpperCase().replace(/[\s_-]/g, '');
    audit.typeCounts[rawType] = (audit.typeCounts[rawType] || 0) + 1;

    const month = monthOf(r);
    const group = txt(r['Item Group']) || UNKNOWN;
    const qty = num(r['Qty']);
    const cost = pkgCostOf(r, hasTotalAmount);
    const wh = txt(r['Target Warehouse']) || UNKNOWN;

    if (type === FG) {
      audit.fgRows++;
      if (month === UNKNOWN) audit.unknownMonthRows++;
      if (!num(r['PKG'])) audit.blankPkgRows++;
      if (cost < 0) audit.negativeCostRows++;
      audit.qtyTotal += qty;
      audit.costTotal += cost;

      if (r['PKg Cost'] !== undefined && r['PKg Cost'] !== null && r['PKg Cost'] !== '') {
        audit.hasSourcePkgCost = true;
        const given = num(r['PKg Cost']);
        audit.sourcePkgCostTotal += given;
        const diff = Math.abs(given - cost);
        if (diff > audit.recomputeMaxDiff) audit.recomputeMaxDiff = diff;
      }

      const key = group + '||' + month + '||' + wh;
      let c = cells.get(key);
      if (!c) cells.set(key, c = { group, month, monthNum: MONTH_NUM[month] || 99, targetWh: wh, qty: 0, cost: 0, rows: 0 });
      c.qty += qty; c.cost += cost; c.rows++;
    } else if (type === 'BIPRODUCT') {
      audit.biRows++;
      const key = group + '||' + month + '||' + wh;
      let c = biCells.get(key);
      if (!c) biCells.set(key, c = { group, month, monthNum: MONTH_NUM[month] || 99, targetWh: wh, qty: 0, cost: 0, rows: 0 });
      c.qty += qty; c.cost += cost; c.rows++;
    }
  }

  return { long: [...cells.values()], biLong: [...biCells.values()], audit };
}

/* -------------------------------------------------------------------------
   Month-wise pivot: one row per Item Group, three columns per month.
   ------------------------------------------------------------------------- */
// CONDITION 3
const ratio = (qty, cost) => qty ? cost / qty : 0;

const METRICS = ['Sum of Qty', 'Sum of PKg Cost', 'PKg Cost / Qty'];

/* One pivot builder for both tabs. `idxFields` names the row dimensions, so
   ['group'] gives the original Item Group trend and ['targetWh','group'] gives
   the warehouse-wise one. Cells are summed across whatever dimension is not in
   the index, which is why adding the warehouse grain moved no existing total. */
function buildPivot(long, months, idxFields, idxLabels) {
  const agg = new Map();                 // index tuple + month -> running cell
  const tuples = new Map();
  long.forEach(c => {
    const tuple = idxFields.map(f => c[f]);
    const ik = tuple.join('||');
    if (!tuples.has(ik)) tuples.set(ik, tuple);
    const k = ik + '||' + c.month;
    let o = agg.get(k);
    if (!o) agg.set(k, o = { qty: 0, cost: 0 });
    o.qty += c.qty; o.cost += c.cost;
  });

  const header = idxLabels.slice();
  months.forEach(m => METRICS.forEach(k => header.push(`${m} (${k})`)));
  METRICS.forEach(k => header.push(`Total (${k})`));

  const nIdx = idxFields.length;
  const rowsOut = [...tuples.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  const body = rowsOut.map(([ik, tuple]) => {
    const row = tuple.slice();
    let tq = 0, tc = 0;
    months.forEach(m => {
      const o = agg.get(ik + '||' + m) || { qty: 0, cost: 0 };
      tq += o.qty; tc += o.cost;
      row.push(o.qty, o.cost, ratio(o.qty, o.cost));
    });
    row.push(tq, tc, ratio(tq, tc));
    return row;
  });

  // Grand Total: the ratio is re-derived from the summed columns
  const grand = ['Grand Total'].concat(new Array(nIdx - 1).fill(''));
  for (let c = nIdx; c < header.length; c++) grand.push(body.reduce((sum, r) => sum + r[c], 0));
  for (let c = nIdx; c < header.length; c += 3) grand[c + 2] = ratio(grand[c], grand[c + 1]);
  body.push(grand);

  return { header, body, nIdx };
}

/* Dedicated pivot for BiProduct: only Sum of Qty */
function buildQtyPivot(long, months, idxFields, idxLabels) {
  const agg = new Map();
  const tuples = new Map();
  long.forEach(c => {
    const tuple = idxFields.map(f => c[f]);
    const ik = tuple.join('||');
    if (!tuples.has(ik)) tuples.set(ik, tuple);
    const k = ik + '||' + c.month;
    let o = agg.get(k);
    if (!o) agg.set(k, o = { qty: 0 });
    o.qty += c.qty;
  });

  const header = idxLabels.concat(months.map(m => `${m} (Sum of Qty)`), ['Total (Sum of Qty)']);
  const nIdx = idxFields.length;
  const rowsOut = [...tuples.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  const body = rowsOut.map(([ik, tuple]) => {
    const row = tuple.slice();
    let tq = 0;
    months.forEach(m => {
      const o = agg.get(ik + '||' + m) || { qty: 0 };
      tq += o.qty;
      row.push(o.qty);
    });
    row.push(tq);
    return row;
  });

  const grand = ['Grand Total'].concat(new Array(nIdx - 1).fill(''));
  for (let c = nIdx; c < header.length; c++) grand.push(body.reduce((sum, r) => sum + r[c], 0));
  body.push(grand);

  return { header, body, nIdx };
}

/* -------------------------------------------------------------------- load */
/* Score every sheet so the right one is chosen even in a workbook full of
   working tabs. Total Cost is weighted heavily on purpose: it is the column
   that separates the analysis extract from a raw stock-entry dump, and the
   two produce very different answers. Row count only breaks ties. */
function scoreSheets(wb) {
  const out = [];
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (!ws || !ws['!ref']) continue;
    const head = (XLSX.utils.sheet_to_json(ws, { header: 1, range: 0, blankrows: false })[0] || [])
      .map(h => txt(h));
    const have = NEEDED.filter(n => head.includes(n));
    let score = have.length * 10;
    if (head.includes('Total Amount')) score += 6;
    else if (head.includes('Total Cost')) score += 5;
    else if (head.includes('Amount')) score += 4;
    else if (head.includes('Value In FG') && head.includes('Additional Cost')) score += 1;
    if (head.includes('Date')) score += 2;
    else if (head.includes('Month')) score += 1;
    if (head.includes('PKg Cost')) score += 1;              // already carries the answer
    out.push({
      name, score,
      rows: XLSX.utils.decode_range(ws['!ref']).e.r,
      usable: have.length === NEEDED.length &&
              (head.includes('Total Amount') || head.includes('Total Cost') || head.includes('Amount') ||
               (head.includes('Value In FG') && head.includes('Additional Cost'))) &&
              (head.includes('Date') || head.includes('Month'))
    });
  }
  out.sort((a, b) => b.score - a.score || b.rows - a.rows);
  return out;
}

function run(buffer, opts) {
  post('Reading workbook', 8);
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true, dense: false });
  const sheets = scoreSheets(wb);
  const chosen = (opts.sheet && wb.SheetNames.includes(opts.sheet))
    ? opts.sheet
    : (sheets[0] ? sheets[0].name : wb.SheetNames[0]);
  const picked = { name: chosen };
  const ws = wb.Sheets[picked.name];

  post('Parsing rows', 26);
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
  if (!rows.length) throw new Error(`Sheet "${picked.name}" has no data rows.`);

  const columns = Object.keys(rows[0]).map(c => txt(c));
  // normalise keys once so lookups never trip over stray spaces
  const rawKeys = Object.keys(rows[0]);
  const needsTrim = rawKeys.some(k => k !== k.trim());
  if (needsTrim) {
    for (const r of rows) for (const k of rawKeys) {
      const t = k.trim();
      if (t !== k) { r[t] = r[k]; delete r[k]; }
    }
  }

  const missing = NEEDED.filter(n => !columns.includes(n));
  if (missing.length) {
    throw new Error(
      `Sheet "${picked.name}" is missing required column(s): ${missing.join(', ')}.\n\n` +
      `Columns found:\n${columns.join(', ')}`);
  }
  const hasTotalAmount = columns.includes('Total Amount');
  const hasTotalCost = columns.includes('Total Cost');
  const hasAmount = columns.includes('Amount');
  if (!hasTotalAmount && !hasTotalCost && !hasAmount && !(columns.includes('Value In FG') && columns.includes('Additional Cost'))) {
    throw new Error(
      'Need either a "Total Amount" (or "Total Cost") column, or both "Value In FG" and "Additional Cost" ' +
      'so total amount can be derived.\n\n' + `Columns found:\n${columns.join(', ')}`);
  }
  if (!columns.includes('Date') && !columns.includes('Month')) {
    throw new Error('Need a "Date" column, or a "Month" column, to place rows on the calendar.');
  }

  post('Applying FG filter and PKg Cost', 52);
  const { long, biLong, audit } = aggregate(rows, hasTotalAmount);
  if (!audit.fgRows) throw new Error('No rows with Item Type = FG were found, so there is nothing to analyse.');

  const months = MONTHS.filter(m => long.some(c => c.month === m) || biLong.some(c => c.month === m))
    .concat(long.some(c => c.month === UNKNOWN) || biLong.some(c => c.month === UNKNOWN) ? [UNKNOWN] : []);
  const groups = [...new Set(long.map(c => c.group))].sort((a, b) => a.localeCompare(b));
  const warehouses = [...new Set(long.map(c => c.targetWh))].sort((a, b) => a.localeCompare(b));

  post('Building month-wise trend', 74);
  const trend = buildPivot(long, months, ['group'], ['Item Group']);
  // the added tab: same three metrics, one row per warehouse and item group
  const whTrend = buildPivot(long, months, ['targetWh', 'group'],
    ['Target Warehouse', 'Item Group']);
  const biWhTrend = buildQtyPivot(biLong, months, ['targetWh', 'group'],
    ['Target Warehouse', 'Item Group']);

  const monthTotals = months.map(m => {
    const cs = long.filter(c => c.month === m);
    const qty = cs.reduce((s, c) => s + c.qty, 0), cost = cs.reduce((s, c) => s + c.cost, 0);
    return { month: m, monthNum: MONTH_NUM[m] || 99, qty, cost, ratio: ratio(qty, cost) };
  });
  const groupTotals = groups.map(g => {
    const cs = long.filter(c => c.group === g);
    const qty = cs.reduce((s, c) => s + c.qty, 0), cost = cs.reduce((s, c) => s + c.cost, 0);
    return { group: g, qty, cost, ratio: ratio(qty, cost) };
  }).sort((a, b) => b.cost - a.cost);

  const checks = [
    { what: 'Rows read', ok: true, fig: rows.length.toLocaleString('en-IN') },
    { what: 'FG rows analysed (Item Type = FG)', ok: audit.fgRows > 0, fig: audit.fgRows.toLocaleString('en-IN') },
    { what: 'Non-FG rows excluded', ok: true, fig: (rows.length - audit.fgRows).toLocaleString('en-IN') },
    { what: 'Rows with an unreadable month', ok: audit.unknownMonthRows === 0, fig: audit.unknownMonthRows },
    { what: 'FG rows with blank or zero PKG %', ok: true, fig: audit.blankPkgRows },
    { what: 'FG rows with negative PKg Cost', ok: audit.negativeCostRows === 0, fig: audit.negativeCostRows }
  ];
  if (audit.hasSourcePkgCost) {
    checks.push({
      what: 'Recomputed PKg Cost vs the column in the file',
      ok: audit.recomputeMaxDiff < 1e-6,
      fig: 'max diff ' + audit.recomputeMaxDiff.toExponential(2)
    });
  }

  lastRows = rows; lastColumns = columns;
  lastMeta = { hasTotalAmount, hasTotalCost };

  post('Done', 95);
  return {
    sheetName: picked.name, sheets, rowCount: rows.length, columns, hasTotalAmount, hasTotalCost,
    hasWh: columns.includes('Target Warehouse'),
    months, groups, warehouses, long, biLong, trend, whTrend, biWhTrend,
    monthTotals, groupTotals, audit, checks
  };
}

/* ======================================================================
   OUTPUT WORKBOOK
   Raw Data is tab 1 and every computed cell downstream of it is a live
   formula, so the whole calculation can be audited without leaving Excel.
   ====================================================================== */
const COL = i => XLSX.utils.encode_col(i);
const A1 = (r, c) => XLSX.utils.encode_cell({ r, c });

function setFmt(ws, r, c, z) { const cell = ws[A1(r, c)]; if (cell) cell.z = z; }
function setFormula(ws, r, c, f, v, z) {
  ws[A1(r, c)] = { t: 'n', f, v: isFinite(v) ? v : 0, ...(z ? { z } : {}) };
}

function buildWorkbook(res, opts) {
  const wb = XLSX.utils.book_new();
  const a0 = res.audit;                 // totals reused by the grand-total rows

  const addAoa = (name, header, body, freezeCols) => {
    const ws = XLSX.utils.aoa_to_sheet([header, ...body]);
    ws['!cols'] = header.map(h => ({ wch: Math.min(Math.max(String(h).length + 2, 11), 40) }));
    ws['!autofilter'] = {
      ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: body.length, c: header.length - 1 } })
    };
    XLSX.utils.book_append_sheet(wb, ws, name);
    return ws;
  };

  /* ------------------------------------------------------- Raw Data (tab 1)
     Source rows, laid out the way the reference workbook lays them out:
     Date 1 and Month first, then the source columns, then PKg Cost last.
     All three derived columns are formulas. */
  const RAW = 'Raw Data';
  let rawRefs = null, rawDropped = [], rawOmitted = false;
  try {
    if (!lastRows || !lastRows.length) throw new RawTooBig();
    const cols = lastColumns.filter(c => !['Date 1', 'Month', 'PKg Cost'].includes(c));
    const n = lastRows.length, lastR = n + 1;

    // Measured against this build of SheetJS: sheet XML runs ~52 bytes a cell
    // and the writer dies past roughly 90 MB for the workbook as a whole.
    const BUDGET = 58e6;
    // Target Warehouse joins CORE because the Warehouse Trend tab's SUMIFS
    // criteria point at it - without the column those cells cannot be written.
    const CORE = ['Item Group', 'Item Type', 'Qty', 'PKG']
      .concat(res.hasTotalAmount ? ['Total Amount'] : (res.hasTotalCost ? ['Total Cost'] : ['Value In FG', 'Additional Cost']))
      .concat(cols.includes('Date') ? ['Date'] : [])
      .concat(cols.includes('Target Warehouse') ? ['Target Warehouse'] : []);
    const PRIORITY = CORE.concat(['Total Amount', 'Item Name', 'Workorder', 'Item Code',
      'Amount', 'Basic Rate', 'Source Warehouse']);

    const sample = lastRows.slice(0, 500);
    const costOf = c => {
      let t = 0;
      for (const r of sample) { const v = r[c]; t += (typeof v === 'string') ? v.length + 62 : 32; }
      return (t / sample.length) * n;
    };
    const cost = {}; for (const c of cols) cost[c] = costOf(c);

    // Date 1 (~34), Month (~86, a CHOOSE so it never depends on locale),
    // PKg Cost (~70), and Total Amount when it has to be derived (~40)
    let budget = BUDGET - (190 + (res.hasTotalAmount || res.hasTotalCost ? 0 : 40)) * n;
    const keep = new Set();
    const order = PRIORITY.filter(c => cols.includes(c))
      .concat(cols.filter(c => !PRIORITY.includes(c)));
    for (const c of order) if (cost[c] <= budget) { keep.add(c); budget -= cost[c]; }

    // without these the formula chain cannot be written at all
    if (!CORE.every(c => keep.has(c) || !cols.includes(c))) throw new RawTooBig();

    const kept = cols.filter(c => keep.has(c));
    rawDropped = cols.filter(c => !keep.has(c));

    const derivedTotal = !res.hasTotalAmount && !res.hasTotalCost && !res.hasAmount;
    const totalColName = res.hasTotalAmount ? 'Total Amount' : (res.hasTotalCost ? 'Total Cost' : 'Total Amount');
    const head = ['Date 1', 'Month'].concat(kept, derivedTotal ? [totalColName] : [], ['PKg Cost']);
    const body = lastRows.map(r => {
      const line = kept.map(c => {
        const v = r[c];
        return v instanceof Date ? v : (v === null || v === undefined ? '' : v);
      });
      return [dateOnly(r) || '', monthOf(r)].concat(line, derivedTotal ? [0] : [], [0]);
    });

    const ws = XLSX.utils.aoa_to_sheet([head, ...body], { cellDates: true });
    ws['!cols'] = head.map((h, i) => ({ wch: i < 2 ? 12 : 15 }));
    ws['!autofilter'] = {
      ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: n, c: head.length - 1 } })
    };

    const at = name => head.indexOf(name);
    const L = name => COL(at(name));
    const cDate = at('Date') >= 0 ? L('Date') : null;
    const cPkg = L('PKG');
    const iTotal = at('Total Amount') >= 0 ? at('Total Amount') : (at('Total Cost') >= 0 ? at('Total Cost') : at('Amount'));
    const cTotal = iTotal >= 0 ? COL(iTotal) : COL(at(totalColName));
    const iPkgCost = at('PKg Cost');
    const dateIsReal = lastRows.some(r => r['Date'] instanceof Date || typeof r['Date'] === 'number');

    for (let i = 0; i < n; i++) {
      const r = i + 1, row = r + 1;
      if (cDate && dateIsReal) {
        // Date 1 and Month are derived from Date, so they say so in the cell.
        // CHOOSE beats TEXT(...,"mmm") here: the label can never turn into a
        // localised month name that no longer matches the report columns.
        ws[A1(r, 0)] = { t: 'n', f: `INT($${cDate}${row})`, v: 0, z: 'dd-mm-yyyy' };
        ws[A1(r, 1)] = { t: 's', f: `CHOOSE(MONTH($${cDate}${row}),"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")`, v: body[i][1] };
      } else {
        setFmt(ws, r, 0, 'dd-mm-yyyy');
      }
      if (derivedTotal) {
        setFormula(ws, r, iTotal, `$${L('Value In FG')}${row}+$${L('Additional Cost')}${row}`,
          totalAmountOf(lastRows[i], false), '#,##0.00');
      }
      // CONDITION 1, written into the cell: =N2*U2/100
      setFormula(ws, r, iPkgCost, `$${cTotal}${row}*$${cPkg}${row}/100`,
        pkgCostOf(lastRows[i], res.hasTotalAmount), '#,##0.0000');
    }
    XLSX.utils.book_append_sheet(wb, ws, RAW);

    const rr = c => `'${RAW}'!$${c}$2:$${c}$${lastR}`;
    rawRefs = {
      month: rr(COL(1)), group: rr(L('Item Group')), type: rr(L('Item Type')),
      qty: rr(L('Qty')), pkgCost: rr(COL(iPkgCost)),
      wh: at('Target Warehouse') >= 0 ? rr(L('Target Warehouse')) : null
    };
  } catch (e) {
    // Drop the sheet rather than ship a truncated one that formulas point at:
    // a partial Raw Data tab would make every SUMIFS silently wrong.
    if (!(e instanceof RawTooBig)) throw e;
    rawOmitted = true; rawRefs = null; rawDropped = [];
    if (wb.SheetNames.includes(RAW)) {
      wb.SheetNames = wb.SheetNames.filter(s => s !== RAW);
      delete wb.Sheets[RAW];
    }
  }

  /* ------------------------------------------------------------ Monthly Trend
     Every Qty and PKg Cost cell is a SUMIFS over Raw Data carrying the same
     conditions the dashboard applies: the row's own dimensions, the month, and
     Item Type = "FG". The warehouse tab reuses this writer with one extra
     criterion, so both tabs are driven by identical arithmetic. */
  const writePivot = (name, pivot, itemTypeFilter = "FG") => {
    const ws = addAoa(name, pivot.header, pivot.body, true);
    const nIdx = pivot.nIdx, nm = res.months.length, last = pivot.body.length;
    for (let i = 0; i < pivot.body.length; i++) {
      const r = i + 1, row = r + 1;
      const isGrand = pivot.body[i][0] === 'Grand Total';
      for (let k = 0; k <= nm; k++) {                     // months, then Total
        const c = nIdx + k * 3;                           // Qty column index
        const qL = COL(c), cL = COL(c + 1);
        if (isGrand) {
          setFormula(ws, r, c, `SUM(${qL}2:${qL}${last})`, pivot.body[i][c], '#,##0.00');
          setFormula(ws, r, c + 1, `SUM(${cL}2:${cL}${last})`, pivot.body[i][c + 1], '#,##0.0000');
        } else if (k < nm && rawRefs && pivot.crit.length === nIdx) {
          const dims = pivot.crit.map((ref, d) => `${ref},$${COL(d)}${row}`).join(',');
          const crit = `${dims},${rawRefs.month},"${res.months[k]}",${rawRefs.type},"${itemTypeFilter}"`;
          setFormula(ws, r, c, `SUMIFS(${rawRefs.qty},${crit})`, pivot.body[i][c], '#,##0.00');
          setFormula(ws, r, c + 1, `SUMIFS(${rawRefs.pkgCost},${crit})`, pivot.body[i][c + 1], '#,##0.0000');
        } else if (k === nm) {
          // the Total block sums the month cells beside it
          const parts = res.months.map((_, j) => `${COL(nIdx + j * 3)}${row}`).join(',');
          const partsC = res.months.map((_, j) => `${COL(nIdx + 1 + j * 3)}${row}`).join(',');
          setFormula(ws, r, c, `SUM(${parts})`, pivot.body[i][c], '#,##0.00');
          setFormula(ws, r, c + 1, `SUM(${partsC})`, pivot.body[i][c + 1], '#,##0.0000');
        } else {
          setFmt(ws, r, c, '#,##0.00'); setFmt(ws, r, c + 1, '#,##0.0000');
        }
        // CONDITION 3, and it is re-derived on the Total and Grand Total too
        setFormula(ws, r, c + 2, `IF(${qL}${row}=0,0,${cL}${row}/${qL}${row})`,
          pivot.body[i][c + 2], '#,##0.0000');
      }
    }
    return ws;
  };

  res.trend.crit = rawRefs ? [rawRefs.group] : [];
  writePivot('Monthly Trend', res.trend, 'FG');

  // the added tab: identical metrics, split by Target Warehouse
  if (res.whTrend && res.whTrend.body.length > 1) {
    res.whTrend.crit = (rawRefs && rawRefs.wh) ? [rawRefs.wh, rawRefs.group] : [];
    writePivot('Warehouse Trend', res.whTrend, 'FG');
  }

  const writeQtyPivot = (name, pivot, itemTypeFilter) => {
    const ws = addAoa(name, pivot.header, pivot.body, true);
    const nIdx = pivot.nIdx, nm = res.months.length, last = pivot.body.length;
    for (let i = 0; i < pivot.body.length; i++) {
      const r = i + 1, row = r + 1;
      const isGrand = pivot.body[i][0] === 'Grand Total';
      for (let k = 0; k <= nm; k++) {                     // months, then Total
        const c = nIdx + k;                               // Qty column index
        const qL = COL(c);
        if (isGrand) {
          setFormula(ws, r, c, `SUM(${qL}2:${qL}${last})`, pivot.body[i][c], '#,##0.00');
        } else if (k < nm && rawRefs && pivot.crit.length === nIdx) {
          const dims = pivot.crit.map((ref, d) => `${ref},$${COL(d)}${row}`).join(',');
          const crit = `${dims},${rawRefs.month},"${res.months[k]}",${rawRefs.type},"${itemTypeFilter}"`;
          setFormula(ws, r, c, `SUMIFS(${rawRefs.qty},${crit})`, pivot.body[i][c], '#,##0.00');
        } else if (k === nm) {
          const parts = res.months.map((_, j) => `${COL(nIdx + j)}${row}`).join(',');
          setFormula(ws, r, c, `SUM(${parts})`, pivot.body[i][c], '#,##0.00');
        } else {
          setFmt(ws, r, c, '#,##0.00');
        }
      }
    }
    return ws;
  };

  // BiProduct tab: same grouping for BiProduct items (only Qty)
  if (res.biWhTrend && res.biWhTrend.body.length > 1) {
    res.biWhTrend.crit = (rawRefs && rawRefs.wh) ? [rawRefs.wh, rawRefs.group] : [];
    const biTypeName = Object.keys(res.audit.typeCounts).find(t => t.toUpperCase().replace(/[\s_-]/g, '') === 'BIPRODUCT') || 'BiProduct';
    writeQtyPivot('BiProduct Trend', res.biWhTrend, biTypeName);
  }

  /* -------------------------------------------------------------- Month Totals
     The trend read the other way round: one row per month. */
  {
    const header = ['Month', 'Month No.', 'Sum of Qty', 'Sum of PKg Cost', 'PKg Cost / Qty'];
    const body = res.monthTotals.map(m => [m.month, m.monthNum, m.qty, m.cost, m.ratio]);
    body.push(['Grand Total', '', a0.qtyTotal, a0.costTotal, ratio(a0.qtyTotal, a0.costTotal)]);
    const ws = addAoa('Month Totals', header, body, false);
    const last = body.length;
    for (let i = 0; i < body.length; i++) {
      const r = i + 1, row = r + 1;
      if (body[i][0] === 'Grand Total') {
        setFormula(ws, r, 2, `SUM(C2:C${last})`, res.audit.qtyTotal, '#,##0.00');
        setFormula(ws, r, 3, `SUM(D2:D${last})`, res.audit.costTotal, '#,##0.0000');
      } else if (rawRefs) {
        const crit = `${rawRefs.month},$A${row},${rawRefs.type},"FG"`;
        setFormula(ws, r, 2, `SUMIFS(${rawRefs.qty},${crit})`, body[i][2], '#,##0.00');
        setFormula(ws, r, 3, `SUMIFS(${rawRefs.pkgCost},${crit})`, body[i][3], '#,##0.0000');
      } else {
        setFmt(ws, r, 2, '#,##0.00'); setFmt(ws, r, 3, '#,##0.0000');
      }
      setFormula(ws, r, 4, `IF(C${row}=0,0,D${row}/C${row})`, body[i][4], '#,##0.0000');
    }
  }

  /* -------------------------------------------------------- Item Group Summary */
  {
    const header = ['Item Group', 'Sum of Qty', 'Sum of PKg Cost', 'PKg Cost / Qty', 'Share of PKg Cost'];
    const body = res.groupTotals.map(g =>
      [g.group, g.qty, g.cost, g.ratio, a0.costTotal ? g.cost / a0.costTotal : 0]);
    body.push(['Grand Total', a0.qtyTotal, a0.costTotal,
      ratio(a0.qtyTotal, a0.costTotal), a0.costTotal ? 1 : 0]);
    const ws = addAoa('Item Group Summary', header, body, false);
    const last = body.length, gr = last + 1;
    for (let i = 0; i < body.length; i++) {
      const r = i + 1, row = r + 1;
      if (body[i][0] === 'Grand Total') {
        setFormula(ws, r, 1, `SUM(B2:B${last})`, res.audit.qtyTotal, '#,##0.00');
        setFormula(ws, r, 2, `SUM(C2:C${last})`, res.audit.costTotal, '#,##0.0000');
      } else if (rawRefs) {
        const crit = `${rawRefs.group},$A${row},${rawRefs.type},"FG"`;
        setFormula(ws, r, 1, `SUMIFS(${rawRefs.qty},${crit})`, body[i][1], '#,##0.00');
        setFormula(ws, r, 2, `SUMIFS(${rawRefs.pkgCost},${crit})`, body[i][2], '#,##0.0000');
      } else {
        setFmt(ws, r, 1, '#,##0.00'); setFmt(ws, r, 2, '#,##0.0000');
      }
      setFormula(ws, r, 3, `IF(B${row}=0,0,C${row}/B${row})`, body[i][3], '#,##0.0000');
      setFormula(ws, r, 4, `IF($C$${gr}=0,0,C${row}/$C$${gr})`, body[i][4], '0.0%');
    }
  }

  /* ------------------------------------------------------------ Logic & Audit */
  const a = res.audit;
  const auditRows = [
    ['CONDITIONS APPLIED', '', ''],
    ['1. PKg Cost', 'Total Amount x PKG / 100',
      'PKG is a percent held as a number, so 10.55 means 10.55%. Formulated as =N2*U2/100'],
    ['   Total Amount', res.hasTotalAmount ? 'taken from the file' : 'derived as Total Amount', ''],
    ['   Formula', '=N2*U2/100',
      'Total Amount x PKG / 100'],
    ['2. Row filter', 'Item Type = FG only', 'RM, PKG, BiProduct and any other type are excluded'],
    ['3. Packaging cost per kg', 'SUM(PKg Cost) / SUM(Qty)',
      'Totals re-derive the ratio; never an average of the monthly ratios'],
    ['4. Date 1', 'Derived from Date', 'Time of day stripped'],
    ['5. Month', 'Derived from Date, as a month name',
      'Jan, Feb, Mar - not 1-12; ordered by calendar, not alphabetically'],
    ['', '', ''],
    ['SOURCE', '', ''],
    ['Sheet analysed', res.sheetName, 'Chosen because it carries the required columns'],
    ['Rows read', res.rowCount, ''],
    ['FG rows analysed', a.fgRows, 'Everything below is computed from these rows only'],
    ['Non-FG rows excluded', res.rowCount - a.fgRows, ''],
    ['Item groups', res.groups.length, ''],
    ['Target warehouses', (res.warehouses || []).length,
      res.hasWh ? 'Warehouse Trend splits the same metrics by these'
                : 'No Target Warehouse column in the source'],
    ['Months present', res.months.join(', '), 'Calendar order'],
    ['Raw Data tab', rawOmitted ? 'omitted - source too large to embed'
      : `all ${res.rowCount.toLocaleString('en-IN')} rows`,
      rawOmitted ? 'Report cells hold values instead of formulas' : 'Every report cell sums it'],
    ['Raw Data columns omitted', rawOmitted ? 'n/a' : (rawDropped.length ? rawDropped.join(', ') : 'none'),
      rawDropped.length && !rawOmitted
        ? 'Every row kept; these columns dropped to stay inside the writer size limit'
        : 'Every source column written'],
    ['', '', ''],
    ['TOTALS', '', ''],
    ['Sum of Qty', a.qtyTotal, 'FG rows'],
    ['Sum of PKg Cost', a.costTotal, 'FG rows'],
    ['PKg Cost / Qty', ratio(a.qtyTotal, a.costTotal), 'Re-derived from the two totals above'],
    ['', '', ''],
    ['ROW TYPES SEEN', '', ''],
    ...Object.entries(a.typeCounts).sort((x, y) => y[1] - x[1]).map(([k, v]) => [k, v, 'rows']),
    ['', '', ''],
    ['CHECKS', '', ''],
    ...res.checks.map(c => [c.what, c.ok ? 'PASS' : 'CHECK', String(c.fig)])
  ];
  const wsA = addAoa('Logic & Audit', ['Item', 'Value', 'Note'], auditRows, false);
  wsA['!cols'] = [{ wch: 34 }, { wch: 46 }, { wch: 62 }];
  {
    const rowOf = label => auditRows.findIndex(x => x[0] === label) + 1;
    const qR = rowOf('Sum of Qty'), cR = rowOf('Sum of PKg Cost'), rR = rowOf('PKg Cost / Qty');
    if (rawRefs) {
      setFormula(wsA, qR, 1, `SUMIF(${rawRefs.type},"FG",${rawRefs.qty})`, a.qtyTotal, '#,##0.00');
      setFormula(wsA, cR, 1, `SUMIF(${rawRefs.type},"FG",${rawRefs.pkgCost})`, a.costTotal, '#,##0.0000');
    }
    setFormula(wsA, rR, 1, `IF(B${qR + 1}=0,0,B${cR + 1}/B${qR + 1})`,
      ratio(a.qtyTotal, a.costTotal), '#,##0.0000');
  }

  post('Writing file', 92);
  return XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
}

/* ------------------------------------------------------------------ router */
self.onmessage = e => {
  const { cmd, buffer, opts } = e.data;
  try {
    if (cmd === 'process') {
      const result = run(buffer, opts || {});
      self.postMessage({ type: 'result', result });
    } else if (cmd === 'export') {
      const out = buildWorkbook(e.data.result, opts || {});
      self.postMessage({ type: 'export', buffer: out }, [out.buffer || out]);
    }
  } catch (err) {
    self.postMessage({
      type: 'error',
      message: (err && err.message ? err.message : String(err))
    });
  }
};
