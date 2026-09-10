# FG Packaging Cost Analyser

Upload a production Excel extract, get a month-wise packaging-cost dashboard, and
download a workbook where every computed cell is a live formula.

Static site, no build step. Everything — xlsx parse, calculation, xlsx write —
happens in the browser, so the file never leaves the machine and there is no
upload size limit to fight.

---

## The conditions

These are the whole calculation. Nothing else is applied.

### 1. PKg Cost = Total Cost × PKG / 100

`PKG` is a **percentage held as a plain number**, so `10.55` means 10.55%.
`Total Cost` is `Value In FG + Additional Cost`; if a future export omits the
column it is reconstructed from those two.

In the reference layout this is literally `=Q2*U2/100` — Q is `Total Cost`,
U is `PKG`.

> This is **not** `Value In FG × PKG / 100`. Checked over all 11,503 rows of the
> reference extract:
>
> | Variant | Max difference from the file's own `PKg Cost` column |
> |---|---|
> | `Total Cost × PKG / 100` | **3.6e-15** — exact |
> | `Value In FG × PKG / 100` | 25.75 — wrong |

### 2. Item Type = FG only

Only finished-goods rows enter the analysis. `RM`, `PKG`, `BiProduct` and
anything else are excluded outright. The app reports how many rows it dropped.

### 3. Packaging cost per kg = SUM(Qty) ÷ SUM(PKg Cost)

Kept in that order because it is what the reference pivot shows. Totals
re-derive the ratio from the summed numerator and denominator — **never** an
average of the monthly ratios.

Note the ratio is kg per unit cost, so the name reads inverted relative to the
arithmetic; actual cost per kg would be the reciprocal.

### 4 & 5. Date 1 and Month are derived from Date

`Date 1` is the date with the time stripped. `Month` is a **real month name** —
`Jan`, `Feb`, `Mar` — never `1`–`12`, and every report is ordered by calendar
position rather than alphabetically.

---

## Choosing the sheet

A workbook often holds several working tabs. Sheets are scored on the columns
they carry, and **`Total Cost` is weighted heavily on purpose**: it separates the
analysis extract from a raw stock-entry dump, and the two give very different
answers. Row count only breaks ties.

The chosen sheet is named in the status bar and can be overridden from the
**Sheet** dropdown, which lists every tab that has the required columns.

---

## The export

Five sheets. `Raw Data` is first, and every computed cell downstream of it is a
`SUMIFS` back into it — so the whole calculation can be audited in Excel without
trusting this app.

| Sheet | Contents |
|---|---|
| **Raw Data** | Your source rows, laid out `Date 1`, `Month`, source columns…, `PKg Cost`. All three derived columns are formulas: `=INT($C2)`, a `CHOOSE(MONTH(...))` month name, and `=$Q2*$U2/100`. |
| **Monthly Trend** | One row per item group, three columns per month, plus a Total block. Every cell is a `SUMIFS` carrying all three conditions. |
| **Month Totals** | The same data by month. |
| **Item Group Summary** | The same data by item group, with share of cost. |
| **Logic & Audit** | The conditions restated beside the control totals and checks. |

`Month` is written with `CHOOSE(MONTH(...))` rather than `TEXT(...,"mmm")` so the
label can never turn into a localised month name that no longer matches the
report columns.

Filters on the page do not narrow the export — it always covers the whole file.

### Size limit

The browser build of SheetJS assembles the whole zip in one buffer and fails
past roughly 90 MB. `Raw Data` is therefore budgeted: **every row is always
kept**, and source columns are added in priority order until the budget is
spent. Anything left out is named on `Logic & Audit`. If even the columns the
formula chain needs will not fit, the sheet is dropped rather than truncated —
a partial `Raw Data` tab would make every `SUMIFS` pointing at it silently wrong.

---

## Verified against the reference

Reproduced from `Sheet1` of the reference extract, grouped by month × item group:

- **All 14 item groups** in the Month 8 pivot match `Sheet3` to the last decimal
- Grand totals exact — Jun `668,894.958 / 83,805.66487186884 / 7.981500522938144`,
  Jul `601,023.408 / 81,435.85954810692 / 7.380328657855636`,
  Aug `648,171.687 / 79,413.86777976828`
- Aug's ratio is blank in `Sheet3`; it computes to `8.161946`
- The recomputed `PKg Cost` reproduces the file's own column, max diff `3.55e-15`

---

## Running it

Open `index.html`, or serve the folder:

```bash
python -m http.server 5173
```

## Deploying to Vercel

No build step. Framework preset **Other**, build command empty, output directory
`.`. Or from the folder:

```bash
vercel --prod
```

## Files

```
index.html          markup, tokens and all styles
assets/app.js       UI, filters, hand-rolled SVG charts
assets/worker.js    parse, FG aggregation, pivots, xlsx writer
vercel.json         static config
```

`assets/app.js` carries a `BUILD` constant appended to the worker URL. **Bump it
whenever `worker.js` changes** — otherwise browsers keep running a cached worker,
which silently masks edits.
