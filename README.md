# FG Packaging Cost Analyser

Upload a production Excel extract, get a month-wise packaging-cost dashboard, and
download a workbook where every computed cell is a live formula.

Static site, no build step. Everything — xlsx parse, calculation, xlsx write —
happens in the browser, so the file never leaves the machine and there is no
upload size limit to fight.

---

## The conditions

These are the whole calculation. Nothing else is applied.

### 1. PKg Cost = Total Amount × PKG / 100

`PKG` is a **percentage held as a plain number**, so `10.55` means 10.55%.
`Total Amount` is taken from the file; if a future export omits the column, it is derived from available amount fields or `Value In FG + Additional Cost`.

In the reference layout this is written as `=N2*U2/100` — N is `Total Amount`, U is `PKG`.

### 2. Item Type = FG only

Only finished-goods rows enter the analysis. `RM`, `PKG`, `BiProduct` and
anything else are excluded outright. The app reports how many rows it dropped.

### 3. Packaging cost per kg = SUM(PKg Cost) ÷ SUM(Qty)

Calculated by dividing `Sum of PKg Cost` by `Sum of Qty`. Totals
re-derive the ratio from the summed numerator and denominator — **never** an
average of the monthly ratios.

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

## On the page

Filters for **Month**, **Item group** and **Target warehouse** narrow every tile,
chart and table at once.

The **Compare** toggle picks which metrics are charted — one, two or all three.
It is multi-select, and the last one cannot be switched off. Each metric keeps
its own chart rather than sharing a plot: two measures of different scale on one
axis would need two y-scales, which invents a correlation that is not in the
data. With one selected the chart goes full width, with two it splits, with
three it runs three across.

| On screen | Column in the export | Is |
|---|---|---|
| **FG Qty** | `Sum of Qty` | finished goods produced |
| **PM Cost** | `Sum of PKg Cost` | packing material cost |
| **PM Cost per kg** | `PKg Cost / Qty` | packaging cost per kg |

**Every chart prints its value for every month** — above each column, above each
line marker, and at each bar end. Magnitudes are shortened (`4.7 L`, `52k`), the
ratio is shown to two decimals.

The two charts below follow whichever metric is selected first: the top five item
groups month by month, and the same metric ranked by item group.

The export keeps the `Sum of Qty` / `Sum of PKg Cost` / `PKg Cost / Qty` column
names.

---

## The export

Six sheets. `Raw Data` is first, and every computed cell downstream of it is a
`SUMIFS` back into it — so the whole calculation can be audited in Excel without
trusting this app.

| Sheet | Contents |
|---|---|
| **Raw Data** | Your source rows, laid out `Date 1`, `Month`, source columns…, `PKg Cost`. All three derived columns are formulas: `=INT($C2)`, a `CHOOSE(MONTH(...))` month name, and `=$Q2*$U2/100`. |
| **Monthly Trend** | One row per item group, three columns per month, plus a Total block. Every cell is a `SUMIFS` carrying all three conditions. |
| **Warehouse Trend** | The same pivot with `Target Warehouse` added as the leading column, so the trend can be read one warehouse at a time. Its `SUMIFS` carry the warehouse as a fourth criterion. |
| **Month Totals** | The same data by month. |
| **Item Group Summary** | The same data by item group, with share of cost. |
| **BiProduct Trend** | Multi-dimensional pivot for `Item Type = BiProduct` grouped by Target Warehouse, Item Group, and Month. |
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
