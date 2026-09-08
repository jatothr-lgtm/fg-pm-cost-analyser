# FG / PM Cost Analyser

Upload a production Excel extract, get packaging-material (PM) cost per kg by month,
item group and warehouse, and download the finished workbook.

Static site — no build step, no server, no database. The workbook is parsed and
calculated entirely in the browser (Web Worker + SheetJS), so production cost data
never leaves the machine it is opened on.

## Deploy

```bash
vercel --prod
```

Or connect the repo in the Vercel dashboard. Framework preset: **Other**.
Build command: none. Output directory: `.` (repo root).

## Run locally

Any static file server works. It must be served over HTTP — opening `index.html`
from the filesystem fails, because a Web Worker cannot load from a `file://` origin.

```bash
python -m http.server 5173
```

## Input

The **first sheet** of the uploaded workbook is read, **in file order**. Row order is
load-bearing — see the bucketing rule below. Required columns:

| Column | Used for |
|---|---|
| `Workorder` | the grouping key for everything |
| `Item Group` | the reporting bucket |
| `Item Type` | `PKG` / `FG` / `RM` / `BiProduct` |
| `Qty` | FG quantity |
| `Total Amount` | PM Value, FG Value |
| `Value In FG` | PM Qty |
| `PKG` | PM Qty (a percentage held as a plain number: `2.02` means 2.02%) |
| `Date` or `Month` | the month bucket |
| `Target Warehouse` / `Source Warehouse` | optional; default to `Unknown` |

Leading and trailing spaces in headers are stripped.

## The calculation

### PM Value

```
PM Value(workorder) = SUM(Total Amount) WHERE Item Type = 'PKG'
```

`RM`, `FG` and `BiProduct` rows contribute nothing.

### Bucketing — which Item Group the PM Value is reported under

Switchable in the UI, because the two rules give materially different answers.

**`Work order first row` (default).** Item Group and Month come from the **first row
of that work order in file order** — equivalent to
`VLOOKUP(Workorder, <input data>, Item Group)`. That row is the work order's primary
input line: the RM line for most work orders, the PKG line for work orders that have
no RM line at all (these bucket under `Packaging Material`).

**`FG row`.** Item Group comes from the work order's FG rows. Work orders with zero
or several distinct FG Item Groups cannot be resolved and their PM Value is reported
as unmapped rather than guessed at.

The two differ because a work order frequently consumes one item group and produces
another, and because many work orders have no FG line to read at all. The first-row
rule reproduces the reference pivot exactly; the FG-row rule does not. Both are
available so the difference can be inspected rather than argued about.

### PM Qty and cost per kg

```
PM Qty     = SUM over FG rows of ( Value In FG × PKG / 100 )
FG Qty     = SUM over FG rows of Qty
PM Cost/kg = PM Value ÷ (PM Qty | FG Qty)      ← denominator switchable
```

Totals re-derive the ratio from summed numerator over summed denominator — never an
average of monthly ratios.

### Control total

Every PKG row belongs to exactly one work order, and each work order's PM Value is
attached to exactly one output cell, so:

```
SUM(PM Value across buckets) == SUM(Total Amount WHERE Item Type = 'PKG')
```

is true by construction. The dashboard shows both figures and their difference; the
`PM Value Audit` tab and sheet carry the same numbers.

## Output workbook

| Sheet | Contents |
|---|---|
| `Monthly Summary` | Target Warehouse × FG Item Group × month |
| `Monthly Summary - Source WH` | Source Warehouse × FG Item Group × month |
| `Workorder Summary` | one row per work order, pre-aggregation |
| `Logic & Audit` | every rule applied, the row census, and the reconciliation |
| `Multiple FG Groups` | work orders with more than one distinct FG Item Group |

Only months actually present in the data get columns, in calendar order.

### Live formulas

Every derived cell is written as a **real Excel formula with a cached value**, so the
workbook recalculates and any figure can be traced by clicking the cell. Only the raw
monthly sums are literal numbers; everything downstream of them is a formula:

```
PM Cost/KG        =IF(D2=0,0,E2/D2)          value ÷ denominator, divide-by-zero guarded
Total (PM Value)  =SUM(E2,I2,M2,Q2)          across the month columns
Grand Total       =SUM(E2:E125)              down the item-group rows
Audit difference  =B10-B11-B12               source − allocated − unmapped, must read 0
```

Column widths, number formats and autofilter are also written. Frozen panes and bold
headers are not — the browser build of SheetJS cannot emit them, and adding a server
just to style a header row was not worth the infrastructure.

## Believing the numbers

The **Logic & Audit** tab recomputes eleven reconciliation checks from the loaded file
every time a setting changes, and each one states what would be wrong if it failed:

1. Allocated + unmapped equals the source PKG total — no value vanished.
2. Nothing is left unallocated.
3. The source total re-adds independently from the raw rows.
4. No work order is counted twice — carrier cells equal distinct work orders.
5. The monthly figures re-add to the total.
6. The Target and Source warehouse views total identically.
7. FG Qty ties back to a plain sum of Qty over the FG rows.
8. Every PKG row carries a work order.
9. No negative packaging value.
10. Every row is classified by Item Type.
11. Every row lands in a real month.

Alongside them the tab shows a **row census** (where all the rows went) and **bucket
provenance** (which line each work order took its item group from), so the reported
figures can be walked back to the source rows without leaving the page. The same
content is written to the `Logic & Audit` sheet of the download.

## Notes

- **Source Warehouse is read from the input (RM/PKG) lines, not the FG line.** FG rows
  carry a Target Warehouse and a blank Source Warehouse, so taking it from FG rows
  makes the entire Source WH summary read `Unknown`.
- **A work order's PM Value lands in exactly one cell.** Where a work order spans
  several warehouses or months on its FG rows, the value goes to its largest FG cell
  rather than being repeated against each — repeating it inflates the control total.
