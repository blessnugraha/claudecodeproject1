# IDX XBRL Workbench

One HTML page that opens IDX XBRL filings and an Excel workbook side by side, lets you copy figures from the filing into the sheet, and saves the filled workbook. It was built around `IDX_FSA_Calculator_v2.xlsx` (sheet `1. Input`, columns FY-4 … 1H Prior), but works with any `.xlsx`.

**To use it:** download [`dist/idx-xbrl-workbench.html`](dist/idx-xbrl-workbench.html) and open it in Chrome or Edge. Everything runs in your browser; no file is uploaded anywhere. It also works offline, though without internet it falls back to system fonts.

## What it does

| Area | Behaviour |
|---|---|
| **XBRL in** | XBRL 2.1 instances (`.xbrl`, `.xml`), Inline XBRL (`.xhtml`, `.html`), and `.zip` packages of either (IDX's download). Open several years at once: they merge into one table, and where a later filing restates a prior year, the later figure wins. The cell gets a dot, and its tooltip shows the other values. Label linkbases in the zip give readable names (English/Indonesian switch). |
| **Excel in** | Renders every sheet with its number formats, fills, fonts, merged cells and frozen columns, and recalculates all formulas live as you paste. |
| **Copy → Paste** | Click an XBRL value (or a row name to take every period), click a sheet cell (or a row number), press **Paste**. *Unit* (Auto / ×1 / ÷1,000 / ÷1,000,000 / ÷1,000,000,000), *Sign* (as filed / flip / always positive) and *Mode* (replace / add to cell) apply to each paste, and the bar shows exactly what will be written first. After a paste the target moves to the next input row. **Copy** also puts the value(s) on the clipboard for pasting into Excel or Google Sheets, and Ctrl+V pastes numbers from other sources (e.g. a PDF) into the selected cell. |
| **Periods** | Reads the sheet's header row (`FY-4`, `FY Current`, `1H Prior`, or years such as `2024`) and maps each column to an XBRL period. You can change any mapping in the bar above the grid. |
| **Auto-fill…** | Proposes values for ~60 input lines of the FSA calculator. Each one shows its source: a tagged concept, an accounting identity, or a residual. You untick anything wrong and fill. Cells that already hold a different value are left unticked. |
| **Tie-out** | Compares the sheet's own subtotals (revenue, gross profit, PBT, profit to parent, basic EPS, current/total assets, liabilities, equity, OCF/ICF/CFF) with the totals in the filing, per period. |
| **Safety** | Formula cells are never overwritten. Every paste or edit can be undone (Ctrl+Z). The page warns before closing with unsaved changes. |
| **Save** | Downloads `<name>-filled.xlsx`. The original file is patched, not rewritten. Only the cells you filled, plus the cached formula results, change. Charts, drawings, comments and styles are kept byte-for-byte, and Excel is told to recalculate everything on open. |

### How auto-fill decides

- **Sign conventions.** COGS, tax and NCI come from identities (Revenue − Gross profit, PBT − Profit for the period, Profit − Profit to parent), so the app doesn't have to guess how a filer signed its expense tags. Expense lines the sheet stores as positive numbers are made positive.
- **Residual rows.** "Other current assets", "Other OPEX", "Other income", "Other financing cash flows" and the like take *reported subtotal − lines placed above*. That makes the sheet's subtotals tie to the filing, but anything unmapped lands in those rows. Look at them. For example, *Other income* also absorbs the share of associates' profit and FX gains/losses when there is no tagged operating profit.
- **PP&E.** If the filing tags only net PP&E, it goes into *Gross PP&E* and *Accumulated Depreciation* stays blank, so *Net PP&E* is still right.
- **Segments.** Total revenue goes into *Segment 1*. Split segments by hand: tick **Breakdowns** to see dimensional (segment) facts, then paste them row by row.

### Limits

- **Concept names are unconfirmed.** They are matched on IFRS names and on IDX taxonomy names as best known. They have only been tested on the invented example filings in `tests/fixtures/`, not on a real IDX filing yet. If a line comes up empty, paste it manually; the tie-out panel shows whether the totals still agree.
- **The preview's formula engine is small.** It covers what the calculator uses: `+ − * / ^ & %`, comparisons, `SUM`, `IF`, `IFERROR`, `ABS`, `COLUMN`, plus `MIN`, `MAX`, `AVERAGE`, `ROUND`, `AND`, `OR`, `NOT`, `IS*`, `ROW`. Other functions show `#NAME?` in the preview only. Excel recalculates properly when you open the saved file.
- **Excel and Google Sheets differ on empty text.** Excel treats `""` in arithmetic as `#VALUE!`; Google Sheets treats it as 0. Pick the program you use under **Calculate like**. Files exported from Google Sheets default to Sheets.
- **Copied decimals use a point.** Values copied to the clipboard use `.` as the decimal mark. An Excel set to Indonesian locale may read `490.6` wrongly. Prefer pasting inside the app.
- **No old formats.** `.xls` (pre-2007) is not supported. Save as `.xlsx` first.

## Verification done

- **Formula engine.** It was run on all 1,029 formulas in `IDX_FSA_Calculator_v2.xlsx`. In Google Sheets mode it reproduces the file's cached values for 1,023 of them. In Excel mode it matches LibreOffice's independent recalculation for 1,023. The 6 left over are a broken row in the workbook itself (see below).
- **End-to-end.** The example filings were loaded into the calculator and auto-filled: all 13 tie-out checks pass for FY2022–FY2024, and the balance-sheet check row is 0. In the saved file, only `xl/workbook.xml` and the sheet XMLs differ from the original. All 4 charts are intact, and LibreOffice opens it.
- **`npm test`.** Unit tests (engine, number formats, XBRL number parsing, period lengths) plus a Chromium run on `tests/fixtures/mini-calculator.xlsx`: auto-fill, row paste, undo, formula protection, and a save that keeps the chart.

### Issues found in `IDX_FSA_Calculator_v2.xlsx` while testing

- **`3. Horizontal` row 66 (FCFF Growth).** It refers to row 0 (`'1. input'!d0`), which is not a valid cell. Google Sheets shows `#ERROR!`. It probably should be `=IFERROR((D64-C64)/C64,"")`.
- **`2. Vertical` row 25, "Cash & ST Investments % TA".** It divides only cash (row 43); short-term investments (row 44) are left out.
- **`2. Vertical` row 35, "ST Debt + CPLTD % TL".** It uses only row 64; CPLTD (row 65) is left out.
- **`4. Dashboard` F9.** In Excel, while ROE is still `""` (no data), `"" > 0.12` is TRUE, so it shows "✓ Above CoE" for an empty model.

## Development

```bash
npm install          # ExcelJS 4.4.0 and JSZip 3.10.1
npm run build        # writes dist/idx-xbrl-workbench.html with everything inlined
npm test             # unit tests, plus the browser test if `playwright` is installed
```

`src/index.html` also runs straight from the folder (it loads the libraries from `node_modules/`).

| File | Role |
|---|---|
| `src/xbrl.js` | XBRL / Inline XBRL / zip reader, label linkbases, merging filings into one table |
| `src/workbook.js` | Reads `.xlsx` with ExcelJS for display; saves by patching the original package with JSZip |
| `src/formula.js` | Formula parser and evaluator for the live preview |
| `src/numfmt.js` | Excel number-format rendering |
| `src/automap.js` | Auto-fill rules, period mapping, tie-out checks |
| `src/app.js` | UI |
| `tools/make_sample.py` | Generates the invented example filings (`tests/fixtures/*.xbrl`, `src/sample-data.js`) |
| `tools/make_fixture.py` | Generates `tests/fixtures/mini-calculator.xlsx` (needs openpyxl) |

ExcelJS is used only to read. It drops charts when it writes, and it crashes reading some chart drawings, so it is handed a copy of the workbook with drawings removed. Saving never goes through it.
