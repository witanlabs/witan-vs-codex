# witan xlsx exec vs codex spreadsheet tooling — 8 reproducible test cases

All cases tested 2026-04-20 against witan CLI 0.9.0 (API v2.19.0),
`@oai/artifact-tool` 2.6.9, xlwings 0.35.1, and Microsoft Excel for Mac
(macOS Darwin 25.3.0). Excel is used only as ground truth via xlwings
automation. Codex workbook edits are exercised through the bundled Node runtime.

All fixtures, scripts, and outputs live under `~/dev/witan-vs-codex/`.

Cases 2-7 reuse the same fixtures and Witan baselines as the sibling
comparison in `~/dev/witan-vs-openpyxl/`. Cases 8-10 add Codex-specific and
Witan-specific comparison outputs in this repo.

## Summary

| # | Task | codex | witan |
|---|------|-------|-------|
| 2 | Iterative calc after changing an input | ✗ stale circular outputs (`38181.82`, not `35000`) | ✓ correct `35000` |
| 3 | Read a legacy `.xls` file | ✗ `FileContainsCorruptedData` | ✓ auto-converts and reads |
| 4 | Add / preserve threaded comments | ✗ loses resolved state; exported file is not Excel-openable | ✓ all threads preserved |
| 5 | Write `=UNIQUE(FILTER(...))` dynamic array | ✗ spill looks right in-tool, but opens blank in Excel | ✓ spill evaluates correctly |
| 7 | Write formulas that consume a spill reference (`A1#`) | ✗ internal results already wrong; exported file is not Excel-openable | ✓ computes spill consumers correctly |
| 8 | Rename a sheet referenced by formulas | ✗ formulas still say `Data!…` after rename | ✓ every reference rewritten |
| 9 | Insert / delete rows and columns used by formulas | ✗ ops are unimplemented; workbook stays unchanged | ✓ formulas and spill shift correctly |
| 10 | Calculate special formulas (`TEXT`, 3D refs, `OFFSET`, `INDIRECT`, `MAP`, `REDUCE`, etc.) | ✗ multiple wrong in-tool results; exported file is not Excel-openable | ✗ 16/17 right, but `INDEX/XMATCH` still fails |

Key: ✓ works · ✗ fails.

## Failure classes

Grouped by the *kind* of failure each case surfaces on the Codex side:

- **Silent wrong answer** (agent can return a value, but it is wrong) — 2, 7, 8, 10
- **Cannot complete the task at all** (hard error or missing implementation) — 3, 9
- **File opens in Excel but is semantically wrong** — 2, 5, 8, 9
- **File is not Excel-openable** — 4, 7, 10
- **Broken formula / metadata serialization** — 5, 7, 8, 10

---

## Environment

- Working dir: `~/dev/witan-vs-codex/`
- Codex spreadsheet library: `@oai/artifact-tool 2.6.9`
- Codex Node runtime: `/Users/nuno/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node`
- Python launcher: `uv run --with xlwings python <script>`
- witan CLI: `0.9.0`, production API `v2.19.0`
- Excel: Microsoft Excel for Mac, automated through xlwings 0.35.1
- Fixtures built in this repo live in `fixtures/`

---

## Case 2 — Iterative calculation over a circular reference

**Verdict**
- codex — **✗** Writes the new input, but leaves the circular outputs stale at `54545.45454 / 5454.545454 / 38181.818178`.
- witan — **✓** Recomputes the iterative model and returns `35000`.

**Fixture:** `fixtures/circular.xlsx`

**Prompt:**
> In `circular.xlsx`, change the bonus rate on `Inputs!B4` from 10% to 20%
> and report the new net income from `Model!B7`.

Analytic answer: `B3 = 100000 - 40000 - 0.2·B3`, so `B3 = 50000`,
bonus = `10000`, net income = `35000`.

### codex

- `Inputs!B4 = 0.2` is written successfully.
- After `workbook.recalculate()`, the cached outputs stay stale:
  - `Model!B3 = 54545.45454`
  - `Model!B4 = 5454.545454`
  - `Model!B7 = 38181.818178`
- Excel opens `outputs/case2_codex.xlsx` cleanly and shows the same stale
  values, so the wrong answer is persisted into the saved file.

### witan

- The same task returns `Model!B7 = 35000`, matching the analytic answer.
- That remains the baseline expected result for the comparison.

---

## Case 3 — Read a legacy `.xls` file

**Verdict**
- codex — **✗** Import fails immediately with `FileContainsCorruptedData`; no workbook is produced.
- witan — **✓** Auto-converts `.xls` to `.xlsx` server-side and reads normally.

**Fixture:** `fixtures/formulas.xls`

**Prompt:**
> Read `Sheet1` (or the first sheet) from `formulas.xls` and report cell B3.

### codex

- `SpreadsheetFile.importXlsx(...)` fails on the legacy workbook with
  `FileContainsCorruptedData`.
- Because no workbook is imported, there is no output file for Excel to open.

### witan

- Reads the same file and returns `B3 = 4`.
- The legacy format itself is not a blocker on the Witan path.

---

## Case 4 — Threaded comments

**Verdict**
- codex — **✗** Imports thread bodies, but loses resolved-state metadata on existing threads; the exported workbook is not Excel-openable.
- witan — **✓** Round-trips all threads, including author and resolved state.

**Fixture:** `fixtures/review.xlsx`

**Prompt:**
> In `review.xlsx`, add a resolved threaded comment on `Data!B3` saying
> "Verified against ledger" by author "Auditor", then list every threaded
> comment in the workbook with author, text, and resolved state.

### codex

- Existing threads import, but the already-resolved `Data!B2` thread is loaded
  as plain `status: 1` with `resolvedAt: null` and `resolvedBy: null`.
- The new `B3` thread can be created in memory, but the saved OOXML is
  suspicious:
  - `xl/threadedcomments/threadedcomment.xml` drops the original `done="1"` /
    `done="0"` markers
  - the new thread id is a short non-GUID token instead of a normal thread GUID
  - `xl/persons/person.xml` duplicates `Auditor`
- Excel / xlwings refuses to open `outputs/case4_codex.xlsx` and returns
  `OSERROR -50`.

### witan

- Preserves the threaded comment parts and resolved-state metadata.
- Excel opens the corresponding Witan result cleanly and the thread inventory is
  still intact.

---

## Case 5 — Dynamic array spill (`=UNIQUE(FILTER(...))`)

**Verdict**
- codex — **✗** The spill looks correct in-tool before export, but the saved workbook opens with a blank spill and the formula degrades on Excel round-trip.
- witan — **✓** Emits a proper dynamic-array anchor and the spill evaluates correctly in Excel.

**Fixture:** `fixtures/report.xlsx`

**Prompt:**
> In `report.xlsx`, put `=UNIQUE(FILTER(Raw!A2:A13, Raw!B2:B13>0))` into
> `Summary!D2` as a dynamic array so it spills down. Return the spilled values.

Expected spill: `Food`, `Rent`, `Supplies`, `Travel`.

### codex

- Before save, the runtime reports the right spill values in `D2:D5`.
- In `outputs/case5_codex.xlsx`, Excel opens the file, but `Summary!D2:D6`
  come back blank even though `D2` still holds the formula string.
- The root serialization problem is visible in the saved XML:
  - `FILTER` is namespaced as `_xlfn._xlws.FILTER(...)`
  - `UNIQUE` is written as plain `UNIQUE(...)` instead of `_xlfn.UNIQUE(...)`
- After an Excel save round-trip, the anchor is rewritten to `_xludf.UNIQUE(...)`
  with `#NAME?`, confirming that Excel never recognized the original formula
  correctly.

### witan

- Saves a valid spill anchor and spill children.
- Excel opens the Witan workbook without repair and shows the expected four
  categories.

---

## Case 7 — Spill-reference consumers (`A1#`)

**Verdict**
- codex — **✗** Spill-reference consumers are already wrong before export, and the saved workbook is not Excel-openable.
- witan — **✓** Evaluates the spill consumers directly and returns the expected values.

**Fixture:** `fixtures/report.xlsx`

**Prompt:**
> Build a summary with `Summary!D2 = UNIQUE(FILTER(...))`, then add:
> `COUNTA(Summary!D2#)`, `COUNTIF(Summary!D2#, "Food")`, and
> `TEXTJOIN(", ", TRUE, Summary!D2#)`.

Expected:

- `F2 = 4`
- `G2 = 1`
- `H2 = Food, Rent, Supplies, Travel`

### codex

- Before export, the consumer formulas are already wrong:
  - `F2 = 1`
  - `G2 = 0`
  - `H2 = Error: #NAME?`
- The saved workbook is structurally incomplete:
  - the spill producer has `cm="1"`, but the spill consumers do not
  - there is no `xl/metadata.xml` part
- Excel / xlwings refuses to open `outputs/case7_codex.xlsx` and returns
  `OSERROR -50`.

### witan

- Computes the spill consumers correctly:
  - `COUNTA(D2#) = 4`
  - `COUNTIF(D2#, "Food") = 1`
  - `TEXTJOIN(", ", TRUE, D2#) = Food, Rent, Supplies, Travel`
- The spill-reference path behaves normally in Excel.

---

## Case 8 — Rename a sheet referenced by formulas

**Verdict**
- codex — **✗** Renames the sheet object, but leaves every dependent formula pointing at `Data!…`.
- witan — **✓** Rewrites the scalar and spilled formulas to the new sheet name.

**Fixture:** `fixtures/case8_rename_fixture.xlsx`

**Prompt:**
> Rename sheet `Data` to `Renamed` and preserve formula correctness on `Summary`,
> including both normal formulas and a spilled array formula.

Initial `Summary` formulas:

- `B2 = SUM(Data!B2:B4)` → `6`
- `B3 = SUMPRODUCT(Data!B2:B4, Data!C2:C4)` → `140`
- `D2 = TRANSPOSE(Data!A2:C4)` spilling through `F4`

### codex

- The sheet names become `Renamed`, `Summary`.
- The formulas do not change:
  - `B2` stays `=SUM(Data!B2:B4)`
  - `B3` stays `=SUMPRODUCT(Data!B2:B4, Data!C2:C4)`
  - `D2` stays `=TRANSPOSE(Data!A2:C4)`
- Excel opens `outputs/case8_codex.xlsx`, but the stale references come back as
  blank dependent cells:
  - `B2.value = None`
  - `B3.value = None`
  - `D2.value = None`

### witan

- Rewrites all three formulas to `Renamed!…`.
- Excel opens `outputs/case8_witan.xlsx` and shows the expected values:
  - `B2 = 6`
  - `B3 = 140`
  - spill block = `A B C / 1 2 3 / 10 20 30`

---

## Case 9 — Insert / delete rows and columns used by formulas

**Verdict**
- codex — **✗** The structural operations are unimplemented, so the workbook stays unchanged.
- witan — **✓** Applies the whole row / column edit sequence and shifts both scalar and spilled formulas correctly.

**Fixture:** `fixtures/case9_shift_fixture.xlsx`

**Prompt:**
> Apply row and column insert/delete operations inside the `Data` range and
> preserve correct references on `Summary`, including a spilled `TRANSPOSE(...)`
> formula.

Applied operation sequence:

1. Insert row after `3`, then fill new row `4` with `X / 10 / 100`
2. Insert column after `B`, then fill new `C` with adjustment values
3. Delete row `5`
4. Delete column `C`

Expected final state:

- `Data` rows become `A`, `B`, `X`, `D`
- `B2 = 17`
- `B3 = 170`
- spill values become `A,B,X,D / 1,2,10,4 / 10,20,100,40`

### codex

- Every structural op returns a `TODO: ... not implemented yet.` result:
  - `rows.insert`
  - `columns.insert`
  - `rows.delete`
  - `columns.delete`
- The saved workbook is unchanged from the fixture.
- Excel opens `outputs/case9_codex.xlsx` cleanly, but it still shows the
  original data and formulas:
  - `B2 = 10`
  - `B3 = 100`
  - spill rows remain `A,B,C,D / 1,2,3,4 / 10,20,30,40`

### witan

- Applies the edit sequence and shifts the dependent formulas to the new final
  ranges.
- Excel opens `outputs/case9_witan.xlsx` and shows the expected final state:
  - `Data!A4:C4 = X, 10, 100`
  - `B2 = 17`
  - `B3 = 170`
  - spill rows = `A,B,X,D / 1,2,10,4 / 10,20,100,40`

---

## Case 10 — Special formula calculation correctness

**Verdict**
- codex — **✗** One complex `TEXT` format works, another picks the wrong format section, several modern or indirection-based functions fail in-tool, and the saved workbook is not Excel-openable.
- witan — **✗** Gets 16/17 formulas right, but `INDEX/XMATCH` still fails and comes back blank in Excel.

**Fixture:** `fixtures/case10_formula_fixture.xlsx`

**Prompt:**
> Populate `Summary!C2:C18` with a mixed formula panel covering `TEXT`, 3D
> references, `OFFSET`, `INDIRECT`, `MAP`, `REDUCE`, `LET`, `XLOOKUP`,
> `INDEX/XMATCH`, `TEXTJOIN(MAP(...))`, `SUMPRODUCT`, `SEQUENCE`, `BYROW`,
> `CHOOSECOLS`, `TAKE`, and `DROP`, then preserve correct values on save.

Expected highlights on `Summary!C2:C18`:

- complex `TEXT` outputs:
  - `C2 = "($1,234.57)"`
  - `C3 = "37.5%"`
- scalar results:
  - `C4 = 60`
  - `C5 = 10`
  - `C6 = 10`
  - `C7 = 20`
  - `C8 = 30`
  - `C9 = 300`
  - `C10 = 30`
  - `C11 = 40`
  - `C12 = "A:1,B:2,C:3,D:4"`
  - `C13 = 9`
  - `C14 = 6`
  - `C15 = 110`
  - `C16 = 10`
  - `C17 = 3`
  - `C18 = 7`

### codex

- Correct in-tool:
  - `TEXT(-1234.567, "[Green]$#,##0.00;[Red]($#,##0.00);0.00")`
  - `LET`
  - `XLOOKUP`
  - `INDEX/XMATCH`
  - `SUMPRODUCT`
  - `SUM(SEQUENCE(...))`
  - `SUM(BYROW(...))`
  - `SUM(CHOOSECOLS(...))`
  - `SUM(TAKE(...))`
  - `SUM(DROP(...))`
- Wrong in-tool:
  - `TEXT(0.375, "[>=1]0.0%;[Red](0.0%);0.0%")` returns `(37.5%)`
  - `SUM(Jan:Mar!B2)` returns `#VALUE!`
  - `OFFSET(...)` returns a literal `"OFFSET is not implemented ..."` message
  - `INDIRECT(...)` returns a literal `"INDIRECT is not implemented ..."` message
  - `SUM(MAP(...))` returns a literal `"MAP is not implemented ..."` message
  - `REDUCE(...)` returns a literal `"REDUCE is not implemented ..."` message
  - `TEXTJOIN(MAP(...))` returns a literal `"Error: MAP is not implemented ..."` message
- `outputs/case10_codex.xlsx` does not open in Excel via xlwings and fails
  with `OSERROR -50`.
- The saved XML has several concrete red flags:
  - no `xl/metadata.xml`
  - modern formulas are mostly written without `_xlfn.` prefixes
  - `Summary!C4:C8` are serialized as `t="e"` formula cells with no `<v>` error
    payload

### witan

- Both complex `TEXT` formulas open in Excel with the expected results:
  - `($1,234.57)`
  - `37.5%`
- 16 of the 17 formulas match the expected value after Excel open.
- The one miss is `INDEX(Data!C2:C5, XMATCH("D", Data!A2:A5))`:
  - Witan runtime returns `#VALUE!`
  - Excel opens `outputs/case10_witan.xlsx`, keeps the formula string in `C11`,
    but xlwings reads the cell value back as blank
