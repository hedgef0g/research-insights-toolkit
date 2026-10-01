# Research Insights Toolkit - Release Smoke Tests

Last updated: 2026-05-25

## Purpose

This checklist is the manual release smoke gate for the current post-split MVP.

Use it to validate the current workflow split before a release candidate and after risky changes to detection, selected-range normalization, generated-sheet handling, banner markers, Check flows, Autorun flows, or workbook-generated outputs.

This document is intentionally concrete. It is meant to be executed by a human in Excel, not read as general guidance.

Related references:

- `docs/TEST_CASES.md` - broader manual regression catalog
- `docs/TABLE_STRUCTURE_MATRIX.md` - structure support coverage
- `docs/GOLD_STANDARD_TEST_SUITE.md` - planned gold-standard validation coverage
- `docs/SELECTED_RANGE_NORMALIZATION.md` - selected-range pass-through / normalized / blocked rules

## Release gate

A release candidate is not ready until:

- [ ] `npm test` passes
- [ ] `npm run build` passes
- [ ] `git diff --check` passes
- [ ] all relevant smoke workbook checks below pass on the target Excel environment

## When to run

Run this checklist:

- before every release candidate;
- after changes to Check, Autorun, Content, Run report, selected-range normalization, banner handling, or generated-sheet guards;
- after risky refactors affecting table detection or workbook traversal;
- after Excel integration changes that could affect active-cell or selection-driven behavior.

## Prerequisites

- Excel desktop environment intended for release validation
- add-in built from the candidate branch
- manual tester knows how to:
  - select a numeric data range;
  - place the active cell inside a target table without selecting the whole table;
  - create a Ctrl+Click non-contiguous selection;
  - rerun the same scenario with changed settings
- if workbook fixtures are not yet committed, recreate them from the scenario definitions in this file and mark the local copies used for smoke

## Standard validation commands

Run before or alongside manual smoke:

```powershell
npm test
npm run build
git diff --check
```

Record results here for the current run:

- Date:
- Branch / commit:
- Excel version:
- Tester:
- `npm test`:
- `npm run build`:
- `git diff --check`:

## Smoke workbook pack

The workbook pack below is the minimum release smoke set for the current MVP. If binaries are not stored in the repo, treat each workbook as `to be created` and keep the scenario shape stable between releases.

### 1. `basic-proportions.xlsx` - to be created

Required contents:

- one sheet with one simple proportions table
- clear table title above the banner
- one banner row with `Всего` and at least two segment columns
- one Base row
- data shaped so significance is expected in at least one row
- workbook notes with expected significance letters for the fixture

Used for:

- selected-range `Рассчитать`
- selected-range `Очистить значимости`
- selected-range `Проверить выделение`
- selected range vs active-cell distinction
- current-table Check and current-table Autorun
- Run report title behavior
- banner letters on a simple banner

### 2. `banner-merged-waves.xlsx` - to be created

Required contents:

- one table with multi-row banner
- vertically merged or merged-like banner labels
- banner contains `Всего` / Total
- wave-like banner structure suitable for auto wave detection
- enough data to produce banner markers and data-cell markers
- workbook notes describing expected banner marker placement and rerun cleanup expectations

Used for:

- `Проставлять буквы в баннере`
- `Учитывать структуру баннера`
- `Автоматически определять волны`
- rerun cleanup in vertically merged banner cells
- data-cell marker cleanup after banner-marker runs

### 3. `multi-table-sheet.xlsx` - to be created

Required contents:

- one sheet containing at least two tables
- at least one empty row between tables
- one table with a clear title
- one table without a title
- at least one table banner contains `Всего`
- active cell can be placed clearly inside each table

Used for:

- multi-table selected range blocked
- current-table Check vs selected-range Check distinction
- `Текущий лист`
- `Вся книга`
- Content generation
- safe no-title fallback `Таблица N`
- broad multi-table selection guard

### 4. `mixed-metrics.xlsx` - to be created

Required contents:

- one workbook with proportions block
- one means block with SD or variance and Base
- one NPS block
- shared base or separate bases documented in notes
- at least one table title and one no-title table variant if practical

Used for:

- Check output across mixed table structures
- Run report title or fallback behavior
- blocked/skipped rows visibility in Run report
- workbook traversal smoke across heterogeneous tables

### 5. `generated-sheets-guard.xlsx` - optional file, may be produced during smoke

Required contents:

- a workbook where `Content` and `Run report` either already exist or are created during smoke
- at least one source table available for normal processing

Used for:

- generated-sheet ignore/block behavior
- no duplicate generated sheets
- Content update behavior
- Run report update behavior

## Execution notes

- For `Расчёт`, the main target is the selected range. Keep the active cell inside or outside the table as needed to prove selection-driven behavior.
- For `Проверка -> Текущая таблица` and `Автозапуск -> Текущая таблица`, the main target is the active-cell detected table. Do not rely on a broad selected range unless the step explicitly tests a guard.
- For blocked cases, confirm no Excel mutation happens before the message.
- For rerun cases, confirm stale banner markers and stale data-cell markers are removed before new output is written.
- Where exact significance letters depend on fixture data, use the workbook note as the oracle. Do not invent ad-hoc expected letters during the run.

## Smoke steps by flow

Mark each step `Pass`, `Fail`, or `N/A`, and add a short note with the sheet name used.

### A. `Расчёт` selected-range workflows

Workbook: `basic-proportions.xlsx`

- [ ] Select only the numeric data area and click `Рассчитать`.
Expected: significance is written for the selected-range workflow only; output matches workbook notes; no active-cell-only detection behavior is required.

- [ ] On the same table, click `Очистить значимости`.
Expected: RIT markers and formatting added by Run are removed; cleaned cells are restored per current output behavior; no stale data-cell markers remain.

- [ ] Select only the numeric data area and click `Проверить выделение`.
Expected: check result is based on the selected range, not on whatever table the active cell might imply elsewhere.

- [ ] Put the active cell inside table A, then select the numeric range of table B and run `Проверить выделение`.
Expected: result follows the selected range, proving selected range and active-cell current-table flows are distinct.

Workbook: `multi-table-sheet.xlsx`

- [ ] Select a range spanning two tables with empty rows between them and click `Рассчитать`.
Expected: blocked by the selected-range normalization guard; no table is processed; no writes happen.

- [ ] Repeat the same broad selection and click `Очистить значимости`.
Expected: blocked by the same guard; no cleanup is applied to either table.

- [ ] Create a Ctrl+Click non-contiguous selection and run `Рассчитать` or `Проверить выделение`.
Expected: controlled unsupported-selection message; no partial processing; no Excel mutation.

### B. `Проверка`

Workbook: `basic-proportions.xlsx`

- [ ] Place the active cell inside the table and run `Проверка -> Текущая таблица`.
Expected: active-cell detected table is checked successfully even if no explicit selected range is prepared.

Workbook: `multi-table-sheet.xlsx`

- [ ] Place the active cell inside table 1 and run `Проверка -> Текущая таблица`.
Expected: only table 1 is checked.

- [ ] Place the active cell inside table 2 and run `Проверка -> Текущая таблица`.
Expected: only table 2 is checked.

- [ ] Run `Проверка -> Текущий лист`.
Expected: sheet-level traversal processes supported tables on the sheet and produces understandable output for each table.

- [ ] Run `Проверка -> Вся книга`.
Expected: workbook-level traversal processes supported source tables across the workbook.

- [ ] Place the active cell outside any table and run `Проверка -> Текущая таблица`.
Expected: clear no-table message; no generated output is created from a non-table position.

Workbook: `generated-sheets-guard.xlsx` or the workbook after Content / Run report creation

- [ ] Place the active cell on `Content` and run `Проверка -> Текущая таблица`.
Expected: generated-sheet message; sheet is not treated as a source table.

- [ ] Place the active cell on `Run report` and run `Проверка -> Текущая таблица`.
Expected: generated-sheet message; sheet is not treated as a source table.

Workbook: `basic-proportions.xlsx` and `multi-table-sheet.xlsx`

- [ ] Enable Run report, then run `Проверка -> Текущая таблица` on a titled table.
Expected: Run report row uses the actual table title.

- [ ] Run `Проверка -> Текущая таблица` on a no-title table whose banner contains `Всего`.
Expected: title fallback is safe and does not use `Всего` as the table title.

### C. `Автозапуск`

Workbook: `basic-proportions.xlsx`

- [ ] Place the active cell inside the table and run `Автозапуск -> Текущая таблица -> Рассчитать`.
Expected: current-table scope processes the active-cell detected table only.

- [ ] On the same table, run `Автозапуск -> Текущая таблица -> Очистить значимости`.
Expected: current-table cleanup works on the detected table only.

Workbook: `multi-table-sheet.xlsx`

- [ ] Run `Автозапуск -> Текущий лист -> Рассчитать`.
Expected: supported tables on the active sheet are processed.

- [ ] Run `Автозапуск -> Текущий лист -> Очистить значимости`.
Expected: cleanup runs across the supported tables on the active sheet.

- [ ] Run `Автозапуск -> Вся книга -> Рассчитать`.
Expected: supported source tables across the workbook are processed.

- [ ] Run `Автозапуск -> Вся книга -> Очистить значимости`.
Expected: cleanup runs across supported tables across the workbook.

- [ ] With a broad multi-table selection present, place the active cell inside one table and run `Автозапуск -> Текущая таблица`.
Expected: current-table guard behavior stays controlled; broad selection does not silently turn into multi-table processing.

Workbook: `generated-sheets-guard.xlsx` or the workbook after Content / Run report creation

- [ ] Run current-table, sheet, and workbook Autorun flows with generated sheets present.
Expected: generated sheets are ignored or blocked as appropriate and are not processed as source tables.

### D. `Оглавление`

Workbook: `multi-table-sheet.xlsx` and any multi-sheet workbook used above

- [ ] Run `Оглавление` from a workbook context.
Expected: workbook-only behavior; a `Content` sheet is created if missing.

- [ ] Run `Оглавление` again after the workbook already has `Content`.
Expected: the existing `Content` sheet is updated, not duplicated.

- [ ] Inspect the title column links in `Content`.
Expected: table links are written in the title column and navigate to the intended table.

- [ ] Inspect the entry for a no-title table.
Expected: fallback is safe `Таблица N`; it does not reuse banner label `Всего` as a title.

- [ ] If backlinks are enabled in the current build, follow at least one backlink from a source table.
Expected: backlink is present and navigates back to `Content`.

- [ ] Confirm generated sheets are not listed as source tables in `Content`.
Expected: `Content` and `Run report` are excluded from source-table entries.

### E. Run report

Workbook: `basic-proportions.xlsx`, `multi-table-sheet.xlsx`, `mixed-metrics.xlsx`

- [ ] Enable Run report and execute at least one current-table Check flow.
Expected: `Run report` is created if missing.

- [ ] Execute another eligible flow with Run report still enabled.
Expected: `Run report` is updated and not duplicated.

- [ ] Inspect rows written for titled and no-title tables.
Expected: title column uses the real title where available and safe fallback where not.

- [ ] Inspect rows describing skipped or blocked tables.
Expected: skipped or blocked rows are understandable to a human tester and identify why the table was not processed.

- [ ] Confirm no duplicate generated sheets appear after repeated runs.
Expected: still only one `Run report` and one `Content` if both features were used.

### F. Banner markers

Workbook: `banner-merged-waves.xlsx`

- [ ] Turn on `Проставлять буквы в баннере` and run the relevant calculation flow.
Expected: banner letters are written where the current behavior allows them.

- [ ] Turn on `Учитывать структуру баннера` and rerun.
Expected: banner-aware behavior follows the structured banner rules for the fixture.

- [ ] Turn on `Автоматически определять волны` and rerun.
Expected: wave-aware behavior follows the fixture notes; no stale output remains from the prior run.

- [ ] Rerun the same table with a different combination of those settings.
Expected: vertically merged or merged-like banner cells do not retain stale markers from the previous mode.

- [ ] After banner-marker runs, run `Очистить значимости`.
Expected: data-cell marker cleanup still works and banner-related stale markers are not left behind.

## Pass / fail log

Use one line per workbook or flow:

| Area | Workbook | Result | Notes |
|---|---|---|---|
| `Расчёт` |  |  |  |
| `Проверка` |  |  |  |
| `Автозапуск` |  |  |  |
| `Оглавление` |  |  |  |
| `Run report` |  |  |  |
| Banner markers |  |  |  |

## Release-candidate sign-off

- [ ] Smoke pack executed on target Excel environment
- [ ] Selected-range flows verified separately from active-cell current-table flows
- [ ] Generated sheets confirmed ignored or blocked by table-processing flows
- [ ] `Content` create/update behavior verified
- [ ] `Run report` create/update behavior verified
- [ ] No-title fallback verified as safe `Таблица N`
- [ ] Non-contiguous selection message verified
- [ ] Broad multi-table selection guard verified
- [ ] Banner rerun cleanup verified on vertically merged or merged-like headers
- [ ] No blocker found for release candidate

Release decision notes:

- Candidate:
- Smoke status:
- Blockers:
- Follow-up issues:
