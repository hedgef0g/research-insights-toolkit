# Architecture Overview

## Product Context
Research Insights Toolkit is an Excel-first Office Add-in designed for research tables and statistical significance. It aims to be the most practical spreadsheet significance tool for research managers and insights teams.

## Current Stable Workflow
The current stable workflow relies on manual selected-range calculation. The user selects a table range manually in Excel and clicks "Run". The add-in then detects metrics, calculates significance, and writes markers directly into cells.
Automatic worksheet/workbook scanning is still not implemented.

## High-Level Pipeline
1. **Selection:** Read the user-selected table range from Excel.
2. **Detection:** Detect metric rows, block structures, and any banner definitions.
3. **Calculation:** Perform statistical significance tests based on detected metrics and structures.
4. **Output:** Write the calculated markers, letters, and formatting back to the Excel spreadsheet.

## Module Responsibilities
- `src/taskpane/taskpane.js` — Excel/UI controller. Orchestrates the flow. No statistical calculations. Owns Office.js interaction and selected range reads.
- `src/taskpane/selected-range-interpreter.js` — Excel/taskpane adapter that interprets the user's selected range into a stable object shared by Run and Check. See [Selected-Range Interpretation Adapter](#selected-range-interpretation-adapter) below.
- `src/core/metric-detector.js` — Detects metric rows and constructs block plans.
- `src/core/banner-detector.js` — Platform-independent banner structure detection.
- `src/core/significance.js` — Statistical tests, comparison routing, marker creation, and fill reasons.
- `src/core/stat-thresholds.js` — Determines statistical critical thresholds.
- `src/core/config/dictionary.config.js` — Config-driven metric and banner dictionaries.
- `src/core/excel-writer.js` — Handles writing values and formatting back to Excel. Owns data-cell output behavior (including numeric output preservation).
- Banner detection and banner-letter placement are separate concerns.

## Selected-Range Interpretation Adapter

`src/taskpane/selected-range-interpreter.js` is the Excel/taskpane adapter that converts a raw Office.js range selection into the stable interpretation object consumed by both Run and Check.

### What it owns

- Stripping significance markers from raw values (`cleanedValues`).
- Calling `normalizeSelectedRange` to classify the selection as pass-through, normalized, or blocked.
- Loading left-label values from the worksheet (immediately left of the selection, or from the sheet's leftmost columns when `labelsOnLeftSide` is on).
- Loading banner rows from the worksheet above the data range.
- Detecting embedded label/unit columns that the normalizer left inside the selection.
- Sanitizing banner cell text so previously written RIT markers do not corrupt detection on re-runs.
- Deriving `writeTargetRange` — the Excel Range that Run writes significance markers back into.
- Returning the canonical interpretation object (see invariants below).

### What it must not own

- Statistical calculations or significance tests — those belong in `src/core/significance.js`.
- Metric-row detection — that belongs in `src/core/metric-detector.js`.
- Banner structure detection — that belongs in `src/core/banner-detector.js`.
- Excel write-back (markers, fills) — that belongs in `src/core/excel-writer.js`.
- UI state, status messages, or taskpane DOM access — those belong in `src/taskpane/taskpane.js`.

### Run/Check flow (after PR #120)

Both Run and Check call `interpretSelectedRange()` with the same arguments and receive the same interpretation object. They diverge only after interpretation:

- **Run** — passes `valuesForCalculation`, `leftLabelValues`, and `bannerContext` to detection and significance, then writes markers back to `writeTargetRange` via `excel-writer.js`.
- **Check** — passes the same values to detection and significance, then builds a read-only preview model. It never writes to the sheet.

### Interpretation states

`interpretSelectedRange()` always returns one of three states:

| `state` | Meaning |
|---|---|
| `"passThrough"` | The selection is a clean numeric data area; used as-is. |
| `"normalized"` | A broader selection was safely decomposed to a clean numeric data area. |
| `"blocked"` | The selection is ambiguous or unsafe; `blockingMessage` explains why. Run and Check must stop before any Excel mutation when blocked. |

### Invariants

These invariants hold for every non-blocked interpretation result:

- `valuesForCalculation` width equals `writeTargetRange.columnCount`.
- `valuesForCalculation` width equals `bannerContext.selectedColumnCount` when a banner context is present.
- Label, unit, and header columns are excluded from `valuesForCalculation`.
- Label, unit, and header columns are excluded from the `writeTargetRange` (Run never writes into them).

### Current platform limitation

The adapter intentionally sits at the taskpane level, not in `src/core/`, because it depends on Office.js context and Range objects to load labels and banner rows from the worksheet. A future decomposition could extract a pure platform-neutral core that handles normalization and column detection without any Office.js dependency, leaving only the worksheet I/O calls in the adapter. That refactor should not be done until the adapter's contract is stable and tested. See `docs/DECISIONS.md` for the rationale.

## Core vs Office.js Boundary
- `src/core/*` modules (like `significance.js`, `metric-detector.js`, `banner-detector.js`) are pure JavaScript logic and should remain Office.js-free. They should not contain direct Office.js API calls or UI DOM access.
- **Architectural Constraints:** No statistical calculations in `taskpane.js` (UI orchestration only). No Excel range writing in `significance.js`. No UI DOM access in core modules.

## High-Risk Files
Changes to these files require extra care:
- `src/core/significance.js`
- `src/core/banner-detector.js`
- `src/core/excel-writer.js`
- `src/taskpane/taskpane.js`
- `src/core/metric-detector.js`

## Manual Selected-Range Workflow Guardrails
- **Do not break the manual selected-range workflow.** This is the core stable feature.
- Ensure the selected range bounds are respected during reading and writing.
- Selected range guardrails currently live in taskpane as a warning-only MVP. Future selected range normalization should be designed before implementation.

## Taskpane Module Boundaries

Recent extraction PRs reduced the size of `src/taskpane/taskpane.js`, but taskpane decomposition is not complete. The current goal is disciplined boundaries: keep Excel orchestration and workflow ownership in `taskpane.js`, while extracted helpers stay narrow and do not silently absorb business flow.

### Guardrails

- `src/taskpane/taskpane.js` still owns Office.js orchestration, `Excel.run` flows, `context.sync` sequencing, range and worksheet reads/writes, workflow wiring, and generated sheet writing.
- Generated sheet names `Content` and `Run report` remain taskpane orchestration concerns, not formatter-module concerns.
- Formatter modules should remain pure or near-pure. They may shape strings, labels, and compact presentation models, but must not call `Excel.run`, `context.sync`, or write workbook ranges.
- `taskpane-status.js` and `taskpane-settings.js` may touch DOM and `localStorage`, but must not own Excel workflow logic.
- `src/core/*` modules remain platform-neutral and must not import taskpane modules.
- Do not mix architecture extraction with behavior changes. If an extraction needs logic changes to "make it fit", stop and split that into a separate task.

### `src/taskpane/taskpane.js`

Responsibility:
Owns the taskpane controller layer. It wires buttons and tabs, runs the main Excel workflows, coordinates Run / Clear / Check / inventory flows, and decides when generated sheets are created or updated.

What belongs there:
- `Office.onReady(...)` startup and event wiring.
- `Excel.run(...)` entry points and orchestration helpers.
- Workbook and worksheet reads/writes.
- Flow control that combines settings, selected-range interpretation, core detection/calculation, banner writing, report writing, and status updates.
- Generated sheet orchestration for `Content` and `Run report`.

What must not be moved there:
- Statistical calculation internals from `src/core/significance.js`.
- Reusable pure formatting helpers that only shape text for Check, inventory, or report output.
- Generic settings persistence helpers that can stay isolated in `taskpane-settings.js`.

Examples currently in it:
- `runSignificanceFromSelection`
- `runCheckTable`
- `runWorkbookCheck`
- `runTableInventory`
- `writeBannerMarkersAboveSelectedRangeUsingBannerStructure`

### `src/taskpane/taskpane-status.js`

Responsibility:
Owns small taskpane status-panel and message-formatting helpers for user-visible status, check, inventory, banner, and selected-range guardrail messaging.

What belongs there:
- DOM writes to status, check, and inventory panels.
- Compact message composition for user-facing taskpane text.
- Shared message-code filtering for banner and resolver output.

What must not be moved there:
- `Excel.run` flows, worksheet reads, or workbook writes.
- Detection, normalization, or significance decisions.
- Workflow branching such as deciding whether Run or Check should continue.

Examples currently in it:
- `setStatusMessage`
- `setCheckMessage`
- `setInventoryMessage`
- `formatBannerUserMessages`
- `appendSelectedRangeGuardrailMessages`
- `buildCheckResolverMessage`

### `src/taskpane/taskpane-settings.js`

Responsibility:
Owns taskpane settings metadata, defaults, panel hydration, and local persistence helpers.

What belongs there:
- Canonical settings control metadata and default values.
- Applying settings to DOM controls and reading persisted settings.
- Settings panel collapse state and `localStorage` persistence helpers.

What must not be moved there:
- Excel workflow logic.
- Range interpretation, table scanning, or workbook mutation.
- Cross-flow orchestration decisions about Run / Check / autorun behavior.

Examples currently in it:
- `SETTINGS_CONTROL_CONFIG`
- `DEFAULT_CALCULATION_SETTINGS`
- `loadSavedSettingsIntoPanel`
- `applySettingsToPanel`
- `saveSettingsToLocalStorage`
- `initializeSettingsCollapse`

### `src/taskpane/taskpane-check-formatters.js`

Responsibility:
Formats Check-related summaries into user-visible text fragments and compact display lines for taskpane output.

What belongs there:
- Pure or near-pure string/label formatting.
- Small presentation summaries derived from already-computed blocks, issues, and banner metadata.

What must not be moved there:
- `Excel.run`, `context.sync`, or range access.
- Table detection, banner detection, normalization, or statistical logic.
- Decisions about whether a Check flow should read or write a sheet.

Examples currently in it:
- `formatSkippedCandidateDetail`
- `checkMetricTypesFromBlocks`
- `formatCheckUserVisibleIssues`
- `formatCheckCalculationBlocks`
- `formatCheckBannerSummary`

### `src/taskpane/taskpane-run-report-formatters.js`

Responsibility:
Formats labels and detail strings for generated Run report sheet rows.

What belongs there:
- Pure mapping from status / reason / item metadata into report-facing text.
- Small report-detail assembly helpers used by taskpane orchestration before writing the report sheet.

What must not be moved there:
- Report sheet creation or workbook writes.
- Inventory scanning, selected-range interpretation, or significance calculation.
- Any helper that needs `Excel.run` or worksheet context.

Examples currently in it:
- `runReportSkipReasonLabel`
- `runReportStatusLabel`
- `runReportMetricTypes`
- `formatIssueDetailsForReport`
- `runReportWarningDetails`

### `src/taskpane/localization.js`

Responsibility:
Owns lightweight taskpane UI localization and language persistence.

What belongs there:
- Translation dictionaries and language selection state.
- `data-i18n` and related DOM text refresh helpers.
- Language persistence independent of calculation settings persistence.

What must not be moved there:
- Excel workflow logic.
- Task-specific Run / Check orchestration.
- Formatting helpers whose main purpose is report/check content rather than language lookup.

Examples currently in it:
- `t`
- `setLanguage`
- `loadSavedLanguage`
- `applyI18n`
- `SUPPORTED_LANGUAGES`

### `src/taskpane/active-cell-resolver.js`

Responsibility:
Provides the read-only adapter that resolves "current table" scope from the active cell by scanning the active sheet and mapping the cursor to a candidate table.

What belongs there:
- Active-cell based table resolution.
- Generated-sheet exclusion for `Content` and `Run report`.
- Read-only candidate lookup and ambiguity / blocked results for current-table flows.

What must not be moved there:
- Run / Clear write-back logic.
- Report sheet writing.
- Statistical calculation or workbook mutation.

Examples currently in it:
- `resolveCurrentTableFromActiveCell`
- `RESOLVER_GENERATED_SHEET_NAMES`
- `RESOLVER_SCAN_CELL_LIMIT`

### `src/taskpane/selected-range-interpreter.js`

Responsibility:
Owns the Excel/taskpane adapter that converts a raw Office.js selection into the stable interpretation object shared by Run and Check.

What belongs there:
- Marker stripping and normalized selection interpretation.
- Loading left labels and banner rows from the worksheet when needed.
- Embedded label / helper-column trimming that aligns `valuesForCalculation`, `bannerContext`, and `writeTargetRange`.
- Read-only adapter logic that prepares a stable contract for later workflow steps.

What must not be moved there:
- Statistical calculation, block building, or banner-structure algorithms from core.
- Taskpane DOM updates or status-panel writes.
- Excel write-back, report sheet writing, or broader workflow orchestration.

Examples currently in it:
- `interpretSelectedRange`
- `loadLabelValuesForSelectedRange`
- `detectLeadingEmptyColumns`
- `detectEmbeddedLabelColumns`
- `loadBannerContextForSelectedRange`
- `sanitizeBannerContextForDetection`

## Safe Next Extraction Candidates

Cautious candidates:
- Inventory / Content formatting helpers that only shape row text, summaries, or column values before taskpane writes them.
- Small status or Check display formatters that currently build user-visible strings inside `taskpane.js`.
- Narrow settings tooltip helpers, if extracted without moving settings-state orchestration or Excel workflow branching.

Do not extract yet:
- `runSignificanceFromSelection` / `runSignificanceForRange` unification.
- Generic `Excel.run` wrappers or `context.sync` wrappers.
- Writer or report-sheet writing flows.
- Selected-range normalization behavior.
- Calculation or statistical logic.

## Excel Writer Performance Guardrails
- **Do not return `excel-writer.js` to per-cell writes.** Maintain block-level or optimized writes to avoid degrading Excel performance.

## Current Agent Workflow
- Keep changes small and isolated.
- Prefer one issue = one PR.
- Do not mix refactor and feature work in the same PR.
- Preserve current UX status-message rules.
- If broader changes seem needed, stop and describe them instead of implementing.
- Do not rewrite `significance.js` wholesale.
- Never let agents merge directly; the human owner makes the final merge decision.

## Recommended Future Architecture Direction
- Shared text normalization utilities (planned, replacing raw values).
- “Проверить таблицу” / table preview mode.
- Implementation of a smoke test checklist.
- Adding a base placement setting.
- Eventual worksheet/workbook auto-scan.
- Row-wise comparisons postponed as a separate product mode.
