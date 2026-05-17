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
