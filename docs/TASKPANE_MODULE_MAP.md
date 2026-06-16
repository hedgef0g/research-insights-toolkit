# Taskpane Module Map

This checkpoint records the current post-Run/Clear extraction taskpane layout.
It distinguishes current module ownership from later extraction candidates.

## Current Modules

| Module | Current responsibility |
| --- | --- |
| `src/taskpane/taskpane.js` | High-level Office.js and UI orchestrator. Wires `Office.onReady`, buttons, tabs, settings reads, selected-range entry points, batch/report coordination, generated sheets, and remaining Check/Content/inventory controller logic. |
| `src/taskpane/run-pipeline.js` | Shared per-table Run executor plus Run calculation helpers. Runs interpretation, banner detection, block building, significance calculation, body writes, and returns deferred recolor/footnote jobs for callers. |
| `src/taskpane/run-preflight.js` | Run marker-overflow preflight. Computes exact per-range overflow where needed and provides batch preflight before operation-level writes. |
| `src/taskpane/run-banner-markers.js` | Banner marker writer for Run. Clears stale banner labels and queues/applies banner marker updates separately from data-cell marker writes. |
| `src/taskpane/run-footnotes.js` | Run footnote job builder and application helpers. Builds pure footnote jobs and applies footnotes after table writes. |
| `src/taskpane/clear-pipeline.js` | Clear pipeline for selected range, current table, sheet, and workbook scopes. Owns marker/fill cleanup, banner marker cleanup, footnote removal jobs, and Clear batching helpers. |
| `src/taskpane/taskpane-dialogs.js` | Marker-overflow dialog and shared decider helpers. Keeps dialog state out of Run pipeline logic. |
| `src/taskpane/taskpane-action-warnings.js` | UI warning collection and rendering for incompatible action/settings combinations. |
| `src/taskpane/taskpane-status.js` | Status/check/inventory DOM message helpers, banner/guardrail message formatting, and batch progress text. |
| `src/taskpane/taskpane-performance.js` | Performance timing flags, elapsed-time helpers, logging, and writer/banner diagnostic aggregation. |
| `src/taskpane/taskpane-settings.js` | Settings metadata, defaults, settings panel initialization, tooltips, tabs, local persistence, and settings DOM hydration. |
| `src/taskpane/selected-range-interpreter.js` | Office.js adapter that normalizes a selected/candidate range into pass-through, normalized, or blocked states and loads labels/banner context for Run and Check. |
| `src/taskpane/active-cell-resolver.js` | Read-only current-table resolver from active cell using worksheet inventory scanning. |
| `src/taskpane/taskpane-check-formatters.js` | Pure Check display formatters for issues, calculation blocks, metric types, and banner summaries. |
| `src/taskpane/taskpane-run-report-formatters.js` | Pure Run/Check report row label and detail formatters. It does not write generated sheets. |
| `src/taskpane/taskpane-inventory-formatters.js` | Pure or near-pure Content/inventory text, row, hyperlink, and backlink label formatting helpers. |
| `src/taskpane/taskpane-inventory-scan.js` | Inventory scan constants and small scan/skipped-sheet formatting helpers. |
| `src/taskpane/localization.js` | Taskpane translation dictionaries, language selection, persistence, and DOM text refresh helpers. |

## Current Controller Boundaries

- `taskpane.js` is still intentionally larger than the helper modules because it
  owns high-level workflow orchestration and generated sheet writes.
- Run/Clear extraction boundaries should be considered stable. Future PRs should
  not move functions back into `taskpane.js` or reopen those modules without a
  specific behavior-preserving reason.
- Check and Content/inventory still have substantial orchestration in
  `taskpane.js`; their extracted modules are mostly formatters and small
  adapters today.
- Generated `Content` and `Run report` sheet writes are still taskpane
  orchestration concerns until report/content controller extraction is scoped.
- Core modules under `src/core/` remain platform-neutral and must not import
  taskpane modules.

## Later Extraction Candidates

- Batch controller for workbook/current-sheet/current-table Run and Clear loops.
- Report controller for `Run report` sheet creation and writing.
- Check pipeline/controller for read-only selected-range, current-table,
  current-sheet, and workbook Check flows.
- Content/inventory controller for workbook inventory collection, Content sheet
  generation, and backlink maintenance.
- Final `taskpane.js` size/import audit after the remaining controllers are
  extracted.
- Selected-range normalization edge-case follow-up if unsafe broad selections or
  ambiguous boundaries reappear.
