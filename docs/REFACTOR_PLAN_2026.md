# Architecture Refactor Plan 2026

## Goal

Lighten `src/taskpane/taskpane.js` by extracting distinct taskpane responsibilities
into dedicated modules while preserving the manual selected-range workflow,
selected-range normalization guardrail, Excel writer behavior, and statistical
calculation logic.

## Current Checkpoint: PR6A

The Run/Clear extraction wave is complete enough that the project should treat
the separated modules as current architecture, not as temporary staging. Future
work should avoid reopening already separated Run/Clear modules unless there is
a clear scoped reason.

Completed extraction areas:

- Settings/UI initialization moved to `src/taskpane/taskpane-settings.js`.
- Performance diagnostics helpers moved to `src/taskpane/taskpane-performance.js`.
- Clear execution moved to `src/taskpane/clear-pipeline.js`.
- Marker-overflow dialog helpers moved to `src/taskpane/taskpane-dialogs.js`.
- Action warning helpers moved to `src/taskpane/taskpane-action-warnings.js`.
- Status helpers were tightened in `src/taskpane/taskpane-status.js`.
- Run marker-overflow preflight moved to `src/taskpane/run-preflight.js`.
- The shared per-table Run executor and Run dispatcher helpers moved to
  `src/taskpane/run-pipeline.js`.
- Run footnote job/application helpers moved to `src/taskpane/run-footnotes.js`.
- Run banner marker writing moved to `src/taskpane/run-banner-markers.js`.

`src/taskpane/taskpane.js` remains the high-level UI and Office.js orchestrator.
It still owns startup wiring, major workflow entry points, batch/report flow
coordination, selected-range reads, generated sheet orchestration, and the
remaining Check and Content/inventory controller logic.

For the current module map, see `docs/TASKPANE_MODULE_MAP.md`.

## Current Architecture Principles

- Keep core modules free of Office.js and taskpane imports.
- Keep `src/taskpane/taskpane.js` focused on UI/event wiring and workflow
  orchestration.
- Keep extracted modules narrow: pipeline helpers should not silently absorb
  unrelated UI, report, or settings responsibilities.
- Preserve Run/Clear selected-range normalization before any Excel mutation.
- Preserve `context.sync` boundaries unless a future issue explicitly scopes
  and validates that change.
- Do not mix architecture extraction with behavior changes.

## Remaining Extraction Candidates

These are candidates for later PRs, not current PR6A work:

- **Batch controller extraction:** move workbook/current-sheet/current-table loop
  orchestration out of `taskpane.js` after auditing status timing, fallback
  behavior, report row accumulation, and shared `Excel.run` context usage.
- **Report controller extraction:** move generated `Run report` sheet creation
  and writing after the report row contract is stable.
- **Check pipeline extraction:** move read-only selected-range/current-table/
  sheet/workbook Check orchestration after preserving preview/status/report
  behavior.
- **Content/inventory controller extraction:** move Content sheet generation,
  backlink maintenance, and inventory collection orchestration after separating
  pure formatting from workbook writes.
- **Final `taskpane.js` size/import audit:** once the controller extractions are
  complete, audit imports and remaining helpers for obsolete dependencies.
- **Selected-range normalization edge cases:** if broad-selection or boundary
  ambiguity issues reappear, extend `docs/SELECTED_RANGE_NORMALIZATION.md` and
  the normalizer contract before changing Run/Clear behavior.

## Historical PR Slices

The original execution plan is now historical context. PRs #317, #318, #321,
#323, #324, #325, #326, #327, #328, #329, #331/#332, and #333 completed the
settings, performance, Clear, dialog, warning, status, Run preflight, Run
pipeline, Run footnote, and Run banner marker extraction wave.

Do not treat pre-PR6A "move X to Y" instructions from older drafts as current
open work unless they are repeated in a new issue.

## Validation Expectations

For documentation-only checkpoints:

- `npm test`
- `npm run build`
- `git diff --check`

Manual Excel smoke testing is not required for docs-only changes. Future
behavior-sensitive extraction PRs should identify affected
`docs/GOLD_STANDARD_TEST_SUITE.md` cases and include manual smoke notes for the
affected Run/Clear/Check/Content paths.
