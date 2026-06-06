# Architecture Refactor Plan 2026

## Goal
Significantly lighten `src/taskpane/taskpane.js` by extracting distinct responsibilities into dedicated controllers and pipelines, clarifying module boundaries, and removing duplicated code, all without breaking the manual selected-range workflow or altering statistical logic.

## Current State & Diagnostics

`src/taskpane/taskpane.js` is still the main taskpane orchestrator, but parts of
the original inline responsibilities have been extracted. Run per-table executor
logic now lives in `src/taskpane/run-pipeline.js`, Clear pipeline logic now lives
in `src/taskpane/clear-pipeline.js`, and settings/performance helpers have moved
to their dedicated modules.

### Key Layering Violations to Address
1.  **Remaining Run Split:** `runSignificanceFromSelection` (manual run) still contains its selected-range calculation flow, while the autorun per-table executor lives in `run-pipeline.js`. Further extraction must stay behavior-preserving and must not change selected-range normalization.
2.  **Duplicate Normalization:** `clearSignificanceFromSelection` and `clearSignificanceForRangeInContext` duplicate `detectEmbeddedLabelColumns` and `detectLeadingEmptyColumns` logic found in `selected-range-interpreter.js` to resolve the clear target.
3.  **UI in Orchestrator:** Tooltip mapping, checkbox bindings, and performance aggregators reside in `taskpane.js` instead of their respective modules (`taskpane-settings.js` and `taskpane-performance.js`).

## Proposed Architecture (Taskpane Layer)

-   `src/taskpane/taskpane.js`: Lean UI controller. Wires events (`Office.onReady`), binds button clicks, and orchestrates high-level flows by calling pipelines.
-   `src/taskpane/run-pipeline.js`: Shared per-table Run executor for autorun flows, plus dispatcher helpers used by the manual Run path. Handles interpretation, significance calculations, and formatting for extracted table execution.
-   `src/taskpane/clear-pipeline.js`: Shared executor for manual and auto Clear flows.
-   `src/taskpane/batch-controller.js`: Manages workbook/worksheet loops and batch result accumulation.
-   `src/taskpane/report-controller.js`: Manages the creation and population of `Content` and `Run report` sheets.
-   `src/taskpane/taskpane-settings.js`: Fully owns UI state, bindings, and persistence for settings.
-   `src/taskpane/taskpane-status.js` & `taskpane-check-formatters.js`: Unchanged; continue shaping user-visible text.
-   `src/taskpane/selected-range-interpreter.js`: Unchanged in purpose, but enhanced to provide a shared clear-target resolution helper.

## Execution Plan: PR Slicing

This sequence prioritizes low-risk extractions before tackling the complex Run/Clear pipelines.

### PR 1: Settings and Performance Cleanup (Low Risk)
**Goal:** Extract isolated UI bindings and diagnostics aggregators out of `taskpane.js`.
**Target Files:** `src/taskpane/taskpane.js`, `src/taskpane/taskpane-settings.js`, `src/taskpane/taskpane-performance.js`
**Function Inventory to Move:**
-   Move to `taskpane-settings.js`:
    -   `initializeSettingsPanel`
    -   `initializeSettingsTooltips`
    -   `initializePreviousColumnTotalWarningPlacement`
    -   `bindMutuallyExclusiveCheckboxes`
-   Move to `taskpane-performance.js`:
    -   `createAggregatedWriterPerfDetails`
    -   `mergeWriterPerfDetails`
    -   `mergeRangeAreasProjection`
    -   `mergeFillRangeAreasProjection`
    -   `createAggregatedBannerPerfDetails`
    -   `roundBannerDiagnosticRatio`
    -   `bannerScanAreasOverlap`
    -   `mergeBannerScanAreaDiagnostics`
    -   `mergeBannerPerfDetails`
**Testing:** Manual UI verification of the settings panel and tooltips.

### PR 2: Unify Target Range Resolution (Medium Risk)
**Goal:** DRY up the clearing logic by sharing the target resolution logic.
**Target Files:** `src/taskpane/taskpane.js`, `src/taskpane/selected-range-interpreter.js`
**Action:**
1.  Extract the bounding logic from `clearSignificanceFromSelection` and `clearSignificanceForRangeInContext` (which checks `detectEmbeddedLabelColumns` and `detectLeadingEmptyColumns`).
2.  Create `resolveClearTargetBodyRange` in `selected-range-interpreter.js`.
3.  Update the clear functions in `taskpane.js` to use this new helper.
**Testing:** Ensure clearing works correctly on sloppily selected tables (with headers/labels included) and strict numeric selections.

### PR 3: Extract Clear Pipeline (Medium Risk)
**Goal:** Move clearing logic out of `taskpane.js`.
**Target Files:** `src/taskpane/taskpane.js`, `src/taskpane/clear-pipeline.js` (new)
**Action:**
1.  Extract `clearSignificanceForRangeInContext` and the core clearing loop from `clearSignificanceFromSelection` into `src/taskpane/clear-pipeline.js`.
2.  Update `taskpane.js` to call the pipeline.

### PR 4: Extract Run Pipeline (High Risk, partially completed by PR #328/#329)
**Goal:** Consolidate the calculation loop for manual and autorun flows.
**Target Files:** `src/taskpane/taskpane.js`, `src/taskpane/run-pipeline.js` (new)
**Contract:** See `docs/RUN_PIPELINE_CONTRACT.md` before moving Run code. That
document records the current entry points, write barriers, job objects, report
row expectations, and `context.sync` boundaries that extraction must preserve.
**Action:**
1.  PR #328 extracted the shared per-table executor to `run-pipeline.js`.
2.  PR #329 moved `calculateBlockResults` and `getFirstBannerStructureError` to `run-pipeline.js`.
3.  Remaining Run cleanup should be small, audited, and behavior-preserving before any further helper extraction.
**Critical Guardrails:** The `markerOverflowDecider` logic must remain capable of pausing Excel execution via UI dialog without breaking batch promises. Do not change `context.sync` placement.
**Testing:** Exhaustive manual smoke testing of Run and Autorun. Run all existing unit tests.

### PR 5: Extract Batch & Report Controllers (Medium Risk)
**Goal:** Extract sheet looping and generated sheet writes.
**Target Files:** `src/taskpane/taskpane.js`, `src/taskpane/batch-controller.js` (new), `src/taskpane/report-controller.js` (new)
**Action:**
1.  Move `runAutoSignificance`, `runAutoCurrentTableSignificance`, `runCurrentSheetSignificance`, `runWorkbookCheck` to `batch-controller.js`.
2.  Move `writeRunReportSheet`, `ensureRunReportWorksheet`, `writeInventoryContentSheet`, `ensureInventoryContentWorksheet` to `report-controller.js`.

## Pre-Implementation Requirements
-   **Unit Tests:** Ensure `selected-range-interpreter.js` has coverage for the leading column stripping algorithms before PR 2.
-   **Regression Testing:** Use `docs/GOLD_STANDARD_TEST_SUITE.md` to verify base/block detection and footnote placements are unaffected by PR 3/4.
