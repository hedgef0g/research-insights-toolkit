# Architecture Refactor Plan 2026

## Goal
Significantly lighten `src/taskpane/taskpane.js` by extracting distinct responsibilities into dedicated controllers and pipelines, clarifying module boundaries, and removing duplicated code, all without breaking the manual selected-range workflow or altering statistical logic.

## Current State & Diagnostics

`src/taskpane/taskpane.js` is currently acting as a "God Object," inline-implementing the sequence of core logic calls for Run, Check, and Clear, while also managing UI initialization and performance aggregation.

### Key Layering Violations to Address
1.  **Duplicate Pipelines:** `runSignificanceFromSelection` (manual run) and `runSignificanceForRangeInContext` (auto run) contain duplicated logic for interpreting ranges, checking overflow, detecting metrics, and calculating blocks.
2.  **Duplicate Normalization:** `clearSignificanceFromSelection` and `clearSignificanceForRangeInContext` duplicate `detectEmbeddedLabelColumns` and `detectLeadingEmptyColumns` logic found in `selected-range-interpreter.js` to resolve the clear target.
3.  **UI in Orchestrator:** Tooltip mapping, checkbox bindings, and performance aggregators reside in `taskpane.js` instead of their respective modules (`taskpane-settings.js` and `taskpane-performance.js`).

## Proposed Architecture (Taskpane Layer)

-   `src/taskpane/taskpane.js`: Lean UI controller. Wires events (`Office.onReady`), binds button clicks, and orchestrates high-level flows by calling pipelines.
-   `src/taskpane/run-pipeline.js`: Shared executor for manual and auto Run flows. Handles interpretation, significance calculations, and formatting.
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

### PR 4: Extract Run Pipeline (High Risk)
**Goal:** Consolidate the calculation loop for manual and autorun flows.
**Target Files:** `src/taskpane/taskpane.js`, `src/taskpane/run-pipeline.js` (new)
**Action:**
1.  Identify the ~100 lines of shared logic between `runSignificanceFromSelection` and `runSignificanceForRangeInContext`.
2.  Extract into `runPipeline(context, range, values, text, interpretation, settings, decider)` in `run-pipeline.js`.
3.  Update `taskpane.js` orchestrators to call this pipeline.
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
