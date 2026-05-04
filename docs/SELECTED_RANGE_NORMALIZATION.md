# Selected Range Normalization Spec

## Overview

The Research Insights Toolkit (RIT) currently relies on a manual selected-range workflow where the user is expected to precisely select the numeric data area in Excel. As a temporary safety net, warning-only selected range guardrails exist.

The desired end-state is a **forgiving full-table selection** model: the user selects the visible table, and RIT identifies the data, labels, and banner internally.

**Important Note:** This document is an architecture spec. Selected range normalization is **not** currently implemented. The existing manual selected-range workflow remains the stable baseline.

## 1. Supported Selection Patterns

The normalization engine must recognize and correctly partition the following selection patterns:
- **Strict Data Selection:** The current baseline. The user selects only the numeric body.
- **Data + Row Labels:** The user selects the numeric body and the left-hand text labels.
- **Data + Banner:** The user selects the numeric body and the top banner headers.
- **Full Table Selection:** The user selects the entire visual table, including row labels, banner headers, and the numeric body.
- **Full Table + Title/Headers:** The user selects the entire visual table along with preceding title/header rows.

## 2. Normalized Selection Data Model

The normalized model separates the user's raw visual selection into logical areas required by the core engines.

- `originalRange`: The raw Excel coordinates.
- `normalizedDataWindow`: The subset of `originalRange` containing the numeric data body.
- `rowLabelArea`: The left-hand column(s) containing metric names.
- `bannerArea`: The top row(s) containing group definitions.
- `titleArea`: (Optional) Leading rows containing report titles or metadata, ignored in calculations.

## 3. Original Selection vs. Normalized Data Window

- **Original Selection:** The exact cells the user highlighted. This is passed from `taskpane.js` to the core.
- **Normalized Data Window:** A computed sub-grid of the original selection. All subsequent block construction, statistical calculations (`significance.js`), and cell writes (`excel-writer.js`) operate *exclusively* within this window.

## 4. Row Label Area

- Defined as the leftmost column(s) within the `originalRange` but outside the `normalizedDataWindow`.
- Usually contains text identifying the metric type (e.g., "NPS", "Mean", "%").
- Normalization must identify the boundary between string-heavy columns and numeric-heavy columns.

## 5. Banner Area

- Defined as the top row(s) within the `originalRange` but outside the `normalizedDataWindow` (and below any `titleArea`).
- Contains categorical column headers.
- Normalization must detect where header strings stop and the numeric data matrix begins.

## 6. Title/Header Area

- Defined as any rows at the very top of the `originalRange` that span the table but do not participate in the banner hierarchy.
- Often contain single-cell strings or wide merged cells (e.g., "Q4 Brand Tracker").
- These rows are excluded from both the `bannerArea` and the `normalizedDataWindow`.

## 7. Coordinate Mapping Back to Worksheet

Because core modules like `significance.js` will operate on the `normalizedDataWindow` grid, any markers or fills generated must map back to the correct absolute Excel coordinates.
- The system must maintain a mapping function: `(normalizedRow, normalizedCol) -> (excelRow, excelCol)`.
- `excel-writer.js` must be updated to use this mapping when translating the logical calculation blocks into physical Excel cell updates.

## 8. Confidence and Ambiguity Policy

- **High Confidence:** If the normalization engine clearly identifies the boundaries (e.g., solid numeric block surrounded by text), it proceeds silently.
- **Low Confidence / Ambiguity:** If boundaries are unclear (e.g., heavy mix of numbers and text in label columns), the system should fall back to the **Strict Data Selection** assumption or halt and warn the user.
- **Rule of Least Surprise:** When in doubt, prefer the current manual behavior over guessing incorrectly.

## 9. Relationship to Current Warning-Only Guardrails

- Current warning-only guardrails alert the user if they suspect a bad selection (e.g., "Selection looks like a full table").
- These guardrails will be phased out and replaced by the normalization engine, which will handle the selection gracefully instead of warning.

## 10. Relationship to Table Preview Model

- Normalization is a prerequisite for a robust Table Preview feature.
- The output of the normalization model (`normalizedDataWindow`, `rowLabelArea`, `bannerArea`) will populate the Preview UI, allowing the user to verify the boundaries before running calculations.

## 11. Implementation Phases

We recommend a **model-first implementation path**:
1.  **Phase 1: Pure Model & Telemetry.** Build the normalization engine. Run it silently alongside the current workflow. Log its output versus the actual user selections to gauge accuracy without affecting the UI or calculations.
2.  **Phase 2: Preview UI Integration.** Expose the normalized model in a read-only Table Preview UI. Allow users to confirm the engine's guesses.
3.  **Phase 3: Opt-in Execution.** Add a setting to "Auto-detect table boundaries" that wires the normalized model into the core execution flow.
4.  **Phase 4: Default Behavior.** Make full-table selection the default, removing the old warning-only guardrails.

## 12. First Safe Coding Issue

**Issue: Create the Normalization Engine (Silent Mode)**
- Implement the boundary detection logic in a new file (e.g., `src/core/range-normalizer.js`).
- Write unit tests for the engine using mock Excel data grids covering the supported selection patterns.
- Do *not* wire this into `taskpane.js` execution yet.

## 13. Major Regression Risks

Implementing this feature carries significant risks:
- **Calculation Shifts:** If the `normalizedDataWindow` is off by one row or column, all subsequent significance calculations will be performed on the wrong cells, potentially overwriting labels with markers.
- **Coordinate Mapping Failures:** Incorrect mapping back to the worksheet could result in formatting or letters being applied to the wrong areas of the spreadsheet.
- **Banner Detection Breakage:** `banner-detector.js` might fail if fed an incorrect `bannerArea`.
- **Metric Detection Breakage:** `metric-detector.js` relies heavily on left-aligned row labels; an incorrect `rowLabelArea` will break metric type identification.

## 14. Smoke-Test Scenarios

Before any normalization logic is merged into the execution path, these scenarios must pass:
1.  **Baseline strict data:** Select only numbers. Ensure calculations map perfectly as they do today.
2.  **Clean full table:** Select labels, banner, and data. Ensure markers land only in the data area.
3.  **Full table with title:** Select a table including a top-row report title. Ensure the title is ignored and the banner is correctly identified below it.
4.  **Numeric labels:** A table where the row labels are numeric (e.g., years: 2021, 2022). The engine must not consume these as the data body.
5.  **Small base tables:** Ensure small-base formatting correctly targets the `normalizedDataWindow` and doesn't bleed into the `rowLabelArea`.
