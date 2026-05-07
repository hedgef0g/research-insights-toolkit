# Gold Standard Test Suite Plan

**Status:** Plan only. No Excel workbooks or expected-output files exist yet.

**Source of truth for coverage:** docs/TABLE_STRUCTURE_MATRIX.md

---

## 1. Purpose

Manual smoke tests and regression checklists (docs/TEST_CASES.md) are sufficient for
catching obvious regressions but cannot verify statistical correctness. A gold standard
suite adds:

- fixed input tables with exact numeric values;
- pre-computed expected significance results;
- explicit expected marker assignments per cell;
- explicit expected row-type assignments per row.

The suite is used to:

- confirm that significance calculations match expected results after any logic change;
- confirm that metric/banner detection assigns the correct type to every row;
- confirm that marker writing places markers in the correct cells only;
- provide a regression target for agents performing code review or PR sign-off.

The suite does NOT replace manual smoke testing. Both are required.

---

## 2. Scope and exclusions

### In scope

- Table structures with status SUPPORTED or PARTIAL in TABLE_STRUCTURE_MATRIX.md.
- All metric types: proportions, means, NPS variants.
- Banner structures: no banner, one-level, two-level, wave.
- Total modes: all-vs-all, first-column-is-Total, compare-only-with-Total, exclude-Total.
- Comparison modes: all-vs-all, previous-column.
- Settings that affect calculation: confidence level, one-tailed/two-tailed.
- Edge cases: small base, zero base, out-of-range proportion, messy selection.

### Out of scope for this plan

- Automated test runner implementation (future issue).
- Excel workbook files (.xlsx) are not added in this issue.
- Expected-output files are not added in this issue; they are placeholders.
- Runtime code is not modified.
- Existing test behavior is not changed.
- UNSUPPORTED and FUTURE structures in TABLE_STRUCTURE_MATRIX.md are noted but not
  assigned test cases until the structures are implemented.

---

## 3. Test case format

Each test case is a self-contained record. The canonical form uses the following fields.

### 3.1. Header fields

    Case ID        A unique identifier in the format GST-NNN (e.g. GST-001).
    Title          One-line description of what the case tests.
    Matrix ref     Section(s) in TABLE_STRUCTURE_MATRIX.md that this case exercises.
    Table type     Metric type(s) present: Proportion, Mean, NPS, Mixed.
    Metric config  Active settings: confidence level, tailed mode, Total mode,
                   previous-column, small-base threshold, banner mode.
    Status         PLANNED | WORKBOOK PENDING | COMPLETE

### 3.2. Input table

A markdown table representing the data area the user selects in Excel.

The first column is always the row label (outside the selected range).
Column headers represent the banner row above the selection (if any).

Example:

    Label       | Col A  | Col B  | Col C
    ------------|--------|--------|------
    %           |  0.50  |  0.70  |  0.51
    Base        |  500   |  500   |  500

All values must be exact so expected results can be independently verified.

### 3.3. Expected row-type assignments

A list mapping each row label to the type the detector must assign.

    %       -> PROPORTION
    Base    -> BASE

### 3.4. Expected calculation blocks

A list of blocks the detector must form, each specifying:

- block type (PROPORTION, MEAN, NPS);
- which rows belong to it;
- which row provides the base.

### 3.5. Expected significance results

A table of comparison pairs and their expected outcome at the given confidence level.

    Pair        | Direction | Expected result
    ------------|-----------|----------------
    Col A vs B  | B > A     | SIGNIFICANT
    Col A vs C  | --        | NOT SIGNIFICANT
    Col B vs C  | B > C     | SIGNIFICANT

Direction is the column expected to receive the marker. "--" means no directional claim.

### 3.6. Expected marker assignments

A cell-by-cell table listing which markers must appear in each data cell.

    Cell        | Expected marker
    ------------|----------------
    % / Col A   | b
    % / Col B   | (none)
    % / Col C   | (none)
    Base / *    | (none)

### 3.7. Workbook reference

    Workbook file:  (pending) test-fixtures/GST-NNN.xlsx
    Sheet:          Sheet1
    Selection:      B2:D3  (data area only; labels are in column A)

This field is left as "pending" until the workbook is created.

### 3.8. Expected output artifact

    Output file:  (pending) test-fixtures/GST-NNN-expected.json

The JSON schema is defined in section 5 below.

---

## 4. Initial case inventory

The cases below are planned. None have workbook files or expected-output files yet.

Coverage is derived from TABLE_STRUCTURE_MATRIX.md sections 1-8.
Each case references the matrix section(s) it exercises.

### 4.1. Proportion cases

    GST-001  Single proportion row + Base, all-vs-all, 95%, two-tailed.
             Matrix ref: 1.1 (Single proportion row + Base)
             Covers: basic pooled z-test, marker assignment to proportion row,
                     no markers on Base row.

    GST-002  Multiple proportion rows + shared Base, all-vs-all, 95%, two-tailed.
             Matrix ref: 1.1 (Multiple proportion rows + shared Base)
             Covers: multiple proportion rows sharing one Base, correct block formation.

    GST-003  Proportion expressed as decimal (0.42-style values).
             Matrix ref: 1.1 (Proportion expressed as decimal)
             Covers: normalization before test, display convention preservation.

    GST-004  Proportion expressed as percent (42%-style values).
             Matrix ref: 1.1 (Proportion expressed as percent)
             Covers: Excel percent format preserved in output.

    GST-005  Borderline proportion: values chosen so result changes between 90% and 95%.
             Matrix ref: 1.1, settings section 8
             Covers: confidence level selector effect on significance outcome.

    GST-006  One-tailed vs two-tailed: same borderline values, both modes.
             Matrix ref: 1.1, settings section 8
             Covers: tailed mode selector effect.

    GST-007  Out-of-range proportion value (>1 after normalization), PARTIAL.
             Matrix ref: 1.1 (Out-of-range proportion)
             Covers: test proceeds; results are noted as potentially invalid.
             Note: expected result is "test runs without crash"; statistical
                   validity is explicitly flagged as not guaranteed.

### 4.2. Mean cases

    GST-010  Mean + SD + Base, all-vs-all, 95%, two-tailed.
             Matrix ref: 1.2 (Mean + SD + Base)
             Covers: Welch's t-test, markers on Mean row only, SD and Base receive
                     no markers.

    GST-011  Mean + Variance + Base, all-vs-all, 95%, two-tailed.
             Matrix ref: 1.2 (Mean + Variance + Base)
             Covers: variance converted to SD internally, same marker rules as GST-010.

    GST-012  Multiple mean blocks in one selection (each block has its own Base).
             Matrix ref: 1.2 (Multiple mean blocks in one selection)
             Covers: each triplet detected as separate block.

### 4.3. NPS cases

    GST-020  NPS-first: NPS / Promoters / Detractors / Base.
             Matrix ref: 1.3 (NPS-first)
             Covers: NPS recalculated from Promoters - Detractors, NPS receives NPS
                     markers, Promoters and Detractors receive proportion markers,
                     Base receives no marker.

    GST-021  NPS-first with Neutral: NPS / Promoters / Neutral / Detractors / Base.
             Matrix ref: 1.3 (NPS-first with Neutral)
             Covers: Neutral receives ordinary proportion markers in addition to
                     Promoters/Detractors.

    GST-022  Extended NPS: Scale rows / NPS / Base.
             Matrix ref: 1.3 (Extended NPS)
             Covers: scale rows receive proportion markers, NPS receives NPS markers,
                     Base receives no marker.

    GST-023  NPS + SD / Base (spread path).
             Matrix ref: 1.3 (NPS + SD/Base)
             Covers: spread significance path, SD and Base receive no markers.

    GST-024  NPS + Variance / Base (spread path).
             Matrix ref: 1.3 (NPS + Variance/Base)
             Covers: same as GST-023 with variance row.

### 4.4. Total mode cases

    GST-030  First-column-is-Total, all-vs-all, 95%.
             Matrix ref: 2.2, 7 (Total comparison)
             Covers: T/t markers, Total column receives no markers, segment markers
                     indexed from `a`.

    GST-031  Compare only with Total.
             Matrix ref: 7 (Compare only with Total)
             Covers: only T/t markers; no ordinary segment letters.

    GST-032  Exclude Total from comparisons.
             Matrix ref: 7 (Exclude Total from comparisons)
             Covers: no T/t; ordinary segment comparisons only.

    GST-033  Previous-column mode, no Total.
             Matrix ref: 7 (Previous-column comparison)
             Covers: arrows written into right column only; no letter markers.

    GST-034  Previous-column with small-base column in chain.
             Matrix ref: 7, 4 (Previous-column + small base)
             Covers: chain does not skip over excluded column.

### 4.5. Banner structure cases

    GST-040  No banner (one-level fallback group), all-vs-all.
             Matrix ref: 2.1 (No banner)
             Covers: default group; comparisons within group.

    GST-041  One-level banner, two groups, no Total.
             Matrix ref: 2.1 (One-level banner)
             Covers: group-local marker indexing; no cross-group comparisons.

    GST-042  One-level banner, two groups, each with local Total.
             Matrix ref: 2.1, 2.2 (One-level banner + local Total)
             Covers: T/t within each group; local Total not indexed.

    GST-043  Two-level banner with repeated labels (e.g. Gender / Gender / Age / Age).
             Matrix ref: 2.1 (Two-level banner with repeated labels)
             Covers: groups detected correctly; no cross-group comparisons.

    GST-044  Two-level banner with merged-like cells (empty upper cells).
             Matrix ref: 2.1 (Two-level banner with merged-like cells)
             Covers: reconstructed span fills in empty upper cells correctly.

    GST-045  Global Total spanning all groups.
             Matrix ref: 2.2 (Global Total column)
             Covers: global Total is the only Total reference; local Totals become
                     ordinary target columns.

    GST-046  Wave group alongside non-wave group.
             Matrix ref: 2.3 (Wave group in banner)
             Covers: wave group uses auto previous-column; non-wave group uses
                     ordinary comparisons; no cross-group comparisons.

    GST-047  Manual previous-column overrides wave auto mode.
             Matrix ref: 2.3 (Manual previous-column overrides wave auto mode)
             Covers: global previous-column applies to all groups.

### 4.6. Small base cases

    GST-050  Small base in proportion block (segment column excluded).
             Matrix ref: 4 (Base row with small base)
             Covers: excluded column receives no markers, receives small-base fill
                     across block.

    GST-051  Small base in Total column stops calculation.
             Matrix ref: 4 (Base row is Total column with small base)
             Covers: calculation stops; status message shown.

    GST-052  Small base in mean block.
             Matrix ref: 4 (Base row with small base), 1.2
             Covers: excluded column receives small-base fill on Mean, SD, and Base
                     rows.

    GST-053  Small base in NPS block.
             Matrix ref: 4 (Base row with small base), 1.3
             Covers: excluded column receives small-base fill on all NPS block rows.

### 4.7. Base edge cases

    GST-060  Zero base in one column (missing base PARTIAL case).
             Matrix ref: 6 (Missing Base = 0 or blank)
             Covers: no crash; result written as non-significant; behavior noted.

    GST-061  Multiple Base rows in one block (first used, PARTIAL).
             Matrix ref: 4 (Multiple Base rows in one block)
             Covers: first Base row is used; behavior noted.

### 4.8. Messy table cases

    GST-070  Selection includes empty rows between blocks (PARTIAL).
             Matrix ref: 3.1 (Selection includes empty rows between blocks)
             Covers: detector attempts to skip empty rows; block formation is noted.

    GST-071  Selection includes title row above data (PARTIAL / guardrail warning).
             Matrix ref: 3.1 (Selection includes title / question text row)
             Covers: guardrail warning emitted; Run proceeds; selection not trimmed.

### 4.9. Numeric output and display cases

    GST-080  Plain numeric values (28), percent values (28%), decimal values (0.28)
             mixed in one table.
             Matrix ref: STATUS.md numeric output rules
             Covers: unmarked cells preserve original Excel numeric values;
                     marker-bearing cells show correct display convention.

    GST-081  Round cell values off vs on.
             Matrix ref: Section 8 (Round cell values)
             Covers: display decimal precision per metric type; calculations use
                     original values.

### 4.10. Mixed block cases

    GST-090  Table with Proportion + Mean + NPS blocks sharing one Base row.
             Matrix ref: 1.1, 1.2, 1.3 (mixed)
             Covers: all three block types detected and calculated correctly;
                     shared Base resolved; no markers on service rows.

---

## 5. Expected output artifact schema

When a workbook is created and run through RIT, the result is captured in a JSON file
with the following structure.

    {
      "caseId": "GST-001",
      "runDate": "YYYY-MM-DD",
      "settings": {
        "confidenceLevel": 95,
        "oneTailed": false,
        "roundCellValues": false,
        "firstColumnIsTotal": false,
        "compareOnlyWithTotal": false,
        "excludeTotalFromComparisons": false,
        "compareWithPreviousColumn": false,
        "respectBannerStructure": false,
        "excludeSmallBases": false,
        "smallBaseThreshold": 50
      },
      "detectedRowTypes": [
        { "rowLabel": "%",    "detectedType": "PROPORTION" },
        { "rowLabel": "Base", "detectedType": "BASE" }
      ],
      "detectedBlocks": [
        {
          "blockType": "PROPORTION",
          "rows": ["%"],
          "baseRow": "Base"
        }
      ],
      "significanceResults": [
        {
          "pair": ["Col A", "Col B"],
          "rowLabel": "%",
          "direction": "Col B > Col A",
          "significant": true
        }
      ],
      "markerAssignments": [
        { "rowLabel": "%",    "column": "Col A", "marker": "b" },
        { "rowLabel": "%",    "column": "Col B", "marker": "" },
        { "rowLabel": "Base", "column": "Col A", "marker": "" },
        { "rowLabel": "Base", "column": "Col B", "marker": "" }
      ],
      "statusMessage": "Calculation complete. Blocks processed: 1."
    }

Notes:

- `marker` is an empty string when no marker is written.
- `significanceResults` lists every comparison pair per row; `significant` is boolean.
- `detectedRowTypes` must list every row in the selected range.
- `detectedBlocks` must list every block formed by the detector.
- `statusMessage` is the verbatim text shown in the taskpane status panel.

This schema is a draft. It may be revised when the first workbooks are created.

---

## 6. Agent workflow

When a PR changes logic in any of the following files, the agent must check whether
any gold standard cases are affected:

    src/core/significance.js
    src/core/metric-detector.js
    src/core/banner-detector.js
    src/core/excel-writer.js
    src/taskpane/taskpane.js

### 6.1. PR requirement for logic changes

When expected-output files exist for a case affected by the PR:

- The agent must run the relevant cases against the updated code.
- The agent must compare actual output against the expected JSON.
- The agent must report any differences as part of the PR description.
- If a difference is intentional (behavior change), the agent must update the
  expected-output file and explain the change in the PR description.
- If a difference is unintentional, the PR must not be merged until the regression
  is resolved.

### 6.2. Adding a new case

When a new structure is implemented or an existing PARTIAL structure is hardened:

1. Determine the appropriate GST case ID from section 4.
2. Create the workbook file in test-fixtures/.
3. Run the workbook through RIT with the settings specified in the case.
4. Capture the output as a JSON file matching the schema in section 5.
5. Save the expected-output file in test-fixtures/.
6. Update the case Status field from PLANNED to COMPLETE.
7. Open a PR referencing the case ID.

### 6.3. Reviewing a case

A case is considered valid when:

- The workbook file exists and matches the input table in section 4.
- The expected-output JSON file exists and was produced by a clean run.
- The run was performed on a known-good commit (tagged or release branch).
- The case ID, title, and matrix reference are consistent across this document and
  the workbook/output file.

### 6.4. What agents must not do

- Do not create a workbook or expected-output file from fabricated data.
  All workbooks and output files must come from an actual RIT run on real input.
- Do not mark a case COMPLETE without a real workbook and a real expected-output file.
- Do not modify expected-output files to match broken code.
  Update the code first; then re-run and recapture the output.

---

## 7. Relationship to other documents

    docs/TABLE_STRUCTURE_MATRIX.md   Structure coverage source. Cases in section 4
                                     above are derived from the matrix statuses.
                                     When the matrix is updated, this document
                                     should be reviewed for new or changed cases.

    docs/TEST_CASES.md               Manual regression test cases. Gold standard
                                     cases complement but do not replace manual
                                     tests. Manual tests cover UI, fill color,
                                     and interaction behavior; gold standard cases
                                     cover numeric and marker correctness.

    docs/STATUS.md                   Canonical list of implemented features and
                                     known debt. Cases marked PLANNED must not be
                                     marked COMPLETE until the relevant feature
                                     status in STATUS.md is SUPPORTED or better.

    docs/SELECTED_RANGE_NORMALIZATION.md
                                     Spec for range normalization guardrails.
                                     Cases GST-070 and GST-071 depend on the
                                     behavior described in that spec.

---

## 8. Known limitations of this plan

- No workbook files or expected-output files exist yet. All cases in section 4 are
  PLANNED status.
- The expected-output schema in section 5 is a draft. Field names and structure may
  change when the first workbooks are produced.
- UNSUPPORTED and FUTURE structures from TABLE_STRUCTURE_MATRIX.md are not assigned
  case IDs here. They will be added when the structures are implemented.
- The weighted base (PARTIAL) and Effective Base (FUTURE) cases are not yet assigned
  because the detection and calculation paths are not finalized.
- Count row handling (FUTURE in matrix section 5.1) is excluded until the behavior
  is defined.
- Google Sheets and cloud settings are excluded as they are not implemented.
- This document does not prescribe a test runner. Automation tooling is a future
  issue once the case inventory and schema are validated against real runs.
