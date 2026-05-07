# Multi-Column Row Labels and Weighted / Effective Bases

**Status:** Specification / discovery. No runtime code, tests, or detector
behavior is changed by this document. All status flags and base priority rules
described below remain proposals until landed in their respective phases.

**Issue:** #17

**Source documents:**
- docs/TABLE_STRUCTURE_MATRIX.md
- docs/GOLD_STANDARD_TEST_SUITE.md
- docs/PR_TECHNICAL_DEBT_REPORT.md
- docs/ROADMAP.md
- docs/ARCHITECTURE.md
- docs/table-preview-model.md
- docs/SELECTED_RANGE_NORMALIZATION.md

**Scope of this document:** define the product, data, and validation contract
for two related research-table features that the existing matrix marks as
PARTIAL or FUTURE:

1. Multi-column row labels (table-understanding feature).
2. Weighted, unweighted, and effective bases (correctness feature).

This document does NOT prescribe implementation details for runtime modules
beyond what is required to keep the spec testable and reviewable. Detection,
calculation, and writer changes are scheduled in the phased plan in section 11.

---

## 1. Problem statement

### 1.1. Multi-column row labels

Research tables exported from Confirmit, SPSS, Dimensions, and similar tools
frequently encode the row label as more than one text column. A common pattern:

    | Question / topic   | Answer / row | Col A | Col B | Col C |
    |--------------------|--------------|-------|-------|-------|
    | Awareness          | %            |  0.42 |  0.51 |  0.49 |
    | Awareness          | Base         |  500  |  500  |  500  |
    | Brand A consideration | %         |  0.31 |  0.40 |  0.29 |
    | Brand A consideration | Base       |  500  |  500  |  500  |

Today the detector reads only one label column (LABEL_SCAN_COLUMNS_LEFT in
src/core/metric-detector.js is set to 2 for scanning, but only the rightmost
non-empty label part is used as the row label for type detection and block
formation; see docs/TABLE_STRUCTURE_MATRIX.md section 3.2 row "Multi-column
label area"). The second label column is effectively discarded for type
detection and block boundary decisions.

Practical consequences:

- The "topic" or "question" column above is never used to disambiguate two
  metric blocks that share the same primary label (e.g. two `%` rows that
  belong to different questions).
- The Check table / preview feature cannot present a hierarchical row label
  to the user even though src/core/table-preview-model.js already exposes
  `labelParts`, `primaryLabel`, `secondaryLabel`, and `combinedLabel`.
- Block detection can merge or split rows incorrectly when the secondary
  label is the only signal that a new block starts.

### 1.2. Weighted, unweighted, and effective bases

The existing detector treats every Base-keyword row identically and uses the
first matching row as the denominator for the z-test. The matrix
(docs/TABLE_STRUCTURE_MATRIX.md section 4) marks:

- Unweighted Base: SUPPORTED.
- Weighted Base: PARTIAL. No dedicated path. Value is fed to the z-test as-is.
- Effective Base: FUTURE. No keyword, no detection, no logic.

This is a correctness gap, not a premium feature. When a researcher provides
a weighted percentage but the only base in the table is a weighted base, a
naive z-test on the weighted base produces an underestimated standard error
and inflated significance. When an Effective Base is available it should be
used because it already accounts for the design effect.

Treating the three base types as equivalent is the root cause of:

- silent over-detection of significance on weighted tables;
- inability to run a proper test on tables that contain both an unweighted
  and a weighted base;
- inability to honour an Effective Base value even when the researcher has
  taken the trouble to compute it.

### 1.3. Why bundle these two topics

Both topics are about how the row label is read.

- Multi-column labels widen the label-reading surface (across text columns).
- Weighted / unweighted / effective bases narrow the label-reading surface
  (subtype within a single keyword family).

The detection paths share a common dependency: a richer row-label data model
that records every label part, normalizes them, and lets downstream code
(metric detection, block formation, preview, calculation) make subtype
decisions on the structured data instead of a single string.

---

## 2. Supported input patterns

This section enumerates table layouts that this spec commits to handle once
the phased plan in section 11 is implemented. It does not change current
detector behavior.

### 2.1. Multi-column labels

S-MCL-1. **Two-column label, primary on the right (closest to data).**

    | Question / topic | Row    | Col A | Col B |
    | Awareness        | %      |  0.42 |  0.51 |
    | Awareness        | Base   |  500  |  500  |

    Primary label   = "%" (right column, closest to data).
    Secondary label = "Awareness" (left column, category header).

S-MCL-2. **Two-column label with the secondary part empty in continuation
rows.** Common when the secondary column is visually merged.

    | Question / topic | Row    | Col A | Col B |
    | Awareness        | %      |  0.42 |  0.51 |
    |                  | Base   |  500  |  500  |
    | Consideration    | %      |  0.31 |  0.40 |
    |                  | Base   |  500  |  500  |

    Empty secondary cells inherit the most recent non-empty secondary value.
    This is span-fill, identical in spirit to banner span reconstruction.

S-MCL-3. **Two-column label where the secondary part disambiguates blocks
sharing the same primary label.**

    | Question / topic | Row    | Col A | Col B |
    | Brand A          | %      |  0.31 |  0.40 |
    | Brand A          | Base   |  500  |  500  |
    | Brand B          | %      |  0.22 |  0.30 |
    | Brand B          | Base   |  500  |  500  |

    Two distinct blocks must be formed even though both primary labels are
    `%` and both base labels are `Base`.

S-MCL-4. **Russian / English mixed secondary labels.** Detection must remain
dictionary-driven on the primary label; the secondary label is treated as
free text.

### 2.2. Bases

S-BASE-1. **Single Unweighted Base.** Current behavior. Used as denominator.

S-BASE-2. **Single Weighted Base detected by label keyword.** Used as
denominator with a "weighted base" warning surfaced to the user. Calculation
proceeds; the warning is not blocking.

S-BASE-3. **Single Effective Base detected by label keyword.** Used as
denominator. No warning. No fallback unless the Effective Base value is
zero, blank, or below the small-base threshold.

S-BASE-4. **Unweighted Base + Weighted Base in the same block.** Unweighted
Base is selected. Weighted Base is recorded in the row diagnostic but not
used. No warning.

S-BASE-5. **Effective Base + Unweighted Base in the same block.** Effective
Base is selected. Unweighted Base and any Weighted Base are recorded but
not used. No warning.

S-BASE-6. **Effective Base + Unweighted Base + Weighted Base in the same
block.** Effective Base wins. Other bases recorded only.

S-BASE-7. **No Base row.** Existing behavior: block is not formed.
Calculation skipped. No change from this spec.

S-BASE-8. **Small base on the selected test base.** Existing small-base
exclusion rules apply to the selected base, regardless of which subtype it
is. See section 6.2.

### 2.3. Combined patterns

S-COMB-1. Two-column labels combined with mixed Base subtypes within one
selection. Each block resolves its own base selection independently.

S-COMB-2. Two-column labels in NPS structures. The primary label still
drives NPS keyword detection (NPS / Promoters / Detractors / Neutral / Base).
The secondary label is recorded for preview and disambiguation only.

---

## 3. Unsupported / future patterns

The following are explicitly out of scope for the phased plan in section 11.
They may be addressed in follow-up specs but must not be implicitly enabled
by changes made for this spec.

U-MCL-1. Three or more label columns. Labels beyond the first two columns
to the left of the data area remain ignored. Detector continues to scan only
LABEL_SCAN_COLUMNS_LEFT columns.

U-MCL-2. Hierarchical totals across multi-column labels (e.g. a row whose
primary label is empty and whose secondary label is "Subtotal"). Such rows
are detected as `unknownText` until a dedicated spec is written.

U-MCL-3. Cross-block aggregations driven by secondary labels (e.g. computing
a NET across all blocks whose secondary label matches a tag). This is a
calculation feature, not a label feature, and is not part of this spec.

U-BASE-1. Derived bases from count rows (count / Base -> proportion). See
docs/TABLE_STRUCTURE_MATRIX.md section 5.1 row "Count row" and
docs/GOLD_STANDARD_TEST_SUITE.md case GST-102 (BLOCKED-BY-SPEC). A separate
opt-in spec is required.

U-BASE-2. Per-cell effective base. Some research tables include an effective
base value in every column independently. Today the detector reads one base
row per block; this remains the rule. A per-cell effective base layout is a
future spec.

U-BASE-3. Time-window / wave-specific effective bases. Wave banner support
exists today; per-wave Effective Base columns are out of scope.

U-BASE-4. Researcher-supplied design effect column or formula. Out of scope.

---

## 4. Proposed data model

This section defines the new fields and shapes that detection, preview, and
calculation code will use. The goal is to commit to the contract before
writing any runtime code in phases 2 and later.

### 4.1. Row label parts

Every detected row carries a structured label record:

    {
      labelParts:         string[],    // all non-empty parts, left-to-right
      normalizedLabelParts: string[],  // same length, normalized
      primaryLabel:       string,      // rightmost non-empty part
      secondaryLabel:     string|null, // next part outward, or null
      combinedLabel:      string,      // joined for display, e.g. "Topic / Row"
      labelRole:          string|null, // see 4.2 below; reserved for future
    }

This shape is already produced by src/core/table-preview-model.js. The spec
commits to:

- promoting these fields to the metric detector's row record so block
  formation can read them;
- defining `labelRole` as a reserved enumeration whose values are deferred
  to phase 3 of the plan in section 11.

### 4.2. Row role enumeration (reserved)

`labelRole` is a coarse classification of the secondary label's function in
the table. Reserved values, none of them required for phase 1 / 2:

    "topic"     -- secondary label names a question or topic
    "category"  -- secondary label names a metric category (e.g. "Awareness")
    "totalRow"  -- secondary label marks a total / subtotal row
    "block"     -- secondary label is a block divider
    null        -- not classified (default)

Detection of `labelRole` is a phase-3 spec extension and is not required for
the base priority work to land.

### 4.3. Base subtype

Every row currently classified as `rowType: "base"` carries a subtype:

    rowSubtype:  "effective" | "unweighted" | "weighted" | null

`null` means no specific subtype was detected. The row is still treated as a
plain Base for backward compatibility (see section 5).

This field is already reserved on the preview model
(docs/table-preview-model.md, RowDiagnostic). The spec commits to populating
it from dictionary keywords once phase 2 lands.

### 4.4. Block-level base selection record

Every calculation block records which base was used and why:

    {
      baseRowIndex:     number,
      baseSubtype:      "effective" | "unweighted" | "weighted" | null,
      baseSelection: {
        chosen:         "effective" | "unweighted" | "plain" | "weighted",
        candidates:     [
          { rowIndex: number, subtype: string, value: number|null }
        ],
        reason:         string,    // human-readable, used in status messages
        warning:        string|null
      }
    }

The `baseSelection` record is the audit trail consumed by:

- the Check table / preview UI;
- the taskpane status panel;
- gold-standard expected-output JSON.

`reason` and `warning` are surface-text fields. Their wording is a phase-2
UI decision and is not pinned by this spec.

### 4.5. Dictionary additions

The dictionary in src/core/config/dictionary.config.js gains two optional
keyword groups, both in Russian and English. Exact keyword lists are
proposed here and finalized in phase 2.

    base.subtypes.effective:
      English:  "effective base", "eff base", "eff. base"
      Russian:  Russian phrases for "effective base" and "eff. base"
                (Cyrillic, finalized in phase 2 alongside the dictionary
                update)

    base.subtypes.weighted:
      English:  "weighted base"
      Russian:  Russian phrase for "weighted base"
                (Cyrillic, finalized in phase 2)

    base.subtypes.unweighted:
      English:  "unweighted base"
      Russian:  Russian phrase for "unweighted base"
                (Cyrillic, finalized in phase 2)

The plain `base` entry already exists; it remains the fallback when no
subtype keyword matches. Under no circumstance does a subtype keyword stop
matching the existing `base` rowType -- it refines it.

---

## 5. Base selection rules

This section is the contract every code path that picks a base must follow.

### 5.1. Priority order

For each calculation block:

    1. Effective Base
    2. Unweighted Base
    3. Plain Base (rowSubtype === null)
    4. Weighted Base  (with warning, see 5.4)

The block selects the highest-priority base that meets the validity rules in
section 5.2. If none qualify, the block is treated as having no base and the
existing "no base" handling applies (block not formed).

### 5.2. Validity rules

A candidate base row qualifies when ALL of the following are true:

- The base value in the column under test is a finite number greater than 0
  (existing behavior; division-by-zero handling is not part of this spec).
- The base value passes the small-base threshold for that column (see 6.2).
- The row is inside the selected range (existing behavior).

A candidate is otherwise non-qualifying and the next-priority candidate is
considered.

### 5.3. Tie-breaking within a single subtype

If multiple rows match the same subtype within one block (e.g. two
"Effective Base" rows), the first matching row wins. This matches the
current "first matching base" behavior and is documented as PARTIAL in
docs/TABLE_STRUCTURE_MATRIX.md section 4.

### 5.4. Weighted-base fallback warning

When the chosen base is the Weighted Base (priority 4), the calculation
proceeds and a warning is surfaced:

- as a status message in the taskpane;
- as a `dataQualityIssue` of severity `warning` in the preview model;
- as a field in the block's `baseSelection.warning`.

The warning text is finalized in phase 2; it must explicitly mention that
the weighted base is being used because no Unweighted or Effective Base was
available, and that statistical validity depends on the weighting method.

The warning is not blocking. Run is not stopped.

### 5.5. Displayed values are not recalculated

Percentages, means, and NPS values shown in the table are read as-is.
This spec does NOT introduce any recalculation of displayed values from
counts and bases. The selected test base (per priority order above) is used
ONLY to compute significance.

This rule is a hard product constraint and applies to all phases.

### 5.6. Small-base threshold uses the selected test base

The small-base threshold is evaluated against the value of the base that
was selected for the test, not the displayed weighted base or any other
base candidate.

Practical consequence: a column whose displayed Weighted Base is 487 may
still be excluded from comparisons if the selected Effective Base for that
column is 32 and 32 is below the threshold.

### 5.7. Base selection is per block, not per selection

Two blocks within the same selected range may resolve to different base
subtypes. The selection is recorded per block in the `baseSelection` field
defined in section 4.4.

---

## 6. Check table / preview implications

The Check table feature defined in docs/table-preview-model.md is the
primary user-facing surface for this spec. Preview must remain read-only;
no Excel writes, no significance calculation.

### 6.1. RowDiagnostic additions

`rowSubtype` is populated for rows of `rowType: "base"`:

    rowSubtype = "effective" | "unweighted" | "weighted" | null

`labelRole` is reserved and remains `null` for phase 1 / 2.

### 6.2. PreviewBlock additions

Each block exposes the `baseSelection` record from section 4.4. The preview
UI uses `baseSelection.chosen` and `baseSelection.reason` as the primary
indicator of which base will be used in calculation.

When `baseSelection.warning` is set, it is surfaced as a Data Quality issue
of severity `warning` with code `WEIGHTED_BASE_FALLBACK` (proposed code,
finalized in phase 2).

### 6.3. Multi-column label rendering

`combinedLabel` is the default label string in preview. The preview UI may
render `secondaryLabel` and `primaryLabel` as a two-line cell when present.

### 6.4. New data quality codes

Proposed new entries for the table in docs/table-preview-model.md
section "DataQualityIssue":

    WEIGHTED_BASE_FALLBACK
        Severity: warning
        Condition: Weighted Base used because no Unweighted or Effective
                   Base was found.

    EFFECTIVE_BASE_BELOW_THRESHOLD
        Severity: warning
        Condition: Effective Base value passes priority but the column is
                   excluded for being below the small-base threshold.
                   Distinct from the existing small-base treatment because
                   the column might display a much larger displayed base.

    AMBIGUOUS_BASE_LABEL
        Severity: info
        Condition: A Base row label matches multiple subtype keywords (e.g.
                   "Weighted effective base"). Detection picks the first
                   match in priority order; user is informed.

These codes are proposals. Final names and severities are settled in phase
2 alongside the dictionary additions.

### 6.5. Preview must remain read-only

The Check table / preview pipeline does not write to Excel and does not
trigger significance calculation. This rule is unchanged from
docs/table-preview-model.md and from the project rules in AGENTS.md.

---

## 7. Calculation implications

This section is the contract between the detector and src/core/significance.js.
It does not prescribe how significance.js implements the contract; it only
fixes the inputs.

### 7.1. Inputs to the test

For every block, the calculation receives:

- a value row (or rows) with displayed values, untouched;
- a base value per column, drawn from the selected base (per section 5);
- the existing settings (confidence level, tailed mode, etc.).

The calculation does NOT receive multiple bases. Subtype selection is done
upstream, in the detector / block builder, and is committed before
significance is computed.

### 7.2. No formula changes

This spec does NOT change the z-test, t-test, or NPS formulas in
src/core/significance.js. The z-test continues to use the standard pooled
formula with the supplied base as denominator. Project rule from AGENTS.md
("Do not rewrite significance.js wholesale") is preserved.

The only change in calculation is which value is supplied as the base.

### 7.3. NPS, means, and proportions

The base selection rules apply identically to:

- proportion blocks (z-test);
- mean blocks (Welch's t-test, where Base is the n);
- NPS blocks (both NPS-first and spread variants).

No metric-specific carve-outs.

### 7.4. Multi-block selections

Each block independently resolves its base. Two proportion blocks in the
same selection may use different base subtypes (e.g. one uses Effective
Base, the other falls back to Plain Base). The taskpane status message
should aggregate base-selection notes per block.

### 7.5. Markers and writer behavior unchanged

The Excel writer continues to follow the rules in AGENTS.md:

- Cells with appended markers may be text;
- Cells without markers preserve numeric Excel values;
- Display conventions (28, 28%, 0.28) are preserved;
- Block-level writes are not regressed to per-cell writes.

This spec does not propose new marker types, new fill colors, or new
banner-letter behavior.

---

## 8. UI implications

This section is intentionally minimal. The current taskpane is a high-risk
file (AGENTS.md). Phase 4 of the plan in section 11 contains UI changes;
they are surfaced here only as constraints.

### 8.1. Preserve manual selected-range workflow

No change to the manual selected-range workflow. Run still requires the user
to select the data range and click Run. No auto-trim, no auto-scan, no
implicit selection adjustment. This is restated to make sure phase 4 stays
within the existing UX.

### 8.2. Taskpane status messages

Phase 4 may add status-panel text for:

- "Effective Base used in block N." (info)
- "Weighted Base used in block N. Statistical validity depends on the
  weighting method." (warning)
- "Block N: Effective Base column M excluded for small base." (info)

Wording is finalized in phase 4. Status messages must not introduce blocking
modal dialogs.

### 8.3. Check table panel

The Check table feature surfaces `baseSelection.chosen` and the related
warning per block. No new modal dialog is introduced by this spec.

### 8.4. Multi-column label display

The preview UI may render the secondary label as a smaller line above the
primary label, using `combinedLabel` as the accessible string. Current Run
output does not change because Run does not display labels in the taskpane.

---

## 9. Risks and open questions

### 9.1. Risks

R-1. **False positive subtype detection.** Adding "weighted" / "effective"
keywords to the dictionary risks misclassifying ordinary Base rows whose
labels happen to contain these words in unrelated contexts. Mitigation:
require the keyword to appear in conjunction with the existing `base`
keyword, not as a standalone match. Phase 2 task.

R-2. **Block boundary regression from richer labels.** Promoting
`labelParts` into block formation could change where blocks split or merge.
Mitigation: phase 1 does not change block formation; phase 2 introduces
secondary-label disambiguation behind a feature flag and is gated by
gold-standard cases that exercise both layouts.

R-3. **Significance.js drift.** Even though calculation formulas are
unchanged, supplying a different denominator changes outcomes. Mitigation:
gold-standard cases GST-062 to GST-066 (planned) and a new GST case for
S-COMB-1 must be added before phase 3 can be marked SUPPORTED.

R-4. **Excel writer regressions.** This spec does not touch the writer, but
new status messages and warnings may interact with the existing run loop.
Mitigation: AGENTS.md rules about excel-writer.js are restated in section 7.5.

R-5. **Small-base threshold semantics change.** Section 5.6 clarifies that
the threshold applies to the selected test base. For tables that previously
used a Weighted Base as-is, the threshold result may change. Mitigation:
documented as an explicit behavior change in the phase-3 release notes.

R-6. **Counts confused with Bases.** The count-row vs base-row distinction
remains a separate FUTURE topic. This spec is careful to keep the dictionary
extension scoped to the existing `base` rowType and not to introduce a new
rowType that aliases `count`. See docs/TABLE_STRUCTURE_MATRIX.md section 5.1.

### 9.2. Open questions

OQ-1. Should the Effective Base, when present, override the displayed
Weighted Base entry in the preview's Base column, or should both rows be
shown side by side? Default proposal: show both rows; mark which one was
selected.

OQ-2. Should the dictionary subtype keywords be required to match as
phrases (e.g. "weighted base") or as standalone tokens ("weighted") near a
base keyword? Default proposal: phrase match for English, token match for
Russian where adjective inflection makes phrase match brittle. Phase 2.

OQ-3. When both an Effective Base and an Unweighted Base are present and
the Effective Base column for a particular cell is blank or zero but the
Unweighted Base is fine, should the cell fall back to Unweighted for that
column only, or should the entire column be excluded? Default proposal:
exclude the column for that block (consistent with how a single missing
base would be handled today). Revisit in phase 3.

OQ-4. Should multi-column label disambiguation also affect the metric-row
detector for non-Base rows (e.g. two `%` rows in different blocks sharing
the same primary label but different secondary labels)? Default proposal:
yes, but only for block boundary purposes; row-type detection still uses
the primary label. Phase 2.

OQ-5. Does the Check table preview need a dedicated icon or affordance for
Weighted-Base fallback warnings? UX decision, phase 4.

OQ-6. Should base subtype detection also write a marker or a fill into the
base row's cells in Run output? Current default: no, the Base row receives
no markers (AGENTS.md). This rule is preserved.

---

## 10. Test coverage

Existing planning artefacts already cover much of the test surface. This
spec commits to using them as the validation contract.

### 10.1. Gold-standard cases

The following docs/GOLD_STANDARD_TEST_SUITE.md cases are directly relevant
and must be updated as the phases land:

    GST-062  Weighted Base only.
             Status today: PLANNED. Update at phase 3 to verify warning.

    GST-063  Unweighted Base alongside Weighted Base in same selection.
             Status today: PLANNED. Update at phase 3 to verify priority.

    GST-064  Effective Base present (highest priority).
             Status today: FUTURE. Promote to PLANNED at phase 2 once the
             dictionary keywords are added; verify at phase 3.

    GST-065  Plain Base fallback when no Unweighted or Effective Base
             is present. Status today: PLANNED. No change expected; the
             fallback path is preserved at every phase.

    GST-066  Weighted Base as lowest-priority fallback with warning.
             Status today: BLOCKED-BY-SPEC. This document is the spec.
             Promote to PLANNED at phase 2.

### 10.2. New cases proposed

Add the following cases in phase 1, with status PLANNED:

    GST-067  Effective Base + Unweighted Base + Weighted Base in one block.
             Verifies S-BASE-6 and section 5.1 priority.

    GST-068  Multi-column label disambiguates two blocks sharing primary
             label "%". Verifies S-MCL-3.

    GST-069  Two-column label with span-fill on continuation rows.
             Verifies S-MCL-2.

    GST-072  Small base on Effective Base while displayed Weighted Base is
             above threshold. Verifies section 5.6.

GST-070 and GST-071 (messy table cases) are not affected by this spec.

### 10.3. Matrix updates

docs/TABLE_STRUCTURE_MATRIX.md updates planned by this spec:

- Section 3.2 row "Multi-column label area": move from UNSUPPORTED to
  PARTIAL once phase 2 lands; to SUPPORTED once phase 4 lands.
- Section 4 row "Effective Base": move from FUTURE to PARTIAL once phase 2
  lands; to SUPPORTED once phase 3 lands.
- Section 4 row "Weighted Base": move from PARTIAL to SUPPORTED once phase
  3 lands and the warning is in place.

These transitions are gated by the corresponding GST cases reaching at
least PLANNED status with manual verification, per GOLD_STANDARD_TEST_SUITE.md
section 6.

### 10.4. Manual smoke areas

Phase 3 must add to the manual smoke checklist in AGENTS.md:

- Weighted base fallback warning shown.
- Effective Base used when present.
- Multi-block selection with mixed base subtypes.
- Multi-column label disambiguates two `%` blocks.

---

## 11. Phased implementation plan

Each phase is an independent, scoped PR set. Phases are ordered so that
each phase is reviewable in isolation and so that runtime risk increases
gradually.

### Phase 0 -- this document

- Add docs/MULTI_COLUMN_LABELS_AND_WEIGHTED_BASES.md (this file).
- No runtime code changes.
- No matrix changes.
- No GST status changes.

Acceptance: PR merged with sign-off from the human owner.

### Phase 1 -- discovery and gold-standard expansion

- Add the proposed GST cases (GST-067, GST-068, GST-069, GST-072) to
  docs/GOLD_STANDARD_TEST_SUITE.md as PLANNED.
- Update docs/TABLE_STRUCTURE_MATRIX.md notes to reference this spec.
- No runtime code changes.

Acceptance: docs PR merged. No high-risk file is touched.

### Phase 2 -- detection and preview model

Scope (split into two PRs if size requires it):

- src/core/config/dictionary.config.js: add Effective / Unweighted /
  Weighted keyword groups. Plain `base` keyword unchanged.
- src/core/metric-detector.js: populate `rowSubtype` on Base rows. Block
  formation reads `labelParts` for secondary-label disambiguation
  (S-MCL-3). LABEL_SCAN_COLUMNS_LEFT remains 2.
- src/core/table-preview-model.js: surface `rowSubtype` and the new
  `baseSelection` record on PreviewBlock (without yet driving calculation).
- New data quality codes added to the preview model:
  WEIGHTED_BASE_FALLBACK, EFFECTIVE_BASE_BELOW_THRESHOLD,
  AMBIGUOUS_BASE_LABEL.

Out of scope:

- src/core/significance.js: not modified.
- src/core/excel-writer.js: not modified.
- src/taskpane/taskpane.js: not modified except to forward the existing
  preview pipeline if already wired (it is currently not wired).

Acceptance:

- All existing GST cases pass manual smoke (per GOLD_STANDARD_TEST_SUITE.md).
- New preview output includes the new fields for each block.
- No change to Run output for any existing layout.

### Phase 3 -- calculation wiring

Scope:

- src/core/metric-detector.js / block builder: select base per the priority
  order in section 5. Record `baseSelection` per block.
- src/core/significance.js: receives the selected base via the existing
  block input. No formula changes. No wholesale rewrite. No rerouting of
  comparisons.
- Status panel: display per-block base-selection notes and the weighted-
  base fallback warning.

Out of scope:

- src/core/excel-writer.js (no marker, fill, or numeric output changes).
- Banner detection (no change).
- New comparison modes.

Acceptance:

- GST-062 to GST-066 progress per section 10.1.
- GST-067, GST-072 added in phase 1 reach PLANNED -> verified manually.
- Manual smoke covers the new entries in AGENTS.md.

### Phase 4 -- UI polish

Scope:

- Check table / preview UI: surface `baseSelection.chosen`, the warning,
  and the multi-column label display defined in section 6.
- Optional taskpane status-panel updates per section 8.2.

Out of scope:

- Modifying the Run pipeline.
- New settings.

Acceptance:

- Preview shows base selection rationale per block.
- Manual smoke verifies UX.

### Phase 5 -- matrix promotion and follow-ups

Scope:

- Update docs/TABLE_STRUCTURE_MATRIX.md per section 10.3.
- Update docs/ROADMAP.md to mark "Multi-column row labels and
  weighted/effective bases remain discovery/spec first" as resolved.
- File follow-up issues for U-MCL-1, U-MCL-2, U-BASE-1, U-BASE-2,
  U-BASE-3, U-BASE-4 (section 3) if any are still relevant.

Acceptance:

- Matrix reflects post-phase-4 reality.
- Follow-up issues exist for every U- entry that is still desired.

---

## 12. Cross-document references

- docs/TABLE_STRUCTURE_MATRIX.md
  - Section 3.2 (Multi-Column Labels)
  - Section 4 (Base Structures), including planned priority order note
  - Section 5.1 (Count row, distinct from Base)
- docs/GOLD_STANDARD_TEST_SUITE.md
  - Section 4.7 (Base structure and edge cases): GST-062 to GST-066
  - Proposed additions: GST-067, GST-068, GST-069, GST-072 (this spec)
  - Section 6 (Agent workflow) for promoting case statuses
- docs/PR_TECHNICAL_DEBT_REPORT.md
  - Sections 2 (Shortcuts) and 4 (Fragile Areas) must be filled in for
    every PR that lands a phase from section 11.
- docs/ROADMAP.md
  - "Multi-column row labels and weighted/effective bases remain
    discovery/spec first" entry is the parent roadmap item this spec
    addresses.
- docs/ARCHITECTURE.md
  - High-risk files list (significance.js, excel-writer.js, banner-detector.js,
    metric-detector.js, taskpane.js) governs phase boundaries.
- docs/table-preview-model.md
  - RowDiagnostic.rowSubtype (already reserved) is populated in phase 2.
  - DataQualityIssue table gains the new codes proposed in section 6.4.
- docs/SELECTED_RANGE_NORMALIZATION.md
  - No direct overlap; this spec assumes the manual selected-range workflow
    from AGENTS.md is preserved at every phase.
- AGENTS.md
  - Project rules quoted in sections 5.5 (no recalculation), 7.2 (no
    significance rewrite), 7.5 (writer rules), and 8.1 (manual selected
    range) must be preserved throughout the implementation.

---

## 13. Assumptions and limitations

A-1. The detector continues to scan only LABEL_SCAN_COLUMNS_LEFT (currently
2) columns to the left of the data area. Extending the scan width is not
part of this spec.

A-2. The dictionary additions in section 4.5 are proposals; final keyword
lists are decided in phase 2.

A-3. `labelRole` is reserved but not populated in phases 1-3.

A-4. Wording for status messages and warnings is finalized in phase 4. This
document does not pin any user-facing string verbatim.

A-5. The base priority order in section 5.1 is a product decision, not a
statistical result. Researchers who want a different order will need a
configuration mechanism in a future spec.

A-6. This spec presumes the existing Run pipeline (manual selected-range,
Run button, marker-and-fill output) and does not add new entry points.

A-7. Google Sheets and the future automatic worksheet/workbook scanner are
out of scope for this spec, but this spec must not block platform-neutral
implementation later.
