# Selected Range Normalization Spec

## Overview

The Research Insights Toolkit (RIT) currently relies on a manual selected-range workflow where the user is expected to precisely select the numeric data area in Excel. As a temporary safety net, warning-only selected range guardrails exist.

The desired end-state is a **forgiving full-table selection** model: the user selects the visible table, and RIT identifies the data, labels, and banner internally.

**Important Note:** `src/core/range-normalizer.js` implements Phase 1 (pure model module). The existing manual selected-range workflow remains the stable baseline and is not changed.

---

## 1. Three Output States

The normalized selection model must distinguish three states. Callers **must** branch on these before using any output.

### State 1 — Pass-through (normalization not needed)

```
normalizationNeeded:  false
normalizationApplied: false
blockingReasons:      []
```

**Meaning:** The selection already looks like numeric data only. Future integrations **must** preserve the existing strict workflow unchanged. No coordinate offsets or sliced grids are produced.

### State 2 — Normalization succeeded

```
normalizationNeeded:  true
normalizationApplied: true
blockingReasons:      []
```

**Meaning:** The selection appeared to contain extra structure (title rows, banner rows, label columns). Decomposition succeeded. Future integrations may use `valuesForCalculation`, `leftLabelValues`, `bannerContext`, `dataRowOffset`, and `dataColOffset`.

### State 3 — Normalization needed but blocked

```
normalizationNeeded:  true
normalizationApplied: false
blockingReasons:      ["BODY_TOO_SHORT", ...]   ← one or more codes
blockingMessage:      "..."                     ← human-readable message
```

**Meaning:** The selection appeared broad / full-table-like, but the normalizer could not safely decompose it.

**Future integrations MUST stop and show `blockingMessage`.** They **must not** silently fall back to running the strict workflow on the original broad selection. Running significance on a broad unpartitioned selection can write markers to title cells, label cells, or banner cells — that is silent data corruption, not a safe degradation.

---

## 2. Normalized Selection Data Model

```
{
  // ── State flags ─────────────────────────────────────────────────────────
  normalizationNeeded:  boolean
  normalizationApplied: boolean

  // ── Input dimensions ────────────────────────────────────────────────────
  originalRowCount:    number
  originalColumnCount: number

  // ── Classification (indices relative to rawValues) ──────────────────────
  titleRows:    number[]    // sparse text-only rows stripped from top
  subtitleRows: number[]    // second sparse text-only row (if present)
  bannerRows:   number[]    // wide text-heavy header rows
  labelColumns: number[]    // row-label columns (0, or [0,1])
  dataColumns:  number[]    // numeric data columns
  bodyRows:     number[]    // rows after title/subtitle/banner stripping

  // ── Sliced grids (populated only when normalizationApplied: true) ────────
  valuesForCalculation:        Array[][]   // bodyRows × dataColumns, numeric values
  textForCalculation:          Array[][]   // bodyRows × dataColumns, text representation
  leftLabelValues:             Array[][]   // bodyRows × labelColumns
  bannerContext: {
    scanRows:    Array[][]   // bannerRows × data columns (aligned with valuesForCalculation)
    columnCount: number      // data column count
  }

  // ── Coordinate mapping (populated only when normalizationApplied: true) ──
  dataRowOffset: number   // selectedRange row index of valuesForCalculation[0]
  dataColOffset: number   // selectedRange col index of valuesForCalculation[0][0]

  // ── Confidence ───────────────────────────────────────────────────────────
  confidence: "high" | "medium" | "low"

  // ── Diagnostics ──────────────────────────────────────────────────────────
  warnings: {
    code:         string    // TITLE_ROWS_DETECTED | SUBTITLE_ROWS_DETECTED |
                            // BANNER_ROWS_DETECTED | LABEL_COLUMNS_DETECTED |
                            // NORMALIZATION_CONFIDENCE_MEDIUM
    severity:     "info" | "warning"
    rowIndexes?:  number[]
    columnIndexes?: number[]
  }[]

  blockingReasons: string[]   // blocking reason codes (empty when not blocked)
  blockingMessage: string     // human-readable message (empty when not blocked)
}
```

---

## 3. Detection Strategy

### Step 0 — Pass-through gate

Structural analysis requires at minimum a 3×3 selection. Smaller selections return State 1 immediately.

For qualifying selections, four patterns trigger normalization:

1. **Left-column pattern:** col[0] has ≥70% text AND cols[1..N] have ≥70% numeric content.
2. **Wide top-row pattern:** row[0] has ≥60% text AND rows[1..N] have ≥70% numeric content.
3. **Sparse title-row pattern:** row[0] is title-like (sparse, text-only, ≤3 non-empty cells) AND rows[1..N] have ≥70% numeric content. This covers merged-like title rows above numeric bodies.
4. **Overall text-heavy fallback:** ≥12 cells total, ≥25% text fraction, ≥50% numeric fraction.

If none fire → State 1 (pass-through). If any fire → proceed to structural decomposition.

### Step 1 — Title / subtitle row detection

A row is **title-like** when:
- At least one non-empty non-numeric text cell is present
- No numeric cells are present
- Non-empty cell count is ≤ 3 (sparse — distinguishes title from wide banner)

Applied to the top of the selection: up to one title row, then optionally one subtitle row. Both must appear consecutively at the very top.

### Step 2 — Banner / header row detection

Applied immediately after title/subtitle rows. A row is a **banner row** when:
- Non-empty cells span ≥ 50% of the total column count (wide — distinguishes from sparse title rows)
- Text fraction ≥ 60%

Consecutive qualifying rows are all collected. Multi-level banners are captured naturally.

These rows are extracted into `bannerContext.scanRows` for future use by `detectBannerStructure()`.

### Step 3 — Label column detection

Applied over body rows only (after title/subtitle/banner stripping). Inspects leftmost columns:

| col[0] text fraction | col[1] text fraction | Result |
|---|---|---|
| < 0.5 | — | 0 label columns, confident |
| 0.5 – 0.6 | — | 0 label columns, **uncertain** |
| ≥ 0.6 | ≥ 0.6 | 2 label columns, confident |
| ≥ 0.6 | < 0.6 | 1 label column, confident |

Never consumes a column with < 50% text content as a label column.

### Step 4 — Body validation

After stripping, the candidate body (body rows × data columns) is validated:

| Code | Condition |
|---|---|
| `BODY_TOO_SHORT` | < 2 body rows remain after stripping |
| `DATA_TOO_NARROW` | < 2 data columns remain after label stripping |
| `BODY_APPEARS_MULTI_TABLE` | An all-empty row gap exists inside the body |
| `NO_NUMERIC_BODY` | No numeric cells found in the body at all |
| `LABEL_SPLIT_BLOCKING` | Label split is uncertain AND body text fraction ≥ 40% |

Any of these returns State 3 (blocked).

### Step 5 — Confidence scoring

Computed on the body numeric fraction only when validation passes:

| Body numeric fraction | Label split | Confidence |
|---|---|---|
| ≥ 70% | confident | **high** |
| 50–70% | confident | **medium** |
| ≥ 70% | uncertain | **medium** |
| < 50% | any | **low** → `LOW_CONFIDENCE` blocking |

`low` confidence adds `LOW_CONFIDENCE` to `blockingReasons` and returns State 3. `medium` returns State 2 with a `NORMALIZATION_CONFIDENCE_MEDIUM` warning.

---

## 4. Blocking Conditions Summary

| Code | Trigger | State |
|---|---|---|
| `BODY_TOO_SHORT` | < 2 body rows after stripping | 3 — Blocked |
| `DATA_TOO_NARROW` | < 2 data columns after label stripping | 3 — Blocked |
| `BODY_APPEARS_MULTI_TABLE` | All-empty row gap inside body | 3 — Blocked |
| `NO_NUMERIC_BODY` | No numeric cells in body | 3 — Blocked |
| `LABEL_SPLIT_BLOCKING` | Uncertain split + body text fraction ≥ 40% | 3 — Blocked |
| `LOW_CONFIDENCE` | Body numeric fraction < 50% | 3 — Blocked |

---

## 5. Warning-only Diagnostics (non-blocking)

Produced only in State 2 (normalization succeeded):

| Code | Meaning |
|---|---|
| `TITLE_ROWS_DETECTED` | One title row was stripped |
| `SUBTITLE_ROWS_DETECTED` | One subtitle row was stripped |
| `BANNER_ROWS_DETECTED` | Banner/header rows were extracted |
| `LABEL_COLUMNS_DETECTED` | Label column(s) were separated from data |
| `NORMALIZATION_CONFIDENCE_MEDIUM` | Decomposition succeeded but with reduced certainty |

---

## 6. Behavior Matrix

| State | `normalizationNeeded` | `normalizationApplied` | `blockingReasons` | Required caller action |
|---|---|---|---|---|
| Pure numeric-only | `false` | `false` | `[]` | Preserve existing strict workflow |
| Normalized (high/medium) | `true` | `true` | `[]` | Use normalized grids; show any warnings |
| Low confidence | `true` | `false` | `["LOW_CONFIDENCE"]` | **Stop. Show blockingMessage. Do not run.** |
| Structurally blocked | `true` | `false` | `["BODY_TOO_SHORT", ...]` | **Stop. Show blockingMessage. Do not run.** |

---

## 7. Coordinate Mapping

When a future integration wires `normalizeSelectedRange()` into Run significance:

```
absoluteExcelRow = selectedRange.rowIndex    + dataRowOffset + normalizedBodyRow
absoluteExcelCol = selectedRange.columnIndex + dataColOffset + normalizedBodyCol
```

`dataRowOffset` and `dataColOffset` are pre-computed in every State 2 result. This mapping is required because `excel-writer.js` currently assumes the result matrix is 1:1 aligned with `selectedRange`. With normalization, the result matrix covers only the body, so the writer must apply the offset.

Coordinate mapping integration is explicitly scoped to Phase 3 (opt-in Run significance). It must not be implemented in the Phase 1 or Phase 2 PRs.

---

## 8. Relationship to Existing Components

### `detectSelectedRangeGuardrails` (taskpane.js)

The pass-through gate in `range-normalizer.js` uses the same structural heuristics as the existing guardrails (left-column text fraction, top-row text fraction, overall text-heavy fallback). In future phases, the guardrail warnings can be replaced by normalization info messages when normalization succeeds. They remain in place for Phase 1.

### `buildTablePreviewModel` (table-preview-model.js)

Accepts `{values, leftLabelValues, bannerContext, settings}`. The normalized model's `valuesForCalculation`, `leftLabelValues`, and `bannerContext` map directly onto these parameters. Phase 2 integration feeds the normalized grids to `buildTablePreviewModel` with no changes to `table-preview-model.js`.

### `detectBannerStructure` (banner-detector.js)

Expects a banner context built from rows above the selected range. Phase 2/3 integration will pass `bannerContext.scanRows` from the normalized model instead of (or in addition to) the rows loaded above the selection. No changes to `banner-detector.js`.

### `table-inventory-scanner.js`

Uses the same cell classification and text fraction heuristics independently. The normalizer implements its own copies of these primitives; shared extraction is a future refactor, not part of this issue.

---

## 9. Accepted / Rejected Selection Examples

### Accepted — State 1 (pass-through)

| Selection | Result |
|---|---|
| 5×4 all numeric | Pass-through; strict workflow unchanged |
| 4×2 grid (too small for structural analysis) | Pass-through; too small |

### Accepted — State 2 (normalization succeeded)

| Selection content | Detected layout | Confidence |
|---|---|---|
| Col A text labels + cols B–E numeric | `labelColumns=[0]`, `dataColOffset=1` | high |
| Cols A–B text + cols C–E numeric | `labelColumns=[0,1]`, `dataColOffset=2` | high |
| Row 1 sparse title + rows 2–10 body | `titleRows=[0]`, `dataRowOffset=1` | high |
| Row 1 title + row 2 subtitle + body | `titleRows=[0]`, `subtitleRows=[1]`, `dataRowOffset=2` | high |
| Row 1 title + row 2 wide banner + col A labels + body | Full decomposition | high |
| Full NPS table (title + banner + labels + metric rows) | All areas extracted | high |

### Rejected — State 3 (blocked, caller must stop)

| Selection content | Blocking reason | Message |
|---|---|---|
| Two tables, empty row between them | `BODY_APPEARS_MULTI_TABLE` | "...пустые строки...несколько таблиц..." |
| After stripping, 1 body row | `BODY_TOO_SHORT` | "Недостаточно строк данных..." |
| After stripping, 1 data column | `DATA_TOO_NARROW` | "Недостаточно столбцов данных..." |
| No numeric cells in body | `NO_NUMERIC_BODY` | "...не найдено числовых данных..." |
| Ambiguous label split + text-heavy body | `LABEL_SPLIT_BLOCKING` | "...Выделите только числовую область данных." |
| Body < 50% numeric | `LOW_CONFIDENCE` | "...Выделите только числовую область данных." |

---

## 10. Implementation Phases

### Phase 1 — Pure module (complete)

`src/core/range-normalizer.js` exports `normalizeSelectedRange(rawValues, rawText, options)`.

No runtime wiring. No changes to high-risk files. No test framework added.

### Phase 2 — Check table preview integration

Wire `normalizeSelectedRange()` into `runCheckTable()` in `taskpane.js`.

When `normalizationApplied: true`, pass the normalized grids to `buildTablePreviewModel()`. When blocked, show `blockingMessage` in the Check table result panel.

Allowed file: `taskpane.js` (`runCheckTable()` function only). No Excel writes. No coordinate mapping needed.

### Phase 3 — Opt-in Run significance integration

Wire `normalizeSelectedRange()` into `runSignificanceFromSelection()` behind an opt-in setting. Implement coordinate mapping so `writeCellResultsToSelectedRange()` offsets by `(dataRowOffset, dataColOffset)`.

Allowed files: `taskpane.js` (Run path), `excel-writer.js` (offset parameter).

---

## 11. Regression Risks

- **Calculation shifts:** If `dataRowOffset` or `dataColOffset` is off by one, markers land in wrong cells.
- **Coordinate mapping failures:** Incorrect offset in `excel-writer.js` corrupts the entire output range.
- **Banner detection breakage:** `detectBannerStructure()` may fail on an incorrect `bannerContext`.
- **Metric detection breakage:** `detectMetricRowsFromLeftLabels()` relies on label text being in `leftLabelValues`; an incorrect label split breaks row-type detection.
- **Strict workflow regression:** The pass-through gate must fire correctly for pure-numeric selections or the existing behavior changes.

---

## 12. Known Limitations (Phase 1)

- A sparse title row (1 text cell, ≤3 non-empty) in an otherwise pure-numeric selection may not trigger the pass-through gate; the normalizer will not detect it. This is an acceptable gap for Phase 1 — such a selection is nearly-numeric and the existing strict workflow handles it with the same limitation it has today.
- `rawText` is optional; if omitted, `textForCalculation` is an empty array.
- `options` parameter is reserved and unused in Phase 1.
- No test framework is included. Manual validation uses inline console checks or the fixture approach described in the implementation issue.

