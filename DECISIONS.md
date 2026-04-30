# Decisions

## Current product decisions

### Product direction

- Research Insights Toolkit is an Excel-first significance testing add-in for market research and consumer insights tables.
- Google Sheets support remains a future direction, not part of the current MVP.
- The main user flow is a single button: **Запустить**.
- New user-facing calculation behavior should be added through the settings panel rather than by adding separate buttons.
- Real research tables are prioritized over idealized academic examples.

### Architecture

- `taskpane.js` is the Excel/UI controller.
- Statistical calculation logic must not be moved back into `taskpane.js`.
- Excel rendering is handled by `core/excel-writer.js`.
- Metric detection is handled by `core/metric-detector.js`.
- Statistical and comparison logic is handled by `core/significance.js`.
- Value normalization is handled by `core/normalizers.js`.
- Metric dictionary terms live in `core/dictionary.config.js`.
- Banner detection lives in `core/banner-detector.js`.
- Statistical threshold logic lives in `core/stat-thresholds.js`.

### Data/privacy

- Calculations run locally in the Excel add-in WebView.
- Selected spreadsheet data should not be sent to a server unless an explicit future feature requires it and the user is informed.

---

## Metric detection decisions

- Single-plan detection is deprecated in favor of block-plan detection.
- A selected range may contain multiple independent calculation blocks.
- Base rows may be:
  - dedicated to one metric;
  - shared by several metrics.
- Proportion rows may wait for the next available base row.
- Supported calculation blocks:
  - proportions + Base;
  - Mean + SD + Base;
  - Mean + Variance + Base;
  - NPS + Promoters + Detractors + Base;
  - NPS + SD + Base;
  - NPS + Variance + Base.

---

## Output rules

- Significance markers may appear only in actual value rows:
  - proportions;
  - Mean;
  - NPS.
- Service rows must never receive significance markers:
  - Base;
  - SD;
  - Variance;
  - Promoters;
  - Detractors.
- Service rows may still receive formatting when relevant, for example small-base fill.

---

## Statistical decisions

### Proportions

- Proportions are compared using a pooled z-test.
- Input percentages/shares are normalized before calculation.
- Calculations use normalized original values, not rounded display values.

### Means

- Means are compared using Welch’s t-test.
- Spread input may be SD or variance.
- Welch’s t-test is preferred because it does not assume equal variances.

### NPS

Two NPS modes are supported:

1. NPS from structure:
   - NPS;
   - Promoters;
   - Detractors;
   - Base.

2. NPS from spread:
   - NPS;
   - SD or variance;
   - Base.

For NPS from structure:

- NPS is recalculated from Promoters minus Detractors.
- NPS is treated as a mean on the scale:
  - promoter = +1;
  - passive = 0;
  - detractor = -1.

### Confidence levels and one-tailed tests

- Supported confidence levels:
  - 99%;
  - 95%;
  - 90%;
  - 80%;
  - 66.6%.
- Tests are two-tailed by default.
- The **Односторонний тест** setting changes only the critical threshold.
- Direction and markers remain based on the observed difference.
- One-tailed mode applies to:
  - proportions;
  - means;
  - NPS;
  - Total comparisons;
  - previous-column comparisons;
  - banner-aware comparisons.
- Critical values are calculated through modular distribution quantile functions rather than static approximation tables.

---

## UI / settings decisions

- Settings panel is the main place for user-facing behavior.
- The settings block is collapsible.
- Settings can be reset to defaults.
- Local settings persistence uses `localStorage`.
- Cloud settings storage is reserved for future implementation.
- Dependent UI states should prevent invalid combinations where possible.

### Previous-column mode

- Previous-column mode writes arrows instead of ordinary letters.
- Arrow is written only into the right/current column.
- `↑` means current column is significantly higher than previous column.
- `↓` means current column is significantly lower than previous column.
- Previous-column fill is enabled by default when previous-column mode is selected.
- User may manually disable previous-column fill.
- Previous-column mode disables banner letter writing.
- Previous-column mode is incompatible with compare-only-with-Total.

### Fill priority

Fill priority is:

1. small base fill;
2. lower than Total fill;
3. normal significance fill;
4. no fill.

### Small bases

- Small-base checks happen before significance calculations.
- If a column base is below threshold:
  - the column is excluded from comparisons for that calculation block;
  - the whole affected block column is filled;
  - the base row is filled too.
- If a manual first-column Total has a small base, calculation stops with a clear error.

---

## Label detection decisions

- Default behavior: search metric labels immediately to the left of the selected range.
- Optional behavior: search metric labels from the leftmost sheet columns.
- Numeric values between labels and selected data are skipped.
- This supports tables where the user excludes Total or selects only the right side of a wide table.

---

## Total comparison decisions

### Without banner structure

- Manual Total mode is primarily `first-column-is-total`.
- In first-column Total mode:
  - column 0 is Total;
  - Total receives no ordinary letters;
  - segments may be compared with Total;
  - segments may also be compared with each other unless compare-only-with-Total is active.
- `T` means segment is significantly higher than Total.
- `t` means segment is significantly lower than Total.
- `compare-only-with-total` creates only Total-vs-segment pairs.
- `exclude-total-from-comparisons` removes Total from all comparisons.

### With banner structure

- Manual Total placement settings are disabled.
- Total placement is detected from the banner.
- `compare-only-with-total` and `exclude-total-from-comparisons` remain valid.
- Ordinary group comparisons never include Total columns.
- Total columns never receive ordinary letter labels.

---

## Banner engine decisions

### Scope

The MVP banner engine supports:

- one-level banners;
- two-level banners;
- repeated adjacent group labels;
- reconstructed span detection for merged-like headers;
- local Totals;
- global Total;
- group-aware comparisons;
- banner-aware previous-column;
- wave group auto previous-column;
- banner-local letter writing.

### Banner detection

- Lower banner level is the row directly above the selected range.
- Meaningful group level is the nearest valid group level above the lower level.
- Numeric banner labels are valid.
- Plain empty rows or non-meaningful rows are ignored where possible.
- Reconstructed span detection is used for merged-like headers where Office.js does not expose full merged-area metadata reliably.

### Group-aware comparisons

- When `respect-banner-structure` is enabled, columns from different banner groups are not compared.
- Marker labels are local to each group.
- Total columns are excluded from ordinary group comparisons.

### Local Total

- Local Total is used as group reference only when no global Total exists.
- Local Total columns do not receive ordinary letters.
- Local Total columns are excluded from ordinary segment-vs-segment comparisons.

### Global Total

- When global Total is detected, it becomes the only Total reference.
- Local Totals are not used as intra-group references.
- Local Totals may be compared with global Total as target columns.
- Local Totals may receive `T/t` markers from global Total comparisons.
- Local Totals still never receive ordinary letter labels.
- Status explains global Total behavior.

### Multiple local Totals

- If one group contains multiple local Totals and no global Total resolves the reference, calculation stops with a clear error.

### Banner letters

- Banner letters are written only into the lowest banner level.
- Upper banner levels are not modified.
- Letters are local to each group.
- Total columns never receive ordinary banner letters.
- Old trailing markers are replaced or removed as needed.
- If the selected range starts in the first row and banner letters are enabled, calculation stops with a warning.

### Wave groups

Wave-like groups are detected by group label keywords:

- `wave`
- `waves`
- `волна`
- `волны`
- `period`
- `periods`
- `период`
- `периоды`
- `замер`
- `замеры`

Decisions:

- plain numeric lower labels like `1, 2, 3` are not enough to classify a group as wave-like;
- if a wave group is detected and global previous-column mode is off, previous-column comparison is applied only inside the wave group;
- non-wave groups keep ordinary group comparisons;
- the UI checkbox is not toggled;
- previous-column fill is applied automatically to auto wave comparisons;
- banner letters are not written for auto wave groups;
- a concise status message explains the automatic behavior.

---

## User-facing status decisions

- Normal successful calculations should show concise status only.
- Technical banner diagnostics must not appear in regular user status.
- User-visible banner messages are restricted to relevant events:
  - global Total used;
  - auto previous-column applied for wave groups;
  - compare-only-with-Total produced no valid pairs;
  - multiple local Totals in one group;
  - malformed/unsupported banner structure;
  - missing rows above selection.
- Developer diagnostics may remain in helper functions but must not be called in normal user flow.

---

## Known limitations / future decisions

- Native `window.confirm()` in Office Add-in WebView is unreliable; future confirm flows should use custom HTML UI.
- Automatic row insertion above first-row selections is not implemented.
- Full three-level or deeper banner support is not implemented.
- Report-title detection and broad table-boundary detection require further hardening.
- Total outside selected range is specified conceptually but may need more edge-case handling.
- Google Sheets support is not implemented.
- Cloud settings storage is not implemented.
