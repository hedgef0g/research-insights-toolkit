# Changelog

## [Current Development Stage]

### Stabilization Release

#### Added
- Warning-only selected range guardrails (temporary safety net).
- Shared agent instructions added to `AGENTS.md` and referenced in `CLAUDE.md`.

#### Improved
- Numeric output preservation for unmarked cells (avoids converting the whole selected range to text).
- Preservation of display conventions: `28` remains plain, `28%` remains percent, `0.28` remains decimal-share.
- Clear significance numeric restoration (`21% b` clears to `21%`, `28.1 b` clears to `28.1`, `0.281 b` clears to `0.281`).
- Run / Clear actions moved near the top of the taskpane.

#### Fixed
- Banner-aware comparisons now handle the first pair inside a multi-column group more correctly.
- Multi-row / merged-like banner group detection was improved.
- Banner letters are now written to visible banner labels when labels are located above the row directly adjacent to the data.

#### Notes
- Selected range normalization is not implemented yet.
- Worksheet/workbook auto-scan is not implemented.
- Full-table selection support is future product/spec work.


### 30.04.2026 — Banner-aware MVP and completed UI settings logic

#### Added

- Banner structure detection for Excel crosstab-style tables:
  - one-level banners;
  - two-level banners;
  - repeated adjacent group labels;
  - reconstructed span detection for merged-like headers.
- Banner-aware comparison logic:
  - comparisons are limited to detected banner groups;
  - columns from different groups are not compared;
  - marker indexing is local to each banner group.
- Local Total logic inside banner groups:
  - local Total is used as the group reference when no global Total exists;
  - local Total columns are excluded from ordinary segment-vs-segment comparisons;
  - local Total columns never receive ordinary letter labels.
- Global Total logic:
  - global Total can be detected from the banner structure;
  - global Total becomes the only Total reference when present;
  - local Totals are compared with global Total as target columns;
  - local Totals are not used as intra-group references when global Total exists.
- Banner-aware Total modes:
  - compare only with Total;
  - exclude Total from comparisons;
  - concise status message when compare-only-with-Total produces no valid pairs;
  - error stop when a group contains multiple local Totals and no global Total resolves the ambiguity.
- Banner-aware previous-column mode:
  - previous-column comparisons stay inside detected banner groups;
  - previous-column comparisons do not cross group boundaries;
  - Total columns are excluded from previous-column chains under banner structure;
  - small-base excluded columns are not skipped over.
- Automatic previous-column mode for wave-like banner groups:
  - wave groups use previous-column comparisons automatically;
  - non-wave groups keep ordinary group comparisons;
  - UI checkbox state is not changed;
  - previous-column fill is applied automatically for auto wave comparisons;
  - banner letters are not written for auto wave groups.
- Wave-group detection based on group labels:
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
- Banner letter writing with banner structure:
  - letters are written only into the lowest banner level;
  - upper banner levels are not modified;
  - letters are local to each detected group;
  - Total columns never receive ordinary banner letters;
  - old trailing banner markers are replaced or removed as needed.
- One-tailed test setting:
  - new checkbox next to the confidence level selector;
  - applies to proportions, means, NPS, Total comparisons, previous-column comparisons, and banner-aware comparisons.
- Modular statistical threshold calculation:
  - threshold logic moved to a dedicated module;
  - z and Student's t critical values are calculated through distribution quantiles rather than static approximation tables.
- Local settings persistence:
  - `settings-storage-mode = none`;
  - `settings-storage-mode = local`;
  - reset button restores default settings and clears saved local settings.
- Small-base handling:
  - columns with bases below threshold are excluded before significance calculation;
  - small-base fill is applied to the full relevant calculation block;
  - small-base fill has the highest fill priority.
- Clean user-facing status messages:
  - technical banner diagnostics are hidden from normal status output;
  - only relevant banner messages are shown to the user.
- Optional wave banner auto-detection setting.
  - Wave auto previous-column mode is now disabled by default.

#### Improved

- Completed the functionality originally reserved by the UI settings panel.
- Previous-column fill is now enabled by default when previous-column mode is selected.
- Previous-column mode disables banner letters and compare-only-with-Total.
- Banner structure mode disables manual Total placement checkboxes, because Total placement is detected by the banner engine.
- Total-related settings remain valid under banner structure:
  - compare only with Total;
  - exclude Total from comparisons.
- User status output is now concise:
  - simple successful calculations show only the success message;
  - wave/global Total/banner errors add only user-relevant messages.
- Updated Russian help page.

#### Fixed

- Prevented Total columns from receiving ordinary group-local letters.
- Prevented Total columns from participating in ordinary banner group comparisons.
- Fixed group-local marker propagation through row comparison objects.
- Fixed banner-aware previous-column behavior with excluded Total columns.
- Fixed old banner markers remaining on columns that should no longer receive letters.
- Fixed excessive technical banner diagnostics appearing in the user status panel.
- Fixed UI state conflicts between previous-column mode, Total settings, banner letters, and banner structure mode.

---

### 28.04.2026

#### Added

- Unified significance engine with automatic metric detection.
- Support for significance testing of proportions.
- Support for significance testing of means using:
  - standard deviation;
  - variance.
- Support for significance testing of NPS using:
  - Promoters / Detractors structure;
  - standard deviation;
  - variance.
- Automatic pairwise all-vs-all comparison across selected columns.
- Automatic significance letters appended directly into Excel cells.
- Pale green highlight + bold formatting for significant cells.
- Automatic center alignment for selected ranges.
- Separate button for clearing significance markers.
- Metric type diagnostics tool.
- Shared-base support across mixed metric tables.
- Support for mixed tables containing:
  - proportions;
  - means;
  - NPS;
  - any combination of them.

#### Improved

- Replaced single-plan detection with block-plan detector.
- Smarter handling of tables where one base row serves several metric blocks.
- Better compatibility with real research tables.
- Added protection against Excel converting values into time format after adding letters.
- Improved project launch workflow with automatic VS Code startup tasks.

#### Fixed

- NPS significance routing issues in auto mode.
- Incorrect significance letters appearing in Promoters / Detractors rows.
- Branch execution conflicts between mean / NPS / fallback logic.
- Repeated marker formatting inconsistencies.

---

### 29.04.2026 — Refactoring pass 1

Unified automatic metric detection through `buildCalculationBlocks` became stable, so old manual calculation modes and their dedicated parsers were removed.

#### Changed

- `taskpane.js`:
  - removed handlers and event listeners for explicit mean and NPS calculation buttons;
  - removed unused imports.
- `core/metric-detector.js`:
  - removed deprecated `buildAutoCalculationPlan`, fully replaced by `buildCalculationBlocks`.
- `core/significance.js`:
  - removed legacy MVP v0.2 `compareAllRowsUsingBottomBases`;
  - removed wrappers for retired manual calculation flows.

#### Notes

- The statistical core remained intact.
- Metric diagnostics remained available.

---

### 29.04.2026 — Refactoring pass 2

#### Changed

- Excel rendering was moved from `taskpane.js` into `core/excel-writer.js`.
- Dirty Excel value normalization was moved from `significance.js` into `core/normalizers.js`.
- Metric dictionary arrays were moved into `core/dictionary.config.js`.
- `taskpane.js` now acts primarily as the Excel/UI controller.
- `significance.js` contains statistical and comparison logic.
- `metric-detector.js` contains detection and block-plan construction.

---

### 29.04.2026 — Feature pass 1

#### Added

- Configurable confidence level selector:
  - 99%;
  - 95%;
  - 90%;
  - 80%;
  - 66.6%.
- Confidence level applies to:
  - proportions;
  - means;
  - NPS from spread;
  - NPS from Promoters / Detractors structure.

#### Improved

- Re-running the calculation fully clears old markers and formatting before recalculation.
- Block-plan detector improved for:
  - shared bases;
  - mixed metric tables;
  - avoiding skipped metric blocks after Mean / NPS.

---

### 29.04.2026 — Feature pass 2

#### Added

- Collapsible **Settings** block.
- Settings UI foundation for:
  - banner letters;
  - banner structure;
  - left-side label detection;
  - Total comparison modes;
  - Total placement modes;
  - fill colors;
  - small bases;
  - settings storage.
- **Round cell values** setting.
- Banner letter writing above selected range:
  - marker format: `(a)`, `(b)`, `(c)`;
  - existing trailing marker is replaced;
  - duplicate trailing marker is not added.
- Optional metric label lookup from the leftmost sheet columns.
- Warning when banner letter writing is enabled but the selected range starts in the first worksheet row.

#### Improved

- Main calculation button renamed to **Запустить**.
- Display rounding before marker insertion:
  - default proportions / NPS / Promoters / Detractors: 1 decimal;
  - default means / SD / variance: 2 decimals;
  - with rounding enabled, proportions / NPS / Promoters / Detractors: integer;
  - with rounding enabled, means / SD / variance: 1 decimal.
- Improved label lookup to skip numeric intermediate columns.

#### Fixed

- Numeric values between labels and selected data are no longer mistaken for metric labels.
