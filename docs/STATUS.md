# STATUS.md

## Current status

Research Insights Toolkit MVP functionality is implemented for the Excel-first workflow.

The add-in currently supports:

- automatic metric block detection;
- proportions significance testing;
- means significance testing with SD or variance;
- NPS significance testing from promoters/detractors structure;
- NPS significance testing from SD or variance;
- confidence level selector;
- one-tailed / two-tailed testing;
- local settings persistence;
- reset to default settings;
- small-base exclusion and fill;
- Total comparison modes;
- previous-column comparison mode;
- banner-aware structure detection;
- banner-aware group comparisons;
- banner-aware local and global Total logic;
- wave banner auto previous-column mode;
- banner-local marker indexing;
- banner letter writing into the lowest banner level;
- clean user-facing status messages.

## Implemented UI settings

### Significance settings

Implemented:

- `confidence-level`
- `one-tailed-test`
- `round-cell-values`

### Previous-column comparison

Implemented:

- `compare-with-previous-column`
- `apply-previous-column-fill`

Behavior:

- previous-column mode writes arrows instead of letter markers;
- arrow is written only into the right/current column;
- upward difference uses `↑`;
- downward difference uses `↓`;
- fill is enabled by default when previous-column mode is selected;
- user may manually disable previous-column fill;
- previous-column mode disables banner letter writing;
- previous-column mode is incompatible with compare-only-with-Total.

### Banner settings

Implemented:

- `write-banner-letters`
- `respect-banner-structure`

Behavior:

- without banner structure, banner letters follow selected-column indexing;
- with banner structure, banner letters are written only to the lowest banner level;
- upper banner levels are not modified;
- labels are local to each detected banner group;
- Total columns never receive ordinary banner letters;
- wave groups using auto previous-column mode do not receive banner letters.

### Label detection

Implemented:

- default label lookup immediately to the left of selected range;
- optional lookup from the left side of the sheet via `labels-on-left-side`;
- numeric columns between selected data and real text labels are skipped.

### Total comparison settings

Implemented:

- `compare-only-with-total`
- `exclude-total-from-comparisons`
- `first-column-is-total`
- `total-in-each-banner`

Current behavior:

- without banner structure, supported manual Total mode is primarily `first-column-is-total`;
- with banner structure enabled, manual Total placement checkboxes are disabled;
- Total placement is then detected from banner structure;
- compare-only-with-Total and exclude-Total remain meaningful with banner structure.

### Fill settings

Implemented:

- `significant-fill-color`
- `lower-than-total-fill-color`
- `fill-only-total-comparisons`
- `small-base-fill-color`

Fill priority:

1. small base fill
2. lower than Total fill
3. normal significance fill
4. no fill

### Small bases

Implemented:

- `exclude-small-bases`
- `small-base-threshold`

Behavior:

- columns with base lower than threshold are excluded before significance calculation;
- small-base fill is applied to the whole affected column within the calculation block;
- the base row itself is also filled;
- if a manual first-column Total has a small base, calculation stops with an error.

### Settings storage

Implemented:

- `settings-storage-mode = none`
- `settings-storage-mode = local`
- reset button

Behavior:

- local settings are stored in `localStorage`;
- reset restores default settings and clears saved local settings;
- cloud storage remains reserved for future implementation.

## Banner engine status

Implemented MVP:

- one-level banner detection;
- two-level banner detection;
- repeated group label detection;
- reconstructed span detection for merged-like headers;
- local Total detection;
- global Total detection;
- group-aware ordinary comparisons;
- group-local cell markers;
- local Total as group reference when no global Total exists;
- global Total as the only Total reference when detected;
- local Totals compared with global Total when global Total exists;
- local Totals not used as group references when global Total exists;
- Total columns excluded from ordinary group comparisons;
- previous-column comparison inside banner groups;
- automatic previous-column mode for wave groups;
- banner-aware lower-level letter writing;
- user-visible banner messages filtered to important events only.

## Wave banner behavior

Wave-like groups are detected from group labels such as:

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

When a wave group is detected and global previous-column mode is not manually enabled:

- previous-column comparison is applied only inside wave groups;
- non-wave groups continue to use ordinary group comparisons;
- UI checkbox state is not changed;
- previous-column fill is applied automatically for wave groups;
- banner letters are not written for wave groups;
- status message explains that auto previous-column was applied.

Plain numeric labels such as `1, 2, 3` are not used as wave signals.

## Statistical engine status

Implemented:

- pooled z-test for proportions;
- Welch’s t-test for means;
- NPS from promoter/detractor structure;
- NPS from spread;
- one-tailed and two-tailed threshold modes;
- modular threshold functions using distribution quantiles.

Notes:

- product-specific comparison routing remains custom;
- external/statistical libraries are used only for threshold calculation;
- formulas remain under project control.

## User-facing status messages

Implemented:

- default success status is concise;
- technical banner diagnostics are hidden;
- only user-relevant banner messages are shown.

Visible banner message types include:

- global Total used;
- auto previous-column applied for wave groups;
- compare-only-with-Total produced no valid Total pairs;
- multiple local Totals in one group;
- malformed or unsupported banner structure;
- no banner rows above selection.

## Known technical debt

- Some old diagnostic helpers for merge/span investigation may remain in `taskpane.js` behind no active call path.
- `formatBannerDetectionDiagnostics()` remains useful for development diagnostics but should not be used in normal user status.
- Google Sheets support is not implemented.
- Cloud settings storage is not implemented.
- Full multi-level banner support beyond MVP is not implemented.
- Report-title detection and broader table-boundary detection are not fully implemented.
- Total outside selection is specified but may need additional edge-case hardening.
- The add-in remains Excel-first.