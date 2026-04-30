# Research Insights Toolkit

Excel-first significance testing toolkit for market research and consumer insights teams.

Research Insights Toolkit helps users run significance testing directly in Excel tables without manually preparing data for SPSS or other statistical software.

## Core workflow

1. Select a table range in Excel.
2. Click **Запустить**.
3. The add-in:
   - detects metric rows from labels;
   - builds calculation blocks;
   - calculates statistical significance;
   - writes markers directly into cells;
   - applies formatting;
   - optionally writes group-local letters into the banner row.

All calculations run locally inside the Excel add-in.

## Supported metric types

The unified auto-detection mode supports:

- proportions / percentages;
- means with standard deviation;
- means with variance;
- NPS from Promoters / Detractors structure;
- NPS from standard deviation;
- NPS from variance.

Mixed tables are supported. One selected range may contain several metric blocks, including shared-base layouts.

## Statistical methods

Implemented methods:

- pooled z-test for proportions;
- Welch’s t-test for means;
- NPS structure calculation through promoter/passive/detractor mean logic;
- NPS spread calculation through mean logic.

Supported confidence levels:

- 99%;
- 95%;
- 90%;
- 80%;
- 66.6%.

Tests are two-tailed by default. The **Односторонний тест** setting switches threshold calculation to one-tailed mode while keeping marker direction based on the observed difference.

Critical thresholds are calculated through modular distribution quantile functions.

## Output

The add-in can write:

- ordinary significance letters;
- `T` / `t` markers for Total comparisons;
- `↑` / `↓` arrows for previous-column comparisons;
- bold formatting;
- configurable fill colors;
- small-base fill;
- optional banner letters.

Marker meanings:

- ordinary letters indicate the column that the marked value is significantly higher than;
- `T` means the value is significantly higher than Total;
- `t` means the value is significantly lower than Total;
- `↑` means the current column is significantly higher than the previous column;
- `↓` means the current column is significantly lower than the previous column.

## Settings

Implemented settings include:

- confidence level;
- one-tailed test;
- cell value rounding;
- previous-column comparison;
- previous-column fill;
- banner letter writing;
- banner structure detection;
- left-side label lookup;
- compare only with Total;
- exclude Total from comparisons;
- first column is Total;
- Total in each banner;
- significant fill color;
- lower-than-Total fill color;
- fill only for Total comparisons;
- small-base exclusion;
- small-base threshold;
- small-base fill color;
- settings storage mode;
- reset to defaults.

Local settings persistence is supported through `localStorage`.

## Total comparison logic

Without banner structure:

- the primary supported manual mode is **first column is Total**;
- Total receives no ordinary letters;
- `compare only with Total` keeps only Total comparisons;
- `exclude Total` removes Total from all comparisons.

With banner structure:

- manual Total placement settings are disabled;
- Total placement is detected from the banner;
- local Total can act as a group reference;
- global Total, when detected, becomes the only Total reference;
- local Totals are compared with global Total when global Total exists;
- Total columns never receive ordinary letter labels;
- Total columns are excluded from ordinary group comparisons.

## Banner structure support

The banner engine supports:

- one-level banners;
- two-level banners;
- repeated adjacent group labels;
- reconstructed span detection for merged-like headers;
- local Total detection;
- global Total detection;
- group-aware comparisons;
- group-local markers;
- banner-aware previous-column mode;
- automatic previous-column mode for wave groups;
- banner-local letter writing.

When banner structure is enabled:

- comparisons are limited to detected groups;
- columns from different groups are not compared;
- ordinary marker labels restart inside each group;
- banner letters are written only to the lowest banner level;
- upper banner levels are not modified;
- Total columns never receive ordinary banner letters.

## Wave groups

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

When a wave group is detected:

- previous-column comparison is applied automatically only inside that group;
- non-wave groups keep ordinary group comparisons;
- the UI checkbox is not changed;
- previous-column fill is applied automatically;
- banner letters are not written for that wave group;
- status explains the automatic behavior.

Plain numeric labels such as `1, 2, 3` are not used by themselves as wave signals.

## Small bases

Small-base logic runs before statistical testing.

If a column base is below the configured threshold:

- the column is excluded from comparisons in the current calculation block;
- the affected block cells in that column are filled;
- the base row itself is filled;
- the column receives no markers.

Small-base fill has the highest fill priority.

## Status messages

User-facing status messages are intentionally concise.

Technical banner diagnostics are hidden from normal output. The status panel shows only relevant events such as:

- calculation completed;
- global Total used;
- automatic previous-column applied for wave groups;
- no valid Total pairs in compare-only-with-Total mode;
- multiple local Totals in one group;
- missing banner rows above selection.

## Development notes

Primary architecture:

- `taskpane/taskpane.js` — Excel/UI controller;
- `core/metric-detector.js` — metric row detection and block-plan construction;
- `core/dictionary.config.js` — config-driven metric dictionary;
- `core/normalizers.js` — normalization of Excel values;
- `core/significance.js` — statistical tests, comparison routing, markers, fill reasons;
- `core/stat-thresholds.js` — statistical critical thresholds;
- `core/banner-detector.js` — platform-independent banner structure detection;
- `core/excel-writer.js` — writing values and formatting back to Excel.

## Strategic goal

Create the most practical spreadsheet significance tool for research managers and insights teams.

The focus is Excel first. Google Sheets support remains a future direction.

## Roadmap

Potential future work:

- weighted bases;
- additional correction methods for multiple comparisons;
- p-values in diagnostics;
- richer multi-level banner support;
- stronger report-title and table-boundary detection;
- Total outside selection hardening;
- Google Sheets support;
- cloud settings storage;
- additional research utilities such as PSM and table QA tools.
