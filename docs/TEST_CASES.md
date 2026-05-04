# Research Insights Toolkit — Manual Test Cases

Last updated: 2026-04-30

## Purpose

This document lists manual regression test cases for Research Insights Toolkit.

The goal is to protect the current MVP behavior while adding new comparison modes, fill rules, and UI settings.

## General Assumptions

Unless stated otherwise:

- User selects only the numeric data area, not row labels.
- Row labels are located immediately to the left of the selected range.
- Confidence level is 95%.
- “Round cell values” can be off unless a test explicitly covers rounding.
- Old markers and old fills should be cleared on every run before applying new results.
- Small-base threshold is 50 where relevant.
- First selected data column is column index 0 internally.

---

# 1. Basic Proportion Tests

## 1.1. All-vs-all proportions

### Input

| Label | A | B | C |
|---|---:|---:|---:|
| % | 50% | 70% | 51% |
| Base | 500 | 500 | 500 |

### Settings

- First column is Total: off
- Compare only with Total: off
- Exclude Total from comparisons: off
- Compare with previous column: off
- Exclude small bases: off

### Expected

- All column pairs are compared:
  - A vs B
  - A vs C
  - B vs C
- Higher significant cells receive ordinary column markers.
- Base row receives no markers.
- Significant cells receive normal significant fill.

---

## 1.2. Re-run cleanup

### Setup

Run test 1.1 once.

Then change confidence level or input values so that previous significant differences are no longer significant.

### Expected

- Old markers are removed.
- Old fills are removed.
- Old bold formatting is removed.
- New results reflect only the current run.

---

# 2. Mean Tests

## 2.1. Mean + SD + Base

### Input

| Label | A | B | C |
|---|---:|---:|---:|
| Mean | 3.5 | 4.2 | 3.6 |
| SD | 1.1 | 1.0 | 1.1 |
| Base | 300 | 300 | 300 |

### Settings

- Compare with previous column: off
- Exclude small bases: off

### Expected

- Welch’s t-test is used.
- Markers are written only into the Mean row.
- SD and Base rows receive no markers.
- Significant Mean cells receive normal significant fill.

---

# 3. NPS Tests

## 3.1. NPS-first format

### Input

| Label | A | B | C |
|---|---:|---:|---:|
| NPS | 30 | 45 | 25 |
| Promoters | 50% | 60% | 45% |
| Detractors | 20% | 15% | 20% |
| Base | 500 | 500 | 500 |

### Expected

- NPS is recalculated from Promoters - Detractors.
- NPS is treated as mean of:
  - promoter = +1;
  - passive = 0;
  - detractor = -1.
- NPS receives NPS significance markers.
- Promoters receives ordinary proportion markers.
- Detractors receives ordinary proportion markers.
- Base receives no marker.

---

## 3.2. NPS-first with Neutral

### Input

| Label | A | B | C |
|---|---:|---:|---:|
| NPS | 30 | 45 | 25 |
| Promoters | 50% | 60% | 45% |
| Neutral | 30% | 25% | 35% |
| Detractors | 20% | 15% | 20% |
| Base | 500 | 500 | 500 |

### Expected

- NPS is recalculated from Promoters - Detractors.
- NPS is treated as mean of:
  - promoter = +1;
  - passive = 0;
  - detractor = -1.
- NPS receives NPS significance markers.
- Promoters receives ordinary proportion markers.
- Neutral receives ordinary proportion markers.
- Detractors receives ordinary proportion markers.
- Base receives no marker.

---

## 3.3. Extended NPS format

### Input

| Label | A | B | C |
|---|---:|---:|---:|
| Top-3 | 50% | 60% | 45% |
| Detractors | 20% | 15% | 20% |
| Neutral | 30% | 25% | 35% |
| Promoters | 50% | 60% | 45% |
| NPS | 30 | 45 | 25 |
| Base | 500 | 500 | 500 |

### Expected

- NPS is recalculated from Promoters - Detractors.
- NPS is treated as mean of:
  - promoter = +1;
  - passive = 0;
  - detractor = -1.
- Top-3 receive ordinary proportion markers.
- Detractors receives ordinary proportion markers.
- Neutral receives ordinary proportion markers.
- Promoters receives ordinary proportion markers.
- NPS receives NPS significance markers.
- Base receives no marker.

---

## 3.4. NPS + SD / Base

### Input

| Label | A | B | C |
|---|---:|---:|---:|
| NPS | 30 | 45 | 25 |
| SD | 0.8 | 0.7 | 0.8 |
| Base | 500 | 500 | 500 |

### Expected

- NPS spread logic is used.
- NPS receives NPS spread significance markers.
- SD receives no marker.
- Base receives no marker.

---

## 3.5. NPS + variance / Base

### Input

| Label | A | B | C |
|---|---:|---:|---:|
| NPS | 30 | 45 | 25 |
| Variance | 0.64 | 0.49 | 0.64 |
| Base | 500 | 500 | 500 |

### Expected

- NPS spread logic is used.
- NPS receives NPS spread significance markers.
- Variance receives no marker.
- Base receives no marker.

---

# 4. Total Comparison Tests

## 4.1. First column is Total: marker indexing

### Input

| Label | Total | Segment 1 | Segment 2 | Segment 3 |
|---|---:|---:|---:|---:|
| % | 60% | 70% | 50% | 65% |
| Base | 500 | 500 | 500 | 500 |

### Settings

- First column is Total: on
- Compare only with Total: off
- Exclude Total from comparisons: off
- Write banner letters: on

### Expected

- Total column receives no letter index.
- Segment 1 = `a`
- Segment 2 = `b`
- Segment 3 = `c`
- Banner markers are written only above segment columns:
  - Total banner cell receives no marker.
  - Segment 1 receives `(a)`.
  - Segment 2 receives `(b)`.
  - Segment 3 receives `(c)`.
- Total cell values never receive markers.

---

## 4.2. Total comparison markers

### Input

| Label | Total | Segment 1 | Segment 2 | Segment 3 |
|---|---:|---:|---:|---:|
| % | 60% | 70% | 50% | 61% |
| Base | 500 | 500 | 500 | 500 |

### Settings

- First column is Total: on
- Compare only with Total: off
- Exclude Total from comparisons: off

### Expected

- Segment significantly higher than Total receives `T`.
- Segment significantly lower than Total receives `t`.
- Segment not significantly different from Total receives no Total marker.
- Total column receives no markers.
- Segment-vs-segment comparisons can also add ordinary markers.
- If a cell has both Total and segment markers, Total marker comes first:
  - example: `Ta`, `tb`.

---

## 4.3. Compare only with Total

### Input

| Label | Total | Segment 1 | Segment 2 | Segment 3 |
|---|---:|---:|---:|---:|
| % | 60% | 70% | 50% | 80% |
| Base | 500 | 500 | 500 | 500 |

### Settings

- First column is Total: on
- Compare only with Total: on
- Exclude Total from comparisons: off

### Expected

- Only these comparisons are performed:
  - Segment 1 vs Total
  - Segment 2 vs Total
  - Segment 3 vs Total
- Segment-vs-segment comparisons are not performed.
- Output contains only `T` / `t` Total markers.
- No ordinary segment letter markers appear.

---

## 4.4. Compare only with Total without Total location

### Settings

- First column is Total: off
- Compare only with Total: on

### Expected

- Calculation stops.
- Status panel shows warning:
  - user must specify Total location;
  - currently supported option is “First column is Total”.
- No calculation is performed.

---

## 4.5. Exclude Total from comparisons

### Input

| Label | Total | Segment 1 | Segment 2 | Segment 3 |
|---|---:|---:|---:|---:|
| % | 60% | 50% | 70% | 55% |
| Base | 500 | 500 | 500 | 500 |

### Settings

- First column is Total: on
- Exclude Total from comparisons: on
- Compare only with Total: off

### Expected

- Total-vs-segment comparisons are not performed.
- Segment-vs-segment comparisons are performed:
  - Segment 1 vs Segment 2
  - Segment 1 vs Segment 3
  - Segment 2 vs Segment 3
- Total column receives no markers.
- No `T` or `t` markers appear.

---

## 4.6. Exclude Total without Total location

### Settings

- First column is Total: off
- Exclude Total from comparisons: on

### Expected

- Calculation stops.
- Status panel shows warning:
  - user must specify Total location;
  - currently supported option is “First column is Total”.
- No calculation is performed.

---

# 5. Total Fill Tests

## 5.1. Lower-than-Total fill

### Input

| Label | Total | Segment 1 | Segment 2 |
|---|---:|---:|---:|
| % | 60% | 50% | 70% |
| Base | 500 | 500 | 500 |

### Settings

- First column is Total: on
- Lower-than-Total fill color: any visible color
- Fill only for Total comparisons: off

### Expected

- Segment 1 receives `t`.
- Segment 1 receives lower-than-Total fill.
- Segment 2 receives `T`.
- Segment 2 receives normal significant fill.

---

## 5.2. Fill only for Total comparisons

### Input

| Label | Total | Segment 1 | Segment 2 | Segment 3 |
|---|---:|---:|---:|---:|
| % | 60% | 70% | 80% | 65% |
| Base | 500 | 500 | 500 | 500 |

### Settings

- First column is Total: on
- Fill only for Total comparisons: on

### Expected

- Cells significantly higher than Total receive normal significant fill.
- Cells significantly lower than Total receive lower-than-Total fill.
- Cells with only segment-vs-segment markers do not receive normal significant fill.
- Cells with only segment-vs-segment markers remain bold and keep their markers.

Example:

| Meaning | Expected |
|---|---|
| `70T` | normal significant fill |
| `80Ta` | normal significant fill |
| `65a` without `T` | no fill, but bold |

---

# 6. Small Base Tests

## 6.1. Small base in proportion block

### Input

| Label | Total | Segment 1 | Segment 2 |
|---|---:|---:|---:|
| % | 50% | 60% | 70% |
| Base | 500 | 30 | 400 |

### Settings

- Exclude small bases: on
- Small-base threshold: 50
- First column is Total: on

### Expected

- Segment 1 is excluded from all comparisons in this block.
- Segment 1 receives no markers.
- Segment 1 receives small-base fill in:
  - % row;
  - Base row.
- Total and Segment 2 are still compared according to active Total settings.

---

## 6.2. Small base in mean block

### Input

| Label | Total | Segment 1 | Segment 2 |
|---|---:|---:|---:|
| Mean | 3.5 | 4.1 | 3.9 |
| SD | 1.2 | 1.1 | 1.3 |
| Base | 300 | 20 | 250 |

### Settings

- Exclude small bases: on
- Small-base threshold: 50

### Expected

- Segment 1 is excluded from all comparisons in this block.
- Small-base fill is applied to Segment 1 in:
  - Mean row;
  - SD row;
  - Base row.
- Segment 1 receives no markers.

---

## 6.3. Small base in NPS structure block

### Input

| Label | Total | Segment 1 | Segment 2 |
|---|---:|---:|---:|
| NPS | 30 | 40 | 25 |
| Promoters | 50% | 55% | 45% |
| Detractors | 20% | 15% | 20% |
| Base | 500 | 35 | 400 |

### Settings

- Exclude small bases: on
- Small-base threshold: 50

### Expected

- Segment 1 is excluded from all comparisons in this block.
- Small-base fill is applied to Segment 1 in:
  - NPS row;
  - Promoters row;
  - Detractors row;
  - Base row.
- Segment 1 receives no markers.

---

## 6.4. Small base in Total column

### Input

| Label | Total | Segment 1 | Segment 2 |
|---|---:|---:|---:|
| % | 50% | 60% | 70% |
| Base | 30 | 300 | 400 |

### Settings

- First column is Total: on
- Exclude small bases: on
- Small-base threshold: 50

### Expected

- Calculation stops.
- Status panel shows warning that Total has a small base and data should be checked.
- No new significance results are written.

Note:

- Old markers may already be cleared before the warning appears. This is acceptable.

---

## 6.5. Small base does not skip over excluded column in previous-column mode

### Input

| Label | Col 1 | Col 2 | Col 3 | Col 4 |
|---|---:|---:|---:|---:|
| % | 50% | 60% | 70% | 80% |
| Base | 500 | 30 | 500 | 500 |

### Settings

- Compare with previous column: on
- Exclude small bases: on
- Small-base threshold: 50

### Expected

- Col 2 is excluded due to small base.
- Col 2 receives small-base fill in the block.
- Col 2 vs Col 1 is not calculated because Col 2 is excluded.
- Col 3 vs Col 2 is not calculated because Col 2 is excluded.
- Col 4 vs Col 3 is calculated.
- Col 3 is not compared with Col 1.

---

# 7. Previous-Column Comparison Tests

## 7.1. Basic previous-column comparison

### Input

| Label | Col 1 | Col 2 | Col 3 |
|---|---:|---:|---:|
| % | 50% | 60% | 55% |
| Base | 500 | 500 | 500 |

### Settings

- Compare with previous column: on
- Apply previous-column fill: off
- First column is Total: off

### Expected

- Col 2 is compared with Col 1.
- Col 3 is compared with Col 2.
- If Col 2 is significantly higher than Col 1, Col 2 receives `↑`.
- If Col 3 is significantly lower than Col 2, Col 3 receives `↓`.
- Arrows are written only into the right/current column.
- Left/previous columns receive no arrow because of that comparison.
- No ordinary letter markers are used.
- No banner markers are written.
- No fill is applied by default.

---

## 7.2. Previous-column comparison with fill

### Input

| Label | Col 1 | Col 2 | Col 3 |
|---|---:|---:|---:|
| % | 50% | 60% | 55% |
| Base | 500 | 500 | 500 |

### Settings

- Compare with previous column: on
- Apply previous-column fill: on

### Expected

- Significantly higher current column receives `↑`.
- Significantly lower current column receives `↓`.
- `↑` cells receive normal significant fill.
- `↓` cells receive lower-than-Total fill color.
- Small-base fill, if present, still has higher priority.

---

## 7.3. Previous-column mode with first column as Total: Total treated as ordinary previous column

### Input

| Label | Total | Segment 1 | Segment 2 |
|---|---:|---:|---:|
| % | 60% | 70% | 65% |
| Base | 500 | 500 | 500 |

### Settings

- Compare with previous column: on
- First column is Total: on
- Exclude Total from comparisons: off

### Expected

- Warning is shown under “First column is Total”:
  - Total will be treated as an ordinary previous column unless “Exclude Total from comparisons” is enabled.
- Segment 1 is compared with Total as ordinary previous column.
- Segment 2 is compared with Segment 1.
- Output uses arrows, not `T` / `t`.
- No literal Total-comparison logic is used.
- No ordinary letter markers are used.

---

## 7.4. Previous-column mode with first column as Total and Total excluded

### Input

| Label | Total | Segment 1 | Segment 2 | Segment 3 |
|---|---:|---:|---:|---:|
| % | 60% | 70% | 65% | 80% |
| Base | 500 | 500 | 500 | 500 |

### Settings

- Compare with previous column: on
- First column is Total: on
- Exclude Total from comparisons: on

### Expected

- Total is excluded from the previous-column chain.
- Segment 1 receives no arrow.
- Segment 2 is compared with Segment 1.
- Segment 3 is compared with Segment 2.
- Output uses arrows only.
- No `T` / `t` markers appear.
- No ordinary letter markers appear.

---

## 7.5. Previous-column mode incompatible with compare only with Total

### Settings

- Compare with previous column: on
- Compare only with Total: attempt to turn on

### Expected

- UI prevents both checkboxes from being active together.
- “Compare only with Total” becomes disabled when previous-column mode is active.
- Defensive validation stops calculation if this invalid state somehow occurs.

---

## 7.6. Previous-column mode with invalid exclude Total state

### Settings

- Compare with previous column: on
- First column is Total: off
- Exclude Total from comparisons: attempt to turn on

### Expected

- UI prevents this state.
- “Exclude Total from comparisons” is disabled unless “First column is Total” is active.
- Defensive validation stops calculation if this invalid state somehow occurs.

---

## 7.7. Previous-column mode disables banner letters

### Settings

- Compare with previous column: on
- Write banner letters: attempt to turn on

### Expected

- “Write banner letters” is unchecked and disabled.
- No banner markers are written.
- Previous-column arrows are still written into data cells.

---

## 7.8. Previous-column mode disables fill only for Total comparisons

### Settings

- Compare with previous column: on
- Fill only for Total comparisons: attempt to turn on

### Expected

- “Fill only for Total comparisons” is unchecked and disabled.
- Previous-column fill behavior is controlled only by “Apply previous-column fill”.

---

# 8. Banner Tests

## 8.1. Banner markers without Total

### Input

Selected range starts below a banner row.

| Banner | Col A | Col B | Col C |
|---|---:|---:|---:|
| Label | A | B | C |
| % | 50% | 60% | 70% |
| Base | 500 | 500 | 500 |

### Settings

- Write banner letters: on
- First column is Total: off
- Compare with previous column: off

### Expected

- Banner cells receive:
  - `(a)`, `(b)`, `(c)`
- If a marker already exists at the end, it is replaced or preserved according to current marker update rules.
- Data cells receive ordinary significance markers as applicable.

---

## 8.2. Banner markers with first column as Total

### Input

| Banner | Total | Segment 1 | Segment 2 |
|---|---:|---:|---:|
| % | 60% | 70% | 50% |
| Base | 500 | 500 | 500 |

### Settings

- Write banner letters: on
- First column is Total: on

### Expected

- Total banner cell receives no marker.
- Segment 1 banner cell receives `(a)`.
- Segment 2 banner cell receives `(b)`.
- If Total banner cell previously had a marker, it is removed.

---

## 8.3. Banner first-row case

### Setup

Selected range starts in row 1 of the worksheet.

### Settings

- Write banner letters: on

### Expected

- Calculation stops.
- Status panel shows message asking user to add a row above the selected range.
- No automatic row insertion is attempted.

---

# 9. Formatting and Output Tests

## 9.1. Percent sign preservation

### Input

| Label | A | B |
|---|---:|---:|
| % | 42% | 55% |
| Base | 500 | 500 |

### Expected

When markers are appended:

- `42%` remains formatted as `42.0% a` or equivalent depending on rounding settings.
- `%` is not lost.
- Comma decimal input is handled safely.

---

## 9.2. Rounding off

### Settings

- Round cell values: off

### Expected

- Share-like rows:
  - proportions;
  - NPS;
  - promoters;
  - detractors;
  are displayed with 1 decimal place.
- Mean-like rows:
  - mean;
  - SD;
  - variance;
  are displayed with 2 decimal places.

---

## 9.3. Rounding on

### Settings

- Round cell values: on

### Expected

- Share-like rows are displayed with 0 decimal places.
- Mean-like rows are displayed with 1 decimal place.
- Statistical calculations use original normalized values, not rounded display values.

---

## 9.4. Old markers are removed

### Setup

Cells contain old markers:

- ordinary letters;
- `T`;
- `t`;
- `↑`;
- `↓`.

### Expected

- Old markers are removed before recalculation.
- Current run writes only current markers/arrows.

---

# 10. UI Tests

## 10.1. Default status panel

### Initial state

Open task pane.

### Expected

- Status panel is hidden.
- No default Microsoft placeholder text is visible.

---

## 10.2. Status after run

### Action

Click “Запустить”.

### Expected

- Status panel appears.
- Status text wraps within task pane width.
- Long messages do not overflow horizontally.

---

## 10.3. Status after clear

### Action

Click “Очистить значимости”.

### Expected

- Status panel appears.
- Message says significance markers were cleared.

---

## 10.4. Primary and secondary action hierarchy

### Expected

- “Запустить” is visually dominant.
- “Очистить значимости” is secondary and does not dominate the UI.

---

# 11. Label Detection Tests

## 11.1. Labels immediately left of selection

### Setup

Metric labels are located 1–2 columns to the left of selected data.

### Settings

- Labels on left side of sheet: off

### Expected

- Detector reads nearby left-side labels.
- Numeric columns between labels and data are skipped.
- Correct row types are detected.

---

## 11.2. Labels from far-left sheet columns

### Setup

User selects a right-side portion of a wide table.
Metric labels remain in the leftmost sheet columns.

### Settings

- Labels on left side of sheet: on

### Expected

- Detector reads labels from leftmost worksheet columns.
- Correct row types are detected.

---

# 12. Mixed Block Tests

## 12.1. Mixed table with shared base

### Input shape

| Label | A | B | C |
|---|---:|---:|---:|
| % | ... | ... | ... |
| % | ... | ... | ... |
| NPS | ... | ... | ... |
| Promoters | ... | ... | ... |
| Detractors | ... | ... | ... |
| Mean | ... | ... | ... |
| SD | ... | ... | ... |
| Base | ... | ... | ... |

### Expected

- Detector builds calculation blocks correctly.
- Shared base is used for appropriate blocks.
- After detecting Mean / NPS / spread structures, detector does not skip to Base and accidentally miss later blocks.
- Service rows do not receive markers. (Promoters and Detractors are not considered service rows here and can receive markers).
- Small-base fill applies only within the relevant calculation block.

---

# 13. Regression Checklist

Before release, manually verify:

- [ ] All-vs-all proportions still work.
- [ ] Means still use Welch’s t-test.
- [ ] NPS structure still recalculates NPS from promoters and detractors.
- [ ] NPS spread still works.
- [ ] Confidence selector affects all supported metric types.
- [ ] Old markers and old fills are cleared on rerun.
- [ ] Percent signs are preserved.
- [ ] First-column Total indexing works.
- [ ] Total column never receives markers.
- [ ] Banner skips Total column.
- [ ] `T` and `t` markers work.
- [ ] Lower-than-Total fill works.
- [ ] Fill-only-for-Total works.
- [ ] Exclude Total from comparisons works.
- [ ] Compare only with Total works.
- [ ] Small-base filtering runs before calculations.
- [ ] Small-base fill has highest priority.
- [ ] Small-base Total stops calculation.
- [ ] Previous-column mode writes arrows only into right/current columns.
- [ ] Previous-column mode does not write banner letters.
- [ ] Previous-column mode disables incompatible settings.
- [ ] Previous-column mode does not skip over small-base columns.
- [ ] Status panel is hidden by default and wraps long text.
- [ ] “Запустить” remains visually primary.


# Smoke test checklist

## 1. Basic proportions

Table:

| Segment 1 | Segment 2 |
|---|---|
| 40% | 60% |
| 100 | 100 |

Settings:

- `respect-banner-structure = false`
- `first-column-is-total = false`
- `compare-with-previous-column = false`

Expected:

- pairwise significance is calculated;
- significant higher cell receives marker of the lower column;
- normal significance fill is applied.

---

## 2. Confidence level selector

Use the same table with borderline values.

Settings:

- run with `confidence-level = 95`
- run with `confidence-level = 90`
- run with `confidence-level = 80`

Expected:

- lower confidence levels produce more significant differences;
- old markers/fills are cleared before recalculation.

---

## 3. One-tailed test

Use borderline values where two-tailed 95% is not significant but one-tailed 95% is significant.

Settings:

- `one-tailed-test = false`
- then `one-tailed-test = true`

Expected:

- one-tailed mode uses lower critical threshold;
- directions and markers remain based on actual value direction;
- one-tailed mode affects proportions, means, NPS, Total comparisons, and previous-column comparisons.

---

## 4. Rounding

Settings:

- `round-cell-values = false`
- then `round-cell-values = true`

Expected:

Default:

- proportions / NPS / promoters / detractors display with 1 decimal;
- means / SD / variance display with 2 decimals.

When enabled:

- proportions / NPS / promoters / detractors display as integers;
- means / SD / variance display with 1 decimal.

Important:

- statistics are calculated from normalized original values before display rounding.

---

# Metric detection

### 5. Mixed table

Rows:

| Label | Segment 1 | Segment 2 |
|---|---:|---:|
| % row 1 | 40% | 60% |
| % row 2 | 20% | 30% |
| Base | 100 | 100 |
| Mean | 3.5 | 4.1 |
| SD | 1.0 | 1.2 |
| Base | 120 | 130 |
| NPS | 10 | 20 |
| Promoters | 40% | 50% |
| Detractors | 30% | 30% |
| Base | 100 | 100 |

Expected:

- proportions block detected;
- mean block detected;
- NPS structure block detected;
- correct marker rows only (including Promoters and Detractors);
- service rows do not receive markers.

---

## 6. Shared base

Rows:

| Label | Segment 1 | Segment 2 |
|---|---:|---:|
| % row | 40% | 60% |
| NPS | 10 | 20 |
| Promoters | 40% | 50% |
| Detractors | 30% | 30% |
| Mean | 3.5 | 4.1 |
| SD | 1.0 | 1.2 |
| Base | 100 | 100 |

Expected:

- blocks use shared base correctly;
- detector does not skip later metrics after NPS/mean blocks.

---

# Label lookup

## 7. Labels immediately left

Selected range begins immediately after label column.

Expected:

- labels are read from left of selected range;
- numeric intermediate columns are skipped when searching for text labels.

---

## 8. Labels on left side of sheet

Settings:

- `labels-on-left-side = true`

Expected:

- labels are read from leftmost sheet columns;
- useful when selected range is far to the right.

---

# Total modes without banner structure

## 9. First column is Total

Table:

| Total | Segment 1 | Segment 2 |
|---:|---:|---:|
| 50% | 40% | 60% |
| 200 | 100 | 100 |

Settings:

- `first-column-is-total = true`
- `compare-only-with-total = false`
- `exclude-total-from-comparisons = false`

Expected:

- Segment 1 vs Total;
- Segment 2 vs Total;
- Segment 1 vs Segment 2;
- Total column receives no ordinary letters;
- `T` means segment is significantly higher than Total;
- `t` means segment is significantly lower than Total.

---

## 10. Compare only with Total

Settings:

- `first-column-is-total = true`
- `compare-only-with-total = true`

Expected:

- only Segment 1 vs Total and Segment 2 vs Total;
- no segment-vs-segment letters;
- only `T/t` markers.

---

## 11. Exclude Total

Settings:

- `first-column-is-total = true`
- `exclude-total-from-comparisons = true`

Expected:

- Total is excluded from all comparisons;
- only Segment 1 vs Segment 2;
- no `T/t`.

---

## 12. Fill only for Total

Settings:

- `first-column-is-total = true`
- `fill-only-total-comparisons = true`

Expected:

- green significance fill is applied only for cells significantly higher than Total;
- segment-vs-segment significance may still produce markers but not normal fill;
- lower-than-Total fill still applies where relevant.

---

# Small bases

## 13. Small base exclusion

Table:

| Segment 1 | Segment 2 | Segment 3 |
|---:|---:|---:|
| 40% | 60% | 70% |
| 100 | 20 | 100 |

Settings:

- `exclude-small-bases = true`
- `small-base-threshold = 50`

Expected:

- Segment 2 is excluded from calculations;
- Segment 2 receives small-base fill across the current calculation block;
- Segment 1 vs Segment 3 is still calculated;
- Segment 2 does not receive markers.

---

## 14. Small base in manual Total

Table:

| Total | Segment 1 | Segment 2 |
|---:|---:|---:|
| 50% | 40% | 60% |
| 20 | 100 | 100 |

Settings:

- `first-column-is-total = true`
- `exclude-small-bases = true`
- `small-base-threshold = 50`

Expected:

- calculation stops;
- status tells user to check Total base.

---

# Previous-column mode

## 15. Basic previous-column

Table:

| W1 | W2 | W3 |
|---:|---:|---:|
| 40% | 50% | 45% |
| 100 | 100 | 100 |

Settings:

- `compare-with-previous-column = true`

Expected:

- W2 compared with W1;
- W3 compared with W2;
- arrows written only into right/current column;
- `↑` if current is significantly higher;
- `↓` if current is significantly lower;
- ordinary letters are not used;
- banner letters are disabled.

---

## 16. Previous-column fill default

Steps:

1. Turn off previous-column mode.
2. Confirm `apply-previous-column-fill` is hidden/disabled/off.
3. Turn on previous-column mode.

Expected:

- `apply-previous-column-fill` becomes enabled by default;
- user can manually turn it off;
- if previous-column mode is turned off again, fill option resets.

---

## 17. Previous-column and small base

Table:

| W1 | W2 | W3 | W4 |
|---:|---:|---:|---:|
| 40% | 50% | 60% | 70% |
| 100 | 20 | 100 | 100 |

Settings:

- `compare-with-previous-column = true`
- `exclude-small-bases = true`
- `small-base-threshold = 50`

Expected:

- W2 vs W1 is skipped because W2 has small base;
- W3 vs W2 is skipped because W2 has small base;
- W4 vs W3 is calculated;
- no skipping over W2.

---

# Banner structure detection

## 18. One-level banner

Banner:

| Segment 1 | Segment 2 |
|---|---|

Data:

| Segment 1 | Segment 2 |
|---:|---:|
| 40% | 60% |
| 100 | 100 |

Settings:

- `respect-banner-structure = true`

Expected:

- fallback one-level/default group;
- calculation still works;
- no technical banner diagnostics shown in status.

---

## 19. Repeated-label two-level banner

Banner:

| Gender | Gender | Gender | Age | Age | Age |
|---|---|---|---|---|---|
| Total | Male | Female | Total | 18-24 | 25-34 |

Settings:

- `respect-banner-structure = true`

Expected:

- groups detected:
  - Gender: Total, Male, Female
  - Age: Total, 18-24, 25-34
- no cross-group comparisons;
- group-local marker indexing.

---

## 20. Merged-like reconstructed span

Upper banner row:

| Age |  |  |
|---|---|---|
| Total | 18-24 | 25-34 |

Settings:

- `respect-banner-structure = true`

Expected:

- reconstructed span detects Age group;
- Age group includes Total, 18-24, 25-34;
- no technical span diagnostics shown in status.

---

# Banner-aware comparisons

## 21. Group-aware ordinary comparisons

Banner:

| Gender | Gender | Gender | Age | Age | Age |
|---|---|---|---|---|---|
| Total | Male | Female | Total | 18-24 | 25-34 |

Settings:

- `respect-banner-structure = true`
- `compare-only-with-total = false`
- `exclude-total-from-comparisons = false`
- `compare-with-previous-column = false`

Expected:

- ordinary comparisons occur only within Gender or Age;
- no Gender-vs-Age comparisons;
- Total columns are not part of ordinary group comparisons;
- local Total comparisons are generated separately.

---

## 22. Group-local cell markers

Banner:

| Gender | Gender | Gender | Age | Age | Age |
|---|---|---|---|---|---|
| Total | Male | Female | Total | 18-24 | 25-34 |

Expected local labels:

- Gender:
  - Male = `a`
  - Female = `b`
- Age:
  - 18-24 = `a`
  - 25-34 = `b`

Expected:

- if 25-34 is higher than 18-24, 25-34 gets marker `a`, not global marker from the full selected range.

---

## 23. Local Total logic

Banner:

| Gender | Gender | Gender | Age | Age | Age |
|---|---|---|---|---|---|
| Total | Male | Female | Total | 18-24 | 25-34 |

Settings:

- `respect-banner-structure = true`
- no global Total

Expected:

- Gender Total is reference for Male/Female;
- Age Total is reference for 18-24/25-34;
- local Total columns do not receive ordinary letters;
- local Total columns are excluded from ordinary group comparisons;
- `T/t` markers are used for local Total comparisons.

---

## 24. Compare only with Total under banner structure

Same banner as above.

Settings:

- `respect-banner-structure = true`
- `compare-only-with-total = true`

Expected:

- only local Total comparisons;
- no ordinary group comparisons;
- no segment-vs-segment letters;
- only `T/t`.

---

## 25. Exclude Total under banner structure

Same banner as above.

Settings:

- `respect-banner-structure = true`
- `exclude-total-from-comparisons = true`

Expected:

- all detected Total columns excluded from all comparisons;
- no `T/t`;
- ordinary comparisons only among non-Total columns inside each group.

---

## 26. Group without Total

Banner:

| Gender | Gender |
|---|---|
| Male | Female |

Settings:

- `respect-banner-structure = true`

Expected:

- group is valid;
- ordinary group comparisons work;
- no Total comparisons are generated;
- not an error.

---

## 27. Compare only with Total but no Total found

Banner:

| Gender | Gender |
|---|---|
| Male | Female |

Settings:

- `respect-banner-structure = true`
- `compare-only-with-total = true`

Expected:

- no valid comparison pairs;
- status includes message that compare-only-with-Total was enabled but no Total was found.

---

## 28. Multiple local Totals

Banner:

| Gender | Gender | Gender | Gender |
|---|---|---|---|
| Total | Male | Total | Female |

Settings:

- `respect-banner-structure = true`

Expected:

- calculation stops;
- status explains that multiple Totals were found in one group.

---

# Global Total

## 29. Global Total with local Totals

Banner:

| Global Total | Gender | Gender | Gender | Age | Age | Age |
|---|---|---|---|---|---|---|
| Global Total | Total | Male | Female | Total | 18-24 | 25-34 |

Settings:

- `respect-banner-structure = true`

Expected:

- global Total detected;
- global Total is the only Total reference;
- local Totals are not used as group references;
- local Totals are compared with global Total as ordinary target columns in Total-comparison logic;
- local Totals may receive `T/t`;
- local Totals do not receive ordinary banner letters;
- status explains global Total behavior.

---

# Banner letters

## 30. Banner-local letters

Banner:

| Gender | Gender | Gender | Age | Age | Age |
|---|---|---|---|---|---|
| Total | Male | Female | Total | 18-24 | 25-34 |

Settings:

- `respect-banner-structure = true`
- `write-banner-letters = true`

Expected lower banner row:

| Total | Male (a) | Female (b) | Total | 18-24 (a) | 25-34 (b) |
|---|---|---|---|---|---|

Expected:

- upper banner row is unchanged;
- Total columns do not receive letters;
- old trailing markers are replaced;
- old trailing markers are removed from columns that no longer need letters.

---

## 31. Banner letters with global Total

Banner:

| Global Total | Gender | Gender | Gender | Age | Age | Age |
|---|---|---|---|---|---|---|
| Global Total | Total | Male | Female | Total | 18-24 | 25-34 |

Settings:

- `respect-banner-structure = true`
- `write-banner-letters = true`

Expected lower banner row:

| Global Total | Total | Male (a) | Female (b) | Total | 18-24 (a) | 25-34 (b) |
|---|---|---|---|---|---|---|

Expected:

- global Total receives no letter;
- local Totals receive no letters;
- segment labels are group-local.

---

# Banner-aware previous-column

## 32. Previous-column inside banner groups

Banner:

| Gender | Gender | Age | Age |
|---|---|---|---|
| Male | Female | 18-24 | 25-34 |

Settings:

- `respect-banner-structure = true`
- `compare-with-previous-column = true`

Expected:

- Female vs Male;
- 25-34 vs 18-24;
- no 18-24 vs Female cross-group comparison.

---

## 33. Previous-column ignores Total under banner structure

Banner:

| Gender | Gender | Gender | Age | Age | Age |
|---|---|---|---|---|---|
| Total | Male | Female | Total | 18-24 | 25-34 |

Settings:

- `respect-banner-structure = true`
- `compare-with-previous-column = true`

Expected:

- Female vs Male;
- 25-34 vs 18-24;
- no Male vs Gender Total;
- no 18-24 vs Age Total.

---

## 34. Previous-column with exclude Total

Same banner as above.

Settings:

- `respect-banner-structure = true`
- `compare-with-previous-column = true`
- `exclude-total-from-comparisons = true`

Expected:

- same as previous test;
- Total columns are excluded;
- no jumping over Total columns.

---

# Wave banner auto previous-column

## 35. Mixed wave and non-wave groups

Banner:

| Gender | Gender | Wave | Wave | Wave |
|---|---|---|---|---|
| Male | Female | W1 | W2 | W3 |

Settings:

- `respect-banner-structure = true`
- `compare-with-previous-column = false`
- `apply-previous-column-fill = false`
- `write-banner-letters = true`

Expected:

- Gender group uses ordinary group comparison;
- Wave group uses previous-column automatically;
- W2 vs W1;
- W3 vs W2;
- no W1 vs Female;
- no all-vs-all inside Wave group;
- Wave arrows are filled even though UI fill checkbox is off;
- UI previous-column checkbox remains off;
- status mentions auto previous-column for Wave;
- banner letters are written only for Gender group;
- Wave lower labels receive no letters.

---

## 36. Non-wave group does not auto-switch

Banner:

| Gender | Gender | Gender |
|---|---|---|
| Male | Female | Other |

Settings:

- `respect-banner-structure = true`
- `compare-with-previous-column = false`

Expected:

- ordinary group comparisons;
- no arrows;
- no auto previous-column status message.

---

## 37. Manual previous-column overrides mixed mode

Banner:

| Gender | Gender | Wave | Wave | Wave |
|---|---|---|---|---|
| Male | Female | W1 | W2 | W3 |

Settings:

- `respect-banner-structure = true`
- `compare-with-previous-column = true`

Expected:

- previous-column inside all groups;
- Female vs Male;
- W2 vs W1;
- W3 vs W2;
- auto previous-column message not required.

---

## 38. Compare only with Total suppresses wave auto previous-column

Banner:

| Wave | Wave | Wave |
|---|---|---|
| Total | W1 | W2 |

Settings:

- `respect-banner-structure = true`
- `compare-only-with-total = true`
- `compare-with-previous-column = false`

Expected:

- only Total comparisons;
- no auto previous-column;
- no auto previous-column status message.

---

# Settings persistence

## 39. Local save

Steps:

1. Select `settings-storage-mode = local`.
2. Change confidence level.
3. Enable one-tailed test.
4. Enable several checkboxes.
5. Change colors.
6. Restart Excel.

Expected:

- settings are restored;
- dependent UI states are refreshed;
- invalid combinations are normalized.

---

## 40. Reset settings

Steps:

1. Save settings locally.
2. Change several settings.
3. Press reset.

Expected:

- settings return to defaults;
- local storage is cleared;
- status says settings were reset;
- after Excel restart, defaults remain.

---

# Status messages

## 41. Clean status for simple one-level banner

Settings:

- `respect-banner-structure = true`

Expected:

- Расчёт выполнен. Обработано блоков: 1.
- No technical banner diagnostics.

## 42. Status for wave auto previous-column

Expected:

- Расчёт выполнен. Обработано блоков: 1.

- Баннер: для волновых групп автоматически применён режим “Сравнение с предыдущей колонкой”: Wave.

## 43. Status for global Total

Expected:

- success status;
- global Total explanatory message;
- no detailed banner structure dump.

## 44. Status for banner error

Example:

- multiple local Totals in one group.

Expected:

- calculation stops;
- status shows concise user-facing error;
- no technical diagnostics dump.

## Means + SD/Base
Setup: Table containing a Mean row and a Standard Deviation (SD) or Base row directly beneath it.
Expected behavior: The Mean row receives significance markers or spread indicators based on calculations. The SD and Base rows receive no significance markers.
Key regression risk: Markers accidentally appearing on the SD or Base rows instead of the Mean row.

## Means + variance/Base
Setup: Table containing a Mean row and a Variance or Base row directly beneath it.
Expected behavior: The Mean row receives significance markers or spread indicators based on calculations. The Variance and Base rows receive no significance markers.
Key regression risk: Markers accidentally appearing on the Variance or Base rows.

## NPS-first
Setup: Table with an NPS row followed immediately by Promoters, Detractors, and Base rows.
Expected behavior: The NPS row receives specific NPS significance spread markers. Promoters and Detractors receive ordinary proportion markers. The Base row receives no markers.
Key regression risk: The NPS row failing to receive markers, or Promoters/Detractors receiving the wrong marker type.

## extended NPS
Setup: Table with an NPS row, alongside Scale rows, buckets, support rows, and a Base row.
Expected behavior: Scale rows, buckets, and support rows receive ordinary proportion markers. The NPS row receives NPS significance markers. The Base row receives no markers.
Key regression risk: Ordinary proportion markers incorrectly applying to the NPS row instead of the specific NPS markers.

## Run → Clear significance
Setup: A table that has already been calculated (markers and fills exist). Select the range and click the clear significance button.
Expected behavior: All significance markers (letters, arrows) and formatting (bolding, fills) are completely removed from the selected cells.
Key regression risk: Leftover formatting or trailing marker artifacts (e.g. ` b`) remaining in cells.

## numeric output conventions: 28, 28%, 0.28
Setup: A table containing values typed as plain numbers (28), formatted as percentages (28%), and formatted as decimals (0.28). Select the table and Run significance.
Expected behavior: Unmarked cells perfectly preserve their original Excel numeric values and display formats. Marker-bearing cells maintain the visual numeric string convention before the appended letter (e.g., `28% b` instead of `0.28 b`).
Key regression risk: The add-in converting the entire selected range (including unmarked cells) to text, or breaking percentage formatting into plain decimals.

## selected range guardrail warning-only behavior
Setup: Manually select a range that inadvertently includes the table title, question text, or top header row (outside the numeric data area). Click Run.
Expected behavior: The calculation proceeds normally based on the selection. A warning is displayed in the taskpane status output indicating that non-data rows may have been included, but the Run is not blocked or aborted. The selection is not automatically trimmed.
Key regression risk: The guardrail warning blocking the calculation, or the add-in attempting to automatically resize the user's manual selection.
