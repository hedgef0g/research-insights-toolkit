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

## 3.1. NPS from structure

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
- Markers are written only into the NPS row.
- Promoters, Detractors, and Base rows receive no markers.

---

## 3.2. NPS from spread

### Input

| Label | A | B | C |
|---|---:|---:|---:|
| NPS | 30 | 45 | 25 |
| SD | 0.8 | 0.7 | 0.8 |
| Base | 500 | 500 | 500 |

### Expected

- NPS spread logic is used.
- Markers are written only into the NPS row.
- SD and Base rows receive no markers.

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
- Service rows do not receive markers.
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