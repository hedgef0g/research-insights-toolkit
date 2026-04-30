# Research Insights Toolkit — Current Status

Last updated: 2026-04-30

## Implemented

- Unified auto-detection of metric blocks.
- Proportion significance using pooled z-test.
- Mean significance using Welch’s t-test.
- NPS significance from structure and spread.
- Confidence level selector: 99%, 95%, 90%, 80%, 66.6%.
- Marker generation with Latin and Cyrillic letters, excluding t/T/т/Т.
- Re-run cleanup of old markers and formatting.
- Display rounding before marker insertion.
- Preservation of % sign when appending markers.
- Left-side label detection.
- Optional label detection from leftmost worksheet columns.
- Banner marker insertion.
- First-column Total mode:
  - Total has no letter index.
  - Segment indexing starts from the second selected column.
  - Total does not receive banner markers.
  - Total does not receive cell markers.
- Total comparison markers:
  - `T` = segment significantly higher than Total.
  - `t` = segment significantly lower than Total.
- Compare only with Total.
- Exclude Total from comparisons.
- Lower-than-Total fill color.
- Fill only for Total comparisons.
- `cellResultMatrix` replacing plain marker matrix.
- Fill priority system:
  - small base
  - lower than Total
  - normal significance
  - none
- Small-base handling:
  - checked before significance calculations;
  - small-base columns excluded from comparisons;
  - small-base fill applied within the calculation block, including Base row;
  - Total small base stops calculation with warning.

- Total comparison improvements:
  - “Compare only with Total” mode.
  - “Exclude Total from comparisons” mode.
  - First selected column can be treated as Total.
  - Total column does not receive cell markers.
  - Total column does not receive banner markers.
  - Segment column indexing starts after Total.
  - `T` marker = segment is significantly higher than Total.
  - `t` marker = segment is significantly lower than Total.
  - Lower-than-Total fill color is supported.
  - Optional “Fill only for Total comparisons” mode:
    - normal significant fill is applied only to cells significantly higher than Total;
    - segment-vs-segment markers remain visible but do not receive normal significant fill.

- Cell result matrix:
  - Replaced plain marker matrix with `cellResultMatrix`.
  - Cell results can now store:
    - markers;
    - fill reason;
    - fill priority;
    - positive Total comparison metadata;
    - previous-column arrow metadata.
  - Fill priority system:
    1. small base;
    2. lower than Total;
    3. normal significance;
    4. none.

- Small base handling:
  - Small bases are checked before statistical calculations.
  - Columns with base below threshold are excluded from comparisons.
  - Small-base fill is applied within the relevant calculation block.
  - Small-base fill includes the Base row itself.
  - Small-base columns do not receive significance markers.
  - If first column is Total and Total has a small base, calculation stops with a user-facing warning.

- Previous-column comparison mode:
  - Added “Compare with previous column” mode.
  - Compares column 2 with column 1, column 3 with column 2, and so on.
  - Writes arrows into the right/current column only.
  - `↑` = current column is significantly higher than previous column.
  - `↓` = current column is significantly lower than previous column.
  - Does not use ordinary letter markers.
  - Does not write banner letters.
  - Optional previous-column fill:
    - higher values use normal significant fill;
    - lower values use lower-than-Total fill color.
  - Small-base filtering remains active and does not skip over excluded columns.

- UI improvements:
  - Removed default Microsoft/diagnostic text from the task pane.
  - Removed visible “Detect Metric Type” button from the user UI.
  - Renamed “Clear Significance” to “Очистить значимости”.
  - Added hidden-by-default status panel.
  - Status panel appears after running calculation or clearing significance.
  - Status text wraps within the task pane.
  - Primary and secondary actions are visually separated:
    - “Запустить” is the primary button;
    - “Очистить значимости” is a secondary text-style action.
  - Added UI mutual exclusions for incompatible modes.
  - Added warning when previous-column mode treats Total as an ordinary previous column.

- Output formatting improvements:
  - Percent sign is preserved when appending markers.
  - Old arrow markers are cleaned before recalculation.

## Not implemented yet

- Banner structure detection.
- Total in each banner.
- Custom modal/warning panel instead of plain output messages.
- Settings persistence.
- Writer optimization for large tables.
- Help page.
- Google Sheets support.

## Current architectural rule

- `taskpane.js` = UI orchestration only.
- `metric-detector.js` = row detection and calculation block planning.
- `significance.js` = statistics, comparison planning, marker/fill result matrix.
- `excel-writer.js` = Excel output only.
- `normalizers.js` = raw value cleanup.
- `dictionary.config.js` = configurable label dictionary.

No statistical calculations in `taskpane.js`.
No Excel range writing in `significance.js`.
No UI DOM access in core modules.