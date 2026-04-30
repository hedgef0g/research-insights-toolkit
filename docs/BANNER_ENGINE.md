# Research Insights Toolkit — Banner Engine Specification

Last updated: 2026-04-30

## Purpose

The banner engine detects column banner structure above the selected data range.

It is designed as a reusable core module that can be used by multiple macros, not only by significance testing.

The banner engine must not depend on Excel APIs directly. It should work from plain data structures prepared by the Excel adapter.

## Architectural role

Recommended module:

`core/banner-detector.js`

The banner engine should not:

* read Excel ranges directly;
* write to Excel;
* access Office.js objects;
* access DOM/UI;
* calculate statistical significance.

The banner engine should:

* analyze banner rows above the selected range;
* detect column labels;
* detect banner groups;
* detect global and local Total columns;
* detect ambiguous or malformed structures;
* return structured metadata and status messages.

The Excel-specific layer is responsible for:

* reading banner values/text/merge metadata from Excel;
* passing them to the banner engine;
* using returned structure to build comparison pairs;
* writing any output back to Excel.

---

## 1. Scope

### MVP scope

The first product-ready version should support:

* one-level banners;
* two-level banners;
* horizontally merged group headers;
* vertically merged global Total in the leftmost selected column;
* repeated adjacent group labels as fallback if merge metadata is unavailable or absent;
* partial selection inside banner groups, where possible;
* local Total inside each group;
* global Total overriding local Totals as reference;
* accumulated banner status messages.

### Future scope

The internal model should be able to represent:

* three-level or deeper banners;
* nested group paths;
* floating group levels at different heights;
* more complex banner structures.
* auto-detection of wave/time-series banners that can suggest or activate previous-column comparison mode.

However, MVP comparison logic will use the last meaningful group level.

---

## 2. Key terms

### Selected range

The numeric/data area selected by the user.

The selected range does not include row labels and does not include banner rows.

### Lower banner level

The row immediately above the selected data range.

This is the primary column label row.

Example:

| Gender                    | Gender | Gender | Age   | Age   | Age   |
| ------------------------- | ------ | ------ | ----- | ----- | ----- |
| Total                     | Male   | Female | Total | 18-24 | 25-34 |
| selected data starts here |        |        |       |       |       |

In this example, the row `Total | Male | Female | Total | 18-24 | 25-34` is the lower banner level.

### Meaningful group level

The nearest valid group label level above the lower banner level.

Usually this is a horizontally merged row or a repeated adjacent label row.

Example:

| Gender | Gender | Gender | Age   | Age   | Age   |
| ------ | ------ | ------ | ----- | ----- | ----- |
| Total  | Male   | Female | Total | 18-24 | 25-34 |

In this example, `Gender | Gender | Gender | Age | Age | Age` is the meaningful group level.

### Banner path

The full detected label path for a column.

Example:

`Region A / Gender / Male`

MVP may store the full path but use only the last meaningful group level for comparison grouping.

### Comparison group

The group within which ordinary segment-vs-segment comparisons are allowed when `respectBannerStructure` is enabled.

Example:

* `Gender`: Total, Male, Female
* `Age`: Total, 18-24, 25-34

### Global Total

A Total column that applies to the full selected table, not only to one banner group.

Example:

| Global Total | Gender Total | Male | Female | Age Total | 18-24 | 25-34 |
| ------------ | ------------ | ---- | ------ | --------- | ----- | ----- |

A vertically merged cell in the first selected column above the data may indicate a global Total.

### Local Total

A Total column inside a specific banner group.

Example:

| Gender | Gender | Gender |
| ------ | ------ | ------ |
| Total  | Male   | Female |

---

## 3. Inputs

The banner engine should receive normalized platform-independent input.

Recommended input shape:

`detectBannerStructure({ selectedColumnCount, lowerBannerRows, upperScanRows, mergeMetadata, settings })`

### selectedColumnCount

Number of columns in the selected data range.

### lowerBannerRows

At minimum, the row immediately above the selected range.

In MVP this will usually be one row, but the input can support multiple rows if the Excel adapter has already read them.

### upperScanRows

Rows above the lower banner level that may contain meaningful group labels.

These rows should be passed in nearest-to-farthest order or with explicit sheet row indexes.

### mergeMetadata

Platform-independent metadata describing merged cells.

Recommended conceptual shape:

* `rowOffset`
* `columnIndex`
* `value`
* `text`
* `isMerged`
* `mergeArea.topRowOffset`
* `mergeArea.leftColumnIndex`
* `mergeArea.rowCount`
* `mergeArea.columnCount`

The exact shape may evolve, but the detector needs enough information to answer:

* does this cell belong to a merged area?
* is the merged area horizontal?
* is the merged area vertical?
* where does the merged area start and end?
* what is the top-left value/text of the merge area?

### settings

Relevant settings:

* `respectBannerStructure`
* `compareOnlyWithTotal`
* `excludeTotalFromComparisons`
* `compareWithPreviousColumn`

Manual Total placement settings should be ignored/disabled when `respectBannerStructure` is active.

---

## 4. Output

Recommended output shape:

### Top-level fields

* `isDetected`
* `mode`
* `columnDescriptors`
* `groups`
* `globalTotalColumnIndex`
* `totalColumnIndexes`
* `messages`

### Example

```js
{
  isDetected: true,
  mode: "twoLevel",

  columnDescriptors: [
    {
      columnIndex: 0,
      lowerLabel: "Total",
      normalizedLowerLabel: "total",
      bannerPath: ["Gender", "Total"],
      displayLabel: "Gender / Total",

      comparisonGroupKey: "group:gender:0",
      comparisonGroupLabel: "Gender",

      isTotal: true,
      totalType: "local",

      isGlobalTotal: false,
      isLocalTotal: true,

      source: {
        lowerLevelRowOffset: 0,
        groupLevelRowOffset: -1,
        mergeArea: null
      }
    }
  ],

  groups: [
    {
      groupKey: "group:gender:0",
      label: "Gender",
      bannerPath: ["Gender"],
      columnIndexes: [0, 1, 2],
      localTotalColumnIndexes: [0],
      hasLocalTotal: true
    }
  ],

  globalTotalColumnIndex: null,

  totalColumnIndexes: [0, 3],

  messages: [
    {
      severity: "info",
      code: "GLOBAL_TOTAL_USED",
      text: "..."
    }
  ]
}
```

### Notes

* `messages` must be accumulated and shown in the final Status panel after calculation.
* `error` severity can be used to stop calculation.
* `warning` and `info` can allow calculation to continue.

---

## 5. Detection strategy

### 5.1. Lower banner level

The lower banner level is the row immediately above the selected data range.

Column labels are read from this row.

Rules:

* A lower-level banner cell may contain text or numbers.
* Numeric labels are valid banner labels.
* The condition is “non-empty cell”, not “text-like cell”.
* Empty lower labels are allowed but may reduce detection confidence.

### 5.2. Vertical merge above top-left selected cell

If the cell directly above the top-left selected data cell is vertically merged, the detector should enter two-level/global-total-aware detection.

A vertical merge in the first selected column may indicate a global Total.

Rules:

* vertical merge in first selected column can define or contribute to global Total detection;
* vertical merges outside the first selected column are not used as primary grouping structure in MVP;
* suspicious vertical merges elsewhere should produce a warning if they make structure ambiguous.

### 5.3. Search for meaningful group level

If the cell directly above the top-left selected cell is not vertically merged:

1. Start from the row above the lower banner level.
2. Check whether cells on that row define a meaningful group level.
3. A meaningful group level can be:

   * horizontal merged cells;
   * repeated adjacent non-empty labels as fallback.
4. If the row has non-empty values but no horizontal merge/repeated group structure, continue scanning upward.
5. Stop when a valid meaningful group level is found or scan limit is reached.
6. Intermediate rows may be ignored.

This search should be performed per group/span, not only once for the full table, to support floating second levels.

### 5.4. Floating second level

The meaningful group level may be at different row offsets for different column spans.

The detector should build descriptors per column:

* lower label = cell immediately above data;
* group label = nearest valid meaningful group level above this column.

This allows different groups to resolve their group labels at different heights.

### 5.5. Partial selection inside merged group

If the selected range starts inside a horizontally merged group, the detector should still identify the group label if merge metadata is available.

Example:

* full banner: `Gender` merged over columns E:H;
* user selection: columns G:H only.

The detector should use merge area metadata to retrieve the group label from the merge area's top-left cell.

If this cannot be resolved:

* group may be marked unknown;
* if fewer than two selected columns belong to the same detected group, no comparisons are generated for that group;
* status may include a warning if this affects comparisons.

### 5.6. Repeated adjacent labels fallback

If merge metadata does not provide horizontal merged groups, repeated adjacent labels can form a group.

Example:

| Gender | Gender | Gender | Age   | Age   | Age   |
| ------ | ------ | ------ | ----- | ----- | ----- |
| Total  | Male   | Female | Total | 18-24 | 25-34 |

Fallback grouping:

* `Gender`: Total, Male, Female
* `Age`: Total, 18-24, 25-34

Rules:

* repeated adjacent labels form one group span;
* group identity must include position/span, not only text;
* numeric repeated labels are valid.

### 5.7. Scan limit

To avoid accidentally treating report titles as banner levels, the detector should use a maximum scan limit.

Recommended MVP value:

`maxBannerScanRows = 5`

This can later become an advanced setting if needed.

### 5.8. Report title detection

A row above the table can be a report title, not a banner group.

Heuristic:

1. Determine the left and right boundaries of the broader table/banner row.
2. Scan left from the selected range until a blank boundary.
3. Scan right from the selected range until a blank boundary.
4. If the far-left and far-right non-empty cells on the candidate level contain the same value, treat the row as a report title or non-meaningful level.
5. Do not use this level as meaningful group level.
6. Stop at previous valid meaningful level or fallback to one-level mode.

This is a heuristic and may evolve.

---

## 6. Total detection

### 6.1. Total labels

Total detection should use a dedicated Total dictionary, not the metric row dictionary.

Initial keywords:

* `total`
* `итого`
* `всего`
* `all`
* `overall`

Do not treat `Base` as Total.

### 6.2. Manual Total placement settings

When `respectBannerStructure` is enabled:

* `first-column-is-total` is disabled;
* `total-in-each-banner` is disabled;
* Total placement is determined only by the banner engine.

### 6.3. Global Total

A global Total can be detected when:

* a vertically merged Total-like cell appears in the first selected column above the data; or
* a clearly identified global Total column exists outside local banner groups.

MVP should prioritize the vertical-merge-in-first-column scenario.

### 6.4. Local Total

A local Total is a Total-like lower-level label inside a comparison group.

Example:

| Gender | Gender | Gender |
| ------ | ------ | ------ |
| Total  | Male   | Female |

### 6.5. Multiple local Totals in one group

If no global Total exists and a group contains more than one local Total, calculation must stop with a status error.

Example message:

`В группе “Gender” найдено несколько Тоталов. Расчёт остановлен: невозможно однозначно определить колонку для сравнения.`

If a global Total exists, local Totals are not used as references. Multiple local Totals may be reported as a warning, but they do not block calculation because global Total is the only reference.

### 6.6. Total outside selection

If banner structure implies that a Total exists outside the selected range, it cannot be used for calculation.

If Total-based comparison requires that Total, status must tell the user to include the relevant Total column in the selection.

Example message:

`Тотал для группы “Gender” находится вне выделенного диапазона или не найден. Для сравнения с Тоталом выделите колонку Тотал вместе с группой.`

---

## 7. Total comparison rules with banner structure

When `respectBannerStructure` is enabled, Total placement is detected by the banner engine.

### 7.1. Exclude Total from comparisons

If `excludeTotalFromComparisons` is enabled:

* all detected Total columns are excluded from all comparisons;
* this applies to both global Total and local/group Totals;
* Total columns are not used as references;
* Total columns are not compared as ordinary columns;
* remaining non-Total columns are compared within their detected groups.

Example:

| Global Total | Gender Total | Male | Female | Age Total | 18-24 | 25-34 |
| ------------ | ------------ | ---- | ------ | --------- | ----- | ----- |

Comparisons:

* Male vs Female
* 18-24 vs 25-34

### 7.2. Global Total exists

If a global Total is detected and Total comparisons are not excluded:

* global Total is the only Total reference;
* local Totals are not used as group references;
* local Totals may still be compared with global Total as ordinary columns;
* status must explicitly say that global Total is used as the only Total reference and local Totals are treated as ordinary columns.

Example:

| Global Total | Gender Total | Male | Female | Age Total | 18-24 | 25-34 |
| ------------ | ------------ | ---- | ------ | --------- | ----- | ----- |

Total comparisons:

* Gender Total vs Global Total
* Male vs Global Total
* Female vs Global Total
* Age Total vs Global Total
* 18-24 vs Global Total
* 25-34 vs Global Total

Do not perform:

* Male vs Gender Total
* Female vs Gender Total
* 18-24 vs Age Total
* 25-34 vs Age Total

Status message:

`Найден глобальный Тотал. Сравнение с Тоталом выполняется только относительно глобального Тотала. Локальные Тоталы обрабатываются как обычные колонки и также могут сравниваться с глобальным Тоталом.`

### 7.3. No global Total, local Totals exist

If no global Total is detected:

* local/group Totals are used as references within their own banner groups;
* groups without local Total simply do not produce Total-comparison pairs;
* groups without local Total are not considered an error.

Example:

* `Gender`: Total, Male, Female
* `Age`: Total, 18-24, 25-34

Total comparisons:

* Male vs Gender Total
* Female vs Gender Total
* 18-24 vs Age Total
* 25-34 vs Age Total

### 7.4. Compare only with Total

If `compareOnlyWithTotal` is enabled:

* only Total-comparison pairs are created;
* if global Total exists, all non-global-Total columns are compared with global Total, including local Totals;
* if no global Total exists, columns are compared with their local group Total where available;
* groups without local Total produce no pairs;
* if no valid Total-comparison pairs exist at all, status must include a message explaining that no Total was found for the selected banner structure.

Example message:

`Режим “Сравнивать только с Тотал” включён, но в выделенном баннере не найден глобальный или локальный Тотал. Сравнения не выполнены.`

### 7.5. Ordinary group comparisons

If `compareOnlyWithTotal` is not enabled:

* ordinary segment-vs-segment comparisons are performed within each detected comparison group;
* columns from different groups are not compared;
* Total comparisons may additionally be performed according to the rules above;
* if `excludeTotalFromComparisons` is enabled, Total columns are removed from ordinary group comparisons too.

---

## 8. Previous-column mode with banner structure

`compareWithPreviousColumn` is a separate comparison mode.

If `compareWithPreviousColumn` is enabled:

* no ordinary letter markers are used;
* no banner letters are written;
* arrows are written into the right/current column only;
* small-base filtering still applies;
* excluded columns are not skipped over.

When `respectBannerStructure` is also enabled, previous-column comparison should respect detected groups.

Recommended rule:

* compare a column only with the previous selected column if both columns belong to the same comparison group;
* if previous selected column belongs to another group, do not create a pair;
* if either column is excluded by small base or Total exclusion, do not create a pair.

Example:

| Gender | Gender | Age   | Age   |
| ------ | ------ | ----- | ----- |
| Male   | Female | 18-24 | 25-34 |

Pairs:

* Female vs Male
* 25-34 vs 18-24

No pair:

* 18-24 vs Female

because that crosses group boundary.

---

## 9. Banner letters

When writing banner letters:

* letters are written into the lowest banner level;
* the lowest banner level is the row immediately above the selected range;
* upper banner/group levels are not modified;
* Total columns do not receive banner markers if they are excluded from indexing;
* global Total cells or vertical merged cells should not be modified.

Example:

| Gender | Gender | Gender |
| ------ | ------ | ------ |
| Total  | Male   | Female |

Markers should be written to:

| Gender | Gender   | Gender     |
| ------ | -------- | ---------- |
| Total  | Male (a) | Female (b) |

not to the upper `Gender` row.

---

## 10. Error and status handling

The banner engine returns accumulated messages.

The final Status panel should include all relevant banner messages after calculation.

Message shape:

```js
{
  severity: "info",
  code: "GLOBAL_TOTAL_USED",
  text: "Найден глобальный Тотал..."
}
```

### Error messages stop calculation

Examples:

* malformed/inconsistent merged structure;
* multiple local Totals in one group when no global Total exists;
* required Total is outside selection in Total-only mode.

### Info/warning messages do not necessarily stop calculation

Examples:

* global Total found and used as only reference;
* local Totals ignored because global Total exists;
* group has no local Total and therefore no Total-comparison pairs;
* partial merged group could not be fully resolved.

---

## 11. Integration with comparison pair builder

Currently comparison pair logic is driven by:

`buildColumnComparisonPairs(columnCount, calculationSettings, excludedColumnIndexes)`

With banner structure, this should evolve to accept optional banner metadata:

`buildColumnComparisonPairs(columnCount, calculationSettings, excludedColumnIndexes, bannerStructure)`

If `bannerStructure` is absent, current behavior remains unchanged.

If `bannerStructure` is present and `respectBannerStructure` is enabled:

* ordinary pairs are generated only within groups;
* Total pairs are generated according to global/local Total rules;
* previous-column pairs do not cross group boundaries;
* excluded columns are ignored;
* Total columns are excluded when `excludeTotalFromComparisons` is enabled.

---

## 12. Integration stages

### Stage 1 — Detection only

Implement `core/banner-detector.js`.

Return `bannerStructure` and messages.

Do not affect calculations yet.

Use status output or development diagnostics to verify detected groups.

### Stage 2 — Banner letters

Use detected lowest banner level for writing banner letters safely.

Do not modify upper banner rows.

Do not write markers into Total/global Total cells.

### Stage 3 — Comparison pairs

Pass `bannerStructure` into comparison pair builder.

Enable group-aware comparisons and banner-aware Total logic.

### Stage 4 — Richer multi-level support

Keep `bannerPath` internally.

Expand support for 3+ level structures later if needed.

---

## 13. Manual examples

### 13.1. One-level banner

| Total | Male | Female | 18-24 | 25-34 |
| ----- | ---- | ------ | ----- | ----- |

Expected:

* one fallback/default group if no group level is found;
* Total detected if label matches Total dictionary;
* if no structure grouping is found, group-aware mode may treat all columns as one group.

### 13.2. Two-level banner with local Totals

| Gender | Gender | Gender | Age   | Age   | Age   |
| ------ | ------ | ------ | ----- | ----- | ----- |
| Total  | Male   | Female | Total | 18-24 | 25-34 |

Expected groups:

* `Gender`: Total, Male, Female
* `Age`: Total, 18-24, 25-34

### 13.3. Two-level banner with global Total

| Global Total | Gender | Gender | Gender | Age   | Age   | Age   |
| ------------ | ------ | ------ | ------ | ----- | ----- | ----- |
| Global Total | Total  | Male   | Female | Total | 18-24 | 25-34 |

Expected:

* first column detected as global Total if vertical merge/structure confirms it;
* local Totals are ordinary columns for global Total comparisons;
* status explains global Total behavior.

### 13.4. Repeated labels without merges

| Gender | Gender | Gender | Age   | Age   | Age   |
| ------ | ------ | ------ | ----- | ----- | ----- |
| Total  | Male   | Female | Total | 18-24 | 25-34 |

Expected:

* repeated adjacent labels form group spans;
* no merge required.

### 13.5. Partial group selection

Full table:

| Gender | Gender | Gender | Age   | Age   | Age   |
| ------ | ------ | ------ | ----- | ----- | ----- |
| Total  | Male   | Female | Total | 18-24 | 25-34 |

Selected:

| Male | Female |
| ---- | ------ |

Expected:

* if both columns resolve to Gender group, compare Male vs Female;
* local Total outside selection is not used;
* if Total-only mode is active, status asks user to include Total in selection.

### 13.6. Group without Total

| Gender | Gender |
| ------ | ------ |
| Male   | Female |

Expected:

* group is valid;
* ordinary group comparisons work;
* Total comparisons are not generated for this group.

### 13.7. Multiple local Totals without global Total

| Gender | Gender | Gender | Gender |
| ------ | ------ | ------ | ------ |
| Total  | Male   | Total  | Female |

Expected:

* calculation stops with error;
* status explains that multiple Totals were found in one group.

### 13.8. Previous-column within groups

| Gender | Gender | Age   | Age   |
| ------ | ------ | ----- | ----- |
| Male   | Female | 18-24 | 25-34 |

Expected previous-column pairs:

* Female vs Male
* 25-34 vs 18-24

No cross-group pair:

* 18-24 vs Female

---

## 14. Open technical questions

### 14.1. Reading merge metadata in Office.js

Need to verify how reliably Office.js exposes merged cell areas for arbitrary cells above the selected range.

Particularly important cases:

* selected range starts inside a merged group;
* value is stored only in the top-left cell of merge area;
* merge area extends outside selected columns.

### 14.2. Detecting global Total

Need to test several real table layouts before finalizing global Total detection.

Known supported MVP signal:

* vertical merge in first selected column above selected data with Total-like label.

### 14.3. Report title heuristic

The left/right boundary heuristic needs real-world testing.

It should prevent treating report titles as banner group levels.

### 14.4. Malformed merge detection

Need to define exact criteria for inconsistent merged structures.

For MVP, ambiguous structure can stop calculation with a clear warning.

## Implementation status

MVP banner engine functionality has been implemented for the Excel add-in.

Implemented:

- one-level banner detection;
- two-level banner detection;
- repeated adjacent group labels;
- reconstructed span detection for merged-like headers;
- local Total detection;
- global Total detection;
- group-aware ordinary comparisons;
- group-local cell markers;
- banner-aware Total comparisons;
- banner-aware previous-column comparisons;
- automatic previous-column mode for wave groups;
- banner-aware lower-level letter writing;
- clean user-facing status messages.

The banner engine remains platform-independent at the core level.

Excel-specific responsibilities remain in `taskpane.js`:

- reading rows above the selected range;
- passing plain banner context into `core/banner-detector.js`;
- writing banner letters back into the lowest banner level.

## Implemented comparison behavior

When `respectBannerStructure` is enabled:

- columns from different banner groups are not compared;
- ordinary group comparisons exclude all Total columns;
- local Total is used as group reference only when no global Total exists;
- global Total, when detected, becomes the only Total reference;
- local Totals are compared with global Total when global Total exists;
- Total columns never receive ordinary letter labels;
  - This applies to:
  - global Total;
  - local/group Totals.
- Local Totals may receive `T/t` markers when compared with global Total, but they never receive ordinary group-local letter labels.
- previous-column comparisons do not cross group boundaries;
- wave groups automatically use previous-column comparison when global previous-column mode is off.

## Implemented wave behavior

Wave groups are detected by group label keywords:

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

- auto previous-column is applied only inside that group;
- non-wave groups keep ordinary group comparisons;
- the UI checkbox is not toggled;
- previous-column fill is applied automatically;
- banner letters are not written for the wave group;
- a user-facing status message is shown.

Plain numeric lower labels such as `1, 2, 3` are not enough to classify a group as wave-like.

## User-facing status

Normal successful calculations should not show technical banner diagnostics.

Visible banner messages are limited to user-relevant cases, including:

- global Total used;
- auto previous-column applied for wave groups;
- compare-only-with-Total produced no valid pairs;
- multiple local Totals;
- malformed banner structure;
- missing banner rows above selection.

Developer diagnostics may remain available through helper functions but should not appear in normal status output.