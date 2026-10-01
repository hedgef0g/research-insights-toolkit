# Table Preview Model

`src/core/table-preview-model.js`

## Purpose

Produces a read-only snapshot of how RIT interprets a selected Excel table.  
Intended for the future **"Проверить таблицу" / "Check table"** feature.

The model tells the user what row types were detected, what calculation blocks were built, and whether any data quality issues are present — without running significance calculations and without touching Excel.

## What it does not do

- Does not write to Excel
- Does not calculate statistical significance
- Does not call `significance.js` or `excel-writer.js`
- Does not scan worksheets or workbooks automatically
- Does not implement weighted base logic
- Does not implement multi-column row label *calculation* (only represents the data for future rendering)
- Has no Office.js dependency — safe to unit-test in Node.js

## Usage

```js
import { buildTablePreviewModel } from "./core/table-preview-model";

const model = buildTablePreviewModel({
  values,          // 2D array [row][col] of cleaned cell values
  leftLabelValues, // 2D array of left-side label cells (LABEL_SCAN_COLUMNS_LEFT columns)
  numberFormats,   // optional — 2D array of Excel number format strings
  bannerContext,   // optional — same shape passed to detectBannerStructure
  settings,        // optional — same settings object used in the main run flow
});
```

`values` should be cleaned (significance markers removed) before passing in, matching the input to `detectMetricRowsFromLeftLabels` in the main run flow.

## Output shape

```js
{
  rowDiagnostics: RowDiagnostic[],
  calculationBlocks: PreviewBlock[],
  bannerStructure: BannerPreview | null,
  dataQualityIssues: DataQualityIssue[],
  qualitySummary: QualitySummary,
  summary: Summary,
  warnings: Warning[],       // flat convenience alias of dataQualityIssues
}
```

### RowDiagnostic

```js
{
  rowIndex: number,
  label: string,             // raw label text (best candidate from left cells)
  labelParts: string[],      // all non-empty label parts from left columns (left-to-right)
  combinedLabel: string,     // joined representation (e.g. "Gender / Male")
  primaryLabel: string,      // rightmost part — closest to data
  secondaryLabel: string | null,  // next part outward (category header)
  normalizedLabel: string,
  normalizedLabelParts: string[],
  rowType: string,           // "proportion" | "base" | "mean" | "nps" | "promoters" |
                             // "detractors" | "standardDeviation" | "variance" |
                             // "empty" | "unknownText"
  rowSubtype: null,          // reserved — future: "weighted" | "unweighted" | "effective" for base rows
  confidence: "high" | "low",
  notes: string[],
}
```

`confidence` is `"high"` for all explicitly matched types and `"low"` for `empty` / `unknownText`.

### PreviewBlock

All fields from every metricType are present; fields that do not apply are `null`.

```js
{
  metricType: "proportion" | "mean" | "npsStructure" | "npsSpread",
  valueRowIndexes: number[] | null,  // array of value row indexes (proportion / convenience copy)
  valueRowIndex: number | null,      // single value row (mean / nps)
  baseRowIndex: number | null,
  promotersRowIndex: number | null,  // npsStructure
  detractorsRowIndex: number | null, // npsStructure
  neutralRowIndex: number | null,    // npsStructure format 2 (NPS / Promoters / Neutral / Detractors)
  sdRowIndex: number | null,         // mean / npsSpread when spreadType === "standardDeviation"
  varianceRowIndex: number | null,   // mean / npsSpread when spreadType === "variance"
  notes: string[],
}
```

`neutralRowIndex` is inferred from the associated proportion block when a single unaccounted-for row exists between Promoters and Detractors in the same base group.

### BannerPreview

```js
{
  isEnabled: boolean,                // true when settings.respectBannerStructure is on
  isDetected: boolean,
  mode: string | null,               // "oneLevel" | "twoLevel" | "fallback"
  groups: Group[],
  totalColumnIndexes: number[],
  globalTotalColumnIndex: number | null,
  hasWaveGroups: boolean,
  recommendedComparisonMode: string | null,
  messages: BannerMessage[],
}
```

### DataQualityIssue

```js
{
  code: string,
  severity: "critical" | "warning" | "info",
  message: string,
  rowIndex: number | null,
  columnIndex: number | null,
  relatedRowIndexes: number[],
  relatedColumnIndexes: number[],
  evidence: object,              // code-specific extra data
}
```

| Code | Severity | Condition |
|---|---|---|
| `NUMERIC_LIKE_LABEL` | warning | Row label looks like an isolated numeric range (e.g. "20-29") but was not matched to a known type; suppressed when neighboring rows also have range-like labels (category block) |
| `SUSPICIOUS_NUMERIC_LABEL` | warning | Row label looks like a single numeric value (e.g. "42", "3.5", "61,00") and may be an uncoded value or export/labeling artifact; suppressed inside a numeric category block (range labels such as 20-29 or single-numeric scale labels such as 1 / 2 / 3 / 4) |
| `SUSPICIOUS_ALL_100_ROW` | warning | Row contains 100% or equivalent (1.0) across all or most columns and is not a recognized base/NPS/mean/SD/variance row; may be a service, test, or uncoded row |
| `MISSING_ROW_LABEL_WITH_DATA` | warning | Row has at least 2 non-empty numeric cells but its label is empty, whitespace-only, or a symbol-only placeholder (e.g. "-", ".", "*"); skips recognized metric/service rows |
| `SUSPICIOUS_ERROR_LABEL` | warning | Row label contains a spreadsheet formula error (#N/A, #VALUE!, #REF!, #DIV/0!, etc.) or a programming error string (null, NaN, undefined, [object Object]); fires on all non-empty rows |
| `SUSPICIOUS_PLACEHOLDER_LABEL` | warning | Row label matches a placeholder or test-row keyword in English (todo, tbd, test¹, dummy, placeholder, temp, delete, remove, ignore, xxx, asdf, qwerty, lorem ipsum) or Russian (тест¹, тестовая строка², удалить, не использовать, заглушка, временно, черновик); skips metric/service rows. ¹ "test"/"тест" fire only when the word starts the label (e.g. "test", "test row") — not when it follows a noun (e.g. "Concept Test", "Product Test"). ² "тестовая" fires only as "тестовая строка/тестовая_строка/тестовая-строка" — not for research labels like "Тестовая концепция" or "Тестовая упаковка" |
| `SUSPICIOUS_CODE_LIKE_LABEL` | warning | Row label looks like a raw variable or code name (e.g. q1_1, Q12_3, var_005, brand_99, d1r3); no spaces allowed; skips metric/service rows. Underscore alone is not sufficient — requires /_\d+$/ (ends with _digits) or /^[a-zA-Z]{1,3}\d+_/ (short-prefix + digits before underscore). Labels like Top_2_Box, No_answer, Brand_A do not fire. |
| `GLOBAL_TOTAL_BASE_TOO_SMALL` | critical | Global Total base < any other column base |
| `LOCAL_TOTAL_BASE_TOO_SMALL` | critical | Local Total base < any member base in its group |
| `LOCAL_TOTAL_BASE_LESS_THAN_SUM` | warning | Local Total base < sum of group member bases |
| `GLOBAL_TOTAL_BASE_LESS_THAN_GROUP_SUM` | warning | Global Total base < sum of any group's member bases |
| `NPS_MISMATCH` | warning / critical | Displayed NPS ≠ Promoters − Detractors (≥ 2pp difference) |

### QualitySummary

```js
{
  criticalCount: number,
  warningCount: number,
  infoCount: number,
  hasBlockingIssues: boolean,   // true when criticalCount > 0
}
```

### Summary

```js
{
  rowCount: number,
  columnCount: number,
  detectedMetricRows: number,   // rows with a recognized type (excludes empty / unknownText)
  detectedBlocks: number,
  baseRows: number,
  hasNps: boolean,
  hasMeans: boolean,
  hasBanner: boolean,
  hasGlobalTotal: boolean,
  hasWaveGroups: boolean,
}
```

## Extension points

### UI integration

Wire `buildTablePreviewModel` into `taskpane.js` behind a new "Check table" button.  
The returned model is JSON-serialisable and can be rendered as-is or passed to a dedicated preview panel.  
`warnings` provides a flat list suitable for a simple message list.

### Weighted / effective base rows

`rowSubtype` is reserved for future weighted, unweighted, and effective base variants.  
When the distinction is implemented, base consistency checks in this module should be extended to validate the relationship between the three base variants.

### Multi-column row labels

`labelParts`, `normalizedLabelParts`, `primaryLabel`, and `secondaryLabel` already represent multi-column label data.  
Future rendering can use `combinedLabel` directly or compose its own view from the parts.

### Base placement setting

Currently the model follows the existing detector/block-building convention for base row placement.
If a "base above" layout is added to the detector, the preview model should inherit it automatically because it delegates to `buildCalculationBlocks`.

### Additional quality checks

The `dataQualityIssues` array is open for extension.  
Possible future additions: missing base row warnings, columns with all-zero values, detected column count mismatch with banner, and row type conflicts.
