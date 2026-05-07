# Table Structure Matrix

**Purpose:** Formal validation artifact describing which research table structures are currently supported, partially supported, or unsupported in Research Insights Toolkit.

**Disclaimer:** This matrix is a validation/control artifact. Support statuses reflect documented/currently known behavior and must be verified against runtime code and smoke tests before relying on them for release decisions.

**How to read this document:**

| Status | Meaning |
|---|---|
| SUPPORTED | Documented/currently known to work in the current workflow; should be verified by smoke tests. |
| PARTIAL | Works in most cases; known gaps or edge-case risks. |
| UNSUPPORTED | Not implemented; will produce wrong results or an error. |
| FUTURE | Identified; no implementation started; design required before coding. |

---

## 1. Basic Table Structures

### 1.1. Proportions

| Structure | Status | Current behavior | Risk / notes |
|---|---|---|---|
| Single proportion row + Base | SUPPORTED | Pooled z-test; markers written into proportion row. | -- |
| Multiple proportion rows + shared Base | SUPPORTED | Each proportion row gets its own markers; all share the Base. | -- |
| Proportion row without Base | UNSUPPORTED | Detector cannot form a calculation block. Calculation stops or block is silently skipped. | Always include a Base row. |
| Proportion expressed as decimal (0.42) | SUPPORTED | Normalized to proportion before test; display convention preserved. | See numeric output rules in STATUS.md. |
| Proportion expressed as percent (42%) | SUPPORTED | Excel percent format preserved in output. | -- |
| Out-of-range proportion (< 0 or > 1 after normalization) | PARTIAL | No dedicated guard; value is passed to z-test. May produce statistically nonsensical results. | Data quality check recommended before running. Follow-up: add data-quality warning. |

### 1.2. Means

| Structure | Status | Current behavior | Risk / notes |
|---|---|---|---|
| Mean + SD + Base | SUPPORTED | Welch's t-test. Marker written to Mean row only. | -- |
| Mean + Variance + Base | SUPPORTED | Variance converted to SD internally. Marker written to Mean row only. | -- |
| Mean + Base (no spread) | UNSUPPORTED | Without SD or Variance the spread test cannot run. Block is not formed. | Provide SD or Variance. |
| SD row receives markers | UNSUPPORTED | SD row never receives markers. | By design. Regression risk: verify SD stays marker-free. |
| Variance row receives markers | UNSUPPORTED | Variance row never receives markers. | By design. Same regression risk as SD. |
| Multiple mean blocks in one selection | SUPPORTED | Each Mean+SD/Variance+Base triplet is detected as a separate block. | Shared Base is resolved correctly. |

### 1.3. NPS Variants

| Structure | Status | Current behavior | Risk / notes |
|---|---|---|---|
| NPS-first (NPS / Promoters / Detractors / Base) | SUPPORTED | NPS recalculated from Promoters minus Detractors. NPS receives NPS markers; Promoters and Detractors receive ordinary proportion markers; Base receives no marker. | -- |
| NPS-first with Neutral (NPS / Promoters / Neutral / Detractors / Base) | SUPPORTED | Same as NPS-first; Neutral receives ordinary proportion markers. | -- |
| Extended NPS (Scale rows / NPS / Base) | SUPPORTED | Scale rows receive ordinary proportion markers; NPS receives NPS markers; Base receives no marker. | -- |
| NPS + SD/Base (spread path) | SUPPORTED | NPS spread significance. NPS receives spread markers; SD and Base receive no markers. | -- |
| NPS + Variance/Base (spread path) | SUPPORTED | Same as SD path. | -- |
| NPS without Promoters/Detractors and without SD/Variance | UNSUPPORTED | Detector cannot determine NPS test path. Block not formed. | Provide at minimum one recognized spread or structure row. |
| Arbitrary NPS ordering (Detractors before Promoters) | PARTIAL | Detector uses keyword matching, not strict row order. Works in most cases but edge-case ordering may fail detection. | Known gap; no issue filed yet. |

---

## 2. Banner Structures

### 2.1. Banner Detection

| Structure | Status | Current behavior | Risk / notes |
|---|---|---|---|
| No banner (one-level fallback group) | SUPPORTED | All non-Total columns form one default group. | -- |
| One-level banner (single header row) | SUPPORTED | Groups detected from column labels in the first row above selection. | -- |
| Two-level banner with repeated labels | SUPPORTED | Upper level repeated-label pattern detected; groups formed. | -- |
| Two-level banner with merged-like cells (empty spans) | SUPPORTED | Reconstructed span detection fills in empty upper cells. | Gaps or irregular merges may mis-detect group boundaries. |
| Three-or-more-level banners | UNSUPPORTED | Only the lowest two levels are processed. Higher levels are ignored. | FUTURE: full multi-level banner support. |
| Selected range starts in row 1 (no room for banner row) | UNSUPPORTED | Calculation stops with a message asking the user to add a row above selection. | -- |
| Banner row with no labels above selection | PARTIAL | Status message emitted; calculation falls back to one default group. | Message may be unclear in some cases. |
| Mixed banner label language (Russian / English) | SUPPORTED | Dictionary-driven detection supports Russian and English label keywords. | -- |

### 2.2. Total Detection Under Banner

| Structure | Status | Current behavior | Risk / notes |
|---|---|---|---|
| Local Total column in each banner group | SUPPORTED | Each group's first column matching a recognized Total keyword is detected as local Total. | -- |
| Global Total column (spanning all groups) | SUPPORTED | Detected from upper banner level; used as the only Total reference. Local Totals fall into ordinary target role. | Status message shown when global Total is used. |
| Group with no Total column | SUPPORTED | Ordinary all-vs-all comparisons within group only. Not an error. | -- |
| Multiple Total columns in one group | UNSUPPORTED | Calculation stops with message about multiple Totals. | User must fix table layout. |
| Total column outside selected range | PARTIAL | Specified in design; may need edge-case hardening. See STATUS.md known debt. | Follow-up issue recommended. |
| Manual Total (first-column-is-Total) without banner | SUPPORTED | First column treated as Total reference. | Disabled when banner structure is enabled. |

### 2.3. Wave Banners

| Structure | Status | Current behavior | Risk / notes |
|---|---|---|---|
| Wave group in banner (label matches wave keywords) | SUPPORTED | Auto previous-column applied inside wave group; ordinary comparisons for non-wave groups. | Status message explains auto previous-column. |
| Manual previous-column overrides wave auto mode | SUPPORTED | When user enables previous-column globally, auto-wave mode is superseded. | -- |
| Compare-only-with-Total suppresses wave auto mode | SUPPORTED | Total comparisons run; auto previous-column not triggered. | -- |
| Numeric-only group labels (1, 2, 3) used as wave signal | UNSUPPORTED | Numeric labels are not treated as wave indicators. | By design; avoids false wave detection. |
| Banner letter writing in wave groups | UNSUPPORTED | Wave groups receive no banner letters. | By design. Previous-column mode disables banner letters for those groups. |

### 2.4. Banner Letter Writing

| Structure | Status | Current behavior | Risk / notes |
|---|---|---|---|
| Banner letters in one-level banner | SUPPORTED | Letters written into lowest banner row per group, excluding Total columns. | -- |
| Banner letters in two-level banner | SUPPORTED | Letters written only to lowest banner level; upper level unchanged. | -- |
| Banner letters with global Total | SUPPORTED | Global Total and local Total columns receive no letters. | -- |
| Banner letters in previous-column mode | UNSUPPORTED | Previous-column mode disables banner letter writing entirely. | By design. |
| Banner letters when selected range is in row 1 | UNSUPPORTED | No room above selection; calculation stops before writing. | -- |

---

## 3. Messy Table Structures

### 3.1. Over-Broad Range Selection

| Structure | Status | Current behavior | Risk / notes |
|---|---|---|---|
| Selection includes title / question text row above data | PARTIAL | Guardrail warning emitted in taskpane; calculation proceeds. Non-data rows may be misidentified as metric rows. | Warning-only; Run is not blocked. Selection is never auto-trimmed. See SELECTED_RANGE_NORMALIZATION.md spec. |
| Selection includes empty rows between blocks | PARTIAL | Detector attempts to skip empty rows; may form incorrect block boundaries. | Manual selection of clean data range is recommended. |
| Selection includes column-label text cells | PARTIAL | Label column is expected to be outside the selected range (to the left). If label cells are included in selection, detection may degrade. | User should select only numeric data area. |
| Selection includes rows below Base | PARTIAL | Extra rows after Base may be interpreted as a new block start or silently ignored. | Guardrail warning covers this case partially. |

### 3.2. Multi-Column Labels

| Structure | Status | Current behavior | Risk / notes |
|---|---|---|---|
| Single label column immediately left of selection | SUPPORTED | Default behavior. | -- |
| Labels separated from data by numeric columns | SUPPORTED | Detector skips numeric intermediate columns when searching for text labels. | -- |
| Labels in leftmost sheet columns (labels-on-left-side mode) | SUPPORTED | Enabled via labels-on-left-side setting; reads labels from left edge of sheet. | -- |
| Multi-column label area (spanning two text columns) | UNSUPPORTED | Only one label column is read. Second label column is ignored. | FUTURE: multi-column label support. Spec: docs/MULTI_COLUMN_LABELS_AND_WEIGHTED_BASES.md. GST coverage: GST-068, GST-069. |
| No label column anywhere | PARTIAL | Detector proceeds without row-type keywords; may fall back to heuristics or misidentify rows. | Metric detection accuracy degrades. Ensure labels are present. |

### 3.3. Adjacent Tables

| Structure | Status | Current behavior | Risk / notes |
|---|---|---|---|
| Two separate tables side-by-side in one selection | UNSUPPORTED | Treated as a single wide table; block detection will be incorrect. | Select each table separately. No multi-table detection. |
| Two stacked metric blocks sharing one Base row | SUPPORTED | Shared Base detection handles this case. | -- |
| Stacked tables with no separating Base row | UNSUPPORTED | Block detector cannot determine where one table ends and another begins without a Base row boundary. | Ensure each block has its own Base row. |

---

## 4. Base Structures

Planned base priority order (not fully implemented yet): Effective Base > Unweighted Base > plain Base > Weighted Base fallback with warning. Effective Base and Weighted Base support are core correctness features, not premium or deferred features. Full spec: docs/MULTI_COLUMN_LABELS_AND_WEIGHTED_BASES.md. GST coverage: GST-062 to GST-067, GST-072.

| Structure | Status | Current behavior | Risk / notes |
|---|---|---|---|
| Unweighted Base (plain integer counts) | SUPPORTED | Standard behavior. | -- |
| Weighted Base (non-integer, e.g. 487.3) | PARTIAL | No explicit weighted-base path. Value is used as-is in z-test denominator. Statistical validity depends on weighting method used by the researcher. | FUTURE: weighted base support planned. Intended role: fallback with warning when no Effective Base is present. Spec: docs/MULTI_COLUMN_LABELS_AND_WEIGHTED_BASES.md. GST: GST-062, GST-063, GST-066. |
| Effective Base | FUTURE | No dedicated effective-base row keyword or logic yet. | Core correctness feature. Planned MVP-core support. Intended as the highest-priority base type. Spec: docs/MULTI_COLUMN_LABELS_AND_WEIGHTED_BASES.md. GST: GST-064, GST-067, GST-072. |
| Shared Base (one Base row for multiple metric rows above) | SUPPORTED | Detector resolves shared Base correctly. | -- |
| Missing Base row | UNSUPPORTED | Block cannot be formed. Calculation skipped for that block. | Always include a Base row per block or as shared Base. |
| Base row with small base (below threshold) | SUPPORTED | Column excluded from comparisons; small-base fill applied across block; Total small-base stops calculation with message. | -- |
| Base row is Total column with small base | SUPPORTED | Calculation stops; status message directs user to check Total base. | -- |
| Multiple Base rows in one block | PARTIAL | Detector uses the first matching Base row. Extra Base rows may cause incorrect block boundaries. | Keep one Base row per block. |

---

## 5. Metric Structures

### 5.1. Proportion Rows

| Structure | Status | Current behavior | Risk / notes |
|---|---|---|---|
| Labeled proportion row (%, Percent, or equivalent Russian keywords) | SUPPORTED | Dictionary-driven keyword detection. | -- |
| Unlabeled proportion row (no matching keyword) | PARTIAL | Heuristic fallback may detect based on value range. May misidentify. | Label rows clearly. |
| NET row (NET, Top, Top-N, or equivalent keywords) | SUPPORTED | Treated as an ordinary proportion row; receives proportion markers. | -- |
| Index row | PARTIAL | No dedicated index-row path. If labeled as proportion, receives proportion markers. Correctness depends on table design. | FUTURE: explicit index row handling. |
| Count row (n=, count-equivalent keywords) | FUTURE | Count rows are a distinct metric type. They are not equivalent to Base rows and must not automatically trigger significance calculations. Count + Base may later support derived proportion calculation, but only with explicit opt-in / trust-first behavior. No significance markers are written to count rows in the current workflow. | Do not treat arbitrary count rows as Base. |

### 5.2. Derived and Support Rows

| Structure | Status | Current behavior | Risk / notes |
|---|---|---|---|
| Promoters row (NPS structure) | SUPPORTED | Receives ordinary proportion markers within NPS block. | -- |
| Detractors row (NPS structure) | SUPPORTED | Receives ordinary proportion markers within NPS block. | -- |
| Neutral row (NPS structure) | SUPPORTED | Receives ordinary proportion markers within NPS block. | -- |
| SD row | SUPPORTED | Never receives markers; part of Mean block structure. | By design. Regression risk: verify no markers appear on SD. |
| Variance row | SUPPORTED | Same as SD; never receives markers. | By design. |
| Base row | SUPPORTED | Never receives markers. | By design. |
| Arbitrary pairwise row-to-row significance inside a column | UNSUPPORTED | Not implemented and not planned in current architecture. RIT compares column values within a row, not row values within a column. | By design. Will not be added without an explicit product decision. |
| Custom row types not in dictionary | PARTIAL | Detector falls back to heuristics; may misidentify or skip. | Add keywords to dictionary.config.js for new row types. |

---

## 6. Data Quality Cases

| Structure | Status | Current behavior | Risk / notes |
|---|---|---|---|
| Out-of-range proportion value (< 0 or > 1 after normalization) | PARTIAL | No dedicated guard; test proceeds. Results may be statistically invalid. | FUTURE: data quality warning for out-of-range values. |
| Missing Base (0 or blank) | PARTIAL | z-test denominator will be 0 or NaN. Division-by-zero handled by JavaScript (Infinity / NaN). Result written as non-significant silently. | FUTURE: explicit zero-base guard. |
| Suspicious Total (Total significantly different from sum of parts) | UNSUPPORTED | RIT does not validate internal table consistency. Total is used as-is. | FUTURE: Check table mode (see docs/table-preview-model.md). |
| Negative Base value | PARTIAL | No explicit guard. Passed into test; statistically invalid. | FUTURE: data quality guard. |
| All columns have the same value (zero variance) | PARTIAL | z-test or t-test will yield non-significant; no divide-by-zero if Base > 0. | Expected behavior; not an error. |
| Non-numeric cells in selected data range | PARTIAL | Cells are read and normalized; non-numeric values default to 0 or are skipped. May silently produce incorrect blocks. | Select only numeric data. |
| Blank rows inside selected range | PARTIAL | Blank rows may be skipped or may break block detection depending on position. | Avoid blank rows inside selected range. |
| Duplicate column headers in banner | PARTIAL | Repeated labels trigger repeated-label banner detection; may group incorrectly if non-wave labels repeat unintentionally. | Verify banner label uniqueness per group. |

---

## 7. Comparison Mode Structures

| Structure | Status | Current behavior | Risk / notes |
|---|---|---|---|
| All-vs-all (default) | SUPPORTED | Every column pair compared within group. | -- |
| Total comparison (first-column-is-Total) | SUPPORTED | Each non-Total column compared with Total. | -- |
| Compare only with Total | SUPPORTED | Only Total comparisons; no segment-vs-segment. | Requires Total location to be specified. |
| Exclude Total from comparisons | SUPPORTED | Total excluded; segment-vs-segment only. | Requires Total location. |
| Previous-column comparison | SUPPORTED | Each column compared with its left neighbor. Arrows used instead of letters. | Incompatible with compare-only-with-Total. |
| Previous-column + Total excluded | SUPPORTED | Total column skipped in previous-column chain. | -- |
| Previous-column + banner structure | SUPPORTED | Previous-column applied inside each banner group; no cross-group chaining. | -- |
| Row-to-row comparisons within a column | UNSUPPORTED | Not implemented. See section 5.2 above. | By design. |
| Custom pairwise specification | UNSUPPORTED | No mechanism for user-defined comparison pairs beyond the above modes. | FUTURE if needed. |

---

## 8. Settings and Configuration

| Structure | Status | Current behavior | Risk / notes |
|---|---|---|---|
| Confidence levels: 66.6%, 80%, 90%, 95%, 99% | SUPPORTED | Selector in taskpane; affects all metric types. | -- |
| One-tailed test | SUPPORTED | Lower critical threshold; affects proportions, means, NPS, and Total comparisons. | -- |
| Two-tailed test (default) | SUPPORTED | Standard behavior. | -- |
| Round cell values off (default) | SUPPORTED | Proportions 1 decimal; means/SD/variance 2 decimals. | -- |
| Round cell values on | SUPPORTED | Proportions 0 decimals; means/SD/variance 1 decimal. Calculations use original values. | -- |
| Small-base exclusion | SUPPORTED | Columns below threshold excluded; fill applied. | -- |
| Local settings persistence | SUPPORTED | Stored in localStorage. | -- |
| Cloud settings persistence | UNSUPPORTED | Reserved for future implementation. | -- |
| Google Sheets | FUTURE | Not implemented in current Excel-first MVP. Planned as a first-class platform after MVP/parity phase; not a deprioritized platform. | No timeline set. |

---

## 9. Platform and Automation

| Structure | Status | Current behavior | Risk / notes |
|---|---|---|---|
| Automatic worksheet/workbook scanning | FUTURE | Not implemented. User must manually select the numeric data range and click Run. | Planned as automatic runner for the MVP path. Required pipeline: scan -> processing plan -> preview/warnings -> controlled write. Must not re-use repeated manual pipeline calls. |
| Google Sheets (platform) | FUTURE | Excel-first only. See section 8. | Planned first-class platform after parity phase. |
| Check-table / data quality validation mode | FUTURE | Not implemented. Specified in docs/table-preview-model.md. | -- |
| Selected range normalization | FUTURE | Spec in docs/SELECTED_RANGE_NORMALIZATION.md; implementation not started. | -- |

---

## Notes and Known Limitations

- **Row-to-row significance** (comparing row values within a single column) is explicitly not supported and not planned in the current architecture. RIT compares column values within a given row.
- **Automatic worksheet/workbook scanning** is not implemented. User must always select the numeric data range manually. Planned as automatic runner with a scan -> processing plan -> preview/warnings -> controlled write pipeline; must not repeat the existing manual pipeline.
- **Total outside selection** is partially designed but may need edge-case hardening.
- **Check-table / data quality validation** mode is specified in docs/table-preview-model.md but not yet implemented.
- **Selected range normalization** spec is in docs/SELECTED_RANGE_NORMALIZATION.md; implementation not started.
- **Effective Base and Weighted Base** support are core correctness features. Planned base priority: Effective > Unweighted > plain Base > Weighted with warning. Full spec: docs/MULTI_COLUMN_LABELS_AND_WEIGHTED_BASES.md.
- **Multi-column label area** support is FUTURE. Spec: docs/MULTI_COLUMN_LABELS_AND_WEIGHTED_BASES.md. GST coverage: GST-068, GST-069.
- This matrix reflects the current stabilization state. Update this document when new structures are implemented or when validation behavior changes.
