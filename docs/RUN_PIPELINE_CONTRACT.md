# Run Pipeline Contract

This document records the current Run behavior that must be preserved before
and after extracting shared Run pipeline code from `src/taskpane/taskpane.js`.

PR #327 documented the contract, PR #328 extracted the shared per-table Run
executor to `src/taskpane/run-pipeline.js`, and PR #329 moved the Run dispatcher
helpers there. The contract remains behavior-preservation first: batch Run
handlers, report writing, footnote writes, banner marker writes, design recolor
execution, and `context.sync` boundaries should not move unless a later task
explicitly scopes that work.

## Current Entry Points

### Manual selected-range Run

`runSignificanceFromSelection` is the stable first-class workflow:

- reads settings from the taskpane;
- reads exactly the user's current selected range;
- blocks unsupported setting combinations before Excel writes;
- rejects non-contiguous and unsafe broad selections;
- interprets the selected range with `interpretSelectedRange`;
- detects banner structure when banner-aware mode is enabled;
- runs marker-overflow preflight before any table write;
- detects calculation blocks and calculates significance;
- writes calculated values/fills back through the Excel writer;
- writes banner markers when banner-aware mode is enabled;
- queues design recolor and footnote work after calculation;
- applies design recolor before footnotes;
- applies footnotes in a separate context after the data body write;
- updates taskpane status and performance diagnostics.

Manual Run must remain selected-range first. It must not scan the worksheet or
workbook unless a future task explicitly changes that product behavior.

### Current-table Run

`runAutoCurrentTableSignificance` resolves the table under the active cell and
delegates the table write to `runSignificanceForRange`. It also performs the
pre-resolver selection guard, writes optional one-row Run report output, and
applies returned design recolor and footnote jobs after the table pipeline
finishes.

### Current-sheet Run

`runCurrentSheetSignificance` scans the active sheet inventory, filters eligible
candidates, runs operation-level marker-overflow preflight before the first write,
then processes tables through a shared-context batch with a per-table fallback.
It accumulates Run report rows, status/progress lines, design recolor jobs, and
footnote jobs. Recolor jobs are applied before footnote jobs after all table
calculations finish.

### Workbook Run

`runAutoSignificance` mirrors current-sheet Run over workbook inventory. It
skips generated sheets and non-eligible candidates, runs operation-level
marker-overflow preflight before the first write, uses a workbook-level shared
context where possible, falls back per table when needed, and writes optional Run
report output.

### Table pipeline

`runSignificanceForRangeInContext` is the current reusable table executor for
batch flows and lives in `src/taskpane/run-pipeline.js`. It must remain callable
inside an existing `Excel.run` context. It loads the source range, interprets the
selected/table range, detects banners and calculation blocks, performs per-table
marker-overflow preflight as a safety net, queues body and banner writes, and
returns pure job objects for deferred design recolor and footnotes.

`runSignificanceForRange` is only a wrapper that provides its own `Excel.run`
context around the extracted per-table executor.

## Boundary Contracts

### Selected range interpretation

Run callers rely on `interpretSelectedRange` to resolve the user's selected range
or candidate range into exactly one supported state before mutation:

- pass-through: the selected/candidate range is already the numeric data body;
- normalized: safe surrounding context is trimmed to a numeric data body;
- blocked: unsafe shape, ambiguous boundary, or unsupported broad selection.

Run extraction must not bypass this guardrail, silently auto-trim unsupported
shapes, or move blocking decisions after any Excel mutation.

### Preflight marker overflow

Marker-overflow preflight is a write barrier.

- Manual Run checks the interpreted table before writing the data body.
- Current-sheet and workbook Run call `preflightBatchMarkerOverflow` once before
  any table write so Stop leaves the operation without partial results.
- `runSignificanceForRangeInContext` keeps per-table preflight as a safety net.
- A shared marker-overflow decider caches Stop/Continue per operation.
- Continue mutates only the operation settings object by setting
  `allowMultiCharacterMarkers = true`.
- Previous-column mode must not prompt because it does not use letter labels.

### Calculation block detection

After interpretation and optional banner detection, Run detects metric rows and
builds calculation blocks from the numeric data body and left labels. Extraction
must preserve supported block behavior for proportions, means with spread/base,
NPS structures, and banner-aware calculation pairs. Statistical formulas and
marker assignment are outside the extraction scope.

### Excel writer

The Excel writer remains the only layer that writes significance values/fills
back to the data body. Extraction must preserve numeric display conventions and
must not convert the full selected range to text. Writer calls must remain after
all preflight/blocking decisions for that table.

### Banner marker updates

Banner marker writing is separate from data-cell marker writing. Run computes or
uses a banner label map, clears stale banner markers for the relevant header
area, and queues replacement banner markers. Extraction must preserve the
current banner-aware mode gates, label-map behavior, and final flush timing.

### Footnote placement/update

Run builds `FootnoteJob` objects from pure geometry and text. It must not insert
footnote rows while there are pending table writes in the same sheet/workbook
batch because row insertion can shift later candidate ranges. Batch flows collect
jobs and apply them bottom-to-top per sheet after all calculations complete.
Current-table Run applies its single footnote job after the table pipeline
returns. Manual Run appends the processed local range suffix; auto-run callers do
not.

### Design recolor jobs

Run builds `DesignRecolorJob` objects from pure geometry. Recolor writes do not
shift rows, so they are applied after calculation and before footnote insertion.
Extraction must keep `labelsOnLeftSide` disabled for recolor and must preserve
the L-shaped rectangle contract: banner rows over data columns, adjacent label
columns over data rows, corner untouched.

### Run report rows

Batch and current-table Run callers own report accumulation and optional report
sheet writing. The table pipeline should return enough structured information to
build a `RunReportRow`; it should not directly decide whether the generated Run
report sheet is enabled. Report writing may stay in the taskpane until a later
report-controller extraction.

### Progress and status updates

Run extraction must keep user-visible status timing:

- running status appears before long work begins;
- current-sheet/workbook progress updates after each eligible table;
- Stop, skipped, blocked, and error messages remain visible;
- final summary includes processed/skipped/error counts and relevant detail
  lines;
- report-write failures append/report separately and do not invalidate completed
  calculations.

### `context.sync` boundaries

The current Run flow depends on explicit Office.js flush points:

- load selected/source range metadata and values before interpretation;
- flush table body, fill, and queued banner marker writes only after preflight
  and calculation have completed;
- keep batch table writes in a shared context where possible;
- run design recolor and footnote jobs in their own contexts after calculation;
- write optional Run report sheets after calculation/report-row accumulation.

Future extraction must treat sync placement as part of the behavior contract, not
an implementation detail to optimize casually.

## Future Object Shapes

The future extraction should use JSDoc/object-shape contracts in plain
JavaScript. These names are proposed contracts, not current exported types.

```js
/**
 * @typedef {object} RunPipelineInput
 * @property {object} context Active Excel.run request context.
 * @property {object} sourceRange Loaded or loadable Excel range proxy.
 * @property {string} sheetName Sheet name for diagnostics/jobs.
 * @property {string} rangeAddress Original selected/candidate range address.
 * @property {object} calculationSettings Mutable per-operation settings.
 * @property {MarkerOverflowDecision} markerOverflowDecider Shared decision cache.
 * @property {RunPipelineCallbacks} callbacks UI/report/perf dependencies.
 * @property {"manual"|"current-table"|"sheet"|"workbook"} scope Caller scope.
 * @property {string} [processedScopeSuffix] Manual Run footnote suffix.
 */

/**
 * @typedef {object} RunPipelineResult
 * @property {"processed"|"skipped"|"blocked"|"stopped"|"error"} status
 * @property {string} sheetName
 * @property {string} rangeAddress Original requested range.
 * @property {string} [processedRangeAddress] Actual normalized write target.
 * @property {number} [blocksProcessed]
 * @property {string} message User-facing or report-facing message.
 * @property {RunTableResult} [table]
 * @property {FootnoteJob|null} [footnoteJob]
 * @property {DesignRecolorJob|null} [recolorJob]
 * @property {object} [_phasesMs] Optional performance diagnostics.
 */

/**
 * @typedef {object} RunTableResult
 * @property {Array<Array<*>>} valuesForCalculation Interpreted data body.
 * @property {Array<Array<*>>} leftLabelValues Row-label matrix.
 * @property {object|null} bannerStructure Detected banner metadata.
 * @property {Array<object>} calculationBlocks Blocks sent to significance logic.
 * @property {object} writeTargetGeometry 0-based data-body coordinates.
 * @property {Array<object>} selectedRangeWarnings Warning-only guardrails.
 */

/**
 * @typedef {object} RunBatchResult
 * @property {number} processed
 * @property {number} skipped
 * @property {number} errors
 * @property {boolean} markerOverflowStopped
 * @property {RunReportRow[]} reportRows
 * @property {FootnoteJob[]} footnoteJobs
 * @property {DesignRecolorJob[]} recolorJobs
 * @property {string[]} detailLines
 */

/**
 * @typedef {object} RunReportRow
 * @property {string} sheetName
 * @property {string} title
 * @property {string} rangeAddress
 * @property {"processed"|"skipped"|"blocked"|"stopped"|"error"} status
 * @property {string} message
 * @property {string} selectedBase
 * @property {string} metricTypes
 * @property {string|number} warnings
 * @property {string|number} critical
 * @property {string} warningDetails
 * @property {string|number} blocksProcessed
 */

/**
 * @typedef {object} FootnoteJob
 * @property {string} sheetName Empty string is allowed for active worksheet jobs.
 * @property {number} tableBottomRowIndex
 * @property {number} tableLeftColIndex
 * @property {number} tableRightColIndex
 * @property {number} dataStartColIndex
 * @property {string} footnoteCellValue Includes the invisible generated marker.
 */

/**
 * @typedef {object} DesignRecolorJob
 * @property {string} sheetName
 * @property {string} color
 * @property {Array<{rowIndex:number,columnIndex:number,rowCount:number,columnCount:number}>} rects
 */

/**
 * @typedef {object} MarkerOverflowDecision
 * @property {() => Promise<"continue"|"stop">} resolve
 * @property {"continue"|"stop"|null} decision
 */

/**
 * @typedef {object} RunPipelineCallbacks
 * @property {(message:string) => void} setStatusMessage
 * @property {(eventName:string, details?:object) => void} [perfLog]
 * @property {(key:string, vars?:object) => string} t
 * @property {(args:object) => string} [buildProgressStatus]
 */
```

## Pure Coverage Added or Confirmed

Existing pure coverage already protects these Run-adjacent contracts:

- marker-overflow decider caching and in-flight prompt reuse;
- non-banner batch marker-overflow Stop/Continue behavior;
- running status and batch progress formatting;
- selected-range guardrail message formatting;
- footnote marker, suffix, span, placement, and removal helpers;
- design recolor label-width and rectangle job helpers;
- banner local significance label maps.

PR #327 strengthened `preflightBatchMarkerOverflow` coverage for mixed candidate
widths and missing inventory metadata without requiring Office.js stubs.

## Missing Test Coverage

These behaviors are intentionally not unit-tested in PR5B because doing so would
require moving production Run code or building brittle Office.js doubles:

- the full manual selected-range Run path;
- `runSignificanceForRangeInContext` end-to-end table execution;
- banner-aware exact marker-overflow preflight;
- report-row creation inside current sheet/workbook Run loops;
- context corruption fallback from shared-context batch to per-table run;
- actual Excel writer, banner marker writer, report-sheet writer, design recolor
  writer, and footnote insertion side effects;
- exact `context.sync` ordering under real Excel.

## Current Extraction Status

PR #328 extracted the shared per-table Run executor behind the object shapes
above, leaving batch loops and generated sheet writers in `taskpane.js`. PR #329
moved `calculateBlockResults` and `getFirstBannerStructureError` into
`src/taskpane/run-pipeline.js`. Later extraction should preserve all existing
entry points and current status/report/job behavior unless explicitly scoped
otherwise.

Future manual smoke coverage for Run extraction must include:

1. Manual Run selected numeric range.
2. Manual Run sloppy/full-table selection normalization.
3. Current-table Run.
4. Current-sheet Run.
5. Workbook Run.
6. Marker overflow Stop/Continue.
7. Previous-column mode.
8. Banner-aware mode.
9. Footnote placement/update.
10. Banner marker writing.
11. Design recolor.
12. Run report output.
