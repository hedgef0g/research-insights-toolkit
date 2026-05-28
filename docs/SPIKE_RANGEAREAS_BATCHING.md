# Spike: RangeAreas batching for body bold/fill formatting

Issue: [#284](https://github.com/hedgef0g/research-insights-toolkit/issues/284).
Scope: diagnosis only — no writer behavior change.
Outcome: **feasible**, recommended as a follow-up implementation issue.

## Problem

Full-formatting result writes on wide workbooks are dominated by per-row-run
formatting commands. Measured baseline (`RIT_PERF`, wide workbook, full
formatting):

| metric | value |
|---|---|
| `totalMs` | ~212.3 s |
| `writeMs` | ~200.2 s |
| `boldCommandCount` | 20,056 |
| `fillCommandCount` | 20,056 |
| `boldRectCommandCountEstimate` | 18,278 |
| `fillRectCommandCountEstimate` | 18,278 |

For comparison, the **markers-only** mode (#282) on the same workbook clocks
~20.8 s with zero bold/fill commands. The values/numberFormat 2D writes are
already efficient; the bottleneck is the ~40 k queued bold/fill operations.

A better row-run / rectangle extractor is a marginal win at best (the rect
estimate is already within 9 % of the row-run count). The real lever is
**batching many ranges into a single Office.js call** via
`worksheet.getRanges(addresses)` → `RangeAreas.format.font.bold` /
`.format.fill.color`.

## API verification

Confirmed against the bundled `@types/office-js@1.0.589` typings:

- `Excel.Worksheet.getRanges(address?: string): Excel.RangeAreas` — **ExcelApi 1.9**
- `Excel.RangeAreas.format` returns a `RangeFormat` — **ExcelApi 1.9**
- `RangeFormat.font.bold`, `RangeFormat.fill.color` are inherited from the
  same `RangeFont`/`RangeFill` used today by `Range.format.*` (ExcelApi 1.1).
  Setting them on a `RangeAreas.format` applies the value uniformly to every
  area in the `RangeAreas`. This matches the existing semantic we need:
  one value per call, but applied to many disjoint rectangles.

Address syntax: comma- or semicolon-separated A1 ranges, optionally with a
sheet prefix (e.g. `"A1:B2, A5:B5"` or `"Sheet1!A1:B2; Sheet1!D1:D4"`).

The address can be unqualified within the worksheet, so we can build it from
A1 references derived from `selectedRange.rowIndex` / `selectedRange.columnIndex`
(both ExcelApi 1.1, already loaded upstream of the writer).

### Host coverage and gating

ExcelApi 1.9 (Sept 2019) is available on:

- Excel on Windows (Microsoft 365 / Office 2019+ / Office 2021)
- Excel on Mac (Microsoft 365 / Office 2019+ / Office 2021)
- Excel on the web
- Excel on iPad

It is **not** available on Excel 2016 perpetual at any updated channel below
1.9. Our `manifest.xml` does not declare a `<Requirements>` block, so the
add-in currently loads at the host's default ExcelApi (1.1). A real
implementation must runtime-gate every RangeAreas call behind:

```js
Office.context.requirements.isSetSupported("ExcelApi", "1.9")
```

…and fall back to the existing per-row-run path when unsupported. Detection
must happen once per Excel.run, not per-cell.

## Practical limits

| limit | source | budget chosen for chunking |
|---|---|---|
| Excel formula-bar address string length | Excel 365: 8,192 chars; older builds: 2,048 chars | 2,000 chars/chunk |
| `RangeAreas` areas per call | Not formally documented; field reports of hundreds-to-low-thousands work | 100 areas/chunk |
| `Worksheet.union(...)` arg count | 30 (firstRange + secondRange + 28 additional) | **N/A** — we use `getRanges(address)`, not `union(...)` |
| `RangeAreas.areaCount` | `int32` max | not a practical concern |

Both caps are conservative. Real wide-workbook chunks stay around ~1.2 KB
each (see projection below), well inside the 2 KB cap and the 2,048-char
floor for legacy builds — so chunks remain valid on every supported host.

## Projected command-count reduction

Diagnostic helpers added to [excel-writer.js](../src/core/excel-writer.js)
turn the row-run span data the writer already produces into the A1 address
chunks a `getRanges`-based writer would consume. They run when
`captureWriterDetails` (RIT_PERF) is enabled and surface as
`writerDetails.boldRangeAreas` and `writerDetails.fillRangeAreas` in the
existing `[RIT perf]` log.

Synthetic projection sized to mirror the wide-workbook baseline (240 rows ×
84 row-runs/row = 20,160 row-runs, single-color fill):

| stage | row-run baseline | projected RangeAreas (100 areas/chunk) | reduction |
|---|---|---|---|
| Bold queued ops | 20,160 | 404 (202 chunks × 2 ops) | ~50× |
| Fill queued ops (1 dominant color) | 20,160 | 404 | ~50× |
| **Both combined** | **~40,320** | **~808** | **~50×** |
| Max chunk address length | n/a | 1,182 chars | inside 2 KB budget |

A two-color split (e.g. significant + lower-than-total) produces nearly the
same total because the chunk count adds, not multiplies — 41 + 162 = 203
chunks for the same area count.

The current row-run writer issues each `getCell(r, c).getResizedRange(0, w-1)
.format.font.bold = true` chain as several queued ops on the proxy; the
RangeAreas writer issues one `getRanges(addr).format.font.bold = true` chain
per chunk. Sync count stays at 1 — the chunked ops queue inside the existing
single `context.sync()`. The expected wall-clock effect is that
`full-formatting writeMs` should converge toward the **markers-only**
baseline (~20 s on the wide workbook) once row-run ops are gone, because
the values/numberFormat 2D writes already dominate that mode.

## Better rectangle extraction — separate verdict

The existing `boldRectCommandCountEstimate` (18,278) is within 9 % of the
row-run count (20,056) on the wide workbook. Even an optimal rectangle
covering would not get below the row-run floor for highly irregular masks,
and the marginal gain is dwarfed by what RangeAreas provides. **Not pursued.**

## Risks and mitigations

- **Host without ExcelApi 1.9.** Mitigated by gating on
  `isSetSupported("ExcelApi", "1.9")` once per write and falling back to
  the current row-run path. Both code paths must remain in the writer.
- **Address-string size on older builds.** Mitigated by the 2,000-char
  chunk cap, which keeps chunks inside the legacy 2,048-char formula-bar
  limit. Projected real max is ~1.2 KB.
- **Visual semantics drift.** The RangeAreas writer must apply the **same**
  bold cells and the **same** fill colors as today. Address chunks are
  built from the same `boldMask` / `fillReasonMask` and the same color
  resolver (`getFillColorForCellResult`). The chunk-grouping is by color,
  so no two colors are ever applied to the same chunk. Visual diff vs
  current writer should be empty.
- **Markers-only path.** Already opted out of bold/fill in #282; this spike
  does not touch that path.
- **Banner-letter writes.** Out of scope; banner write batching is already
  handled by recent banner-marker write changes and uses its own range
  builder.

## Diagnostics surface added in this PR

`writerDetails` (RIT_PERF only) now carries two new fields populated whenever
visual formatting runs and `selectedRange.rowIndex` / `columnIndex` are
loaded:

```jsonc
{
  // existing fields ...
  "boldRangeAreas": {
    "areaCountTotal": 20056,
    "chunkCount": 202,
    "commandCountEstimate": 404,
    "maxAddressLength": 1182
  },
  "fillRangeAreas": {
    "colorCount": 2,
    "areaCountTotal": 20056,
    "chunkCount": 203,
    "commandCountEstimate": 406,
    "maxAddressLength": 1186,
    "perColor": [
      { "color": "#E2F0D9", "areaCount": 16128, "chunkCount": 162, "commandCountEstimate": 324, "maxAddressLength": 1186 },
      { "color": "#FCE4D6", "areaCount": 3928,  "chunkCount":  41, "commandCountEstimate":  82, "maxAddressLength": 1180 }
    ]
  }
}
```

These let the human owner confirm the projection against any real workbook
**before** an implementation PR by comparing `boldCommandCount` /
`fillCommandCount` (today) against `boldRangeAreas.commandCountEstimate` /
`fillRangeAreas.commandCountEstimate` in the same log.

The estimate skips silently if anchor coordinates aren't loaded (no
exception, no log noise), keeping it zero-risk for non-instrumented runs.
The internal `_boldRunSpansByRow` / `_fillRunSpansByRow` fields used to
feed the estimate are deleted from the emitted `writerDetails` so the
RIT_PERF log stays compact on wide workbooks.

## Recommendation

**Implement** as a follow-up issue. Suggested scope:

1. Add a once-per-write `isSetSupported("ExcelApi", "1.9")` cache to the writer.
2. When supported, replace the inner per-row-run setters in
   `applyGroupedBoldFormatting` and `applyGroupedFillFormatting` with
   `selectedRange.worksheet.getRanges(chunkAddress).format.font.bold = true`
   / `.format.fill.color = color`, using `packAddressesIntoChunks` (already
   shipped here as a pure helper).
3. Keep the existing row-run path as the fallback branch — no removal.
4. Reuse the diagnostic helpers' chunking caps (100 areas / 2,000 chars)
   unless field measurement shows we can grow them safely.
5. Validate on the wide workbook against the diagnostic projection.
6. Manual smoke: full-formatting visual output unchanged; markers-only
   unchanged; Clear Significance unchanged.

Expected effect: full-formatting `writeMs` on the wide workbook drops from
~200 s to roughly the markers-only baseline (~10–20 s), with no visual
semantic change.

## Files touched by this spike

- [src/core/excel-writer.js](../src/core/excel-writer.js) — added pure
  helpers (`columnIndexToA1`, `runSpanToA1Address`, `packAddressesIntoChunks`,
  `buildBoldRangeAreasDiagnostics`,
  `buildFillRangeAreasDiagnosticsByColor`) and wired their output into
  `writerDetails` behind the existing `captureWriterDetails` gate. No call
  to Office.js, no formatting behavior change.
- [tests/core/excel-writer-rangeareas.test.mjs](../tests/core/excel-writer-rangeareas.test.mjs)
  — coverage for the pure helpers.
