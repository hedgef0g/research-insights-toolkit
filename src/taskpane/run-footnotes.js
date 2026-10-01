/* global console, Excel */

import {
  collectStatisticTypeLabels,
  buildSignificanceFootnoteCellValue,
  resolveFootnoteSpan,
  resolveFootnotePlacement,
  FOOTNOTE_SCAN_WINDOW_ROWS,
} from "../core/significance-footnote";

import { LABEL_SCAN_COLUMNS_LEFT } from "../core/metric-detector";

/**
 * Builds a footnote job for a single processed table, or null when the footnote
 * setting is off / geometry is unusable.
 *
 * Geometry is sheet-absolute and 0-based:
 *   - dataStartRowIndex / dataStartColIndex: top-left of the data body.
 *   - dataRowCount / dataColCount: data body dimensions.
 *   - leftLabelValues: label columns immediately left of the data body. Their
 *     width sets how far left the merged footnote extends (the "full table width
 *     including label columns").
 *
 * Returns null when the footnote setting is off. The feature is explicitly
 * incompatible with "labels on left side" mode (labels are not adjacent to the
 * data, so a footnote cannot span the table correctly): readCalculationSettings-
 * FromPanel already forces addTableFootnote=false in that mode, and this extra
 * guard makes the rule independent of the caller.
 *
 * processedScopeSuffix is an optional, already-formatted detail (e.g.
 * " Обработано: B12:F34.") appended to the visible footnote text. It is
 * supplied only by Manual Run; auto-run callers omit it so their footnote text
 * is unchanged.
 */
export function buildSignificanceFootnoteJob({
  sheetName,
  dataStartRowIndex,
  dataStartColIndex,
  dataRowCount,
  dataColCount,
  leftLabelValues,
  adjacentLabelColumnCount,
  calculationBlocks,
  calculationSettings,
  processedScopeSuffix,
}) {
  if (!calculationSettings.addTableFootnote) return null;
  if (calculationSettings.labelsOnLeftSide) return null;
  if (!Number.isFinite(dataStartRowIndex) || !Number.isFinite(dataStartColIndex)) return null;
  if (!(dataRowCount > 0) || !(dataColCount > 0)) return null;

  const labelColumns =
    Array.isArray(leftLabelValues) && Array.isArray(leftLabelValues[0]) ? leftLabelValues[0].length : 0;

  const { tableLeftColIndex, tableRightColIndex } = resolveFootnoteSpan({
    dataStartColIndex,
    dataColCount,
    labelColumns,
    adjacentLabelColumnCount,
  });
  const tableBottomRowIndex = dataStartRowIndex + dataRowCount - 1;

  const footnoteCellValue = buildSignificanceFootnoteCellValue({
    confidenceLevel: calculationSettings.confidenceLevel,
    oneTailedTest: calculationSettings.oneTailedTest,
    statisticLabels: collectStatisticTypeLabels(calculationBlocks),
    scopeDetail: processedScopeSuffix,
  });

  return {
    sheetName,
    tableBottomRowIndex,
    tableLeftColIndex,
    tableRightColIndex,
    dataStartColIndex,
    footnoteCellValue,
  };
}

// ─── Significance settings footnote application ──────────────────────────────
//
// When the "Добавлять подпись под таблицей" setting is on, each PROCESSED table
// gets one footnote row inserted directly below it, merged across the full table
// width (label columns + data columns). Skipped/blocked/error tables get nothing.
//
// IMPORTANT ordering rule (see issue #281): footnotes must NOT be inserted while
// calculations are still running, because inserting a worksheet row shifts every
// candidate range below it and corrupts the ranges still queued for calculation.
// Instead, each run flow collects footnote "jobs" (pure geometry + text, computed
// from the already-known ranges) and applies them AFTER all calculations finish,
// bottom-to-top per worksheet so earlier insertions never shift a not-yet-applied
// job's row index.

/**
 * Inserts or updates a single footnote row for one table.
 *
 * A small bounded window of rows below the table body is scanned (values loaded
 * across a span that also covers a couple of columns to the LEFT, so the marker
 * is seen even when a label/data blank-gap column shifts the merged footnote's
 * anchor). Placement is decided by the pure resolveFootnotePlacement helper:
 *
 *  - if a generated RIT footnote already exists for the table (marker-based), it
 *    is updated in place — no new row, no duplicate;
 *  - ordinary user note rows directly below the table are skipped, and the
 *    generated footnote is inserted BELOW them (never overwriting them);
 *  - otherwise a brand-new worksheet row is inserted so nothing is overwritten.
 *
 * The footnote cells are merged across exactly [tableLeftColIndex .. tableRightColIndex].
 */
export async function writeOrInsertSignificanceFootnoteRow(context, worksheet, job) {
  const { tableBottomRowIndex, tableLeftColIndex, tableRightColIndex, dataStartColIndex, footnoteCellValue } = job;
  const width = tableRightColIndex - tableLeftColIndex + 1;
  if (width < 1) return;

  const firstRowBelowTable = tableBottomRowIndex + 1;

  // Scan a span starting a few columns LEFT of the table so a left-offset marker
  // (from a run whose label/gap geometry differed) is still detected.
  const scanStartCol = Math.max(0, tableLeftColIndex - LABEL_SCAN_COLUMNS_LEFT);
  const scanColCount = tableRightColIndex - scanStartCol + 1;
  const dataColRef = Number.isFinite(dataStartColIndex) ? dataStartColIndex : tableLeftColIndex;
  const dataColStartOffset = dataColRef - scanStartCol;

  // Load the bounded window of rows below the table to decide placement.
  const scanRange = worksheet.getRangeByIndexes(
    firstRowBelowTable,
    scanStartCol,
    FOOTNOTE_SCAN_WINDOW_ROWS,
    scanColCount
  );
  scanRange.load("values");
  await context.sync();

  const placement = resolveFootnotePlacement(
    scanRange.values,
    firstRowBelowTable,
    dataColStartOffset
  );
  const footnoteRowIndex = placement.rowIndex;

  if (placement.mode === "insert") {
    // Insert a new blank row so nothing existing below the table is overwritten.
    worksheet
      .getRangeByIndexes(footnoteRowIndex, 0, 1, 1)
      .getEntireRow()
      .insert(Excel.InsertShiftDirection.down);
    await context.sync();
  } else {
    // Updating an existing generated footnote row in place: clear the full scan
    // span first so a prior merge and any left-offset marker cell are removed
    // before the row is rewritten at the current geometry. The row is RIT's own
    // generated footnote row, so clearing its span never touches user content.
    const priorRange = worksheet.getRangeByIndexes(footnoteRowIndex, scanStartCol, 1, scanColCount);
    priorRange.unmerge();
    priorRange.clear(Excel.ClearApplyTo.contents);
    await context.sync();
  }

  const footnoteRange = worksheet.getRangeByIndexes(footnoteRowIndex, tableLeftColIndex, 1, width);

  // Clear any prior merge first: a merged range rejects multi-cell writes, and an
  // existing footnote from a previous run may already be merged across this span.
  footnoteRange.unmerge();
  await context.sync();

  const leftCell = worksheet.getRangeByIndexes(footnoteRowIndex, tableLeftColIndex, 1, 1);
  leftCell.values = [[footnoteCellValue]];

  if (width > 1) {
    footnoteRange.merge();
  }

  footnoteRange.format.font.italic = true;
  footnoteRange.format.font.bold = false;
  footnoteRange.format.fill.clear();
  footnoteRange.format.horizontalAlignment = "Left";
  footnoteRange.format.verticalAlignment = "Center";
  await context.sync();
}

/**
 * Applies a list of footnote jobs. Jobs are grouped per worksheet and applied
 * bottom-to-top (descending table bottom row) so that inserting a footnote row
 * for a lower table never shifts the row index of a higher table still pending.
 *
 * Tolerant: a failure on one footnote is logged and skipped; it never aborts the
 * Run itself (calculations have already completed and been written at this point).
 */
export async function applySignificanceFootnoteJobs(context, jobs) {
  if (!Array.isArray(jobs) || jobs.length === 0) return;

  const jobsBySheet = new Map();
  for (const job of jobs) {
    if (!job || !job.sheetName) continue;
    if (!jobsBySheet.has(job.sheetName)) jobsBySheet.set(job.sheetName, []);
    jobsBySheet.get(job.sheetName).push(job);
  }

  for (const [sheetName, sheetJobs] of jobsBySheet) {
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const sorted = sheetJobs
      .slice()
      .sort((a, b) => b.tableBottomRowIndex - a.tableBottomRowIndex);
    for (const job of sorted) {
      try {
        await writeOrInsertSignificanceFootnoteRow(context, worksheet, job);
      } catch (footnoteErr) {
        console.warn("RIT: не удалось записать подпись под таблицей.", footnoteErr);
      }
    }
  }
}

/**
 * Convenience wrapper that applies footnote jobs inside their own Excel.run.
 * Used by run flows whose calculations completed in a different context.
 */
export async function applySignificanceFootnoteJobsInOwnContext(jobs) {
  if (!Array.isArray(jobs) || jobs.length === 0) return;
  try {
    await Excel.run(async (context) => {
      await applySignificanceFootnoteJobs(context, jobs);
    });
  } catch (footnoteErr) {
    console.warn("RIT: не удалось записать подписи под таблицами.", footnoteErr);
  }
}
