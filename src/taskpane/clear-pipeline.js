/* global console, Excel */

import {
  removeSignificanceMarkersFromText,
  isSignificanceMarkerLabel,
} from "../core/significance";
import { LABEL_SCAN_COLUMNS_LEFT } from "../core/metric-detector";
import { resolveNumericOutput } from "../core/excel-writer";
import {
  resolveFootnoteRemovalRow,
  FOOTNOTE_SCAN_WINDOW_ROWS,
} from "../core/significance-footnote";
import { resolveClearTargetBodyRange } from "./selected-range-interpreter";
import { perfNow, perfElapsed } from "./taskpane-performance";

/**
 * Inner clear-significance pipeline for a single named range, using a
 * caller-supplied Office.js RequestContext.
 *
 * Extracted so workbook/sheet batch clears can share one Excel.run context
 * across all tables on the same worksheet, amortising per-context initialisation
 * overhead.  Callers that need a self-contained execution unit use
 * clearSignificanceForRange below, which wraps this in its own Excel.run.
 *
 * Returns { status, message, _clearDetails }.
 * status: "cleared" | "skipped"
 * _clearDetails is always populated on a "cleared" result; null on "skipped".
 */
export async function clearSignificanceForRangeInContext(context, sheetName, rangeAddress) {
  const _t0 = perfNow();
  const worksheet = context.workbook.worksheets.getItem(sheetName);
  const sourceRange = worksheet.getRange(rangeAddress);

  sourceRange.load(["values", "text"]);

  await context.sync();

  const selectedValues = sourceRange.values;
  const selectedText = sourceRange.text;

  if (!selectedValues || selectedValues.length < 1 || !selectedValues[0] || selectedValues[0].length < 1) {
    return { status: "skipped", message: "нет данных в диапазоне", _clearDetails: null };
  }

  const clearTarget = resolveClearTargetBodyRange({
    values: selectedValues,
    text: selectedText,
  });

  if (clearTarget.state === "blocked") {
    const codes =
      clearTarget.blockingReasons && clearTarget.blockingReasons.length > 0
        ? ` [${clearTarget.blockingReasons.join(", ")}]`
        : "";
    return { status: "skipped", message: `${clearTarget.blockingMessage}${codes}`, _clearDetails: null };
  }

  if (clearTarget.state === "empty") {
    return { status: "skipped", message: "нет данных после нормализации", _clearDetails: null };
  }

  let clearTargetRange;

  if (clearTarget.usesFullSelection) {
    clearTargetRange = sourceRange;
  } else {
    clearTargetRange = sourceRange
      .getCell(clearTarget.rowOffset, clearTarget.colOffset)
      .getResizedRange(clearTarget.rowCount - 1, clearTarget.colCount - 1);
  }

  clearTargetRange.load(["values", "numberFormat", "rowIndex", "columnIndex", "columnCount"]);

  await context.sync();

  const targetValues = clearTargetRange.values;
  const targetNumberFormats = clearTargetRange.numberFormat;
  const _knownDims = {
    rowIndex: clearTargetRange.rowIndex,
    columnIndex: clearTargetRange.columnIndex,
    columnCount: clearTargetRange.columnCount,
  };

  const nextValues = [];
  const nextNumberFormats = [];
  let bodyCellsChanged = 0;
  let bodyHasValueChange = false;
  let bodyHasFormatChange = false;

  for (let rowIndex = 0; rowIndex < targetValues.length; rowIndex++) {
    const valueRow = [];
    const formatRow = [];

    for (let columnIndex = 0; columnIndex < targetValues[rowIndex].length; columnIndex++) {
      const rawValue = targetValues[rowIndex][columnIndex];
      const currentFormat = targetNumberFormats[rowIndex][columnIndex];

      if (typeof rawValue === "number") {
        valueRow.push(rawValue);
        formatRow.push(currentFormat);
        continue;
      }

      const cleanedText = removeSignificanceMarkersFromText(rawValue);
      const resolved = resolveNumericOutput(cleanedText);

      if (resolved !== null) {
        if (resolved.value !== rawValue) bodyHasValueChange = true;
        if (resolved.format !== currentFormat) bodyHasFormatChange = true;
        if (resolved.value !== rawValue || resolved.format !== currentFormat) bodyCellsChanged++;
        valueRow.push(resolved.value);
        formatRow.push(resolved.format);
      } else {
        if (cleanedText !== rawValue) { bodyHasValueChange = true; bodyCellsChanged++; }
        valueRow.push(cleanedText);
        formatRow.push("@");
      }
    }

    nextValues.push(valueRow);
    nextNumberFormats.push(formatRow);
  }

  // Only write body data when something actually changed, avoiding a full
  // matrix write on a re-clear of an already-clean workbook.
  if (bodyHasFormatChange) clearTargetRange.numberFormat = nextNumberFormats;
  if (bodyHasValueChange) clearTargetRange.values = nextValues;

  clearTargetRange.format.font.bold = false;
  clearTargetRange.format.fill.clear();

  // Probe the row immediately below the cleared table for a generated footnote.
  // Queued before the banner read so its sync (or the final sync) flushes it —
  // no extra round-trip. Deletion happens later, bottom-to-top, by the caller.
  const { range: footnoteProbeRange, firstRowBelowTable, dataColStartOffset } = buildFootnoteProbe(
    worksheet,
    _knownDims,
    targetValues.length
  );
  footnoteProbeRange.load("values");

  // No body sync here — deferred to the banner read sync inside
  // clearBannerMarkerUpdatesForRange, which flushes all queued writes.
  const bannerDetails = await clearBannerMarkerUpdatesForRange(context, clearTargetRange, _knownDims);

  await context.sync();

  const footnoteRemovalRow = resolveFootnoteRemovalRow(
    footnoteProbeRange.values,
    firstRowBelowTable,
    dataColStartOffset
  );
  const footnoteRemovalJob =
    footnoteRemovalRow !== null ? { sheetName, footnoteRowIndex: footnoteRemovalRow } : null;

  const bodyCellsRead = targetValues.length * (targetValues[0] ? targetValues[0].length : 0);

  return {
    status: "cleared",
    message: "очищено",
    footnoteRemovalJob,
    _clearDetails: {
      bodyCellsRead,
      bodyCellsChanged,
      bannerCellsRead: bannerDetails ? bannerDetails.cellsRead : 0,
      bannerCellsChanged: bannerDetails ? bannerDetails.cellsChanged : 0,
      bannerWriteCommands: bannerDetails ? bannerDetails.writeCommands : 0,
      totalMs: perfElapsed(_t0),
    },
  };
}

/**
 * Clears significance markers for a single named range on a named sheet.
 *
 * Thin wrapper around clearSignificanceForRangeInContext that provides its own
 * Excel.run context.  Used by clearAutoCurrentTableSignificance and as a
 * per-table fallback when the shared-context batch in sheet/workbook clear
 * encounters an Office.js error that corrupts the shared context.
 *
 * Returns { status, message, _clearDetails }.
 * status: "cleared" | "skipped" | "error"
 */
export async function clearSignificanceForRange(sheetName, rangeAddress, options = {}) {
  const _t0 = perfNow();
  const result = await Excel.run((context) =>
    clearSignificanceForRangeInContext(context, sheetName, rangeAddress)
  );
  if (result._clearDetails) {
    options.perfLog?.("clearSignificanceForRange", {
      rangeAddress,
      ...result._clearDetails,
      totalMs: perfElapsed(_t0),
    });
  }
  return result;
}

/**
 * Groups an array of candidate objects by their sheetName property.
 * Returns a Map<sheetName, candidate[]> preserving insertion order.
 */
export function groupCandidatesBySheet(candidates) {
  const map = new Map();
  for (const c of candidates) {
    if (!map.has(c.sheetName)) map.set(c.sheetName, []);
    map.get(c.sheetName).push(c);
  }
  return map;
}

/**
 * Accumulates _clearDetails from a single table result into a running aggregate.
 */
export function accumulateClearDetails(aggregate, details) {
  if (!aggregate || !details) return;
  aggregate.bodyCellsRead += details.bodyCellsRead || 0;
  aggregate.bodyCellsChanged += details.bodyCellsChanged || 0;
  aggregate.bannerCellsRead += details.bannerCellsRead || 0;
  aggregate.bannerCellsChanged += details.bannerCellsChanged || 0;
  aggregate.bannerWriteCommands += details.bannerWriteCommands || 0;
}

/**
 * Staged 4-phase Clear pipeline for all tables on one sheet.
 *
 * Reduces sync count from 4×N (per-table serial) to 4 (whole-sheet batch) by
 * loading and writing all N tables across four single round-trips:
 *
 *   Phase 1 — sync 1: load source range values+text for every table.
 *   Phase 2 — sync 2: compute clear-target ranges in JS; load values+format+dims.
 *   Phase 3 — sync 3: queue body writes; load banner scan text for every table.
 *                     Single flush covers all body writes + all banner reads.
 *   Phase 4 — sync 4: diff banner texts, queue banner writes, final flush.
 *
 * Must be called inside an Excel.run callback (shared context, no nested run).
 *
 * @param {Excel.RequestContext} context
 * @param {Excel.Worksheet} worksheetRef  Worksheet proxy for this sheet.
 * @param {Array<{rangeAddress: string}>} candidates  Tables on this sheet.
 * @returns {Promise<{
 *   results: Array<{status, message, _clearDetails}>,
 *   sheetDiags: {syncPhases, sourceRangesLoaded, targetRangesLoaded, bannerRangesLoaded}
 * }>}
 */
export async function clearSignificanceForSheetBatched(context, worksheetRef, candidates) {
  const BANNER_UPPER_SCAN_LIMIT = 5;

  // ── Per-table record ──────────────────────────────────────────────────────
  const records = candidates.map((c) => ({
    rangeAddress: c.rangeAddress,
    status: "pending", // "pending" | "skipped" | "cleared"
    message: "",
    // Phase 1
    sourceRange: null,
    // Phase 2 inputs (computed in JS after sync 1)
    clearTargetRange: null,
    // Phase 2 outputs (loaded after sync 2)
    knownDims: null, // { rowIndex, columnIndex, columnCount }
    // Phase 3 — body
    bodyCellsRead: 0,
    bodyCellsChanged: 0,
    bodyHasValueChange: false,
    bodyHasFormatChange: false,
    nextValues: null,
    nextNumberFormats: null,
    // Phase 3 — banner scan setup
    bannerScanRange: null,
    bannerScanRowCount: 0,
    bannerScanColCount: 0,
    bannerStartColIndex: 0,
    bannerStartRowAbs: 0,
    // Phase 4 — banner results
    bannerCellsRead: 0,
    bannerCellsChanged: 0,
    bannerWriteCommands: 0,
  }));

  // ══════════════════════════════════════════════════════════════════════════
  // PHASE 1 — Load source range values+text for every table
  // ══════════════════════════════════════════════════════════════════════════
  for (const rec of records) {
    const sr = worksheetRef.getRange(rec.rangeAddress);
    sr.load(["values", "text"]);
    rec.sourceRange = sr;
  }
  await context.sync(); // Sync 1

  // ══════════════════════════════════════════════════════════════════════════
  // PHASE 2 — Compute clear-target ranges in JS; load values+format+dims
  // ══════════════════════════════════════════════════════════════════════════
  for (const rec of records) {
    if (rec.status !== "pending") continue;

    const selectedValues = rec.sourceRange.values;
    const selectedText = rec.sourceRange.text;

    if (
      !selectedValues ||
      selectedValues.length < 1 ||
      !selectedValues[0] ||
      selectedValues[0].length < 1
    ) {
      rec.status = "skipped";
      rec.message = "нет данных в диапазоне";
      continue;
    }

    const clearTarget = resolveClearTargetBodyRange({
      values: selectedValues,
      text: selectedText,
    });

    if (clearTarget.state === "blocked") {
      const codes =
        clearTarget.blockingReasons && clearTarget.blockingReasons.length > 0
          ? ` [${clearTarget.blockingReasons.join(", ")}]`
          : "";
      rec.status = "skipped";
      rec.message = `${clearTarget.blockingMessage}${codes}`;
      continue;
    }

    if (clearTarget.state === "empty") {
      rec.status = "skipped";
      rec.message = "нет данных после нормализации";
      continue;
    }

    let clearTargetRange;

    if (clearTarget.usesFullSelection) {
      clearTargetRange = rec.sourceRange;
    } else {
      clearTargetRange = rec.sourceRange
        .getCell(clearTarget.rowOffset, clearTarget.colOffset)
        .getResizedRange(clearTarget.rowCount - 1, clearTarget.colCount - 1);
    }

    rec.clearTargetRange = clearTargetRange;
    clearTargetRange.load(["values", "numberFormat", "rowIndex", "columnIndex", "columnCount"]);
  }
  await context.sync(); // Sync 2

  // ══════════════════════════════════════════════════════════════════════════
  // PHASE 3 — Body cleanup + queue writes + set up banner reads
  // ══════════════════════════════════════════════════════════════════════════
  for (const rec of records) {
    if (rec.status !== "pending") continue;

    const targetValues = rec.clearTargetRange.values;
    const targetNumberFormats = rec.clearTargetRange.numberFormat;

    rec.knownDims = {
      rowIndex: rec.clearTargetRange.rowIndex,
      columnIndex: rec.clearTargetRange.columnIndex,
      columnCount: rec.clearTargetRange.columnCount,
    };

    const nextValues = [];
    const nextNumberFormats = [];
    let bodyCellsChanged = 0;
    let bodyHasValueChange = false;
    let bodyHasFormatChange = false;

    for (let r = 0; r < targetValues.length; r++) {
      const valueRow = [];
      const formatRow = [];

      for (let c = 0; c < targetValues[r].length; c++) {
        const rawValue = targetValues[r][c];
        const currentFormat = targetNumberFormats[r][c];

        if (typeof rawValue === "number") {
          valueRow.push(rawValue);
          formatRow.push(currentFormat);
          continue;
        }

        const cleanedText = removeSignificanceMarkersFromText(rawValue);
        const resolved = resolveNumericOutput(cleanedText);

        if (resolved !== null) {
          if (resolved.value !== rawValue) bodyHasValueChange = true;
          if (resolved.format !== currentFormat) bodyHasFormatChange = true;
          if (resolved.value !== rawValue || resolved.format !== currentFormat) bodyCellsChanged++;
          valueRow.push(resolved.value);
          formatRow.push(resolved.format);
        } else {
          if (cleanedText !== rawValue) {
            bodyHasValueChange = true;
            bodyCellsChanged++;
          }
          valueRow.push(cleanedText);
          formatRow.push("@");
        }
      }

      nextValues.push(valueRow);
      nextNumberFormats.push(formatRow);
    }

    rec.bodyCellsRead = targetValues.length * (targetValues[0] ? targetValues[0].length : 0);
    rec.bodyCellsChanged = bodyCellsChanged;
    rec.bodyHasValueChange = bodyHasValueChange;
    rec.bodyHasFormatChange = bodyHasFormatChange;
    rec.nextValues = nextValues;
    rec.nextNumberFormats = nextNumberFormats;

    // Queue body writes (only when something actually changed).
    if (bodyHasFormatChange) rec.clearTargetRange.numberFormat = nextNumberFormats;
    if (bodyHasValueChange) rec.clearTargetRange.values = nextValues;
    rec.clearTargetRange.format.font.bold = false;
    rec.clearTargetRange.format.fill.clear();

    // Queue banner scan load (uses sheet row index, so no banner sync needed).
    const { rowIndex: targetStartRowIndex, columnIndex: targetStartColumnIndex, columnCount: targetColumnCount } =
      rec.knownDims;

    if (targetStartRowIndex > 0 && targetColumnCount >= 1) {
      const totalScanRowCount = Math.min(BANNER_UPPER_SCAN_LIMIT + 1, targetStartRowIndex);
      if (totalScanRowCount >= 1) {
        const bannerScanRange = worksheetRef.getRangeByIndexes(
          targetStartRowIndex - totalScanRowCount,
          targetStartColumnIndex,
          totalScanRowCount,
          targetColumnCount
        );
        bannerScanRange.load("text");
        rec.bannerScanRange = bannerScanRange;
        rec.bannerScanRowCount = totalScanRowCount;
        rec.bannerScanColCount = targetColumnCount;
        rec.bannerStartColIndex = targetStartColumnIndex;
        rec.bannerStartRowAbs = targetStartRowIndex - totalScanRowCount;
      }
    }

    // Probe the rows below this table for a generated footnote (flushed by sync 3).
    const probe = buildFootnoteProbe(worksheetRef, rec.knownDims, targetValues.length);
    probe.range.load("values");
    rec.footnoteProbeRange = probe.range;
    rec.footnoteFirstRowBelow = probe.firstRowBelowTable;
    rec.footnoteDataColStartOffset = probe.dataColStartOffset;
  }
  await context.sync(); // Sync 3 — flushes ALL body writes; reads ALL banner scan texts

  // ══════════════════════════════════════════════════════════════════════════
  // PHASE 4 — Diff banner texts; queue banner writes; final flush
  // ══════════════════════════════════════════════════════════════════════════
  const footnoteRemovalRows = [];
  for (const rec of records) {
    if (rec.status !== "pending") continue;

    rec.status = "cleared"; // body writes already flushed in sync 3

    // Generated footnote in the trailing area below this cleared table → schedule
    // its row for deletion (ordinary user notes are skipped, not removed).
    if (rec.footnoteProbeRange) {
      const removalRow = resolveFootnoteRemovalRow(
        rec.footnoteProbeRange.values,
        rec.footnoteFirstRowBelow,
        rec.footnoteDataColStartOffset
      );
      if (removalRow !== null) footnoteRemovalRows.push(removalRow);
    }

    if (!rec.bannerScanRange) continue;

    const bannerTexts = rec.bannerScanRange.text;
    const { bannerScanRowCount, bannerScanColCount, bannerStartColIndex, bannerStartRowAbs } = rec;

    rec.bannerCellsRead = bannerScanRowCount * bannerScanColCount;

    const writesByRow = new Map();
    let cellsChanged = 0;

    for (let r = 0; r < bannerScanRowCount; r++) {
      const row = bannerTexts[r] || [];
      for (let c = 0; c < bannerScanColCount; c++) {
        const cur = row[c] || "";
        const nxt = cur && getTrailingBannerMarker(cur) ? removeTrailingBannerMarker(cur) : cur;
        if (cur === nxt) continue;
        const absRow = bannerStartRowAbs + r;
        const absCol = bannerStartColIndex + c;
        if (!writesByRow.has(absRow)) writesByRow.set(absRow, []);
        writesByRow.get(absRow).push({ colIndex: absCol, text: nxt });
        cellsChanged++;
      }
    }

    rec.bannerCellsChanged = cellsChanged;

    if (cellsChanged > 0) {
      let writeCommands = 0;
      for (const [rowIndex, items] of writesByRow) {
        items.sort((a, b) => a.colIndex - b.colIndex);
        let i = 0;
        while (i < items.length) {
          let j = i;
          while (j + 1 < items.length && items[j + 1].colIndex === items[j].colIndex + 1) j++;
          const texts = items.slice(i, j + 1).map((x) => x.text);
          worksheetRef.getRangeByIndexes(rowIndex, items[i].colIndex, 1, j - i + 1).values = [texts];
          writeCommands++;
          i = j + 1;
        }
      }
      rec.bannerWriteCommands = writeCommands;
    }
  }
  await context.sync(); // Sync 4 — flushes ALL banner writes

  // ── Build result array and sheet-level diagnostics ───────────────────────
  let targetRangesLoaded = 0;
  let bannerRangesLoaded = 0;
  for (const rec of records) {
    if (rec.clearTargetRange !== null) targetRangesLoaded++;
    if (rec.bannerScanRange !== null) bannerRangesLoaded++;
  }

  const results = records.map((rec) => {
    if (rec.status === "cleared") {
      return {
        status: "cleared",
        message: "очищено",
        _clearDetails: {
          bodyCellsRead: rec.bodyCellsRead,
          bodyCellsChanged: rec.bodyCellsChanged,
          bannerCellsRead: rec.bannerCellsRead,
          bannerCellsChanged: rec.bannerCellsChanged,
          bannerWriteCommands: rec.bannerWriteCommands,
          totalMs: 0,
        },
      };
    }
    return { status: rec.status, message: rec.message, _clearDetails: null };
  });

  return {
    results,
    footnoteRemovalRows,
    sheetDiags: {
      syncPhases: 4,
      sourceRangesLoaded: candidates.length,
      targetRangesLoaded,
      bannerRangesLoaded,
    },
  };
}

/**
 * Workbook-level staged 4-phase Clear pipeline for all eligible candidates.
 *
 * Generalises clearSignificanceForSheetBatched to span multiple worksheets
 * inside a single Excel.run context.  A single Office.js RequestContext can
 * queue loads and writes against ranges on any number of worksheets, so the
 * same 4-phase approach that reduces per-sheet sync count to 4 can be applied
 * across the whole workbook — reducing total syncs from 4×N_sheets to 4.
 *
 *   Phase 1 — sync 1: load values+text for ALL source ranges (all sheets).
 *   Phase 2 — sync 2: compute clear targets in JS; load values+format+dims.
 *   Phase 3 — sync 3: queue ALL body writes; load ALL banner scan texts.
 *   Phase 4 — sync 4: queue ALL banner writes; final flush.
 *
 * Must be called inside an Excel.run callback (shared context, no nested run).
 *
 * @param {Excel.RequestContext} context
 * @param {Array<{sheetName: string, rangeAddress: string}>} eligible
 *   All candidates across all sheets (each carries its own sheetName).
 * @returns {Promise<{
 *   results: Array<{status, message, _clearDetails}>,
 *   batchDiags: {syncPhases, sourceRangesLoaded, targetRangesLoaded, bannerRangesLoaded}
 * }>}
 */
export async function clearSignificanceForWorkbookBatched(context, eligible) {
  const BANNER_UPPER_SCAN_LIMIT = 5;

  // One worksheet proxy per unique sheet — avoids redundant getItem calls.
  const worksheetRefs = new Map();
  for (const c of eligible) {
    if (!worksheetRefs.has(c.sheetName)) {
      worksheetRefs.set(c.sheetName, context.workbook.worksheets.getItem(c.sheetName));
    }
  }

  // ── Per-table record ──────────────────────────────────────────────────────
  const records = eligible.map((c) => ({
    sheetName: c.sheetName,
    rangeAddress: c.rangeAddress,
    status: "pending", // "pending" | "skipped" | "cleared"
    message: "",
    // Phase 1
    sourceRange: null,
    // Phase 2 inputs (computed in JS after sync 1)
    clearTargetRange: null,
    // Phase 2 outputs (loaded after sync 2)
    knownDims: null, // { rowIndex, columnIndex, columnCount }
    // Phase 3 — body
    bodyCellsRead: 0,
    bodyCellsChanged: 0,
    bodyHasValueChange: false,
    bodyHasFormatChange: false,
    nextValues: null,
    nextNumberFormats: null,
    // Phase 3 — banner scan setup
    bannerScanRange: null,
    bannerScanRowCount: 0,
    bannerScanColCount: 0,
    bannerStartColIndex: 0,
    bannerStartRowAbs: 0,
    // Phase 4 — banner results
    bannerCellsRead: 0,
    bannerCellsChanged: 0,
    bannerWriteCommands: 0,
  }));

  // ══════════════════════════════════════════════════════════════════════════
  // PHASE 1 — Load source range values+text for every table (all sheets)
  // ══════════════════════════════════════════════════════════════════════════
  for (const rec of records) {
    const sr = worksheetRefs.get(rec.sheetName).getRange(rec.rangeAddress);
    sr.load(["values", "text"]);
    rec.sourceRange = sr;
  }
  await context.sync(); // Sync 1

  // ══════════════════════════════════════════════════════════════════════════
  // PHASE 2 — Compute clear-target ranges in JS; load values+format+dims
  // ══════════════════════════════════════════════════════════════════════════
  for (const rec of records) {
    if (rec.status !== "pending") continue;

    const selectedValues = rec.sourceRange.values;
    const selectedText = rec.sourceRange.text;

    if (
      !selectedValues ||
      selectedValues.length < 1 ||
      !selectedValues[0] ||
      selectedValues[0].length < 1
    ) {
      rec.status = "skipped";
      rec.message = "нет данных в диапазоне";
      continue;
    }

    const clearTarget = resolveClearTargetBodyRange({
      values: selectedValues,
      text: selectedText,
    });

    if (clearTarget.state === "blocked") {
      const codes =
        clearTarget.blockingReasons && clearTarget.blockingReasons.length > 0
          ? ` [${clearTarget.blockingReasons.join(", ")}]`
          : "";
      rec.status = "skipped";
      rec.message = `${clearTarget.blockingMessage}${codes}`;
      continue;
    }

    if (clearTarget.state === "empty") {
      rec.status = "skipped";
      rec.message = "нет данных после нормализации";
      continue;
    }

    let clearTargetRange;

    if (clearTarget.usesFullSelection) {
      clearTargetRange = rec.sourceRange;
    } else {
      clearTargetRange = rec.sourceRange
        .getCell(clearTarget.rowOffset, clearTarget.colOffset)
        .getResizedRange(clearTarget.rowCount - 1, clearTarget.colCount - 1);
    }

    rec.clearTargetRange = clearTargetRange;
    clearTargetRange.load(["values", "numberFormat", "rowIndex", "columnIndex", "columnCount"]);
  }
  await context.sync(); // Sync 2

  // ══════════════════════════════════════════════════════════════════════════
  // PHASE 3 — Body cleanup + queue writes + set up banner reads (all sheets)
  // ══════════════════════════════════════════════════════════════════════════
  for (const rec of records) {
    if (rec.status !== "pending") continue;

    const targetValues = rec.clearTargetRange.values;
    const targetNumberFormats = rec.clearTargetRange.numberFormat;

    rec.knownDims = {
      rowIndex: rec.clearTargetRange.rowIndex,
      columnIndex: rec.clearTargetRange.columnIndex,
      columnCount: rec.clearTargetRange.columnCount,
    };

    const nextValues = [];
    const nextNumberFormats = [];
    let bodyCellsChanged = 0;
    let bodyHasValueChange = false;
    let bodyHasFormatChange = false;

    for (let r = 0; r < targetValues.length; r++) {
      const valueRow = [];
      const formatRow = [];

      for (let c = 0; c < targetValues[r].length; c++) {
        const rawValue = targetValues[r][c];
        const currentFormat = targetNumberFormats[r][c];

        if (typeof rawValue === "number") {
          valueRow.push(rawValue);
          formatRow.push(currentFormat);
          continue;
        }

        const cleanedText = removeSignificanceMarkersFromText(rawValue);
        const resolved = resolveNumericOutput(cleanedText);

        if (resolved !== null) {
          if (resolved.value !== rawValue) bodyHasValueChange = true;
          if (resolved.format !== currentFormat) bodyHasFormatChange = true;
          if (resolved.value !== rawValue || resolved.format !== currentFormat) bodyCellsChanged++;
          valueRow.push(resolved.value);
          formatRow.push(resolved.format);
        } else {
          if (cleanedText !== rawValue) {
            bodyHasValueChange = true;
            bodyCellsChanged++;
          }
          valueRow.push(cleanedText);
          formatRow.push("@");
        }
      }

      nextValues.push(valueRow);
      nextNumberFormats.push(formatRow);
    }

    rec.bodyCellsRead = targetValues.length * (targetValues[0] ? targetValues[0].length : 0);
    rec.bodyCellsChanged = bodyCellsChanged;
    rec.bodyHasValueChange = bodyHasValueChange;
    rec.bodyHasFormatChange = bodyHasFormatChange;
    rec.nextValues = nextValues;
    rec.nextNumberFormats = nextNumberFormats;

    // Queue body writes (only when something actually changed).
    if (bodyHasFormatChange) rec.clearTargetRange.numberFormat = nextNumberFormats;
    if (bodyHasValueChange) rec.clearTargetRange.values = nextValues;
    rec.clearTargetRange.format.font.bold = false;
    rec.clearTargetRange.format.fill.clear();

    // Queue banner scan load for this table's sheet.
    const wsRef = worksheetRefs.get(rec.sheetName);
    const { rowIndex: targetStartRowIndex, columnIndex: targetStartColumnIndex, columnCount: targetColumnCount } =
      rec.knownDims;

    if (targetStartRowIndex > 0 && targetColumnCount >= 1) {
      const totalScanRowCount = Math.min(BANNER_UPPER_SCAN_LIMIT + 1, targetStartRowIndex);
      if (totalScanRowCount >= 1) {
        const bannerScanRange = wsRef.getRangeByIndexes(
          targetStartRowIndex - totalScanRowCount,
          targetStartColumnIndex,
          totalScanRowCount,
          targetColumnCount
        );
        bannerScanRange.load("text");
        rec.bannerScanRange = bannerScanRange;
        rec.bannerScanRowCount = totalScanRowCount;
        rec.bannerScanColCount = targetColumnCount;
        rec.bannerStartColIndex = targetStartColumnIndex;
        rec.bannerStartRowAbs = targetStartRowIndex - totalScanRowCount;
      }
    }

    // Probe the rows below this table for a generated footnote (flushed by sync 3).
    const probe = buildFootnoteProbe(wsRef, rec.knownDims, targetValues.length);
    probe.range.load("values");
    rec.footnoteProbeRange = probe.range;
    rec.footnoteFirstRowBelow = probe.firstRowBelowTable;
    rec.footnoteDataColStartOffset = probe.dataColStartOffset;
  }
  await context.sync(); // Sync 3 — flushes ALL body writes; reads ALL banner scan texts

  // ══════════════════════════════════════════════════════════════════════════
  // PHASE 4 — Diff banner texts; queue banner writes; final flush (all sheets)
  // ══════════════════════════════════════════════════════════════════════════
  const footnoteRemovalJobs = [];
  for (const rec of records) {
    if (rec.status !== "pending") continue;

    rec.status = "cleared"; // body writes already flushed in sync 3

    // Generated footnote in the trailing area below this cleared table → schedule
    // its row for deletion (ordinary user notes are skipped, not removed).
    if (rec.footnoteProbeRange) {
      const removalRow = resolveFootnoteRemovalRow(
        rec.footnoteProbeRange.values,
        rec.footnoteFirstRowBelow,
        rec.footnoteDataColStartOffset
      );
      if (removalRow !== null) {
        footnoteRemovalJobs.push({ sheetName: rec.sheetName, footnoteRowIndex: removalRow });
      }
    }

    if (!rec.bannerScanRange) continue;

    const bannerTexts = rec.bannerScanRange.text;
    const { bannerScanRowCount, bannerScanColCount, bannerStartColIndex, bannerStartRowAbs } = rec;

    rec.bannerCellsRead = bannerScanRowCount * bannerScanColCount;

    const writesByRow = new Map();
    let cellsChanged = 0;

    for (let r = 0; r < bannerScanRowCount; r++) {
      const row = bannerTexts[r] || [];
      for (let c = 0; c < bannerScanColCount; c++) {
        const cur = row[c] || "";
        const nxt = cur && getTrailingBannerMarker(cur) ? removeTrailingBannerMarker(cur) : cur;
        if (cur === nxt) continue;
        const absRow = bannerStartRowAbs + r;
        const absCol = bannerStartColIndex + c;
        if (!writesByRow.has(absRow)) writesByRow.set(absRow, []);
        writesByRow.get(absRow).push({ colIndex: absCol, text: nxt });
        cellsChanged++;
      }
    }

    rec.bannerCellsChanged = cellsChanged;

    if (cellsChanged > 0) {
      const wsRef = worksheetRefs.get(rec.sheetName);
      let writeCommands = 0;
      for (const [rowIndex, items] of writesByRow) {
        items.sort((a, b) => a.colIndex - b.colIndex);
        let i = 0;
        while (i < items.length) {
          let j = i;
          while (j + 1 < items.length && items[j + 1].colIndex === items[j].colIndex + 1) j++;
          const texts = items.slice(i, j + 1).map((x) => x.text);
          wsRef.getRangeByIndexes(rowIndex, items[i].colIndex, 1, j - i + 1).values = [texts];
          writeCommands++;
          i = j + 1;
        }
      }
      rec.bannerWriteCommands = writeCommands;
    }
  }
  await context.sync(); // Sync 4 — flushes ALL banner writes (all sheets)

  // ── Build result array and batch-level diagnostics ────────────────────────
  let targetRangesLoaded = 0;
  let bannerRangesLoaded = 0;
  for (const rec of records) {
    if (rec.clearTargetRange !== null) targetRangesLoaded++;
    if (rec.bannerScanRange !== null) bannerRangesLoaded++;
  }

  const results = records.map((rec) => {
    if (rec.status === "cleared") {
      return {
        status: "cleared",
        message: "очищено",
        _clearDetails: {
          bodyCellsRead: rec.bodyCellsRead,
          bodyCellsChanged: rec.bodyCellsChanged,
          bannerCellsRead: rec.bannerCellsRead,
          bannerCellsChanged: rec.bannerCellsChanged,
          bannerWriteCommands: rec.bannerWriteCommands,
          totalMs: 0,
        },
      };
    }
    return { status: rec.status, message: rec.message, _clearDetails: null };
  });

  return {
    results,
    footnoteRemovalJobs,
    batchDiags: {
      syncPhases: 4,
      sourceRangesLoaded: eligible.length,
      targetRangesLoaded,
      bannerRangesLoaded,
    },
  };
}

/**
 * Removes significance markers and formatting from the selected range.
 *
 * PURPOSE:
 * User-facing cleanup button.
 */
export async function clearSignificanceFromSelection(options) {
  const { setStatusMessage, runningStatusMessage, nonContiguousSelectionMessage, t, perfLog } = options;
  const _t0 = perfNow();
  setStatusMessage(runningStatusMessage("clear"));
  await Excel.run(async (context) => {
    // getSelectedRange() throws a RichApi.Error for non-contiguous (Ctrl+Click
    // multi-area) selections. Catch that error here and surface a user-facing
    // message rather than letting the runtime error propagate.
    let selectedRange;
    try {
      selectedRange = context.workbook.getSelectedRange();
      // Read-only load: needed to decide whether to operate on the whole
      // selection (strict numeric case) or only on the detected data body
      // (forgiving full-table case). No writes happen before the target is known.
      selectedRange.load(["values", "text"]);
      await context.sync();
    } catch (_selectionErr) {
      setStatusMessage(nonContiguousSelectionMessage());
      return;
    }

    const selectedValues = selectedRange.values;
    const selectedText = selectedRange.text;

    if (
      !selectedValues ||
      selectedValues.length < 1 ||
      !selectedValues[0] ||
      selectedValues[0].length < 1
    ) {
      setStatusMessage("Нет данных в выделенном диапазоне.");
      return;
    }

    const clearTarget = resolveClearTargetBodyRange({
      values: selectedValues,
      text: selectedText,
    });

    // State 3: broad/full-table-like selection but decomposition failed.
    // Block and return without mutating anything.
    if (clearTarget.state === "blocked") {
      const codes =
        clearTarget.blockingReasons && clearTarget.blockingReasons.length > 0
          ? ` [${clearTarget.blockingReasons.join(", ")}]`
          : "";
      setStatusMessage(`${clearTarget.blockingMessage}${codes}`);
      return;
    }

    if (clearTarget.state === "empty") {
      setStatusMessage("Нет данных в выделенном диапазоне.");
      return;
    }

    let clearTargetRange;

    if (clearTarget.usesFullSelection) {
      clearTargetRange = selectedRange;
    } else {
      clearTargetRange = selectedRange
        .getCell(clearTarget.rowOffset, clearTarget.colOffset)
        .getResizedRange(clearTarget.rowCount - 1, clearTarget.colCount - 1);
    }

    clearTargetRange.load(["values", "numberFormat", "rowIndex", "columnIndex", "columnCount"]);

    await context.sync();

    const targetValues = clearTargetRange.values;
    const targetNumberFormats = clearTargetRange.numberFormat;
    const _knownDims = {
      rowIndex: clearTargetRange.rowIndex,
      columnIndex: clearTargetRange.columnIndex,
      columnCount: clearTargetRange.columnCount,
    };

    const nextValues = [];
    const nextNumberFormats = [];
    let bodyCellsChanged = 0;
    let bodyHasValueChange = false;
    let bodyHasFormatChange = false;

    for (let rowIndex = 0; rowIndex < targetValues.length; rowIndex++) {
      const valueRow = [];
      const formatRow = [];

      for (let columnIndex = 0; columnIndex < targetValues[rowIndex].length; columnIndex++) {
        const rawValue = targetValues[rowIndex][columnIndex];
        const currentFormat = targetNumberFormats[rowIndex][columnIndex];

        if (typeof rawValue === "number") {
          valueRow.push(rawValue);
          formatRow.push(currentFormat);
          continue;
        }

        const cleanedText = removeSignificanceMarkersFromText(rawValue);
        const resolved = resolveNumericOutput(cleanedText);

        if (resolved !== null) {
          if (resolved.value !== rawValue) bodyHasValueChange = true;
          if (resolved.format !== currentFormat) bodyHasFormatChange = true;
          if (resolved.value !== rawValue || resolved.format !== currentFormat) bodyCellsChanged++;
          valueRow.push(resolved.value);
          formatRow.push(resolved.format);
        } else {
          if (cleanedText !== rawValue) { bodyHasValueChange = true; bodyCellsChanged++; }
          valueRow.push(cleanedText);
          formatRow.push("@");
        }
      }

      nextValues.push(valueRow);
      nextNumberFormats.push(formatRow);
    }

    // Only write body data when something actually changed, avoiding a full
    // matrix write on a re-clear of an already-clean workbook.
    if (bodyHasFormatChange) clearTargetRange.numberFormat = nextNumberFormats;
    if (bodyHasValueChange) clearTargetRange.values = nextValues;

    clearTargetRange.format.font.bold = false;
    clearTargetRange.format.fill.clear();

    // Probe the row immediately below the cleared table for a generated footnote
    // (queued before the banner read so its sync flushes it — no extra round-trip).
    const { range: footnoteProbeRange, firstRowBelowTable, dataColStartOffset } = buildFootnoteProbe(
      clearTargetRange.worksheet,
      _knownDims,
      targetValues.length
    );
    footnoteProbeRange.load("values");

    // No body sync here — deferred to the banner read sync inside
    // clearBannerMarkerUpdatesForRange, which flushes all queued writes.
    const bannerDetails = await clearBannerMarkerUpdatesForRange(context, clearTargetRange, _knownDims);

    await context.sync();

    // Silently delete the generated footnote row in the trailing area below the
    // cleared table. Single table → no bottom-to-top ordering concern. Only a row
    // holding the marker is deleted; ordinary user notes (even above the footnote)
    // are skipped and preserved.
    const footnoteRemovalRow = resolveFootnoteRemovalRow(
      footnoteProbeRange.values,
      firstRowBelowTable,
      dataColStartOffset
    );
    if (footnoteRemovalRow !== null) {
      clearTargetRange.worksheet
        .getRangeByIndexes(footnoteRemovalRow, 0, 1, 1)
        .getEntireRow()
        .delete(Excel.DeleteShiftDirection.up);
      await context.sync();
    }

    perfLog("clearSignificanceFromSelection", {
      bodyCellsRead: targetValues.length * (targetValues[0] ? targetValues[0].length : 0),
      bodyCellsChanged,
      bodyHasValueChange,
      bodyHasFormatChange,
      bannerDetails: bannerDetails || null,
      totalMs: perfElapsed(_t0),
    });
    setStatusMessage(t("status.clearDone"));
  });
}

/**
 * Combined clear-only banner pass for Clear Significance.
 *
 * Clear-only analog of applyBannerMarkerUpdatesForRange — reads the banner scan
 * area once (flushing any pending body writes queued by the caller), strips all
 * RIT trailing markers in-memory, diffs against the original texts, and queues
 * writes only for cells that actually changed.  No marker placement, no
 * numberFormat writes.
 *
 * Sync behaviour: issues 1 context.sync() to read the banner area (and flush
 * pending writes).  Queued banner writes are NOT flushed here — the caller must
 * issue its own final context.sync().  When cellsChanged === 0, no writes are
 * queued and the caller's final sync is a cheap no-op.
 *
 * Returns null when targetStartRowIndex === 0 or targetColumnCount < 1.
 * Otherwise returns a compact diagnostics object.
 *
 * @param {Excel.RequestContext} context
 * @param {Excel.Range} targetRange  The clear-target range (data body).
 * @param {{ rowIndex, columnIndex, columnCount }} [knownDimensions]
 *   Pre-loaded sheet coordinates of targetRange.  When provided, skips the
 *   extra load+sync that would otherwise be needed to resolve them.
 */
async function clearBannerMarkerUpdatesForRange(context, targetRange, knownDimensions) {
  const BANNER_UPPER_SCAN_LIMIT = 5;
  const BANNER_SCAN_ROW_COUNT = BANNER_UPPER_SCAN_LIMIT + 1;

  let targetStartRowIndex, targetStartColumnIndex, targetColumnCount;
  if (knownDimensions) {
    targetStartRowIndex = knownDimensions.rowIndex;
    targetStartColumnIndex = knownDimensions.columnIndex;
    targetColumnCount = knownDimensions.columnCount;
  } else {
    targetRange.load(["rowIndex", "columnIndex", "columnCount"]);
    await context.sync();
    targetStartRowIndex = targetRange.rowIndex;
    targetStartColumnIndex = targetRange.columnIndex;
    targetColumnCount = targetRange.columnCount;
  }

  if (targetStartRowIndex === 0 || targetColumnCount < 1) {
    return null;
  }

  const totalScanRowCount = Math.min(BANNER_SCAN_ROW_COUNT, targetStartRowIndex);
  if (totalScanRowCount < 1) {
    return null;
  }

  const scanStartColIndex = targetStartColumnIndex;
  const scanColCount = targetColumnCount;

  const bannerScanRange = targetRange.worksheet.getRangeByIndexes(
    targetStartRowIndex - totalScanRowCount,
    scanStartColIndex,
    totalScanRowCount,
    scanColCount
  );

  const readSyncStartMs = perfNow();
  bannerScanRange.load("text");
  // This sync also flushes any pending body writes queued by the caller.
  await context.sync();
  const readSyncEndMs = perfNow();
  let syncCount = 1;

  const bannerTexts = bannerScanRange.text;

  // Strip RIT trailing markers from every cell.  Keeps non-RIT parenthesised
  // text (e.g. "Wave (quarter)") intact — getTrailingBannerMarker only matches
  // single-character significance labels.
  const desiredTexts = new Array(totalScanRowCount);
  for (let r = 0; r < totalScanRowCount; r++) {
    const row = bannerTexts[r] || [];
    const destRow = new Array(scanColCount);
    for (let c = 0; c < scanColCount; c++) {
      const txt = row[c] || "";
      destRow[c] = (txt && getTrailingBannerMarker(txt)) ? removeTrailingBannerMarker(txt) : txt;
    }
    desiredTexts[r] = destRow;
  }

  // Diff against original texts and queue only changed cells.
  const writesByRow = new Map();
  let cellsChanged = 0;

  for (let r = 0; r < totalScanRowCount; r++) {
    for (let c = 0; c < scanColCount; c++) {
      const cur = (bannerTexts[r] && bannerTexts[r][c]) || "";
      const nxt = desiredTexts[r][c] || "";
      if (cur === nxt) continue;
      const absRow = targetStartRowIndex - totalScanRowCount + r;
      const absCol = scanStartColIndex + c;
      if (!writesByRow.has(absRow)) writesByRow.set(absRow, []);
      writesByRow.get(absRow).push({ colIndex: absCol, text: nxt });
      cellsChanged++;
    }
  }

  const changedRows = writesByRow.size;
  const queueWriteStartMs = perfNow();
  let writeCommands = 0;

  if (cellsChanged > 0) {
    const worksheet = targetRange.worksheet;
    for (const [rowIndex, items] of writesByRow) {
      items.sort((a, b) => a.colIndex - b.colIndex);
      let i = 0;
      while (i < items.length) {
        let j = i;
        while (j + 1 < items.length && items[j + 1].colIndex === items[j].colIndex + 1) j++;
        const texts = items.slice(i, j + 1).map((x) => x.text);
        worksheet.getRangeByIndexes(rowIndex, items[i].colIndex, 1, j - i + 1).values = [texts];
        writeCommands++;
        i = j + 1;
      }
    }
    // Intentionally no context.sync() here — caller's final sync flushes queued writes.
  }

  const queueWriteEndMs = perfNow();

  return {
    rowsScanned: totalScanRowCount,
    cellsRead: totalScanRowCount * scanColCount,
    cellsChanged,
    writeCommands,
    changedRows,
    readSyncMs: readSyncEndMs - readSyncStartMs,
    planMs: queueWriteStartMs - readSyncEndMs,
    queueWriteMs: queueWriteEndMs - queueWriteStartMs,
    syncCount,
  };
}

// ─── Significance footnote removal (Clear) ──────────────────────────────────────
//
// Clear Significance silently removes a RIT-generated footnote row directly below
// a cleared table. Removal is unconditional (no separate setting) but strictly
// guarded: only a row whose detected cell starts with SIGNIFICANCE_FOOTNOTE_MARKER
// is deleted, so ordinary user comments/notes under a table are never touched.
//
// Detection scans the same bounded trailing area below the cleared data body as
// the insert/update model (FOOTNOTE_SCAN_WINDOW_ROWS rows, a few columns to the
// LEFT as well so a left-offset marker is seen), skipping ordinary user note
// rows and stopping at a blank / next-table boundary. Only a row holding the
// SIGNIFICANCE_FOOTNOTE_MARKER is removed, so ordinary user notes are preserved
// even when the generated footnote sits below them. Detection piggybacks on
// existing Clear syncs; deletions are applied afterwards, bottom-to-top per
// worksheet, so they never shift ranges still being cleared.

/**
 * Builds the footnote probe range (a bounded window of rows below the cleared
 * table) plus the metadata the pure removal resolver needs.
 *
 * @param worksheet         Excel.Worksheet proxy.
 * @param knownDims         { rowIndex, columnIndex, columnCount } of the cleared body.
 * @param dataBodyRowCount  row count of the cleared body.
 * @returns { range, firstRowBelowTable, dataColStartOffset }
 */
function buildFootnoteProbe(worksheet, knownDims, dataBodyRowCount) {
  const firstRowBelowTable = knownDims.rowIndex + dataBodyRowCount; // bottom row + 1
  const startCol = Math.max(0, knownDims.columnIndex - LABEL_SCAN_COLUMNS_LEFT);
  const colCount = knownDims.columnIndex + knownDims.columnCount - startCol;
  const dataColStartOffset = knownDims.columnIndex - startCol;
  const range = worksheet.getRangeByIndexes(
    firstRowBelowTable,
    startCol,
    FOOTNOTE_SCAN_WINDOW_ROWS,
    Math.max(1, colCount)
  );
  return { range, firstRowBelowTable, dataColStartOffset };
}

/**
 * Deletes generated footnote rows for cleared tables. Jobs are grouped per sheet
 * and deleted bottom-to-top (descending row) so each deletion never shifts the
 * index of a not-yet-deleted footnote above it. Identical rows are de-duplicated.
 * Only ever called with rows already confirmed to start with the footnote marker.
 */
async function applySignificanceFootnoteRemovalJobs(context, jobs) {
  if (!Array.isArray(jobs) || jobs.length === 0) return;

  const jobsBySheet = new Map();
  for (const job of jobs) {
    if (!job || !job.sheetName || !Number.isFinite(job.footnoteRowIndex)) continue;
    if (!jobsBySheet.has(job.sheetName)) jobsBySheet.set(job.sheetName, []);
    jobsBySheet.get(job.sheetName).push(job);
  }

  let queuedAnyDeletion = false;
  for (const [sheetName, sheetJobs] of jobsBySheet) {
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const seenRows = new Set();
    const sortedRows = sheetJobs
      .map((j) => j.footnoteRowIndex)
      .sort((a, b) => b - a);
    for (const rowIndex of sortedRows) {
      if (seenRows.has(rowIndex)) continue;
      seenRows.add(rowIndex);
      worksheet
        .getRangeByIndexes(rowIndex, 0, 1, 1)
        .getEntireRow()
        .delete(Excel.DeleteShiftDirection.up);
      queuedAnyDeletion = true;
    }
  }

  if (queuedAnyDeletion) await context.sync();
}

/**
 * Applies footnote removal jobs inside a dedicated Excel.run. Safe to call after
 * the Clear pipeline has committed — row indices stay valid because clearing never
 * inserts or deletes rows.
 */
export async function applySignificanceFootnoteRemovalJobsInOwnContext(jobs) {
  if (!Array.isArray(jobs) || jobs.length === 0) return;
  try {
    await Excel.run(async (context) => {
      await applySignificanceFootnoteRemovalJobs(context, jobs);
    });
  } catch (footnoteErr) {
    console.warn("RIT: не удалось удалить подписи под таблицами.", footnoteErr);
  }
}

/**
 * Removes trailing banner marker.
 *
 * Used when a column should no longer have a banner marker,
 * for example because Total is excluded.
 */
function removeTrailingBannerMarker(rawText) {
  if (rawText === null || rawText === undefined) {
    return "";
  }

  const text = String(rawText);
  const currentMarker = getTrailingBannerMarker(text);

  if (!currentMarker) {
    return text.trim();
  }

  return text.slice(0, currentMarker.start).trim();
}

function getTrailingBannerMarker(rawText) {
  const text = rawText === null || rawText === undefined ? "" : String(rawText);

  // Require the marker token to be preceded by whitespace or appear at the
  // very start of the cell.  This prevents parenthesised fragments inside
  // words — e.g. "сам(а)" — from being mistaken for RIT significance markers
  // even when the single letter inside happens to be a valid label (Cyrillic
  // "а" is the first Cyrillic entry in generateSignificanceLabels()).
  const markerMatch = text.match(/(^|\s)\(([^()]*)\)\s*$/);

  if (!markerMatch) {
    return null;
  }

  const markerLabel = markerMatch[2]; // group 2: label inside parens

  // Recognise any token RIT may have written as a marker — single characters
  // from any historical alphabet, or multi-character overflow markers — so
  // banner markers are cleaned up on re-runs regardless of the current
  // Cyrillic/overflow settings.
  if (!isSignificanceMarkerLabel(markerLabel)) {
    return null;
  }

  return {
    label: markerLabel,
    start: markerMatch.index,
  };
}
