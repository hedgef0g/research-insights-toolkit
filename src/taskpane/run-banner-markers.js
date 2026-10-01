import {
  generateSignificanceLabels,
  buildBannerLocalSignificanceLabelMap,
  isSignificanceMarkerLabel,
} from "../core/significance";

import {
  BANNER_SCAN_AREA_STATS,
  perfNow,
  perfBannerWriteProfileEnabled,
  roundBannerDiagnosticRatio,
} from "./taskpane-performance";

/**
 * Combined banner-marker clear + write for a Run pass.
 *
 * Uses a single read/plan/write phase.  Reads the full banner scan area
 * once, strips all RIT trailing markers in-memory (the clear phase), then
 * applies the desired markers (lower banner row when non-empty, otherwise
 * nearest non-empty cell above; uses `bannerStructure` labels when respect-
 * banner-structure is on, otherwise sequential significance labels with the
 * first-column-is-total adjustment).  Diffs the desired matrix against the
 * read texts and writes back only cells that actually changed, batched into
 * contiguous same-row runs and split by markered/clear-only flag so the
 * "@" number-format guard is preserved on marker writes.
 *
 * Sync count in normal mode: 1 for the banner-area read (also flushes any
 * pending data writes from the caller) and 0 for the writes — the caller is
 * expected to issue its own final `context.sync()` to flush queued banner
 * writes.
 *
 * Development-only banner write profiling can split numberFormat and values
 * into separate syncs to measure host flush cost. It is disabled by default.
 *
 * Returns null when no banner work was needed.  Otherwise returns a perf
 * details object describing the work performed.
 */
export async function applyBannerMarkerUpdatesForRange(
  context,
  writeTargetRange,
  calculationSettings,
  bannerStructure,
  knownDimensions,
  sourceStartColIndex
) {
  if (!calculationSettings.writeBannerLetters) {
    return null;
  }

  const BANNER_UPPER_SCAN_LIMIT = 5;
  const BANNER_SCAN_ROW_COUNT = BANNER_UPPER_SCAN_LIMIT + 1;

  let targetStartRowIndex, targetStartColumnIndex, targetColumnCount;
  if (knownDimensions) {
    targetStartRowIndex = knownDimensions.rowIndex;
    targetStartColumnIndex = knownDimensions.columnIndex;
    targetColumnCount = knownDimensions.columnCount;
  } else {
    writeTargetRange.load(["rowIndex", "columnIndex", "columnCount"]);
    await context.sync();
    targetStartRowIndex = writeTargetRange.rowIndex;
    targetStartColumnIndex = writeTargetRange.columnIndex;
    targetColumnCount = writeTargetRange.columnCount;
  }

  if (targetStartRowIndex === 0 || targetColumnCount < 1) {
    return null;
  }

  const totalScanRowCount = Math.min(BANNER_SCAN_ROW_COUNT, targetStartRowIndex);
  if (totalScanRowCount < 1) {
    return null;
  }

  // Widen the scan left to absorb stale-left marker cleanup when the source
  // selection extends further left than the actual write target (#272).
  const scanStartColIndex =
    sourceStartColIndex !== undefined && sourceStartColIndex < targetStartColumnIndex
      ? sourceStartColIndex
      : targetStartColumnIndex;
  const scanColCount = targetStartColumnIndex + targetColumnCount - scanStartColIndex;

  if (scanColCount < 1) {
    return null;
  }

  const bannerScanRange = writeTargetRange.worksheet.getRangeByIndexes(
    targetStartRowIndex - totalScanRowCount,
    scanStartColIndex,
    totalScanRowCount,
    scanColCount
  );
  const readSyncStartMs = perfNow();
  bannerScanRange.load("text");
  await context.sync();
  const readSyncEndMs = perfNow();
  let syncCount = 1;

  const bannerTexts = bannerScanRange.text;

  // Desired-text matrix.  Initialised from current texts, then mutated in
  // place by step 1 (strip RIT markers) and step 3 (place markers).
  const desiredTexts = new Array(totalScanRowCount);
  for (let r = 0; r < totalScanRowCount; r++) {
    const row = bannerTexts[r] || [];
    const destRow = new Array(scanColCount);
    for (let c = 0; c < scanColCount; c++) {
      destRow[c] = row[c] || "";
    }
    desiredTexts[r] = destRow;
  }

  // Step 1: strip RIT trailing markers from every cell in the scan area.
  // Keeps non-RIT parenthesised text (e.g. "Wave (quarter)") intact because
  // getTrailingBannerMarker only matches single-character significance labels.
  for (let r = 0; r < totalScanRowCount; r++) {
    const row = desiredTexts[r];
    for (let c = 0; c < scanColCount; c++) {
      const t = row[c];
      if (t && getTrailingBannerMarker(t)) {
        row[c] = removeTrailingBannerMarker(t);
      }
    }
  }

  // Step 2: compute the label for each data column.
  const useStructure = !!(calculationSettings.respectBannerStructure && bannerStructure);
  let labelMap;
  if (useStructure) {
    labelMap = buildBannerLocalSignificanceLabelMap(bannerStructure, calculationSettings);
  } else {
    const significanceLabels = generateSignificanceLabels({
      useCyrillicMarkers: Boolean(calculationSettings.useCyrillicMarkers),
      allowMultiCharacterMarkers: Boolean(calculationSettings.allowMultiCharacterMarkers),
      minimumCount: targetColumnCount,
    });
    labelMap = new Map();
    for (let dataCol = 0; dataCol < targetColumnCount; dataCol++) {
      if (calculationSettings.firstColumnIsTotal && dataCol === 0) {
        // Total column 0: step 1 above already stripped any RIT marker from
        // the lower row and the upper rows, so no additional label work is
        // required for this column.
        continue;
      }
      const markerIndex = calculationSettings.firstColumnIsTotal ? dataCol - 1 : dataCol;
      const marker = significanceLabels[markerIndex];
      if (marker) {
        labelMap.set(dataCol, marker);
      }
    }
  }

  // Step 3: place each marker into the lower banner row (immediately above
  // the data) or, if that cell is blank, the nearest non-empty cell above.
  const lowerRowIdx = totalScanRowCount - 1;
  const dataColOffsetInScan = targetStartColumnIndex - scanStartColIndex;
  const markedCellKeys = new Set();
  let cellsPlanned = 0;

  for (let dataCol = 0; dataCol < targetColumnCount; dataCol++) {
    const label = labelMap.get(dataCol);
    if (!label) continue;
    cellsPlanned++;

    const scanCol = dataColOffsetInScan + dataCol;

    // After step 1 the lower-cell text has any RIT marker stripped.  Use that
    // post-strip text to decide whether the cell is "blank" — matches the
    // legacy writers, which read the same pre-existing text and decided via
    // appendOrReplaceTrailingBannerMarker (which also operates on a stripped
    // basis when a marker is already present).
    const lowerStripped = desiredTexts[lowerRowIdx][scanCol] || "";
    if (lowerStripped !== "") {
      desiredTexts[lowerRowIdx][scanCol] = appendOrReplaceTrailingBannerMarker(lowerStripped, label);
      markedCellKeys.add(`${lowerRowIdx},${scanCol}`);
      continue;
    }

    // Lower banner cell is blank — walk upward to the nearest non-empty cell
    // and place the marker there.  This is the legacy "vertically merged /
    // multi-row banner" behaviour: when the lower row is empty because the
    // visible header lives in a higher row, the marker follows the visible
    // text.
    let placed = false;
    for (let r = lowerRowIdx - 1; r >= 0; r--) {
      const upperStripped = desiredTexts[r][scanCol] || "";
      if (upperStripped !== "") {
        desiredTexts[r][scanCol] = appendOrReplaceTrailingBannerMarker(upperStripped, label);
        markedCellKeys.add(`${r},${scanCol}`);
        placed = true;
        break;
      }
    }

    if (!placed) {
      // Fallback when no non-empty upper cell exists: write the bare marker
      // text into the lower banner row.  Both legacy writers reach this
      // branch — structure-mode via its explicit "fall back to the lower
      // banner row" arm, non-structure via fall-through past the upper-scan
      // loop into the bottom "queue write to lower banner row" branch.  The
      // @ number-format distinction between the two paths is preserved by
      // the use of `useStructure` later when batching writes.
      desiredTexts[lowerRowIdx][scanCol] = appendOrReplaceTrailingBannerMarker("", label);
      markedCellKeys.add(`${lowerRowIdx},${scanCol}`);
    }
  }

  // Step 4: diff against the original texts and queue only changed cells.
  const writesByRow = new Map();
  let cellsChanged = 0;
  let skippedNoOpWrites = 0;

  for (let r = 0; r < totalScanRowCount; r++) {
    for (let c = 0; c < scanColCount; c++) {
      const cur = (bannerTexts[r] && bannerTexts[r][c]) || "";
      const nxt = desiredTexts[r][c] || "";
      const key = `${r},${c}`;

      if (cur === nxt) {
        if (markedCellKeys.has(key)) {
          skippedNoOpWrites++;
        }
        continue;
      }

      const absRow = targetStartRowIndex - totalScanRowCount + r;
      const absCol = scanStartColIndex + c;
      if (!writesByRow.has(absRow)) writesByRow.set(absRow, []);
      const isMarkered = markedCellKeys.has(key);
      writesByRow.get(absRow).push({
        colIndex: absCol,
        text: nxt,
        markered: isMarkered,
        needsNF: isMarkered && needsNumberFormatForBannerMarker(nxt),
      });
      cellsChanged++;
    }
  }

  const changedRows = writesByRow.size;
  let maxChangedCellsInRow = 0;
  for (const items of writesByRow.values()) {
    if (items.length > maxChangedCellsInRow) {
      maxChangedCellsInRow = items.length;
    }
  }

  const queueWriteStartMs = perfNow();
  let writeCommands = 0;
  let changedCellRuns = 0;
  let oneCellWriteCommands = 0;
  let multiCellWriteCommands = 0;
  let markeredWriteCommands = 0;
  let clearOnlyWriteCommands = 0;
  let numberFormatCommands = 0;
  let numberFormatCells = 0;
  let valueWriteCells = 0;
  let markerRowsUsed = 0;
  let maxRunLength = 0;
  let numberFormatSyncMs = 0;
  let valueWriteSyncMs = 0;
  let profileSyncCount = 0;
  const writeProfileEnabled = perfBannerWriteProfileEnabled();

  const profiledValueWrites = [];
  const markerRowIndexes = new Set();
  let queueCommandEndMs = 0;
  if (cellsChanged > 0) {
    const worksheet = writeTargetRange.worksheet;
    for (const [rowIndex, items] of writesByRow) {
      items.sort((a, b) => a.colIndex - b.colIndex);
      let i = 0;
      while (i < items.length) {
        let j = i;
        // Group adjacent columns with the same markered and needsNF flags so
        // the "@" number format is applied only to bare-marker cells —
        // non-bare marker text like "Wave 1 (a)" does not need "@".
        while (
          j + 1 < items.length &&
          items[j + 1].colIndex === items[j].colIndex + 1 &&
          items[j + 1].markered === items[j].markered &&
          items[j + 1].needsNF === items[j].needsNF
        ) {
          j++;
        }
        const slice = items.slice(i, j + 1);
        const texts = slice.map((x) => x.text);
        const range = worksheet.getRangeByIndexes(rowIndex, items[i].colIndex, 1, j - i + 1);
        const runLength = texts.length;

        if (useStructure && slice[0].markered && slice[0].needsNF) {
          range.numberFormat = [texts.map(() => "@")];
          numberFormatCommands++;
          numberFormatCells += runLength;
        }

        if (writeProfileEnabled) {
          profiledValueWrites.push({ range, texts });
        } else {
          range.values = [texts];
        }

        writeCommands++;
        changedCellRuns++;
        valueWriteCells += runLength;
        if (runLength === 1) {
          oneCellWriteCommands++;
        } else {
          multiCellWriteCommands++;
        }
        if (slice[0].markered) {
          markeredWriteCommands++;
          markerRowIndexes.add(rowIndex);
        } else {
          clearOnlyWriteCommands++;
        }
        if (runLength > maxRunLength) {
          maxRunLength = runLength;
        }
        i = j + 1;
      }
    }
    markerRowsUsed = markerRowIndexes.size;
    queueCommandEndMs = perfNow();

    if (writeProfileEnabled) {
      if (numberFormatCommands > 0) {
        const numberFormatSyncStartMs = perfNow();
        await context.sync();
        numberFormatSyncMs = perfNow() - numberFormatSyncStartMs;
        syncCount++;
        profileSyncCount++;
      }

      const valueWriteSyncStartMs = perfNow();
      for (const pending of profiledValueWrites) {
        pending.range.values = [pending.texts];
      }
      await context.sync();
      valueWriteSyncMs = perfNow() - valueWriteSyncStartMs;
      syncCount++;
      profileSyncCount++;
    }
    // In normal mode, intentionally NO context.sync() here.  Caller's final
    // context.sync() flushes queued banner writes alongside other pending ops.
  }

  const queueWriteEndMs = queueCommandEndMs || perfNow();
  const details = {
    rowsScanned: totalScanRowCount,
    cellsRead: totalScanRowCount * scanColCount,
    cellsPlanned,
    cellsChanged,
    writeCommands,
    changedCellRuns,
    avgRunLength: changedCellRuns
      ? roundBannerDiagnosticRatio(cellsChanged / changedCellRuns)
      : 0,
    maxRunLength,
    oneCellWriteCommands,
    multiCellWriteCommands,
    markeredWriteCommands,
    clearOnlyWriteCommands,
    numberFormatCommands,
    changedRows,
    avgChangedCellsPerChangedRow: changedRows
      ? roundBannerDiagnosticRatio(cellsChanged / changedRows)
      : 0,
    maxChangedCellsInRow,
    skippedNoOpWrites,
    readSyncMs: readSyncEndMs - readSyncStartMs,
    planMs: queueWriteStartMs - readSyncEndMs,
    queueWriteMs: queueWriteEndMs - queueWriteStartMs,
    syncCount,
    writeProfileEnabled,
    numberFormatSyncMs,
    valueWriteSyncMs,
    profileSyncCount,
    numberFormatCells,
    valueWriteCells,
    markerRowsUsed,
  };

  Object.defineProperty(details, BANNER_SCAN_AREA_STATS, {
    value: {
      rowIndex: targetStartRowIndex - totalScanRowCount,
      columnIndex: scanStartColIndex,
      rowCount: totalScanRowCount,
      columnCount: scanColCount,
    },
    enumerable: false,
  });

  return details;
}

/**
 * Returns true when a banner marker write needs numberFormat = "@".
 *
 * Only bare-marker cells risk Excel auto-formatting and need the guard.
 * A cell is "bare" when the only content is the marker itself, e.g. "(a)".
 * Cells with actual text like "Wave 1 (a)" are already text-safe without "@".
 */
function needsNumberFormatForBannerMarker(text) {
  return removeTrailingBannerMarker(text) === "";
}

/**
 * Appends or replaces trailing banner marker.
 *
 * Examples:
 * - "Male" + "a" -> "Male (a)"
 * - "Male (b)" + "a" -> "Male (a)"
 * - "Male (a)" + "a" -> "Male (a)"
 */
function appendOrReplaceTrailingBannerMarker(rawText, label) {
  const text = rawText === null || rawText === undefined ? "" : String(rawText).trim();

  const marker = `(${label})`;
  const currentMarker = getTrailingBannerMarker(text);

  if (currentMarker) {
    return `${text.slice(0, currentMarker.start).trim()} ${marker}`.trim();
  }

  return `${text} ${marker}`.trim();
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
