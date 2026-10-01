import { removeSignificanceMarkersFromText } from "./significance";

/**
 * Writes cell results into selected range.
 *
 * PURPOSE:
 * Shared writer for proportions, means, NPS, Total comparisons, and small bases.
 * It writes markers and applies fill formatting according to fillReason priority.
 */
export function writeCellResultsToSelectedRange(
  selectedRange,
  selectedText,
  cellResultMatrix,
  detectionResult,
  calculationSettings,
  options = {}
) {
  const rowCount = cellResultMatrix.length;
  const columnCount = cellResultMatrix[0] ? cellResultMatrix[0].length : 0;

  if (!rowCount || !columnCount) {
    return null;
  }

  const shouldCaptureWriterDetails = Boolean(options.captureWriterDetails);
  const writerDetails = shouldCaptureWriterDetails ? createWriterDetails() : null;
  const buildMatricesStartedAt = shouldCaptureWriterDetails ? Date.now() : 0;
  const shouldApplyVisualFormatting = calculationSettings.resultFormattingLevel !== "markersOnly";

  const rowTypeByIndex = buildDetectedRowTypeByIndexMap(detectionResult);

  const nextValues = [];
  const nextNumberFormats = [];

  const boldMask = [];
  const fillReasonMask = [];

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
    const valueRow = [];
    const numberFormatRow = [];
    const boldRow = [];
    const fillReasonRow = [];

    for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      const cellResult = cellResultMatrix[rowIndex][columnIndex];

      const currentText =
        selectedText && selectedText[rowIndex] && selectedText[rowIndex][columnIndex] !== undefined
          ? selectedText[rowIndex][columnIndex]
          : "";

      if (!cellResult) {
        const resolvedCurrent = resolveNumericOutput(currentText);

        if (resolvedCurrent !== null) {
          valueRow.push(resolvedCurrent.value);
          numberFormatRow.push(resolvedCurrent.format);
        } else {
          valueRow.push(currentText);
          numberFormatRow.push("@");
        }

        boldRow.push(false);
        fillReasonRow.push("none");
        continue;
      }

      const markers = cellResult.markers || "";
      const previousColumnArrow = cellResult.previousColumnArrow || "";
      const fillReason = cellResult.fillReason || "none";

      const displayedValueWithoutMarkers = removeSignificanceMarkersFromText(currentText);

      const roundedDisplayedValue = formatDisplayedValueForOutput(
        displayedValueWithoutMarkers,
        rowIndex,
        rowTypeByIndex,
        calculationSettings
      );

      const outputMarkerText = previousColumnArrow || markers;

      const nextValue = outputMarkerText
        ? `${roundedDisplayedValue} ${outputMarkerText}`.trim()
        : roundedDisplayedValue;

      if (outputMarkerText) {
        valueRow.push(nextValue);
        numberFormatRow.push("@");
      } else {
        const resolvedRounded = resolveNumericOutput(roundedDisplayedValue);

        if (resolvedRounded !== null) {
          valueRow.push(resolvedRounded.value);
          numberFormatRow.push(resolvedRounded.format);
        } else {
          valueRow.push(roundedDisplayedValue);
          numberFormatRow.push("@");
        }
      }

      boldRow.push(
        Boolean(
          markers ||
          previousColumnArrow ||
          fillReason === "significant" ||
          fillReason === "lowerThanTotal"
        )
      );

      fillReasonRow.push(fillReason);
    }

    nextValues.push(valueRow);
    nextNumberFormats.push(numberFormatRow);
    boldMask.push(boldRow);
    fillReasonMask.push(fillReasonRow);
  }

  if (writerDetails) {
    writerDetails.buildMatricesMs = Date.now() - buildMatricesStartedAt;
  }

  // Main performance win:
  // one values write + one numberFormat write instead of per-cell writes.
  const numberFormatWriteStartedAt = writerDetails ? Date.now() : 0;
  selectedRange.numberFormat = nextNumberFormats;
  if (writerDetails) {
    writerDetails.numberFormatWriteMs = Date.now() - numberFormatWriteStartedAt;
  }

  const valuesWriteStartedAt = writerDetails ? Date.now() : 0;
  selectedRange.values = nextValues;
  if (writerDetails) {
    writerDetails.valuesWriteMs = Date.now() - valuesWriteStartedAt;
  }

  if (shouldApplyVisualFormatting) {
    const formattingContext = resolveFormattingContext(selectedRange, options);

    applyGroupedBoldFormatting(selectedRange, boldMask, formattingContext, writerDetails);
    applyGroupedFillFormatting(
      selectedRange,
      fillReasonMask,
      cellResultMatrix,
      calculationSettings,
      formattingContext,
      writerDetails
    );

    if (writerDetails) {
      writerDetails.formattingPath = shouldUseRangeAreasPath(formattingContext)
        ? "rangeAreas"
        : "rowRuns";
      populateRangeAreasDiagnostics(writerDetails, formattingContext);
    }
  }

  return writerDetails;
}

/**
 * Diagnostics-only: projects how many queued Office.js operations a chunked
 * `worksheet.getRanges(...)` RangeAreas strategy (ExcelApi 1.9) uses for the
 * bold and fill masks the writer just processed. On the RangeAreas path, the
 * projection matches the actual queued ops. On the row-run fallback path,
 * the projection still shows what the chunked path would have done for
 * before/after comparison.
 *
 * Requires resolved `anchorRowIndex` / `anchorColumnIndex` from the formatting
 * context. If neither caller-supplied options nor loaded selectedRange
 * properties provided them, the projection is skipped silently.
 */
function populateRangeAreasDiagnostics(writerDetails, formattingContext) {
  const anchorRowIndex = formattingContext.anchorRowIndex;
  const anchorColumnIndex = formattingContext.anchorColumnIndex;

  if (typeof anchorRowIndex !== "number" || typeof anchorColumnIndex !== "number") {
    // No anchor available — drop the internal span arrays anyway so the
    // RIT_PERF log stays compact on wide workbooks.
    delete writerDetails._boldRunSpansByRow;
    delete writerDetails._fillRunSpansByRow;
    return;
  }

  const boldRunSpansByRow = writerDetails._boldRunSpansByRow || [];
  const fillRunSpansByRow = writerDetails._fillRunSpansByRow || [];

  writerDetails.boldRangeAreas = buildBoldRangeAreasDiagnostics(
    boldRunSpansByRow,
    anchorRowIndex,
    anchorColumnIndex
  );

  writerDetails.fillRangeAreas = buildFillRangeAreasDiagnosticsByColor(
    fillRunSpansByRow,
    anchorRowIndex,
    anchorColumnIndex
  );

  // The per-row span arrays are an internal detail used to feed the RangeAreas
  // estimate. They can be large on wide workbooks, so drop them from the
  // emitted writerDetails to keep the perf log lean.
  delete writerDetails._boldRunSpansByRow;
  delete writerDetails._fillRunSpansByRow;
}

/**
 * Resolves the formatting context for a single writer invocation:
 *   - whether the host advertises ExcelApi 1.9 (`worksheet.getRanges(...)`);
 *   - the sheet-absolute anchor row/column of the write range.
 *
 * Support is evaluated once per write so the gate cost is paid at most one
 * time across both bold and fill formatting passes.
 *
 * Anchor coords are preferred from caller-supplied options
 * (`anchorRowIndex` / `anchorColumnIndex`) because the autorun path computes
 * them from the already-loaded source range and never loads them on the write
 * target proxy. If options are missing, the writer falls back to reading
 * `selectedRange.rowIndex` / `selectedRange.columnIndex` (loaded by the
 * manual selected-range path). Either resolves to numeric coords, or both
 * stay null and the RangeAreas path is disabled for this write.
 */
function resolveFormattingContext(selectedRange, options = {}) {
  let anchorRowIndex = null;
  let anchorColumnIndex = null;

  if (typeof options.anchorRowIndex === "number" && typeof options.anchorColumnIndex === "number") {
    anchorRowIndex = options.anchorRowIndex;
    anchorColumnIndex = options.anchorColumnIndex;
  } else {
    try {
      const candidateRow = selectedRange.rowIndex;
      const candidateColumn = selectedRange.columnIndex;
      if (typeof candidateRow === "number" && typeof candidateColumn === "number") {
        anchorRowIndex = candidateRow;
        anchorColumnIndex = candidateColumn;
      }
    } catch (_) {
      // PropertyNotLoaded or similar — RangeAreas path is disabled for this
      // write and the row-run fallback is used.
    }
  }

  return {
    isExcelApi19Supported: isExcelApi19Supported(),
    anchorRowIndex,
    anchorColumnIndex,
  };
}

/**
 * Returns true if the host advertises ExcelApi 1.9 (`worksheet.getRanges(...)`
 * → `RangeAreas`). Safe to call in test environments without an Office global.
 */
function isExcelApi19Supported() {
  try {
    if (
      typeof Office === "undefined" ||
      !Office ||
      !Office.context ||
      !Office.context.requirements ||
      typeof Office.context.requirements.isSetSupported !== "function"
    ) {
      return false;
    }
    return Boolean(Office.context.requirements.isSetSupported("ExcelApi", "1.9"));
  } catch (_) {
    return false;
  }
}

/**
 * Pure branch-selection helper exported for tests. Returns true when the
 * writer should issue the chunked `worksheet.getRanges(...)` RangeAreas path
 * and false when it should fall back to the per-row-run path.
 */
export function shouldUseRangeAreasPath(formattingContext) {
  return Boolean(
    formattingContext &&
    formattingContext.isExcelApi19Supported &&
    typeof formattingContext.anchorRowIndex === "number" &&
    typeof formattingContext.anchorColumnIndex === "number"
  );
}

/**
 * Returns fill color for cell result according to fill reason and settings.
 *
 * RULE:
 * If fillOnlyTotalComparisons is enabled, normal significant fill is applied
 * only to cells that are significantly higher than Total.
 */
function getFillColorForCellResult(cellResult, calculationSettings) {
  const fillReason = cellResult.fillReason || "none";

  if (fillReason === "smallBase") {
    return calculationSettings.smallBaseFillColor || "#D0D0D0";
  }

  if (fillReason === "lowerThanTotal") {
    return calculationSettings.lowerThanTotalFillColor || "#FCE4D6";
  }

  if (fillReason === "significant") {
    if (calculationSettings.fillOnlyTotalComparisons && !cellResult.hasPositiveTotalComparison) {
      return "";
    }

    return calculationSettings.significantFillColor || "#E2F0D9";
  }

  return "";
}

/**
 * Formats displayed cell value before significance markers are appended.
 *
 * PURPOSE:
 * Real data often contains long decimal values.
 * We round values before adding significance letters to keep cells readable.
 */
function formatDisplayedValueForOutput(
  displayedValue,
  rowIndex,
  rowTypeByIndex,
  calculationSettings
) {
  const parsedValue = parseOutputNumber(displayedValue);

  if (parsedValue === null) {
    return displayedValue;
  }

  const rowType = rowTypeByIndex.get(rowIndex) || null;
  const decimalPlaces = getDecimalPlacesForRowType(rowType, calculationSettings);

  if (decimalPlaces === null) {
    return displayedValue;
  }

  const roundedValue = parsedValue.numericValue.toFixed(decimalPlaces);

  return parsedValue.hasPercentSign ? `${roundedValue}%` : roundedValue;
}

/**
 * Parses a displayed spreadsheet value into number.
 *
 * PURPOSE:
 * Supports comma decimals and percent signs.
 */
function parseOutputNumber(displayedValue) {
  if (displayedValue === null || displayedValue === undefined || displayedValue === "") {
    return null;
  }

  const rawTextValue = String(displayedValue).trim();
  const hasPercentSign = rawTextValue.endsWith("%");

  const textValue = rawTextValue.replace("%", "").replace(",", ".");

  const numericValue = Number(textValue);

  if (Number.isNaN(numericValue)) {
    return null;
  }

  return {
    numericValue,
    hasPercentSign,
  };
}

/**
 * Returns detected row type for selected range row index.
 */
function buildDetectedRowTypeByIndexMap(detectionResult) {
  const rowTypeByIndex = new Map();

  if (!detectionResult || !detectionResult.rowDiagnostics) {
    return rowTypeByIndex;
  }

  for (const rowDiagnostic of detectionResult.rowDiagnostics) {
    rowTypeByIndex.set(rowDiagnostic.rowIndex, rowDiagnostic.rowType);
  }

  return rowTypeByIndex;
}

/**
 * Returns number of decimal places for visible output.
 *
 * DEFAULT:
 * - proportions / NPS / promoters / detractors: 1 decimal
 * - means / SD / variance: 2 decimals
 *
 * IF roundCellValues is enabled:
 * - proportions / NPS / promoters / detractors: 0 decimals
 * - means / SD / variance: 1 decimal
 */
function getDecimalPlacesForRowType(rowType, calculationSettings) {
  const shouldRoundCellValues = calculationSettings && calculationSettings.roundCellValues;

  const shareLikeRowTypes = ["proportion", "nps", "promoters", "detractors"];

  const meanLikeRowTypes = ["mean", "standardDeviation", "variance"];

  if (shareLikeRowTypes.includes(rowType)) {
    return shouldRoundCellValues ? 0 : 1;
  }

  if (meanLikeRowTypes.includes(rowType)) {
    return shouldRoundCellValues ? 1 : 2;
  }

  return null;
}

/**
 * Resolves numeric value and Excel format for an unmarked output cell.
 *
 * PURPOSE:
 * Unmarked cells should be written as numeric Excel values to avoid
 * "number stored as text" warnings. The format is derived from the display
 * string itself so the visible convention is preserved exactly:
 * - strings with "%" use percent format and divide by 100 for Excel storage;
 * - strings without "%" use a plain decimal format matching the string's precision;
 * - integers use "General";
 * - non-numeric strings (labels, empty) return null → caller falls back to text.
 *
 * This ensures "28%" stays "28%" and "0.281" stays "0.281" regardless of row type.
 */
export function resolveNumericOutput(displayString) {
  const parsed = parseOutputNumber(displayString);

  if (parsed === null) {
    return null;
  }

  const cleanText = String(displayString).trim().replace("%", "").trim();
  const dotIndex = cleanText.indexOf(".");
  const decimalPlaces = dotIndex >= 0 ? cleanText.length - dotIndex - 1 : 0;

  if (parsed.hasPercentSign) {
    const format = decimalPlaces === 0 ? "0%" : `0.${"0".repeat(decimalPlaces)}%`;
    return { value: parsed.numericValue / 100, format };
  }

  if (decimalPlaces === 0) {
    return { value: parsed.numericValue, format: "General" };
  }

  return { value: parsed.numericValue, format: `0.${"0".repeat(decimalPlaces)}` };
}

function applyGroupedBoldFormatting(selectedRange, boldMask, formattingContext, writerDetails = null) {
  const formatStartedAt = writerDetails ? Date.now() : 0;

  // Walk the mask once and materialize per-row spans so both the row-run and
  // RangeAreas writers can consume the same span data without re-walking.
  const boldRunSpansByRow = collectBoldRunSpansByRow(boldMask);
  const boldRunsByRow = boldRunSpansByRow.map((rowSpans) => rowSpans.length);

  const useRangeAreas = shouldUseRangeAreasPath(formattingContext);
  let appliedCommandEstimate = 0;

  if (useRangeAreas) {
    appliedCommandEstimate = applyBoldSpansViaRangeAreas(
      selectedRange,
      boldRunSpansByRow,
      formattingContext.anchorRowIndex,
      formattingContext.anchorColumnIndex
    );
  } else {
    applyBoldSpansViaRowRuns(selectedRange, boldRunSpansByRow);
  }

  if (writerDetails) {
    writerDetails.boldFormatMs = Date.now() - formatStartedAt;
    writerDetails.boldRunCountByRow = boldRunsByRow;
    writerDetails.boldCommandCount = sumCounts(boldRunsByRow);
    writerDetails.boldRectCommandCountEstimate = estimateRectangleCommandCount(boldRunSpansByRow);
    writerDetails._boldRunSpansByRow = boldRunSpansByRow;
    writerDetails.boldRangeAreasAppliedCommandEstimate = appliedCommandEstimate;
  }
}

function collectBoldRunSpansByRow(boldMask) {
  const out = [];

  for (let rowIndex = 0; rowIndex < boldMask.length; rowIndex++) {
    const row = boldMask[rowIndex];
    const rowSpans = [];
    let runStart = null;

    for (let columnIndex = 0; columnIndex <= row.length; columnIndex++) {
      const shouldBeBold = columnIndex < row.length ? row[columnIndex] : false;

      if (shouldBeBold && runStart === null) {
        runStart = columnIndex;
        continue;
      }

      if ((!shouldBeBold || columnIndex === row.length) && runStart !== null) {
        rowSpans.push({ start: runStart, end: columnIndex - 1 });
        runStart = null;
      }
    }

    out.push(rowSpans);
  }

  return out;
}

function applyBoldSpansViaRowRuns(selectedRange, boldRunSpansByRow) {
  for (let rowIndex = 0; rowIndex < boldRunSpansByRow.length; rowIndex++) {
    const rowSpans = boldRunSpansByRow[rowIndex];
    for (const runSpan of rowSpans) {
      const runWidth = runSpan.end - runSpan.start + 1;
      selectedRange
        .getCell(rowIndex, runSpan.start)
        .getResizedRange(0, runWidth - 1).format.font.bold = true;
    }
  }
}

function applyBoldSpansViaRangeAreas(
  selectedRange,
  boldRunSpansByRow,
  anchorRowIndex,
  anchorColumnIndex
) {
  const addresses = [];
  for (let rowIndex = 0; rowIndex < boldRunSpansByRow.length; rowIndex++) {
    const rowSpans = boldRunSpansByRow[rowIndex];
    for (const runSpan of rowSpans) {
      addresses.push(runSpanToA1Address(runSpan, rowIndex, anchorRowIndex, anchorColumnIndex));
    }
  }

  if (addresses.length === 0) {
    return 0;
  }

  const chunks = packAddressesIntoChunks(addresses);
  const worksheet = selectedRange.worksheet;

  for (const chunk of chunks) {
    worksheet.getRanges(chunk.address).format.font.bold = true;
  }

  // Each chunk costs one getRanges + one .format.font.bold setter.
  return chunks.length * 2;
}

function applyGroupedFillFormatting(
  selectedRange,
  fillReasonMask,
  cellResultMatrix,
  calculationSettings,
  formattingContext,
  writerDetails = null
) {
  const formatStartedAt = writerDetails ? Date.now() : 0;

  const fillRunSpansByRow = collectFillRunSpansByRow(
    fillReasonMask,
    cellResultMatrix,
    calculationSettings
  );
  const fillRunsByRow = fillRunSpansByRow.map((rowSpans) => rowSpans.length);

  const useRangeAreas = shouldUseRangeAreasPath(formattingContext);
  let appliedCommandEstimate = 0;

  if (useRangeAreas) {
    appliedCommandEstimate = applyFillSpansViaRangeAreas(
      selectedRange,
      fillRunSpansByRow,
      formattingContext.anchorRowIndex,
      formattingContext.anchorColumnIndex
    );
  } else {
    applyFillSpansViaRowRuns(selectedRange, fillRunSpansByRow);
  }

  if (writerDetails) {
    writerDetails.fillFormatMs = Date.now() - formatStartedAt;
    writerDetails.fillRunCountByRow = fillRunsByRow;
    writerDetails.fillCommandCount = sumCounts(fillRunsByRow);
    writerDetails.fillRectCommandCountEstimate = estimateRectangleCommandCount(fillRunSpansByRow);
    writerDetails._fillRunSpansByRow = fillRunSpansByRow;
    writerDetails.fillRangeAreasAppliedCommandEstimate = appliedCommandEstimate;
  }
}

function collectFillRunSpansByRow(fillReasonMask, cellResultMatrix, calculationSettings) {
  const out = [];

  for (let rowIndex = 0; rowIndex < fillReasonMask.length; rowIndex++) {
    const row = fillReasonMask[rowIndex];
    const rowSpans = [];

    let runStart = null;
    let currentRunColor = "";

    for (let columnIndex = 0; columnIndex <= row.length; columnIndex++) {
      const cellResult =
        columnIndex < row.length && cellResultMatrix[rowIndex]
          ? cellResultMatrix[rowIndex][columnIndex]
          : null;

      const fillColor = cellResult
        ? getFillColorForCellResult(cellResult, calculationSettings)
        : "";

      const continuesRun = fillColor && runStart !== null && fillColor === currentRunColor;

      if (fillColor && runStart === null) {
        runStart = columnIndex;
        currentRunColor = fillColor;
        continue;
      }

      if (continuesRun) {
        continue;
      }

      if (runStart !== null) {
        rowSpans.push({
          start: runStart,
          end: columnIndex - 1,
          color: currentRunColor,
        });

        runStart = null;
        currentRunColor = "";
      }

      if (fillColor) {
        runStart = columnIndex;
        currentRunColor = fillColor;
      }
    }

    out.push(rowSpans);
  }

  return out;
}

function applyFillSpansViaRowRuns(selectedRange, fillRunSpansByRow) {
  for (let rowIndex = 0; rowIndex < fillRunSpansByRow.length; rowIndex++) {
    const rowSpans = fillRunSpansByRow[rowIndex];
    for (const runSpan of rowSpans) {
      const runWidth = runSpan.end - runSpan.start + 1;
      selectedRange
        .getCell(rowIndex, runSpan.start)
        .getResizedRange(0, runWidth - 1).format.fill.color = runSpan.color;
    }
  }
}

function applyFillSpansViaRangeAreas(
  selectedRange,
  fillRunSpansByRow,
  anchorRowIndex,
  anchorColumnIndex
) {
  // RangeAreas.format.fill.color is a single uniform value per call, so spans
  // are grouped by color and chunked independently for each color. Different
  // colors never share a chunk.
  const addressesByColor = new Map();

  for (let rowIndex = 0; rowIndex < fillRunSpansByRow.length; rowIndex++) {
    const rowSpans = fillRunSpansByRow[rowIndex];
    for (const runSpan of rowSpans) {
      const color = runSpan.color || "";
      if (!color) continue;

      const address = runSpanToA1Address(runSpan, rowIndex, anchorRowIndex, anchorColumnIndex);

      if (!addressesByColor.has(color)) {
        addressesByColor.set(color, []);
      }
      addressesByColor.get(color).push(address);
    }
  }

  if (addressesByColor.size === 0) {
    return 0;
  }

  let totalChunkCount = 0;
  const worksheet = selectedRange.worksheet;

  for (const [color, addresses] of addressesByColor) {
    const chunks = packAddressesIntoChunks(addresses);
    totalChunkCount += chunks.length;

    for (const chunk of chunks) {
      worksheet.getRanges(chunk.address).format.fill.color = color;
    }
  }

  // Each chunk costs one getRanges + one .format.fill.color setter.
  return totalChunkCount * 2;
}

function createWriterDetails() {
  return {
    buildMatricesMs: 0,
    numberFormatWriteMs: 0,
    valuesWriteMs: 0,
    boldFormatMs: 0,
    fillFormatMs: 0,
    boldCommandCount: 0,
    fillCommandCount: 0,
    boldRunCountByRow: [],
    fillRunCountByRow: [],
    boldRectCommandCountEstimate: 0,
    fillRectCommandCountEstimate: 0,
    boldRangeAreas: null,
    fillRangeAreas: null,
    // formattingPath records which branch of applyGrouped*Formatting actually
    // ran for this write. "rangeAreas" → chunked `worksheet.getRanges(...)`
    // path (ExcelApi 1.9). "rowRuns" → per-row-run fallback. null when visual
    // formatting was skipped (markers-only mode).
    formattingPath: null,
    // Actual queued-op estimates for the path the writer took:
    //   rangeAreas path → chunks * 2 (one getRanges + one setter per chunk).
    //   rowRuns path    → 0 (the projection in boldRangeAreas/fillRangeAreas
    //                       still shows the chunked path's count for comparison).
    boldRangeAreasAppliedCommandEstimate: 0,
    fillRangeAreasAppliedCommandEstimate: 0,
  };
}

function sumCounts(counts) {
  return counts.reduce((sum, count) => sum + count, 0);
}

function estimateRectangleCommandCount(runSpansByRow) {
  let rectangleCount = 0;
  let previousRowSignatureCounts = new Map();

  for (const rowRunSpans of runSpansByRow) {
    const currentRowSignatureCounts = new Map();

    for (const runSpan of rowRunSpans) {
      const signature = buildRunSpanSignature(runSpan);
      const nextCount = (currentRowSignatureCounts.get(signature) || 0) + 1;
      currentRowSignatureCounts.set(signature, nextCount);

      const previousCount = previousRowSignatureCounts.get(signature) || 0;

      if (nextCount > previousCount) {
        rectangleCount += 1;
      }
    }

    previousRowSignatureCounts = currentRowSignatureCounts;
  }

  return rectangleCount;
}

function buildRunSpanSignature(runSpan) {
  const colorKey = runSpan.color || "";
  return `${runSpan.start}:${runSpan.end}:${colorKey}`;
}

/**
 * Helpers for the chunked RangeAreas formatting path (issues #284 / #286).
 *
 * These helpers convert the row-run mask data the writer produces into the
 * A1 address strings a `worksheet.getRanges(...)` (RangeAreas, ExcelApi 1.9)
 * call consumes. They are used both:
 *   - by the writer's RangeAreas path (`applyBoldSpansViaRangeAreas` /
 *     `applyFillSpansViaRangeAreas`) when the host advertises ExcelApi 1.9, to
 *     produce the chunk addresses passed to `worksheet.getRanges(...)`;
 *   - by `populateRangeAreasDiagnostics` to record the projection in
 *     `writerDetails.boldRangeAreas` / `writerDetails.fillRangeAreas` for
 *     RIT_PERF logging. On the RangeAreas path the projection matches the
 *     actual queued ops; on the row-run fallback it still shows what the
 *     chunked path would have done for before/after comparison.
 *
 * Pure: no Office.js calls. Anchor coords are passed in as numbers.
 */

const DEFAULT_RANGE_AREAS_CHUNK_OPTIONS = Object.freeze({
  maxAreasPerChunk: 100,
  maxCharsPerChunk: 2000,
});

/**
 * Converts a zero-based column index to an A1-style column reference.
 * 0 → "A", 25 → "Z", 26 → "AA", 51 → "AZ", 52 → "BA".
 */
export function columnIndexToA1(zeroBasedColumnIndex) {
  let n = zeroBasedColumnIndex + 1;
  let result = "";

  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }

  return result;
}

/**
 * Converts a row-run span on the selected range to a sheet-relative A1
 * address. Returns "B3:F3" style. Single-cell runs collapse to "B3".
 */
export function runSpanToA1Address(runSpan, rowIndex, anchorRowIndex, anchorColumnIndex) {
  const absoluteRow = anchorRowIndex + rowIndex + 1;
  const startColumnLetters = columnIndexToA1(anchorColumnIndex + runSpan.start);

  if (runSpan.end === runSpan.start) {
    return `${startColumnLetters}${absoluteRow}`;
  }

  const endColumnLetters = columnIndexToA1(anchorColumnIndex + runSpan.end);
  return `${startColumnLetters}${absoluteRow}:${endColumnLetters}${absoluteRow}`;
}

/**
 * Packs A1 address strings into comma-separated chunks bounded by both an
 * area-count cap and a character-count cap. Either cap may fire first.
 *
 * Returns an array of chunk descriptors so diagnostics can record both the
 * address string and its size without re-walking the runs.
 */
export function packAddressesIntoChunks(addresses, options = {}) {
  const maxAreasPerChunk = options.maxAreasPerChunk || DEFAULT_RANGE_AREAS_CHUNK_OPTIONS.maxAreasPerChunk;
  const maxCharsPerChunk = options.maxCharsPerChunk || DEFAULT_RANGE_AREAS_CHUNK_OPTIONS.maxCharsPerChunk;

  const chunks = [];

  if (addresses.length === 0) {
    return chunks;
  }

  let currentAddresses = [];
  let currentLength = 0;

  for (const address of addresses) {
    const separatorLength = currentAddresses.length === 0 ? 0 : 1;
    const projectedLength = currentLength + separatorLength + address.length;

    const overAreaCap = currentAddresses.length >= maxAreasPerChunk;
    const overCharCap = projectedLength > maxCharsPerChunk;

    if ((overAreaCap || overCharCap) && currentAddresses.length > 0) {
      chunks.push({
        address: currentAddresses.join(","),
        areaCount: currentAddresses.length,
        length: currentLength,
      });
      currentAddresses = [];
      currentLength = 0;
    }

    const isFirstInNewChunk = currentAddresses.length === 0;
    currentAddresses.push(address);
    currentLength += (isFirstInNewChunk ? 0 : 1) + address.length;
  }

  if (currentAddresses.length > 0) {
    chunks.push({
      address: currentAddresses.join(","),
      areaCount: currentAddresses.length,
      length: currentLength,
    });
  }

  return chunks;
}

/**
 * Builds RangeAreas diagnostic data for a single uniform-format mask (e.g.
 * bold). Returns counts and chunk metadata, plus a worked command-count
 * projection assuming one `worksheet.getRanges(...)` + one setter per chunk.
 *
 * runSpansByRow: same shape used by the writer — array indexed by mask row,
 *                each entry is array of { start, end } spans.
 */
export function buildBoldRangeAreasDiagnostics(
  runSpansByRow,
  anchorRowIndex,
  anchorColumnIndex,
  options = {}
) {
  const addresses = [];

  for (let rowIndex = 0; rowIndex < runSpansByRow.length; rowIndex++) {
    const rowSpans = runSpansByRow[rowIndex];
    if (!rowSpans) continue;

    for (const runSpan of rowSpans) {
      addresses.push(runSpanToA1Address(runSpan, rowIndex, anchorRowIndex, anchorColumnIndex));
    }
  }

  const chunks = packAddressesIntoChunks(addresses, options);

  return summarizeRangeAreasChunks(addresses.length, chunks);
}

/**
 * Builds RangeAreas diagnostic data for fill formatting, grouped by color
 * because RangeAreas.format.fill.color is a single uniform value per call.
 *
 * runSpansByRow entries carry `{ start, end, color }`.
 */
export function buildFillRangeAreasDiagnosticsByColor(
  runSpansByRow,
  anchorRowIndex,
  anchorColumnIndex,
  options = {}
) {
  const addressesByColor = new Map();

  for (let rowIndex = 0; rowIndex < runSpansByRow.length; rowIndex++) {
    const rowSpans = runSpansByRow[rowIndex];
    if (!rowSpans) continue;

    for (const runSpan of rowSpans) {
      const color = runSpan.color || "";
      if (!color) continue;

      const address = runSpanToA1Address(runSpan, rowIndex, anchorRowIndex, anchorColumnIndex);

      if (!addressesByColor.has(color)) {
        addressesByColor.set(color, []);
      }

      addressesByColor.get(color).push(address);
    }
  }

  const perColor = [];
  let totalAreaCount = 0;
  let totalChunkCount = 0;
  let maxAddressLength = 0;

  for (const [color, addresses] of addressesByColor) {
    const chunks = packAddressesIntoChunks(addresses, options);
    const summary = summarizeRangeAreasChunks(addresses.length, chunks);

    perColor.push({
      color,
      areaCount: summary.areaCountTotal,
      chunkCount: summary.chunkCount,
      commandCountEstimate: summary.commandCountEstimate,
      maxAddressLength: summary.maxAddressLength,
    });

    totalAreaCount += summary.areaCountTotal;
    totalChunkCount += summary.chunkCount;
    if (summary.maxAddressLength > maxAddressLength) {
      maxAddressLength = summary.maxAddressLength;
    }
  }

  return {
    perColor,
    colorCount: perColor.length,
    areaCountTotal: totalAreaCount,
    chunkCount: totalChunkCount,
    // Each chunk costs one getRanges + one format.fill.color setter.
    commandCountEstimate: totalChunkCount * 2,
    maxAddressLength,
  };
}

function summarizeRangeAreasChunks(areaCountTotal, chunks) {
  let maxAddressLength = 0;
  for (const chunk of chunks) {
    if (chunk.length > maxAddressLength) {
      maxAddressLength = chunk.length;
    }
  }

  return {
    areaCountTotal,
    chunkCount: chunks.length,
    // Each chunk costs one getRanges + one format.font.bold (or .fill.color)
    // setter. Sync count is unchanged: the chunked ops queue inside the
    // existing single sync.
    commandCountEstimate: chunks.length * 2,
    maxAddressLength,
  };
}
