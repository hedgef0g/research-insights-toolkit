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
  calculationSettings
) {
  const rowCount = cellResultMatrix.length;
  const columnCount = cellResultMatrix[0] ? cellResultMatrix[0].length : 0;

  if (!rowCount || !columnCount) {
    return;
  }

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
        const parsedCurrent = parseOutputNumber(currentText);

        if (parsedCurrent !== null) {
          valueRow.push(
            parsedCurrent.hasPercentSign ? parsedCurrent.numericValue / 100 : parsedCurrent.numericValue
          );
          numberFormatRow.push(getNumberFormatForRowType(rowTypeByIndex.get(rowIndex), calculationSettings));
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
        const parsedRounded = parseOutputNumber(roundedDisplayedValue);

        if (parsedRounded !== null) {
          valueRow.push(
            parsedRounded.hasPercentSign ? parsedRounded.numericValue / 100 : parsedRounded.numericValue
          );
          numberFormatRow.push(getNumberFormatForRowType(rowTypeByIndex.get(rowIndex), calculationSettings));
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

  // Main performance win:
  // one values write + one numberFormat write instead of per-cell writes.
  selectedRange.numberFormat = nextNumberFormats;
  selectedRange.values = nextValues;

  applyGroupedBoldFormatting(selectedRange, boldMask);
  applyGroupedFillFormatting(selectedRange, fillReasonMask, cellResultMatrix, calculationSettings);
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
 * Returns an Excel number format string for a given row type.
 *
 * PURPOSE:
 * Unmarked output cells should be written as numeric values rather than text.
 * This helper provides the appropriate display format so the visible output
 * matches what formatDisplayedValueForOutput would have shown as a string.
 */
function getNumberFormatForRowType(rowType, calculationSettings) {
  const shouldRound = calculationSettings && calculationSettings.roundCellValues;

  const shareLikeTypes = ["proportion", "nps", "promoters", "detractors"];

  if (shareLikeTypes.includes(rowType)) {
    return shouldRound ? "0%" : "0.0%";
  }

  if (rowType === "mean" || rowType === "standardDeviation" || rowType === "variance") {
    return shouldRound ? "0.0" : "0.00";
  }

  return "General";
}

function applyGroupedBoldFormatting(selectedRange, boldMask) {
  for (let rowIndex = 0; rowIndex < boldMask.length; rowIndex++) {
    const row = boldMask[rowIndex];

    let runStart = null;

    for (let columnIndex = 0; columnIndex <= row.length; columnIndex++) {
      const shouldBeBold = columnIndex < row.length ? row[columnIndex] : false;

      if (shouldBeBold && runStart === null) {
        runStart = columnIndex;
        continue;
      }

      if ((!shouldBeBold || columnIndex === row.length) && runStart !== null) {
        const runEnd = columnIndex - 1;
        const runWidth = runEnd - runStart + 1;

        selectedRange
          .getCell(rowIndex, runStart)
          .getResizedRange(0, runWidth - 1).format.font.bold = true;

        runStart = null;
      }
    }
  }
}

function applyGroupedFillFormatting(
  selectedRange,
  fillReasonMask,
  cellResultMatrix,
  calculationSettings
) {
  for (let rowIndex = 0; rowIndex < fillReasonMask.length; rowIndex++) {
    const row = fillReasonMask[rowIndex];

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
        const runEnd = columnIndex - 1;
        const runWidth = runEnd - runStart + 1;

        selectedRange
          .getCell(rowIndex, runStart)
          .getResizedRange(0, runWidth - 1).format.fill.color = currentRunColor;

        runStart = null;
        currentRunColor = "";
      }

      if (fillColor) {
        runStart = columnIndex;
        currentRunColor = fillColor;
      }
    }
  }
}
