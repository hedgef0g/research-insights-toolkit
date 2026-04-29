import { removeSignificanceMarkersFromText } from "./significance";

/**
 * Writes significance markers into selected range.
 *
 * PURPOSE:
 * Shared writer for proportions, means, and NPS blocks.
 * It also protects cells from Excel auto-formatting values as time.
 */
export function writeMarkersToSelectedRange(
  selectedRange,
  selectedText,
  markerMatrix,
  detectionResult,
  calculationSettings
) {
  const rowCount = markerMatrix.length; // Number of rows in marker matrix.
  const columnCount = markerMatrix[0] ? markerMatrix[0].length : 0; // Number of columns.

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      const markers = markerMatrix[rowIndex][columnIndex]; // Marker letters for current cell.

      if (!markers) {
        continue;
      }

      const currentCell = selectedRange.getCell(rowIndex, columnIndex); // Target cell.
      const displayedValueWithoutMarkers = removeSignificanceMarkersFromText(
        selectedText[rowIndex][columnIndex]
      );

      const roundedDisplayedValue = formatDisplayedValueForOutput(
        displayedValueWithoutMarkers,
        rowIndex,
        detectionResult,
        calculationSettings
      );

      currentCell.numberFormat = [["@"]];

      currentCell.values = [[`${roundedDisplayedValue} ${markers}`.trim()]];

      currentCell.format.font.bold = true;
      currentCell.format.fill.color = "#E2F0D9";
    }
  }
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
  detectionResult,
  calculationSettings
) {
  const numericValue = parseOutputNumber(displayedValue);

  if (numericValue === null) {
    return displayedValue;
  }

  const rowType = getDetectedRowTypeByIndex(detectionResult, rowIndex);
  const decimalPlaces = getDecimalPlacesForRowType(
    rowType,
    calculationSettings
  );

  if (decimalPlaces === null) {
    return displayedValue;
  }

  return numericValue.toFixed(decimalPlaces);
}

/**
 * Parses a displayed spreadsheet value into number.
 *
 * PURPOSE:
 * Supports comma decimals and percent signs.
 */
function parseOutputNumber(displayedValue) {
  if (
    displayedValue === null ||
    displayedValue === undefined ||
    displayedValue === ""
  ) {
    return null;
  }

  const textValue = String(displayedValue)
    .trim()
    .replace("%", "")
    .replace(",", ".");

  const numericValue = Number(textValue);

  if (Number.isNaN(numericValue)) {
    return null;
  }

  return numericValue;
}

/**
 * Returns detected row type for selected range row index.
 */
function getDetectedRowTypeByIndex(detectionResult, rowIndex) {
  if (!detectionResult || !detectionResult.rowDiagnostics) {
    return null;
  }

  const rowDiagnostic = detectionResult.rowDiagnostics.find(
    (diagnostic) => diagnostic.rowIndex === rowIndex
  );

  return rowDiagnostic ? rowDiagnostic.rowType : null;
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
  const shouldRoundCellValues =
    calculationSettings && calculationSettings.roundCellValues;

  const shareLikeRowTypes = [
    "proportion",
    "nps",
    "promoters",
    "detractors",
  ];

  const meanLikeRowTypes = [
    "mean",
    "standardDeviation",
    "variance",
  ];

  if (shareLikeRowTypes.includes(rowType)) {
    return shouldRoundCellValues ? 0 : 1;
  }

  if (meanLikeRowTypes.includes(rowType)) {
    return shouldRoundCellValues ? 1 : 2;
  }

  return null;
}