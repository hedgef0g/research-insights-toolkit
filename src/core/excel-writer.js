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

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      const cellResult = cellResultMatrix[rowIndex][columnIndex];

      if (!cellResult) {
        continue;
      }

      const markers = cellResult.markers || "";
      const previousColumnArrow = cellResult.previousColumnArrow || "";
      const fillReason = cellResult.fillReason || "none";

      if (!markers && !previousColumnArrow && fillReason === "none") {
        continue;
      }

      const currentCell = selectedRange.getCell(rowIndex, columnIndex);

      const outputMarkerText = previousColumnArrow || markers;

      if (outputMarkerText) {
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
        currentCell.values = [[`${roundedDisplayedValue} ${outputMarkerText}`.trim()]];

        if (previousColumnArrow) {
          applyPreviousColumnArrowFontColorExperimental(
            currentCell,
            cellResult,
            calculationSettings
          );
        }
      }

      const fillColor = getFillColorForCellResult(cellResult, calculationSettings);

      if (fillColor) {
        currentCell.format.fill.color = fillColor;
      }

      if (
        markers ||
        previousColumnArrow ||
        fillReason === "significant" ||
        fillReason === "lowerThanTotal"
      ) {
        currentCell.format.font.bold = true;
      }
    }
  }
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
  detectionResult,
  calculationSettings
) {
  const parsedValue = parseOutputNumber(displayedValue);

  if (parsedValue === null) {
    return displayedValue;
  }

  const rowType = getDetectedRowTypeByIndex(detectionResult, rowIndex);
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
 * Tries to color only the previous-column arrow.
 *
 * IMPORTANT:
 * This is experimental. Some Office.js Excel runtimes may not support
 * character-level formatting inside a cell. If unsupported, we silently
 * fall back to leaving the arrow uncolored instead of breaking the calculation.
 */
function applyPreviousColumnArrowFontColorExperimental(
  currentCell,
  cellResult,
  calculationSettings
) {
  const arrowDirection = cellResult.previousColumnArrowDirection;

  if (!arrowDirection) {
    return;
  }

  const arrowColor =
    arrowDirection === "up"
      ? calculationSettings.significantFillColor || "#70AD47"
      : calculationSettings.lowerThanTotalFillColor || "#C00000";

  try {
    // Placeholder for future rich-text implementation.
    // Standard Excel.Range formatting applies to the whole cell, not a substring.
    // Do not use currentCell.format.font.color here unless you accept coloring
    // the entire cell value.
    //
    // If a reliable rich-text API becomes available in the target runtime,
    // implement it here only, without changing significance.js.
    void currentCell;
    void arrowColor;
  } catch (error) {
    console.warn("Arrow character coloring is not supported in this Excel runtime.", error);
  }
}
