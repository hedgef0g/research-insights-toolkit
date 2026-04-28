/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import {
  createEmptyMarkerMatrix,
  applyComparisonResultsToFullMarkerMatrix,
  keepMarkersOnlyInAllowedRows,
  compareProportionRowsUsingBaseRow,
  compareMeanBlockByRowIndexes,
  compareNpsStructureBlockByRowIndexes,
  compareNpsSpreadBlockByRowIndexes,
  removeSignificanceMarkersFromMatrix,
  buildSignificanceMarkerMatrix,
  removeSignificanceMarkersFromText,
  compareMeansUsingSpreadAndBaseRows,
  compareNpsUsingStructureRows,
  compareNpsUsingSpreadAndBaseRows,
} from "../core/significance";

import {
  LABEL_SCAN_COLUMNS_LEFT,
  detectMetricRowsFromLeftLabels,
  formatMetricDetectionDiagnostics,
  buildCalculationBlocks,
  getAllowedMarkerRowIndexes,
} from "../core/metric-detector";

/**
 * Initializes task pane events after Office is ready.
 *
 * PURPOSE:
 * Connect visible task pane buttons with calculation and utility logic.
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const sideloadMessage = document.getElementById("sideload-msg"); // Default Office template loading message.
    const appBody = document.getElementById("app-body"); // Main task pane body.

    if (sideloadMessage) {
      sideloadMessage.style.display = "none";
    }

    if (appBody) {
      appBody.style.display = "flex";
    }
  }

  const calculateButton = document.getElementById("calculate-significance"); // Unified auto-detection button.
  const clearButton = document.getElementById("clear-significance"); // Button for removing markers.

  const calculateMeanSignificanceSdButton = document.getElementById(
    "calculate-mean-significance-sd"
  ); // Explicit mean + SD button.

  const calculateMeanSignificanceVarianceButton = document.getElementById(
    "calculate-mean-significance-variance"
  ); // Explicit mean + variance button.

  const calculateNpsStructureButton = document.getElementById(
    "calculate-nps-significance-structure"
  ); // Explicit NPS + promoters/detractors button.

  const calculateNpsSdButton = document.getElementById(
    "calculate-nps-significance-sd"
  ); // Explicit NPS + SD button.

  const calculateNpsVarianceButton = document.getElementById(
    "calculate-nps-significance-variance"
  ); // Explicit NPS + variance button.

  const detectMetricTypeButton = document.getElementById("detect-metric-type"); // Diagnostic detector button.

  if (calculateButton) {
    calculateButton.addEventListener("click", runSignificanceFromSelection);
  }

  if (clearButton) {
    clearButton.addEventListener("click", clearSignificanceFromSelection);
  }

  if (calculateMeanSignificanceSdButton) {
    calculateMeanSignificanceSdButton.addEventListener("click", () =>
      runMeanSignificanceFromSelection("standardDeviation")
    );
  }

  if (calculateMeanSignificanceVarianceButton) {
    calculateMeanSignificanceVarianceButton.addEventListener("click", () =>
      runMeanSignificanceFromSelection("variance")
    );
  }

  if (calculateNpsStructureButton) {
    calculateNpsStructureButton.addEventListener(
      "click",
      runNpsSignificanceFromStructureSelection
    );
  }

  if (calculateNpsSdButton) {
    calculateNpsSdButton.addEventListener("click", () =>
      runNpsSignificanceFromSpreadSelection("standardDeviation")
    );
  }

  if (calculateNpsVarianceButton) {
    calculateNpsVarianceButton.addEventListener("click", () =>
      runNpsSignificanceFromSpreadSelection("variance")
    );
  }

  if (detectMetricTypeButton) {
    detectMetricTypeButton.addEventListener(
      "click",
      runMetricDetectionDiagnostics
    );
  }
});

/**
 * Reads selected Excel range, detects metric blocks, calculates pairwise significance,
 * and writes significance letters only into actual value rows.
 *
 * PURPOSE:
 * Unified auto mode for complex tables.
 * One selected range may contain proportions, means, and NPS blocks.
 *
 * SUPPORTED BLOCKS:
 * - Proportions + Base
 * - Mean + SD/Variance + Base
 * - NPS + Promoters + Detractors + Base
 * - NPS + SD/Variance + Base
 */
async function runSignificanceFromSelection() {
  await Excel.run(async (context) => {
    const selectedRange = context.workbook.getSelectedRange(); // Current selected Excel range.

    selectedRange.load(["values", "text", "rowCount", "columnCount"]); // Load raw and displayed cell content.

    // Center all cells in the selected range after macro execution.
    selectedRange.format.horizontalAlignment = "Center";
    selectedRange.format.verticalAlignment = "Center";

    await context.sync();

    const selectedValues = selectedRange.values; // Raw values used for initial validation and cleanup.
    const selectedText = selectedRange.text; // Displayed values used for preserving visible formatting.
    const outputElement = document.getElementById("significance-result"); // Result block in task pane.

    if (
      !selectedValues ||
      selectedValues.length < 2 ||
      selectedValues[0].length < 2
    ) {
      outputElement.textContent =
        "Please select at least 2 columns and 2 rows.";
      return;
    }

    const cleanedValues = removeSignificanceMarkersFromMatrix(selectedValues); // Values without old significance letters.

    // Remove old significance markers before running a new calculation.
    selectedRange.values = cleanedValues;

    await context.sync();

    const leftLabelValues = await loadLeftLabelsForSelectedRange(
      context,
      selectedRange
    ); // Labels located 1-2 columns to the left of the selected data.

    const detectionResult = detectMetricRowsFromLeftLabels(
      cleanedValues,
      leftLabelValues
    ); // Row type diagnostics based on left-side labels.

    const calculationBlocks = buildCalculationBlocks(detectionResult); // List of metric blocks to calculate.

    if (!calculationBlocks || calculationBlocks.length === 0) {
      outputElement.textContent = "Could not detect any calculation blocks.";
      return;
    }

    const fullMarkerMatrix = createEmptyMarkerMatrix(
      cleanedValues.length,
      cleanedValues[0].length
    ); // Full-size marker storage matching the selected range.

    for (const calculationBlock of calculationBlocks) {
      const blockResults = calculateBlockResults(
        cleanedValues,
        calculationBlock
      ); // Calculation result for current block.

      if (!blockResults) {
        continue;
      }

      // Add markers only to comparisonRows.valueRowIndex rows returned by the block calculation.
      applyComparisonResultsToFullMarkerMatrix(
        blockResults,
        fullMarkerMatrix
      );
    }

    const allowedMarkerRows = getAllowedMarkerRowIndexes(calculationBlocks); // Rows where marker letters are allowed.

    // Defensive cleanup: service rows must never receive significance markers.
    keepMarkersOnlyInAllowedRows(fullMarkerMatrix, allowedMarkerRows);

    // Write markers once, through one shared writer.
    writeMarkersToSelectedRange(selectedRange, selectedText, fullMarkerMatrix);

    await context.sync();

    outputElement.textContent = `Significance calculated for ${calculationBlocks.length} detected block(s).`;
  });
}

/**
 * Calculates one detected metric block.
 *
 * PURPOSE:
 * Keeps dispatcher logic out of runSignificanceFromSelection().
 * Each block type is routed to the correct core calculation function.
 */
function calculateBlockResults(cleanedValues, calculationBlock) {
  if (calculationBlock.metricType === "proportion") {
    return compareProportionRowsUsingBaseRow(
      cleanedValues,
      calculationBlock.valueRowIndexes,
      calculationBlock.baseRowIndex
    );
  }

  if (calculationBlock.metricType === "mean") {
    return compareMeanBlockByRowIndexes(
      cleanedValues,
      calculationBlock.valueRowIndex,
      calculationBlock.spreadRowIndex,
      calculationBlock.baseRowIndex,
      calculationBlock.spreadType
    );
  }

  if (calculationBlock.metricType === "npsStructure") {
    return compareNpsStructureBlockByRowIndexes(
      cleanedValues,
      calculationBlock.valueRowIndex,
      calculationBlock.promotersRowIndex,
      calculationBlock.detractorsRowIndex,
      calculationBlock.baseRowIndex
    );
  }

  if (calculationBlock.metricType === "npsSpread") {
    return compareNpsSpreadBlockByRowIndexes(
      cleanedValues,
      calculationBlock.valueRowIndex,
      calculationBlock.spreadRowIndex,
      calculationBlock.baseRowIndex,
      calculationBlock.spreadType
    );
  }

  return null;
}

/**
 * Removes significance markers and formatting from the selected range.
 *
 * PURPOSE:
 * User-facing cleanup button.
 */
async function clearSignificanceFromSelection() {
  await Excel.run(async (context) => {
    const selectedRange = context.workbook.getSelectedRange(); // Current selected Excel range.

    selectedRange.load("values"); // Load current cell values.

    await context.sync();

    const selectedValues = selectedRange.values; // Current values, possibly with markers.
    const cleanedValues = removeSignificanceMarkersFromMatrix(selectedValues); // Values without markers.

    selectedRange.values = cleanedValues;

    // Reset formatting applied by significance macro.
    selectedRange.format.font.bold = false;
    selectedRange.format.fill.clear();

    await context.sync();

    const outputElement = document.getElementById("significance-result");
    outputElement.textContent = "Significance markers removed.";
  });
}

/**
 * Reads selected Excel range and calculates pairwise significance for means.
 *
 * PURPOSE:
 * Explicit mean calculation button.
 * This legacy explicit mode remains useful for testing and manual override.
 *
 * EXPECTED SELECTION:
 * Row 1: means
 * Row 2: standard deviations OR variances
 * Row 3: bases
 *
 * INPUT:
 * spreadType - "standardDeviation" or "variance".
 */
async function runMeanSignificanceFromSelection(spreadType) {
  await Excel.run(async (context) => {
    const selectedRange = context.workbook.getSelectedRange(); // Current selected Excel range.

    selectedRange.load(["values", "text", "rowCount", "columnCount"]); // Load values and displayed text.

    await context.sync();

    const selectedValues = selectedRange.values; // Raw selected values.
    const selectedText = selectedRange.text; // Displayed selected values.
    const outputElement = document.getElementById("significance-result"); // Task pane output.

    if (
      !selectedValues ||
      selectedValues.length < 3 ||
      selectedValues[0].length < 2
    ) {
      outputElement.textContent =
        "Please select at least 3 rows and 2 columns: means, SD/variance, bases.";
      return;
    }

    const cleanedValues = removeSignificanceMarkersFromMatrix(selectedValues); // Remove old markers before recalculation.

    selectedRange.values = cleanedValues;

    selectedRange.format.horizontalAlignment = "Center";
    selectedRange.format.verticalAlignment = "Center";

    await context.sync();

    const allResults = compareMeansUsingSpreadAndBaseRows(
      cleanedValues,
      spreadType
    );

    if (allResults === null) {
      outputElement.textContent = "Could not process selected mean range.";
      return;
    }

    const markerMatrix = buildSignificanceMarkerMatrix(allResults, 1); // Significance letters for mean row.

    writeMarkersToSelectedRange(selectedRange, selectedText, markerMatrix);

    await context.sync();

    outputElement.textContent =
      spreadType === "standardDeviation"
        ? "Mean significance calculated using standard deviations."
        : "Mean significance calculated using variances.";
  });
}

/**
 * Reads selected Excel range and calculates NPS significance from structure.
 *
 * PURPOSE:
 * Explicit NPS structure calculation button.
 *
 * EXPECTED SELECTION:
 * Row 1: NPS
 * Row 2: Promoters %
 * Row 3: Detractors %
 * Row 4: Base
 */
async function runNpsSignificanceFromStructureSelection() {
  await Excel.run(async (context) => {
    const selectedRange = context.workbook.getSelectedRange(); // Current selected Excel range.

    selectedRange.load(["values", "text", "rowCount", "columnCount"]);

    await context.sync();

    const selectedValues = selectedRange.values; // Raw selected values.
    const selectedText = selectedRange.text; // Displayed selected values.
    const outputElement = document.getElementById("significance-result"); // Task pane output.

    if (
      !selectedValues ||
      selectedValues.length < 4 ||
      selectedValues[0].length < 2
    ) {
      outputElement.textContent =
        "Please select at least 4 rows and 2 columns: NPS, Promoters, Detractors, Base.";
      return;
    }

    const cleanedValues = removeSignificanceMarkersFromMatrix(selectedValues);

    selectedRange.values = cleanedValues;

    selectedRange.format.horizontalAlignment = "Center";
    selectedRange.format.verticalAlignment = "Center";

    await context.sync();

    const allResults = compareNpsUsingStructureRows(cleanedValues);

    if (allResults === null) {
      outputElement.textContent = "Could not process selected NPS range.";
      return;
    }

    const markerMatrix = buildSignificanceMarkerMatrix(allResults, 1);

    writeMarkersToSelectedRange(selectedRange, selectedText, markerMatrix);

    await context.sync();

    outputElement.textContent =
      "NPS significance calculated using promoters and detractors.";
  });
}

/**
 * Reads selected Excel range and calculates NPS significance using SD or variance.
 *
 * PURPOSE:
 * Explicit NPS spread calculation button.
 * This is NOT the same as mean significance because NPS values must be normalized:
 * - 40 becomes 0.40
 * - 0.40 stays 0.40
 *
 * EXPECTED SELECTION:
 * Row 1: NPS
 * Row 2: SD or variance
 * Row 3: Base
 */
async function runNpsSignificanceFromSpreadSelection(spreadType) {
  await Excel.run(async (context) => {
    const selectedRange = context.workbook.getSelectedRange(); // Current selected Excel range.

    selectedRange.load(["values", "text", "rowCount", "columnCount"]);

    await context.sync();

    const selectedValues = selectedRange.values; // Raw selected values.
    const selectedText = selectedRange.text; // Displayed selected values.
    const outputElement = document.getElementById("significance-result"); // Task pane output.

    if (
      !selectedValues ||
      selectedValues.length < 3 ||
      selectedValues[0].length < 2
    ) {
      outputElement.textContent =
        "Please select at least 3 rows and 2 columns: NPS, SD/variance, Base.";
      return;
    }

    const cleanedValues = removeSignificanceMarkersFromMatrix(selectedValues);

    selectedRange.values = cleanedValues;

    selectedRange.format.horizontalAlignment = "Center";
    selectedRange.format.verticalAlignment = "Center";

    await context.sync();

    const allResults = compareNpsUsingSpreadAndBaseRows(
      cleanedValues,
      spreadType
    );

    if (allResults === null) {
      outputElement.textContent = "Could not process selected NPS range.";
      return;
    }

    const markerMatrix = buildSignificanceMarkerMatrix(allResults, 1);

    writeMarkersToSelectedRange(selectedRange, selectedText, markerMatrix);

    await context.sync();

    outputElement.textContent =
      spreadType === "standardDeviation"
        ? "NPS significance calculated using standard deviations."
        : "NPS significance calculated using variances.";
  });
}

/**
 * Reads selected Excel range and scans labels to the left of it.
 *
 * PURPOSE:
 * Diagnostic-only step for auto metric detection.
 * This function does not calculate significance and does not change the sheet.
 */
async function runMetricDetectionDiagnostics() {
  await Excel.run(async (context) => {
    const selectedRange = context.workbook.getSelectedRange(); // User-selected data range.

    selectedRange.load(["values", "rowIndex", "columnIndex", "rowCount"]);

    await context.sync();

    const selectedValues = selectedRange.values; // Selected data values.
    const selectedStartRowIndex = selectedRange.rowIndex; // Zero-based first row of selected range.
    const selectedStartColumnIndex = selectedRange.columnIndex; // Zero-based first column of selected range.
    const selectedRowCount = selectedRange.rowCount; // Number of selected rows.

    const outputElement = document.getElementById("significance-result"); // Task pane output area.

    if (selectedStartColumnIndex === 0) {
      const detectionResult = detectMetricRowsFromLeftLabels(
        selectedValues,
        []
      );

      outputElement.textContent =
        formatMetricDetectionDiagnostics(detectionResult) +
        "\n\nNo columns exist to the left of the selected range. Default would be proportions.";

      return;
    }

    const labelColumnCount = Math.min(
      LABEL_SCAN_COLUMNS_LEFT,
      selectedStartColumnIndex
    ); // Scan up to 2 columns left, but not beyond sheet boundary.

    const labelStartColumnIndex = selectedStartColumnIndex - labelColumnCount; // First scanned label column.

    const worksheet = selectedRange.worksheet; // Worksheet containing selected range.

    const leftLabelRange = worksheet.getRangeByIndexes(
      selectedStartRowIndex,
      labelStartColumnIndex,
      selectedRowCount,
      labelColumnCount
    ); // Cells to the left of selected data.

    leftLabelRange.load("values");

    await context.sync();

    const leftLabelValues = leftLabelRange.values; // 2D array of left-side label values.

    const detectionResult = detectMetricRowsFromLeftLabels(
      selectedValues,
      leftLabelValues
    );

    outputElement.textContent =
      formatMetricDetectionDiagnostics(detectionResult);
  });
}

/**
 * Reads left-side labels for selected range.
 *
 * PURPOSE:
 * Shared helper for auto metric detection.
 */
async function loadLeftLabelsForSelectedRange(context, selectedRange) {
  selectedRange.load(["rowIndex", "columnIndex", "rowCount"]);

  await context.sync();

  const selectedStartRowIndex = selectedRange.rowIndex; // First selected row index, zero-based.
  const selectedStartColumnIndex = selectedRange.columnIndex; // First selected column index, zero-based.
  const selectedRowCount = selectedRange.rowCount; // Number of selected rows.

  if (selectedStartColumnIndex === 0) {
    return [];
  }

  const labelColumnCount = Math.min(
    LABEL_SCAN_COLUMNS_LEFT,
    selectedStartColumnIndex
  ); // Scan up to configured number of columns left, but not beyond sheet boundary.

  const labelStartColumnIndex = selectedStartColumnIndex - labelColumnCount; // First scanned label column.

  const leftLabelRange = selectedRange.worksheet.getRangeByIndexes(
    selectedStartRowIndex,
    labelStartColumnIndex,
    selectedRowCount,
    labelColumnCount
  ); // Cells left of selected range.

  leftLabelRange.load("values");

  await context.sync();

  return leftLabelRange.values;
}

/**
 * Writes significance markers into selected range.
 *
 * PURPOSE:
 * Shared writer for proportions, means, and NPS blocks.
 * It also protects cells from Excel auto-formatting values as time.
 */
function writeMarkersToSelectedRange(
  selectedRange,
  selectedText,
  markerMatrix
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

      currentCell.numberFormat = [["@"]]; // Force text format to prevent Excel time conversion.

      currentCell.values = [
        [`${displayedValueWithoutMarkers} ${markers}`.trim()],
      ];

      currentCell.format.font.bold = true;
      currentCell.format.fill.color = "#E2F0D9";
    }
  }
}
