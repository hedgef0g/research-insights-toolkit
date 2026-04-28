/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

import {
  compareAllRowsUsingBottomBases,
  buildSignificanceMarkerMatrix,
  removeSignificanceMarkersFromText,
  removeSignificanceMarkersFromMatrix,
  compareMeansUsingSpreadAndBaseRows,
  compareNpsUsingStructureRows,
  compareNpsUsingSpreadAndBaseRows,
} from "../core/significance";

import {
  LABEL_SCAN_COLUMNS_LEFT,
  detectMetricRowsFromLeftLabels,
  formatMetricDetectionDiagnostics,
  buildAutoCalculationPlan,
} from "../core/metric-detector";

/**
 * Initializes task pane events after Office is ready.
 *
 * PURPOSE:
 * Connect the visible button in the Excel panel with our calculation logic.
 */
Office.onReady(() => {
  const calculateButton = document.getElementById("calculate-significance");
  const clearButton = document.getElementById("clear-significance");

  const calculateMeanSignificanceSdButton = document.getElementById(
    "calculate-mean-significance-sd"
  );

  const calculateMeanSignificanceVarianceButton = document.getElementById(
    "calculate-mean-significance-variance"
  );

  calculateButton.addEventListener("click", runSignificanceFromSelection);

  clearButton.addEventListener("click", clearSignificanceFromSelection);

  calculateMeanSignificanceSdButton.addEventListener("click", () =>
    runMeanSignificanceFromSelection("standardDeviation")
  );

  calculateMeanSignificanceVarianceButton.addEventListener("click", () =>
    runMeanSignificanceFromSelection("variance")
  );

  const calculateNpsStructureButton = document.getElementById(
    "calculate-nps-significance-structure"
  );
  
  const calculateNpsSdButton = document.getElementById(
    "calculate-nps-significance-sd"
  );
  
  const calculateNpsVarianceButton = document.getElementById(
    "calculate-nps-significance-variance"
  );
  
  calculateNpsStructureButton.addEventListener(
    "click",
    runNpsSignificanceFromStructureSelection
  );
  
  calculateNpsSdButton.addEventListener("click", () =>
    runNpsSignificanceFromSpreadSelection("standardDeviation")
  );

  calculateNpsVarianceButton.addEventListener("click", () =>
    runNpsSignificanceFromSpreadSelection("variance")
  );

  const detectMetricTypeButton = document.getElementById("detect-metric-type");

  detectMetricTypeButton.addEventListener(
    "click",
    runMetricDetectionDiagnostics
  );
});

/**
 * Reads selected Excel range, calculates pairwise significance,
 * and writes significance letters directly into value cells.
 *
 * MVP v0.3:
 * - Last selected row is treated as bases.
 * - All rows above are treated as values.
 * - Significant higher values receive labels of lower columns.
 */
async function runSignificanceFromSelection() {
  await Excel.run(async (context) => {
    const selectedRange = context.workbook.getSelectedRange(); // Current selected Excel range.

    selectedRange.load(["values", "text", "rowCount", "columnCount"]); // Load raw and displayed cell content.

    // Center all cells in the selected range after macro execution.
    selectedRange.format.horizontalAlignment = "Center";
    selectedRange.format.verticalAlignment = "Center";

    await context.sync();

    const selectedValues = selectedRange.values; // Raw values used for calculations.
    const cleanedValues = removeSignificanceMarkersFromMatrix(selectedValues);

    // Remove old significance markers before running a new calculation.
    selectedRange.values = cleanedValues;

    await context.sync();
    const selectedText = selectedRange.text; // Displayed values used for visible output.

    const outputElement = document.getElementById("significance-result"); // Result block in task pane.

        const leftLabelValues = await loadLeftLabelsForSelectedRange(
      context,
      selectedRange
    );

    const detectionResult = detectMetricRowsFromLeftLabels(
      selectedValues,
      leftLabelValues
    );

    const autoPlan = buildAutoCalculationPlan(detectionResult);

    if (
      !selectedValues ||
      selectedValues.length < 2 ||
      selectedValues[0].length < 2
    ) {
      outputElement.textContent =
        "Please select at least 2 columns and 2 rows. Last row must contain bases.";
      return;
    }

    if (autoPlan.metricType === "mean") {
  const allResults = compareMeansUsingSpreadAndBaseRows(
    cleanedValues,
    autoPlan.spreadType
  );

  const markerMatrix = buildSignificanceMarkerMatrix(allResults, 1);

  const columnCount = selectedValues[0].length;

  for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
    const markers = markerMatrix[0][columnIndex];

    if (!markers) {
      continue;
    }

    const currentCell = selectedRange.getCell(0, columnIndex);

    const displayedValueWithoutMarkers =
      removeSignificanceMarkersFromText(
        selectedText[0][columnIndex]
      );

    currentCell.values = [
      [`${displayedValueWithoutMarkers} ${markers}`.trim()],
    ];

    currentCell.format.font.bold = true;
    currentCell.format.fill.color = "#E2F0D9";
  }

  await context.sync();

  outputElement.textContent =
    autoPlan.spreadType === "standardDeviation"
      ? "Auto detected: Mean + SD"
      : "Auto detected: Mean + Variance";

  return;
}

    const allResults = compareAllRowsUsingBottomBases(cleanedValues);

    if (allResults === null) {
      outputElement.textContent = "Could not process selected range.";
      return;
    }

    const markerMatrix = buildSignificanceMarkerMatrix(allResults);

    const valueRowCount = allResults.baseRowIndex; // Number of rows above base row.
    const columnCount = selectedValues[0].length; // Number of selected columns.

    for (let rowIndex = 0; rowIndex < valueRowCount; rowIndex++) {
      for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
        const markers = markerMatrix[rowIndex][columnIndex]; // Letters to add to current cell.

        if (!markers) {
          continue;
        }

        const currentCell = selectedRange.getCell(rowIndex, columnIndex); // Cell that receives significance letters.
        const displayedValueWithoutMarkers = removeSignificanceMarkersFromText(
          selectedText[rowIndex][columnIndex]
        );
        
        currentCell.values = [[`${displayedValueWithoutMarkers} ${markers}`.trim()]];

        // Format cells where the value is significantly higher than at least one other column.
        currentCell.format.font.bold = true; // Make the whole cell text bold.
        currentCell.format.fill.color = "#E2F0D9"; // Pale green fill.
      }
    }

    await context.sync();

    outputElement.textContent = "Significance markers added to selected cells.";
  });
}

/**
 * Formats all pairwise comparison results into readable text for the task pane.
 *
 * PURPOSE:
 * Temporary MVP output.
 * Later we will replace this with table markers, colors, or letters.
 *
 * INPUT:
 * allResults - object returned by compareAllRowsUsingBottomBases().
 *
 * OUTPUT:
 * Multiline text for display in the Excel task pane.
 */
function formatAllComparisonsForDisplay(allResults) {
  const outputLines = []; // Final text lines for the task pane.

  outputLines.push("Pairwise significance results");
  outputLines.push(`Base row: ${allResults.baseRowIndex + 1}`);
  outputLines.push("");

  for (const comparisonRow of allResults.comparisonRows) {
    const displayedRowNumber = comparisonRow.valueRowIndex + 1; // Human-readable row number inside selection.

    outputLines.push(`Value row ${displayedRowNumber}:`);

    for (const comparison of comparisonRow.rowComparisons) {
      const firstColumnNumber = comparison.firstColumnIndex + 1; // Human-readable column number inside selection.
      const secondColumnNumber = comparison.secondColumnIndex + 1; // Human-readable column number inside selection.

      if (comparison.result === null) {
        outputLines.push(
          `  Col ${firstColumnNumber} vs Col ${secondColumnNumber}: skipped`
        );
        continue;
      }

      outputLines.push(
        `  Col ${firstColumnNumber} vs Col ${secondColumnNumber}: ` +
          `z=${comparison.result.zScore.toFixed(3)}, ` +
          `sig=${comparison.result.isSignificant ? "YES" : "NO"}, ` +
          `direction=${comparison.result.direction}`
      );
    }

    outputLines.push("");
  }

  return outputLines.join("\n");
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
 * Excel-specific wrapper for mean significance MVP.
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

    const valueRowCount = 1; // For mean MVP, only first row receives markers.
    const columnCount = selectedValues[0].length; // Number of selected columns.

    for (let rowIndex = 0; rowIndex < valueRowCount; rowIndex++) {
      for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
        const markers = markerMatrix[rowIndex][columnIndex]; // Marker letters for current mean cell.

        if (!markers) {
          continue;
        }

        const currentCell = selectedRange.getCell(rowIndex, columnIndex); // Mean cell receiving markers.

        const displayedValueWithoutMarkers = removeSignificanceMarkersFromText(
          selectedText[rowIndex][columnIndex]
        );

        currentCell.values = [[`${displayedValueWithoutMarkers} ${markers}`.trim()]];

        // Highlight significant winners only.
        currentCell.format.font.bold = true;
        currentCell.format.fill.color = "#E2F0D9";
      }
    }

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
    const outputElement = document.getElementById("significance-result");

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

    const valueRowCount = 1; // Only NPS row receives markers.
    const columnCount = selectedValues[0].length;

    for (let rowIndex = 0; rowIndex < valueRowCount; rowIndex++) {
      for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
        const markers = markerMatrix[rowIndex][columnIndex];

        if (!markers) {
          continue;
        }

        const currentCell = selectedRange.getCell(rowIndex, columnIndex);

        const displayedValueWithoutMarkers = removeSignificanceMarkersFromText(
          selectedText[rowIndex][columnIndex]
        );

        currentCell.values = [
          [`${displayedValueWithoutMarkers} ${markers}`.trim()],
        ];

        currentCell.format.font.bold = true;
        currentCell.format.fill.color = "#E2F0D9";
      }
    }

    await context.sync();

    outputElement.textContent =
      "NPS significance calculated using promoters and detractors.";
  });
}

/**
 * Reads selected Excel range and calculates NPS significance using SD or variance.
 *
 * PURPOSE:
 * This is NOT the same as mean significance.
 * NPS values must be normalized:
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
    const outputElement = document.getElementById("significance-result");

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

    // Remove old significance markers before recalculation.
    selectedRange.values = cleanedValues;

    // Center the entire selected range after macro execution.
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

    // Only the first row, the NPS row, should receive markers.
    const markerMatrix = buildSignificanceMarkerMatrix(allResults, 1);

    const columnCount = selectedValues[0].length; // Number of selected columns.

    for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      const markers = markerMatrix[0][columnIndex]; // Marker letters for current NPS cell.

      if (!markers) {
        continue;
      }

      const currentCell = selectedRange.getCell(0, columnIndex); // NPS cell receiving markers.

      const displayedValueWithoutMarkers = removeSignificanceMarkersFromText(
        selectedText[0][columnIndex]
      );

      currentCell.values = [
        [`${displayedValueWithoutMarkers} ${markers}`.trim()],
      ];

      // Highlight significant winners only.
      currentCell.format.font.bold = true;
      currentCell.format.fill.color = "#E2F0D9";
    }

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

    const labelStartColumnIndex =
      selectedStartColumnIndex - labelColumnCount; // First scanned label column.

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

  const selectedStartRowIndex = selectedRange.rowIndex;
  const selectedStartColumnIndex = selectedRange.columnIndex;
  const selectedRowCount = selectedRange.rowCount;

  if (selectedStartColumnIndex === 0) {
    return [];
  }

  const labelColumnCount = Math.min(
    LABEL_SCAN_COLUMNS_LEFT,
    selectedStartColumnIndex
  );

  const labelStartColumnIndex =
    selectedStartColumnIndex - labelColumnCount;

  const leftLabelRange = selectedRange.worksheet.getRangeByIndexes(
    selectedStartRowIndex,
    labelStartColumnIndex,
    selectedRowCount,
    labelColumnCount
  );

  leftLabelRange.load("values");

  await context.sync();

  return leftLabelRange.values;
}