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
  generateSignificanceLabels,
} from "../core/significance";

import {
  LABEL_SCAN_COLUMNS_LEFT,
  detectMetricRowsFromLeftLabels,
  formatMetricDetectionDiagnostics,
  buildCalculationBlocks,
  getAllowedMarkerRowIndexes,
} from "../core/metric-detector";

import { writeMarkersToSelectedRange } from "../core/excel-writer";

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
  const detectMetricTypeButton = document.getElementById("detect-metric-type"); // Diagnostic detector button.

  initializeSettingsPanel();

  if (calculateButton) {
    calculateButton.addEventListener("click", runSignificanceFromSelection);
  }

  if (clearButton) {
    clearButton.addEventListener("click", clearSignificanceFromSelection);
  }

  if (detectMetricTypeButton) {
    detectMetricTypeButton.addEventListener("click", runMetricDetectionDiagnostics);
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
    let selectedRange = context.workbook.getSelectedRange();

    const outputElement = document.getElementById("significance-result");
    const calculationSettings = readCalculationSettingsFromPanel();

    selectedRange.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);

    await context.sync();

    if (calculationSettings.writeBannerLetters && selectedRange.rowIndex === 0) {
      outputElement.textContent =
        "Данные расположены в первой строке. Добавьте строку над выделенным массивом и запустите расчёт повторно.";

      return;
    }

    selectedRange.load(["values", "text", "rowIndex", "columnIndex", "rowCount", "columnCount"]);

    selectedRange.format.horizontalAlignment = "Center";
    selectedRange.format.verticalAlignment = "Center";

    await context.sync();

    const selectedValues = selectedRange.values;
    const selectedText = selectedRange.text;

    if (!selectedValues || selectedValues.length < 2 || selectedValues[0].length < 2) {
      outputElement.textContent = "Please select at least 2 columns and 2 rows.";
      return;
    }

    const cleanedValues = removeSignificanceMarkersFromMatrix(selectedValues);

    selectedRange.values = cleanedValues;

    selectedRange.format.font.bold = false;
    selectedRange.format.fill.clear();
    selectedRange.format.horizontalAlignment = "Center";
    selectedRange.format.verticalAlignment = "Center";

    await context.sync();

    const leftLabelValues = await loadLabelValuesForSelectedRange(
      context,
      selectedRange,
      calculationSettings
    ); // Labels located 1-2 columns to the left of the selected data.

    const detectionResult = detectMetricRowsFromLeftLabels(cleanedValues, leftLabelValues); // Row type diagnostics based on left-side labels.

    const calculationBlocks = buildCalculationBlocks(detectionResult); // List of metric blocks to calculate.

    if (!calculationBlocks || calculationBlocks.length === 0) {
      outputElement.textContent = "Could not detect any calculation blocks.";
      return;
    }

    const fullMarkerMatrix = createEmptyMarkerMatrix(cleanedValues.length, cleanedValues[0].length); // Full-size marker storage matching the selected range.

    for (const calculationBlock of calculationBlocks) {
      const blockResults = calculateBlockResults(
        cleanedValues,
        calculationBlock,
        calculationSettings
      ); // Calculation result for current block.

      if (!blockResults) {
        continue;
      }

      // Add markers only to comparisonRows.valueRowIndex rows returned by the block calculation.
      applyComparisonResultsToFullMarkerMatrix(blockResults, fullMarkerMatrix);
    }

    const allowedMarkerRows = getAllowedMarkerRowIndexes(calculationBlocks); // Rows where marker letters are allowed.

    // Defensive cleanup: service rows must never receive significance markers.
    keepMarkersOnlyInAllowedRows(fullMarkerMatrix, allowedMarkerRows);

    // Write markers once, through one shared writer.
    writeMarkersToSelectedRange(
      selectedRange,
      selectedText,
      fullMarkerMatrix,
      detectionResult,
      calculationSettings
    );

    if (calculationSettings.writeBannerLetters) {
      await writeBannerMarkersAboveSelectedRange(context, selectedRange, calculationSettings);
    }

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
function calculateBlockResults(cleanedValues, calculationBlock, calculationSettings) {
  if (calculationBlock.metricType === "proportion") {
    return compareProportionRowsUsingBaseRow(
      cleanedValues,
      calculationBlock.valueRowIndexes,
      calculationBlock.baseRowIndex,
      calculationSettings
    );
  }

  if (calculationBlock.metricType === "mean") {
    return compareMeanBlockByRowIndexes(
      cleanedValues,
      calculationBlock.valueRowIndex,
      calculationBlock.spreadRowIndex,
      calculationBlock.baseRowIndex,
      calculationBlock.spreadType,
      calculationSettings
    );
  }

  if (calculationBlock.metricType === "npsStructure") {
    return compareNpsStructureBlockByRowIndexes(
      cleanedValues,
      calculationBlock.valueRowIndex,
      calculationBlock.promotersRowIndex,
      calculationBlock.detractorsRowIndex,
      calculationBlock.baseRowIndex,
      calculationSettings
    );
  }

  if (calculationBlock.metricType === "npsSpread") {
    return compareNpsSpreadBlockByRowIndexes(
      cleanedValues,
      calculationBlock.valueRowIndex,
      calculationBlock.spreadRowIndex,
      calculationBlock.baseRowIndex,
      calculationBlock.spreadType,
      calculationSettings
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
 * Reads selected Excel range and scans labels to the left of it.
 *
 * PURPOSE:
 * Diagnostic-only step for auto metric detection.
 * This function does not calculate significance and does not change the sheet.
 */
async function runMetricDetectionDiagnostics() {
  await Excel.run(async (context) => {
    const calculationSettings = readCalculationSettingsFromPanel();
    const selectedRange = context.workbook.getSelectedRange(); // User-selected data range.

    selectedRange.load(["values", "rowIndex", "columnIndex", "rowCount"]);

    await context.sync();

    const selectedValues = selectedRange.values; // Selected data values.
    const selectedStartRowIndex = selectedRange.rowIndex; // Zero-based first row of selected range.
    const selectedStartColumnIndex = selectedRange.columnIndex; // Zero-based first column of selected range.
    const selectedRowCount = selectedRange.rowCount; // Number of selected rows.

    const outputElement = document.getElementById("significance-result"); // Task pane output area.

    if (selectedStartColumnIndex === 0) {
      const detectionResult = detectMetricRowsFromLeftLabels(selectedValues, []);

      outputElement.textContent =
        formatMetricDetectionDiagnostics(detectionResult) +
        "\n\nNo columns exist to the left of the selected range. Default would be proportions.";

      return;
    }

    const labelColumnCount = Math.min(LABEL_SCAN_COLUMNS_LEFT, selectedStartColumnIndex); // Scan up to 2 columns left, but not beyond sheet boundary.

    const leftLabelValues = await loadLabelValuesForSelectedRange(
      context,
      selectedRange,
      calculationSettings
    );

    const labelStartColumnIndex = selectedStartColumnIndex - labelColumnCount; // First scanned label column.

    const worksheet = selectedRange.worksheet; // Worksheet containing selected range.

    await context.sync();

    const detectionResult = detectMetricRowsFromLeftLabels(selectedValues, leftLabelValues);

    outputElement.textContent = formatMetricDetectionDiagnostics(detectionResult);
  });
}

/**
 * Reads label values for selected range.
 *
 * PURPOSE:
 * By default, reads labels immediately to the left of selected data.
 * If labelsOnLeftSide is enabled, reads labels from the leftmost sheet columns.
 */
async function loadLabelValuesForSelectedRange(context, selectedRange, calculationSettings) {
  if (calculationSettings.labelsOnLeftSide) {
    return loadLabelsFromLeftSideOfSheet(context, selectedRange);
  }

  return loadLabelsImmediatelyLeftOfSelection(context, selectedRange);
}

/**
 * Reads labels immediately to the left of selected range.
 *
 * PURPOSE:
 * Default detection mode.
 */
async function loadLabelsImmediatelyLeftOfSelection(context, selectedRange) {
  selectedRange.load(["rowIndex", "columnIndex", "rowCount"]);

  await context.sync();

  const selectedStartRowIndex = selectedRange.rowIndex;
  const selectedStartColumnIndex = selectedRange.columnIndex;
  const selectedRowCount = selectedRange.rowCount;

  if (selectedStartColumnIndex === 0) {
    return [];
  }

  const labelColumnCount = Math.min(LABEL_SCAN_COLUMNS_LEFT, selectedStartColumnIndex);

  const labelStartColumnIndex = selectedStartColumnIndex - labelColumnCount;

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

/**
 * Reads labels from the leftmost columns of the worksheet.
 *
 * PURPOSE:
 * Supports wide horizontal tables where metric labels stay on the far left,
 * while the user selects data columns far to the right.
 */
async function loadLabelsFromLeftSideOfSheet(context, selectedRange) {
  selectedRange.load(["rowIndex", "rowCount"]);

  await context.sync();

  const selectedStartRowIndex = selectedRange.rowIndex;
  const selectedRowCount = selectedRange.rowCount;

  const leftLabelRange = selectedRange.worksheet.getRangeByIndexes(
    selectedStartRowIndex,
    0,
    selectedRowCount,
    LABEL_SCAN_COLUMNS_LEFT
  );

  leftLabelRange.load("values");

  await context.sync();

  return leftLabelRange.values;
}

/**
 * Reads calculation settings from the task pane UI.
 *
 * PURPOSE:
 * Centralized UI settings object for current and future calculation logic.
 */
function readCalculationSettingsFromPanel() {
  const confidenceLevelElement = document.getElementById("confidence-level");

  const smallBaseThresholdElement = document.getElementById("small-base-threshold");

  return {
    confidenceLevel: confidenceLevelElement ? confidenceLevelElement.value : "95",

    roundCellValues: getCheckboxValue("round-cell-values"),

    writeBannerLetters: getCheckboxValue("write-banner-letters"),
    respectBannerStructure: getCheckboxValue("respect-banner-structure"),
    labelsOnLeftSide: getCheckboxValue("labels-on-left-side"),

    compareOnlyWithTotal: getCheckboxValue("compare-only-with-total"),
    excludeTotalFromComparisons: getCheckboxValue("exclude-total-from-comparisons"),

    firstColumnIsTotal: getCheckboxValue("first-column-is-total"),
    totalInEachBanner: getCheckboxValue("total-in-each-banner"),

    significantFillColor: getInputValue("significant-fill-color", "#E2F0D9"),

    lowerThanTotalFillColor: getInputValue("lower-than-total-fill-color", "#FCE4D6"),

    fillOnlyTotalComparisons: getCheckboxValue("fill-only-total-comparisons"),

    excludeSmallBasesFromComparisons: getCheckboxValue("exclude-small-bases"),

    smallBaseThreshold: smallBaseThresholdElement ? Number(smallBaseThresholdElement.value) : 50,

    smallBaseFillColor: getInputValue("small-base-fill-color", "#d0d0d0"),

    settingsStorageMode: getInputValue("settings-storage-mode", "none"),
  };
}

/**
 * Reads checkbox value safely.
 */
function getCheckboxValue(elementId) {
  const element = document.getElementById(elementId);
  return element ? element.checked : false;
}

/**
 * Reads input value safely.
 */
function getInputValue(elementId, fallbackValue) {
  const element = document.getElementById(elementId);
  return element ? element.value : fallbackValue;
}

/**
 * Initializes settings panel UI behavior.
 *
 * PURPOSE:
 * Handles mutually exclusive checkbox groups.
 */
function initializeSettingsPanel() {
  bindMutuallyExclusiveCheckboxes("compare-only-with-total", "exclude-total-from-comparisons");

  bindMutuallyExclusiveCheckboxes("first-column-is-total", "total-in-each-banner");

  const helpLink = document.getElementById("help-link");

  if (helpLink) {
    helpLink.addEventListener("click", (event) => {
      event.preventDefault();
      window.open("https://github.com/YOUR_USER/YOUR_REPO#readme", "_blank");
    });
  }

  initializeSettingsToggle();
}

/**
 * Makes two checkboxes mutually exclusive.
 *
 * PURPOSE:
 * If one checkbox is enabled, the other one is disabled automatically.
 */
function bindMutuallyExclusiveCheckboxes(firstCheckboxId, secondCheckboxId) {
  const firstCheckbox = document.getElementById(firstCheckboxId);
  const secondCheckbox = document.getElementById(secondCheckboxId);

  if (!firstCheckbox || !secondCheckbox) {
    return;
  }

  firstCheckbox.addEventListener("change", () => {
    if (firstCheckbox.checked) {
      secondCheckbox.checked = false;
    }
  });

  secondCheckbox.addEventListener("change", () => {
    if (secondCheckbox.checked) {
      firstCheckbox.checked = false;
    }
  });
}

/**
 * Initializes collapsible settings panel.
 *
 * PURPOSE:
 * Allows user to collapse the settings block and keep the panel compact.
 */
function initializeSettingsToggle() {
  const settingsToggle = document.getElementById("settings-toggle");
  const settingsContent = document.getElementById("settings-content");
  const settingsToggleIcon = document.getElementById("settings-toggle-icon");

  if (!settingsToggle || !settingsContent || !settingsToggleIcon) {
    return;
  }

  settingsToggle.addEventListener("click", () => {
    const isCollapsed = settingsContent.classList.toggle("collapsed");

    settingsToggleIcon.textContent = isCollapsed ? "▸" : "▾";
  });
}

/**
 * Writes column significance labels into the banner row above selected range.
 *
 * PURPOSE:
 * If enabled, every selected data column receives its significance marker
 * in the cell directly above the selected range.
 */
async function writeBannerMarkersAboveSelectedRange(context, selectedRange, calculationSettings) {
  if (!calculationSettings.writeBannerLetters) {
    return;
  }

  const selectedStartRowIndex = selectedRange.rowIndex;
  const selectedStartColumnIndex = selectedRange.columnIndex;
  const selectedColumnCount = selectedRange.columnCount;

  if (selectedStartRowIndex === 0) {
    return;
  }

  const significanceLabels = generateSignificanceLabels();

  const bannerRange = selectedRange.worksheet.getRangeByIndexes(
    selectedStartRowIndex - 1,
    selectedStartColumnIndex,
    1,
    selectedColumnCount
  );

  bannerRange.load("values");

  await context.sync();

  const bannerValues = bannerRange.values[0];

  const updatedBannerValues = bannerValues.map((currentValue, columnIndex) => {
    const marker = significanceLabels[columnIndex];

    if (!marker) {
      return currentValue;
    }

    return updateBannerCellMarker(currentValue, marker);
  });

  bannerRange.values = [updatedBannerValues];
}

/**
 * Adds or replaces significance marker at the end of a banner cell.
 *
 * PURPOSE:
 * Banner cells may already contain an old marker like "Segment A (/b/)".
 * We replace old marker with the current marker for this column.
 */
function updateBannerCellMarker(rawValue, marker) {
  const textValue = rawValue === null || rawValue === undefined ? "" : String(rawValue);

  const expectedMarkerSuffix = `(${marker})`;

  if (textValue.trim().endsWith(expectedMarkerSuffix)) {
    return textValue;
  }

  const markerSuffixPattern = /\s*\([^()]+\)$/;

  const textWithoutOldMarker = textValue.replace(markerSuffixPattern, "").trim();

  if (!textWithoutOldMarker) {
    return expectedMarkerSuffix;
  }

  return `${textWithoutOldMarker} ${expectedMarkerSuffix}`;
}
