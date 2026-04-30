/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import {
  createEmptyCellResultMatrix,
  applyComparisonResultsToFullCellResultMatrix,
  applySmallBaseRulesForCalculationBlock,
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

import { writeCellResultsToSelectedRange } from "../core/excel-writer";

const LOCAL_SETTINGS_STORAGE_KEY = "rit.settings.v1";

const SETTINGS_CONTROL_CONFIG = [
  { id: "confidence-level", type: "value", settingName: "confidenceLevel" },

  { id: "round-cell-values", type: "checked", settingName: "roundCellValues" },

  {
    id: "compare-with-previous-column",
    type: "checked",
    settingName: "compareWithPreviousColumn",
  },
  {
    id: "apply-previous-column-fill",
    type: "checked",
    settingName: "applyPreviousColumnFill",
  },

  { id: "write-banner-letters", type: "checked", settingName: "writeBannerLetters" },
  { id: "respect-banner-structure", type: "checked", settingName: "respectBannerStructure" },
  { id: "labels-on-left-side", type: "checked", settingName: "labelsOnLeftSide" },

  { id: "compare-only-with-total", type: "checked", settingName: "compareOnlyWithTotal" },
  {
    id: "exclude-total-from-comparisons",
    type: "checked",
    settingName: "excludeTotalFromComparisons",
  },

  { id: "first-column-is-total", type: "checked", settingName: "firstColumnIsTotal" },
  { id: "total-in-each-banner", type: "checked", settingName: "totalInEachBanner" },

  { id: "significant-fill-color", type: "value", settingName: "significantFillColor" },
  {
    id: "lower-than-total-fill-color",
    type: "value",
    settingName: "lowerThanTotalFillColor",
  },

  {
    id: "fill-only-total-comparisons",
    type: "checked",
    settingName: "fillOnlyTotalComparisons",
  },

  {
    id: "exclude-small-bases",
    type: "checked",
    settingName: "excludeSmallBasesFromComparisons",
  },
  { id: "small-base-threshold", type: "number", settingName: "smallBaseThreshold" },
  { id: "small-base-fill-color", type: "value", settingName: "smallBaseFillColor" },

  { id: "settings-storage-mode", type: "value", settingName: "settingsStorageMode" },
];

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
  loadSavedSettingsIntoPanel();
  refreshSettingsPanelState();
  initializeSettingsPersistence();

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
    if (calculationSettings.compareWithPreviousColumn && calculationSettings.compareOnlyWithTotal) {
      outputElement.textContent =
        "Режим “Сравнение с предыдущей колонкой” несовместим с режимом “Сравнивать только с Тотал”.";

      return;
    }

    if (
      calculationSettings.compareWithPreviousColumn &&
      calculationSettings.excludeTotalFromComparisons &&
      !calculationSettings.firstColumnIsTotal
    ) {
      outputElement.textContent =
        "Для режима “Не сравнивать с Тотал” нужно указать расположение Тотала. Сейчас поддерживается только вариант “Первая колонка — Тотал”.";

      return;
    }
    if (calculationSettings.compareOnlyWithTotal && !calculationSettings.firstColumnIsTotal) {
      outputElement.textContent =
        "Для режима “Сравнивать только с Тотал” нужно указать расположение Тотала. Сейчас поддерживается только вариант “Первая колонка — Тотал”.";

      return;
    }

    if (
      calculationSettings.excludeTotalFromComparisons &&
      !calculationSettings.firstColumnIsTotal
    ) {
      outputElement.textContent =
        "Для режима “Не сравнивать с Тотал” нужно указать расположение Тотала. Сейчас поддерживается только вариант “Первая колонка — Тотал”.";

      return;
    }

    selectedRange.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);

    await context.sync();

    if (calculationSettings.writeBannerLetters && selectedRange.rowIndex === 0) {
      outputElement.textContent =
        "Данные расположены в первой строке. Добавьте строку над выделенным массивом и запустите расчёт повторно.";

      return;
    }

    if (
      calculationSettings.compareWithPreviousColumn &&
      calculationSettings.fillOnlyTotalComparisons
    ) {
      outputElement.textContent =
        "Режим “Сравнение с предыдущей колонкой” несовместим с настройкой “Заливка только для Тотала”.";

      return;
    }

    selectedRange.load(["values", "text", "rowIndex", "columnIndex", "rowCount", "columnCount"]);

    selectedRange.format.horizontalAlignment = "Center";
    selectedRange.format.verticalAlignment = "Center";

    await context.sync();

    const selectedValues = selectedRange.values;
    const selectedText = selectedRange.text;

    if (!selectedValues || selectedValues.length < 2 || selectedValues[0].length < 2) {
      setStatusMessage("Please select at least 2 columns and 2 rows.");
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
      setStatusMessage("Could not detect any calculation blocks.");
      return;
    }

    const fullCellResultMatrix = createEmptyCellResultMatrix(
      cleanedValues.length,
      cleanedValues[0].length
    ); // Full-size marker storage matching the selected range.

    for (const calculationBlock of calculationBlocks) {
      const smallBaseResult = applySmallBaseRulesForCalculationBlock(
        cleanedValues,
        calculationBlock,
        fullCellResultMatrix,
        calculationSettings
      );

      if (smallBaseResult.errorMessage) {
        setStatusMessage(smallBaseResult.errorMessage);
        return;
      }

      const blockCalculationSettings = {
        ...calculationSettings,
        excludedColumnIndexes: smallBaseResult.excludedColumnIndexes,
      };

      const blockResults = calculateBlockResults(
        cleanedValues,
        calculationBlock,
        blockCalculationSettings
      );

      if (!blockResults) {
        continue;
      }

      applyComparisonResultsToFullCellResultMatrix(
        blockResults,
        fullCellResultMatrix,
        blockCalculationSettings
      );
    }

    const allowedMarkerRows = getAllowedMarkerRowIndexes(calculationBlocks);

    keepMarkersOnlyInAllowedRows(fullCellResultMatrix, allowedMarkerRows);

    writeCellResultsToSelectedRange(
      selectedRange,
      selectedText,
      fullCellResultMatrix,
      detectionResult,
      calculationSettings
    );

    if (calculationSettings.writeBannerLetters) {
      await writeBannerMarkersAboveSelectedRange(context, selectedRange, calculationSettings);
    }

    await context.sync();

    setStatusMessage(`Significance calculated for ${calculationBlocks.length} detected block(s).`);
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
    setStatusMessage("Significance markers removed.");
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
    const selectedStartColumnIndex = selectedRange.columnIndex; // Zero-based first column of selected range.

    const outputElement = document.getElementById("significance-result"); // Task pane output area.

    if (selectedStartColumnIndex === 0 && !calculationSettings.labelsOnLeftSide) {
      const detectionResult = detectMetricRowsFromLeftLabels(selectedValues, []);

      outputElement.textContent =
        formatMetricDetectionDiagnostics(detectionResult) +
        "\n\nNo columns exist to the left of the selected range. Default would be proportions.";

      return;
    }

    const leftLabelValues = await loadLabelValuesForSelectedRange(
      context,
      selectedRange,
      calculationSettings
    );

    const detectionResult = detectMetricRowsFromLeftLabels(selectedValues, leftLabelValues);

    setStatusMessage(formatMetricDetectionDiagnostics(detectionResult));
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

    compareWithPreviousColumn: getCheckboxValue("compare-with-previous-column"),
    applyPreviousColumnFill: getCheckboxValue("apply-previous-column-fill"),

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

  initializePreviousColumnComparisonSettings();

  const helpLink = document.getElementById("help-link");

  if (helpLink) {
    helpLink.addEventListener("click", (event) => {
      event.preventDefault();
      window.open(
        "https://github.com/hedgef0g/research-insights-toolkit/blob/main/README.md",
        "_blank"
      );
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
    if (calculationSettings.firstColumnIsTotal && columnIndex === 0) {
      return removeBannerCellMarker(currentValue);
    }

    const markerIndex = calculationSettings.firstColumnIsTotal ? columnIndex - 1 : columnIndex;

    const marker = significanceLabels[markerIndex];

    if (!marker) {
      return currentValue;
    }

    return updateBannerCellMarker(currentValue, marker);
  });

  /**
   * Removes significance marker from the end of a banner cell.
   *
   * PURPOSE:
   * In firstColumnIsTotal mode, the Total banner cell must not have
   * a significance marker.
   */
  function removeBannerCellMarker(rawValue) {
    const textValue = rawValue === null || rawValue === undefined ? "" : String(rawValue);

    const markerSuffixPattern = /\s*\([^()]+\)$/;

    return textValue.replace(markerSuffixPattern, "").trim();
  }

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

/**
 * Initializes Previous Column comparison UI behavior.
 *
 * RULES:
 * - Previous-column mode disables banner letters.
 * - Previous-column mode is incompatible with compare-only-with-total.
 * - exclude-total-from-comparisons is available in previous-column mode
 *   only when first-column-is-total is enabled.
 * - Previous-column fill option is visible only in previous-column mode.
 * - Warning is shown when previous-column + first-column-is-total is enabled
 *   and exclude-total-from-comparisons is not enabled.
 */
function initializePreviousColumnComparisonSettings() {
  const previousColumnCheckbox = document.getElementById("compare-with-previous-column");
  const firstColumnIsTotalCheckbox = document.getElementById("first-column-is-total");
  const excludeTotalCheckbox = document.getElementById("exclude-total-from-comparisons");

  if (previousColumnCheckbox) {
    previousColumnCheckbox.addEventListener("change", refreshSettingsPanelState);
  }

  if (firstColumnIsTotalCheckbox) {
    firstColumnIsTotalCheckbox.addEventListener("change", refreshSettingsPanelState);
  }

  if (excludeTotalCheckbox) {
    excludeTotalCheckbox.addEventListener("change", refreshSettingsPanelState);
  }

  refreshSettingsPanelState();
}

function setStatusMessage(message) {
  const statusPanel = document.getElementById("status-panel");
  const outputElement = document.getElementById("significance-result");

  if (statusPanel) {
    statusPanel.style.display = "block";
  }

  if (outputElement) {
    outputElement.textContent = message || "";
  }
}

/**
 * Refreshes dependent settings UI state after user changes or saved settings load.
 */
function refreshSettingsPanelState() {
  refreshPreviousColumnComparisonState();
}

/**
 * Applies Previous Column comparison UI rules.
 */
function refreshPreviousColumnComparisonState() {
  const previousColumnCheckbox = document.getElementById("compare-with-previous-column");
  const previousColumnFillWrapper = document.getElementById("previous-column-fill-wrapper");
  const previousColumnFillCheckbox = document.getElementById("apply-previous-column-fill");

  const writeBannerLettersCheckbox = document.getElementById("write-banner-letters");

  const compareOnlyWithTotalCheckbox = document.getElementById("compare-only-with-total");
  const excludeTotalCheckbox = document.getElementById("exclude-total-from-comparisons");
  const firstColumnIsTotalCheckbox = document.getElementById("first-column-is-total");

  const fillOnlyTotalComparisonsCheckbox = document.getElementById("fill-only-total-comparisons");

  const warningElement = document.getElementById("previous-column-total-warning");

  if (!previousColumnCheckbox) {
    return;
  }

  const isPreviousColumnMode = previousColumnCheckbox.checked;
  const firstColumnIsTotal = firstColumnIsTotalCheckbox
    ? firstColumnIsTotalCheckbox.checked
    : false;

  if (previousColumnFillWrapper) {
    previousColumnFillWrapper.style.display = isPreviousColumnMode ? "block" : "none";
  }

  if (!isPreviousColumnMode && previousColumnFillCheckbox) {
    previousColumnFillCheckbox.checked = false;
  }

  if (writeBannerLettersCheckbox) {
    if (isPreviousColumnMode) {
      writeBannerLettersCheckbox.checked = false;
      writeBannerLettersCheckbox.disabled = true;
    } else {
      writeBannerLettersCheckbox.disabled = false;
    }
  }

  if (compareOnlyWithTotalCheckbox) {
    if (isPreviousColumnMode) {
      compareOnlyWithTotalCheckbox.checked = false;
      compareOnlyWithTotalCheckbox.disabled = true;
    } else {
      compareOnlyWithTotalCheckbox.disabled = false;
    }
  }

  if (fillOnlyTotalComparisonsCheckbox) {
    if (isPreviousColumnMode) {
      fillOnlyTotalComparisonsCheckbox.checked = false;
      fillOnlyTotalComparisonsCheckbox.disabled = true;
    } else {
      fillOnlyTotalComparisonsCheckbox.disabled = false;
    }
  }

  if (excludeTotalCheckbox) {
    if (isPreviousColumnMode && !firstColumnIsTotal) {
      excludeTotalCheckbox.checked = false;
      excludeTotalCheckbox.disabled = true;
    } else {
      excludeTotalCheckbox.disabled = false;
    }
  }

  if (warningElement) {
    const shouldShowWarning =
      isPreviousColumnMode &&
      firstColumnIsTotal &&
      excludeTotalCheckbox &&
      !excludeTotalCheckbox.checked;

    warningElement.style.display = shouldShowWarning ? "block" : "none";
  }
}

/**
 * Initializes local settings persistence.
 *
 * RULES:
 * - settings-storage-mode = local: save all settings on every change.
 * - settings-storage-mode = none: delete saved settings.
 * - cloud is reserved for future implementation.
 */
function initializeSettingsPersistence() {
  for (const controlConfig of SETTINGS_CONTROL_CONFIG) {
    const element = document.getElementById(controlConfig.id);

    if (!element) {
      continue;
    }

    element.addEventListener("change", () => {
      refreshSettingsPanelState();
      handleSettingsPersistenceAfterChange();
    });

    element.addEventListener("input", () => {
      if (element.type === "color" || element.type === "number") {
        refreshSettingsPanelState();
        handleSettingsPersistenceAfterChange();
      }
    });
  }
}

/**
 * Saves or clears settings depending on selected storage mode.
 */
function handleSettingsPersistenceAfterChange() {
  const settings = readCalculationSettingsFromPanel();

  if (settings.settingsStorageMode === "local") {
    saveSettingsToLocalStorage(settings);
    return;
  }

  if (settings.settingsStorageMode === "none") {
    clearSavedLocalSettings();
  }
}

/**
 * Loads saved settings into the panel if local saving was enabled.
 */
function loadSavedSettingsIntoPanel() {
  const savedSettings = readSettingsFromLocalStorage();

  if (!savedSettings) {
    return;
  }

  if (savedSettings.settingsStorageMode !== "local") {
    return;
  }

  applySettingsToPanel(savedSettings);
}

/**
 * Reads saved settings from localStorage.
 */
function readSettingsFromLocalStorage() {
  try {
    const rawSettings = localStorage.getItem(LOCAL_SETTINGS_STORAGE_KEY);

    if (!rawSettings) {
      return null;
    }

    return JSON.parse(rawSettings);
  } catch (error) {
    console.warn("Could not read saved RIT settings.", error);
    return null;
  }
}

/**
 * Saves settings to localStorage.
 */
function saveSettingsToLocalStorage(settings) {
  try {
    localStorage.setItem(LOCAL_SETTINGS_STORAGE_KEY, JSON.stringify(settings));
  } catch (error) {
    console.warn("Could not save RIT settings.", error);
  }
}

/**
 * Clears saved local settings.
 */
function clearSavedLocalSettings() {
  try {
    localStorage.removeItem(LOCAL_SETTINGS_STORAGE_KEY);
  } catch (error) {
    console.warn("Could not clear saved RIT settings.", error);
  }
}

/**
 * Applies saved settings to task pane controls.
 */
function applySettingsToPanel(settings) {
  for (const controlConfig of SETTINGS_CONTROL_CONFIG) {
    const element = document.getElementById(controlConfig.id);

    if (!element) {
      continue;
    }

    const value = settings[controlConfig.settingName];

    if (value === undefined || value === null) {
      continue;
    }

    if (controlConfig.type === "checked") {
      element.checked = Boolean(value);
      continue;
    }

    if (controlConfig.type === "number") {
      element.value = String(value);
      continue;
    }

    element.value = String(value);
  }
}
