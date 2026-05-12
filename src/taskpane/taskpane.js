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
  removeSignificanceMarkersFromText,
  generateSignificanceLabels,
  buildBannerLocalSignificanceLabelMap,
} from "../core/significance";

import {
  LABEL_SCAN_COLUMNS_LEFT,
  detectMetricRowsFromLeftLabels,
  formatMetricDetectionDiagnostics,
  buildCalculationBlocks,
  getAllowedMarkerRowIndexes,
} from "../core/metric-detector";

import { writeCellResultsToSelectedRange, resolveNumericOutput } from "../core/excel-writer";

import { buildTablePreviewModel } from "../core/table-preview-model";
import { scanWorksheetForTables } from "../core/table-inventory-scanner";

import { detectBannerStructure, formatBannerDetectionDiagnostics } from "../core/banner-detector";

import { normalizeSelectedRange } from "../core/range-normalizer";

const USER_VISIBLE_BANNER_MESSAGE_CODES = new Set([
  "GLOBAL_TOTAL_USED",
  "BANNER_AUTO_PREVIOUS_COLUMN_APPLIED",
  "BANNER_TOTAL_ONLY_NO_TOTAL_PAIRS",
  "BANNER_MULTIPLE_LOCAL_TOTALS",
  "BANNER_TOTAL_OUTSIDE_SELECTION",
  "BANNER_MALFORMED_STRUCTURE",
  "BANNER_NO_ROWS_ABOVE_SELECTION",
]);

const SELECTED_RANGE_GUARDRAIL_WARNING_TEXT =
  "Похоже, вы выделили лейблы строк или шапку вместе с данными. Сейчас RIT ожидает выделение только числовой части таблицы.";

const SCAN_CELL_LIMIT = 250000;

function formatBannerUserMessages(bannerStructure) {
  if (!bannerStructure || !bannerStructure.messages) {
    return "";
  }

  const visibleMessages = bannerStructure.messages.filter((message) =>
    USER_VISIBLE_BANNER_MESSAGE_CODES.has(message.code)
  );

  if (visibleMessages.length === 0) {
    return "";
  }

  if (visibleMessages.length === 1) {
    return visibleMessages[0].text;
  }

  return ["Сообщения:", ...visibleMessages.map((message) => `- ${message.text}`)].join("\n");
}

function formatSelectedRangeGuardrailMessages(warnings) {
  if (!warnings || warnings.length === 0) {
    return "";
  }

  const uniqueTexts = Array.from(new Set(warnings.map((warning) => warning.text).filter(Boolean)));

  if (uniqueTexts.length === 1) {
    return uniqueTexts[0];
  }

  return ["Предупреждения:", ...uniqueTexts.map((text) => `- ${text}`)].join("\n");
}

function appendSelectedRangeGuardrailMessages(statusMessages, warnings) {
  const guardrailMessage = formatSelectedRangeGuardrailMessages(warnings);

  if (!guardrailMessage) {
    return statusMessages;
  }

  return [...statusMessages, "", guardrailMessage];
}

function formatStatusWithSelectedRangeGuardrails(message, warnings) {
  return appendSelectedRangeGuardrailMessages([message], warnings).join("\n");
}

function formatBannerUserMessagesExcludingCodes(bannerStructure, excludedCodes = []) {
  if (!bannerStructure || !bannerStructure.messages) {
    return "";
  }

  const excludedCodeSet = new Set(excludedCodes);

  const visibleMessages = bannerStructure.messages.filter(
    (message) =>
      USER_VISIBLE_BANNER_MESSAGE_CODES.has(message.code) && !excludedCodeSet.has(message.code)
  );

  if (visibleMessages.length === 0) {
    return "";
  }

  if (visibleMessages.length === 1) {
    return visibleMessages[0].text;
  }

  return ["Сообщения:", ...visibleMessages.map((message) => `- ${message.text}`)].join("\n");
}

const LOCAL_SETTINGS_STORAGE_KEY = "rit.settings.v1";

const SETTINGS_CONTROL_CONFIG = [
  { id: "confidence-level", type: "value", settingName: "confidenceLevel" },

  {
    id: "one-tailed-test",
    type: "checked",
    settingName: "oneTailedTest",
  },

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
  {
    id: "auto-detect-wave-banners",
    type: "checked",
    settingName: "autoDetectWaveBanners",
  },
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

const DEFAULT_CALCULATION_SETTINGS = {
  confidenceLevel: "95",
  oneTailedTest: false,

  roundCellValues: false,

  compareWithPreviousColumn: false,
  applyPreviousColumnFill: false,

  writeBannerLetters: false,
  respectBannerStructure: false,
  autoDetectWaveBanners: false,
  labelsOnLeftSide: false,

  compareOnlyWithTotal: false,
  excludeTotalFromComparisons: false,

  firstColumnIsTotal: false,
  totalInEachBanner: false,

  significantFillColor: "#E2F0D9",
  lowerThanTotalFillColor: "#FCE4D6",

  fillOnlyTotalComparisons: false,

  excludeSmallBasesFromComparisons: false,
  smallBaseThreshold: 50,
  smallBaseFillColor: "#D0D0D0",

  settingsStorageMode: "none",
};

const SETTINGS_TOOLTIPS = {
  "confidence-level":
    "Выберите уровень значимости для статистических тестов. Чем выше уровень, тем строже проверка и тем меньше отличий будут признаны значимыми.",

  "one-tailed-test":
    "Использовать односторонний тест вместо двустороннего. При том же уровне значимости такой тест легче находит отличия, но предполагает проверку различия в одном направлении.",

  "round-cell-values":
    "Округлять отображаемые значения перед добавлением маркеров. Расчёты при этом выполняются по исходным очищенным значениям, а не по округлённым.",

  "compare-with-previous-column":
    "Сравнивать каждую колонку только с колонкой слева: колонка 2 с колонкой 1, колонка 3 с колонкой 2 и так далее. Вместо букв используются стрелки вверх или вниз.",

  "apply-previous-column-fill":
    "Применять заливку к ячейкам со значимыми отличиями в режиме сравнения с предыдущей колонкой. Для роста используется обычная заливка значимости, для снижения — цвет “ниже Total”.",

  "write-banner-letters":
    "Добавлять буквенные индексы колонок в строку над выделенным диапазоном. Например: Segment 1 (a), Segment 2 (b). В режиме учёта структуры баннера буквы ставятся локально внутри групп.",

  "respect-banner-structure":
    "Анализировать структуру баннера над выделенным диапазоном. Это позволяет сравнивать колонки только внутри групп, определять локальные и глобальные Total, а также распознавать волновые баннеры.",

  "auto-detect-wave-banners":
    "Автоматически распознавать волновые группы в баннере, например Wave, Period, Волна, Период, и применять к ним сравнение с предыдущей колонкой. Обычные группы при этом продолжают сравниваться внутри группы обычным способом.",

  "labels-on-left-side":
    "Искать лейблы строк не рядом с выделенным диапазоном, а в самых левых колонках листа. Полезно для широких таблиц, где данные выделены справа, а названия строк находятся далеко слева.",

  "compare-only-with-total":
    "Сравнивать каждую колонку только с колонкой Total. Обычные попарные сравнения между сегментами выполняться не будут.",

  "exclude-total-from-comparisons":
    "Исключить Total из расчётов. Total не будет использоваться как база сравнения и не будет сравниваться с другими колонками.",

  "first-column-is-total":
    "Считать первую колонку выделенного диапазона Total. Она будет использоваться как референс для сравнения с остальными колонками.",

  "total-in-each-banner":
    "Считать, что Total находится внутри каждой группы баннера. При включённом учёте структуры баннера расположение Total определяется автоматически.",

  "significant-fill-color":
    "Цвет заливки для ячеек, которые статистически значимо выше другой сравниваемой ячейки или Total.",

  "lower-than-total-fill-color":
    "Цвет заливки для ячеек, которые статистически значимо ниже Total или ниже предыдущей колонки в режиме сравнения с предыдущей колонкой.",

  "fill-only-total-comparisons":
    "Применять обычную зелёную заливку только к ячейкам, которые значимо выше Total. Отличия между обычными сегментами будут отмечаться буквами, но без зелёной заливки.",

  "exclude-small-bases":
    "Исключать из расчётов колонки, где база меньше заданного порога. Такие колонки не участвуют в статистических сравнениях и получают отдельную заливку.",

  "small-base-threshold":
    "Минимальный допустимый размер базы. Если база колонки меньше этого значения, колонка исключается из расчётов.",

  "small-base-fill-color":
    "Цвет заливки для колонок с маленькой базой. Эта заливка имеет самый высокий приоритет и перекрывает остальные типы заливки.",

  "settings-storage-mode":
    "Выберите, сохранять ли настройки панели. Локальное сохранение работает только на этом устройстве и в этом браузере/Excel WebView.",

  "reset-settings":
    "Сбросить все настройки к значениям по умолчанию и удалить локально сохранённые настройки.",
};

const ENABLE_BANNER_SPAN_DIAGNOSTICS = false;

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
  const checkTableButton = document.getElementById("check-table"); // Read-only table check button.
  const findTablesButton = document.getElementById("find-tables"); // Table inventory button.

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

  if (checkTableButton) {
    checkTableButton.addEventListener("click", runCheckTable);
  }

  if (findTablesButton) {
    findTablesButton.addEventListener("click", runTableInventory);
  }
});

/**
 * Detects high-confidence signs that labels or header rows were included in the data range.
 */
function detectSelectedRangeGuardrails(selectedText, cleanedValues) {
  const values = Array.isArray(cleanedValues) ? cleanedValues : [];
  const rowCount = values.length;
  const columnCount = rowCount > 0 && Array.isArray(values[0]) ? values[0].length : 0;

  if (rowCount < 2 || columnCount < 2) {
    return [];
  }

  const warnings = [];
  const allCells = getSelectedRangeCells(selectedText, values, 0, 0, rowCount, columnCount);
  const allStats = analyzeSelectedRangeCells(allCells);

  if (rowCount >= 3 && columnCount >= 3) {
    const firstColumnStats = analyzeSelectedRangeCells(
      getSelectedRangeCells(selectedText, values, 0, 0, rowCount, 1)
    );
    const rightSideStats = analyzeSelectedRangeCells(
      getSelectedRangeCells(selectedText, values, 0, 1, rowCount, columnCount - 1)
    );

    if (
      firstColumnStats.nonEmptyCount >= Math.max(2, Math.ceil(rowCount * 0.5)) &&
      firstColumnStats.textRatio >= 0.7 &&
      rightSideStats.nonEmptyCount >= Math.max(4, Math.ceil(rowCount * (columnCount - 1) * 0.5)) &&
      rightSideStats.numericRatio >= 0.7
    ) {
      warnings.push({
        code: "SELECTED_RANGE_LIKELY_LEFT_LABEL_COLUMN",
        severity: "warning",
        text: SELECTED_RANGE_GUARDRAIL_WARNING_TEXT,
        rowIndex: null,
        columnIndex: 0,
        evidence: {
          firstColumnTextRatio: roundGuardrailRatio(firstColumnStats.textRatio),
          rightSideNumericRatio: roundGuardrailRatio(rightSideStats.numericRatio),
          firstColumnTextCount: firstColumnStats.textCount,
          rightSideNumericCount: rightSideStats.numericCount,
        },
      });
    }

    const firstRowStats = analyzeSelectedRangeCells(
      getSelectedRangeCells(selectedText, values, 0, 0, 1, columnCount)
    );
    const lowerRowsStats = analyzeSelectedRangeCells(
      getSelectedRangeCells(selectedText, values, 1, 0, rowCount - 1, columnCount)
    );

    if (
      firstRowStats.nonEmptyCount >= Math.max(2, Math.ceil(columnCount * 0.5)) &&
      firstRowStats.textRatio >= 0.6 &&
      lowerRowsStats.nonEmptyCount >= Math.max(4, Math.ceil((rowCount - 1) * columnCount * 0.5)) &&
      lowerRowsStats.numericRatio >= 0.7
    ) {
      warnings.push({
        code: "SELECTED_RANGE_LIKELY_TOP_HEADER_ROW",
        severity: "warning",
        text: SELECTED_RANGE_GUARDRAIL_WARNING_TEXT,
        rowIndex: 0,
        columnIndex: null,
        evidence: {
          firstRowTextRatio: roundGuardrailRatio(firstRowStats.textRatio),
          lowerRowsNumericRatio: roundGuardrailRatio(lowerRowsStats.numericRatio),
          firstRowTextCount: firstRowStats.textCount,
          lowerRowsNumericCount: lowerRowsStats.numericCount,
        },
      });
    }
  }

  if (
    warnings.length === 0 &&
    allStats.totalCount >= 12 &&
    allStats.textCount >= 4 &&
    allStats.textRatio >= 0.25 &&
    allStats.numericRatio >= 0.5
  ) {
    warnings.push({
      code: "SELECTED_RANGE_TEXT_HEAVY",
      severity: "warning",
      text: SELECTED_RANGE_GUARDRAIL_WARNING_TEXT,
      rowIndex: null,
      columnIndex: null,
      evidence: {
        textRatio: roundGuardrailRatio(allStats.textRatio),
        numericRatio: roundGuardrailRatio(allStats.numericRatio),
        textCount: allStats.textCount,
        numericCount: allStats.numericCount,
      },
    });
  }

  return warnings;
}

function getSelectedRangeCells(selectedText, values, startRow, startColumn, rowCount, columnCount) {
  const cells = [];

  for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
    const rowIndex = startRow + rowOffset;

    for (let columnOffset = 0; columnOffset < columnCount; columnOffset++) {
      const columnIndex = startColumn + columnOffset;

      cells.push({
        value: values && values[rowIndex] ? values[rowIndex][columnIndex] : undefined,
        text: selectedText && selectedText[rowIndex] ? selectedText[rowIndex][columnIndex] : undefined,
      });
    }
  }

  return cells;
}

function analyzeSelectedRangeCells(cells) {
  const stats = {
    totalCount: cells.length,
    nonEmptyCount: 0,
    blankCount: 0,
    numericCount: 0,
    textCount: 0,
    numericRatio: 0,
    textRatio: 0,
  };

  for (const cell of cells) {
    const cellValue = getGuardrailCellValue(cell);

    if (isBlankLikeCellValue(cellValue)) {
      stats.blankCount++;
      continue;
    }

    stats.nonEmptyCount++;

    if (isNumericLikeCellValue(cellValue)) {
      stats.numericCount++;
      continue;
    }

    if (isTextLikeCellValue(cellValue)) {
      stats.textCount++;
    }
  }

  if (stats.nonEmptyCount > 0) {
    stats.numericRatio = stats.numericCount / stats.nonEmptyCount;
    stats.textRatio = stats.textCount / stats.nonEmptyCount;
  }

  return stats;
}

function getGuardrailCellValue(cell) {
  if (!cell) {
    return "";
  }

  if (!isBlankLikeCellValue(cell.value)) {
    return cell.value;
  }

  return cell.text;
}

function isBlankLikeCellValue(value) {
  return value === null || value === undefined || String(value).trim() === "";
}

function isNumericLikeCellValue(value) {
  if (isBlankLikeCellValue(value)) {
    return false;
  }

  if (typeof value === "number") {
    return Number.isFinite(value);
  }

  const normalizedValue = String(value)
    .trim()
    .replace(/\s/g, "")
    .replace("%", "")
    .replace(",", ".");

  return normalizedValue !== "" && !Number.isNaN(Number(normalizedValue));
}

function isTextLikeCellValue(value) {
  return !isBlankLikeCellValue(value) && !isNumericLikeCellValue(value);
}

function roundGuardrailRatio(value) {
  return Math.round(value * 100) / 100;
}

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
      calculationSettings.excludeTotalFromComparisons &&
      !calculationSettings.firstColumnIsTotal &&
      !calculationSettings.respectBannerStructure
    ) {
      setStatusMessage(
        "Для режима “Не сравнивать с Тотал” нужно указать расположение Тотала. Сейчас поддерживается вариант “Первая колонка — Тотал” или режим “Учитывать структуру баннера”."
      );

      return;
    }

    if (
      calculationSettings.compareOnlyWithTotal &&
      !calculationSettings.firstColumnIsTotal &&
      !calculationSettings.respectBannerStructure
    ) {
      setStatusMessage(
        "Для режима “Сравнивать только с Тотал” нужно указать расположение Тотала. Сейчас поддерживается вариант “Первая колонка — Тотал” или режим “Учитывать структуру баннера”."
      );

      return;
    }

    selectedRange.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount"]);

    await context.sync();

    if (
      calculationSettings.compareWithPreviousColumn &&
      calculationSettings.fillOnlyTotalComparisons
    ) {
      outputElement.textContent =
        "Режим “Сравнение с предыдущей колонкой” несовместим с настройкой “Заливка только для Тотала”.";

      return;
    }

    selectedRange.load([
      "values",
      "text",
      "rowIndex",
      "columnIndex",
      "rowCount",
      "columnCount",
    ]);

    await context.sync();

    const selectedValues = selectedRange.values;
    const selectedText = selectedRange.text;

    if (!selectedValues || selectedValues.length < 2 || selectedValues[0].length < 2) {
      setStatusMessage("Please select at least 2 columns and 2 rows.");
      return;
    }

    const cleanedValues = removeSignificanceMarkersFromMatrix(selectedValues);
    const normalized = normalizeSelectedRange(cleanedValues, selectedText);

    if (normalized.normalizationNeeded && !normalized.normalizationApplied) {
      const codes =
        normalized.blockingReasons && normalized.blockingReasons.length > 0
          ? ` [${normalized.blockingReasons.join(", ")}]`
          : "";

      setStatusMessage(`${normalized.blockingMessage}${codes}`);
      return;
    }

    const rawSelectedRangeGuardrailWarnings = detectSelectedRangeGuardrails(
      selectedText,
      cleanedValues
    );

    const runModel = {
      values: cleanedValues,
      text: selectedText,
      leftLabelValues: null,
      bannerContext: null,
      writeTargetRange: selectedRange,
      targetStartRowIndex: selectedRange.rowIndex,
      normalizationApplied: false,
      normalizationStatusLines: [],
      selectedRangeGuardrailWarnings: rawSelectedRangeGuardrailWarnings,
    };

    if (normalized.normalizationNeeded && normalized.normalizationApplied) {
      if (
        !normalized.valuesForCalculation ||
        normalized.valuesForCalculation.length < 2 ||
        !normalized.valuesForCalculation[0] ||
        normalized.valuesForCalculation[0].length < 2
      ) {
        setStatusMessage("Please select at least 2 columns and 2 rows.");
        return;
      }

      runModel.values = normalized.valuesForCalculation;
      runModel.text = normalized.textForCalculation;
      runModel.leftLabelValues = normalized.leftLabelValues;
      runModel.bannerContext = normalized.bannerContext;
      runModel.writeTargetRange = selectedRange
        .getCell(normalized.dataRowOffset, normalized.dataColOffset)
        .getResizedRange(
          normalized.valuesForCalculation.length - 1,
          normalized.valuesForCalculation[0].length - 1
        );
      runModel.targetStartRowIndex = selectedRange.rowIndex + normalized.dataRowOffset;
      runModel.normalizationApplied = true;
      runModel.normalizationStatusLines.push(
        "Диапазон нормализован: расчёт выполнен только по области данных."
      );
      runModel.selectedRangeGuardrailWarnings = [];
    } else {
      runModel.leftLabelValues = await loadLabelValuesForSelectedRange(
        context,
        selectedRange,
        calculationSettings
      ); // Labels located 1-2 columns to the left of the selected data.
    }

    if (!runModel.values || runModel.values.length < 2 || runModel.values[0].length < 2) {
      setStatusMessage("Please select at least 2 columns and 2 rows.");
      return;
    }

    runModel.writeTargetRange.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);

    await context.sync();

    runModel.targetStartRowIndex = runModel.writeTargetRange.rowIndex;

    if (calculationSettings.writeBannerLetters && runModel.targetStartRowIndex === 0) {
      outputElement.textContent =
        "Данные расположены в первой строке. Добавьте строку над выделенным массивом и запустите расчёт повторно.";

      return;
    }

    runModel.writeTargetRange.values = runModel.values;

    runModel.writeTargetRange.format.font.bold = false;
    runModel.writeTargetRange.format.fill.clear();
    runModel.writeTargetRange.format.horizontalAlignment = "Center";
    runModel.writeTargetRange.format.verticalAlignment = "Center";

    await context.sync();

    const selectedRangeGuardrailWarnings = runModel.selectedRangeGuardrailWarnings;

    const detectionResult = detectMetricRowsFromLeftLabels(
      runModel.values,
      runModel.leftLabelValues
    ); // Row type diagnostics based on left-side labels.

    const calculationBlocks = buildCalculationBlocks(detectionResult); // List of metric blocks to calculate.

    if (!calculationBlocks || calculationBlocks.length === 0) {
      setStatusMessage(
        formatStatusWithSelectedRangeGuardrails(
          "Could not detect any calculation blocks.",
          selectedRangeGuardrailWarnings
        )
      );
      return;
    }

    let bannerStructure = null;

    if (calculationSettings.respectBannerStructure) {
      const bannerContext =
        buildRunBannerContext(runModel.bannerContext) ||
        (await loadBannerContextForSelectedRange(
          context,
          runModel.writeTargetRange,
          calculationSettings
        ));

      bannerStructure = detectBannerStructure(bannerContext, calculationSettings);

      if (bannerContext.messages && bannerContext.messages.length > 0) {
        bannerStructure.messages = [...bannerContext.messages, ...(bannerStructure.messages || [])];
      }

      /** 
      bannerSpanDiagnostics = await loadBannerSpanDiagnosticsForSelectedRange(
        context,
        selectedRange
      );
      */
    }

    const fullCellResultMatrix = createEmptyCellResultMatrix(
      runModel.values.length,
      runModel.values[0].length
    ); // Full-size marker storage matching the data body being calculated.

    for (const calculationBlock of calculationBlocks) {
      const smallBaseResult = applySmallBaseRulesForCalculationBlock(
        runModel.values,
        calculationBlock,
        fullCellResultMatrix,
        calculationSettings
      );

      if (smallBaseResult.errorMessage) {
        setStatusMessage(
          formatStatusWithSelectedRangeGuardrails(
            smallBaseResult.errorMessage,
            selectedRangeGuardrailWarnings
          )
        );
        return;
      }

      const blockCalculationSettings = {
        ...calculationSettings,
        excludedColumnIndexes: smallBaseResult.excludedColumnIndexes,
        bannerStructure,
      };

      const blockResults = calculateBlockResults(
        runModel.values,
        calculationBlock,
        blockCalculationSettings
      );

      const bannerStructureError = getFirstBannerStructureError(bannerStructure);

      if (bannerStructureError) {
        const statusMessages = [bannerStructureError.text];

        const bannerUserMessages = formatBannerUserMessagesExcludingCodes(bannerStructure, [
          bannerStructureError.code,
        ]);

        if (bannerUserMessages) {
          statusMessages.push("");
          statusMessages.push(bannerUserMessages);
        }

        setStatusMessage(
          appendSelectedRangeGuardrailMessages(
            statusMessages,
            selectedRangeGuardrailWarnings
          ).join("\n")
        );
        return;
      }

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
      runModel.writeTargetRange,
      runModel.text,
      fullCellResultMatrix,
      detectionResult,
      calculationSettings
    );

    if (calculationSettings.writeBannerLetters) {
      if (calculationSettings.respectBannerStructure && bannerStructure) {
        await writeBannerMarkersAboveSelectedRangeUsingBannerStructure(
          context,
          runModel.writeTargetRange,
          bannerStructure,
          calculationSettings
        );
      } else {
        await writeBannerMarkersAboveSelectedRange(
          context,
          runModel.writeTargetRange,
          calculationSettings
        );
      }
    }

    await context.sync();

    const statusMessages = [`Расчёт выполнен. Обработано блоков: ${calculationBlocks.length}.`];

    if (runModel.normalizationStatusLines.length > 0) {
      statusMessages.push("");
      statusMessages.push(...runModel.normalizationStatusLines);
    }

    const bannerUserMessages = formatBannerUserMessages(bannerStructure);

    if (bannerUserMessages) {
      statusMessages.push("");
      statusMessages.push(bannerUserMessages);
    }

    setStatusMessage(
      appendSelectedRangeGuardrailMessages(statusMessages, selectedRangeGuardrailWarnings).join("\n")
    );
  });
}

/**
 * Converts normalized banner context into the shape used by Run banner detection.
 */
function buildRunBannerContext(bannerContext) {
  if (!bannerContext) {
    return null;
  }

  if (bannerContext.selectedColumnCount !== undefined) {
    return bannerContext;
  }

  const scanRows = Array.isArray(bannerContext.scanRows) ? bannerContext.scanRows : [];
  const selectedColumnCount = bannerContext.columnCount || 0;

  if (!selectedColumnCount || scanRows.length === 0) {
    return null;
  }

  return {
    selectedColumnCount,
    lowerBannerRow: scanRows[scanRows.length - 1],
    upperScanRows: scanRows.slice(0, -1).reverse(),
    messages: bannerContext.messages || [],
  };
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
    const selectedRange = context.workbook.getSelectedRange();

    // Read-only load: needed to decide whether to operate on the whole
    // selection (strict numeric case) or only on the detected data body
    // (forgiving full-table case). No writes happen before the target is known.
    selectedRange.load(["values", "text"]);

    await context.sync();

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

    const cleanedValues = removeSignificanceMarkersFromMatrix(selectedValues);
    const normalized = normalizeSelectedRange(cleanedValues, selectedText);

    // State 3: broad/full-table-like selection but decomposition failed.
    // Block and return without mutating anything.
    if (normalized.normalizationNeeded && !normalized.normalizationApplied) {
      const codes =
        normalized.blockingReasons && normalized.blockingReasons.length > 0
          ? ` [${normalized.blockingReasons.join(", ")}]`
          : "";
      setStatusMessage(`${normalized.blockingMessage}${codes}`);
      return;
    }

    // Resolve the clear target:
    //   - State 1 (pass-through): the original selection is numeric-only.
    //   - State 2 (normalized):   only the detected data body subrange.
    let clearTargetRange;

    if (normalized.normalizationNeeded && normalized.normalizationApplied) {
      const bodyRowCount = normalized.valuesForCalculation.length;
      const bodyColCount = normalized.valuesForCalculation[0].length;

      if (bodyRowCount < 1 || bodyColCount < 1) {
        setStatusMessage("Нет данных в выделенном диапазоне.");
        return;
      }

      clearTargetRange = selectedRange
        .getCell(normalized.dataRowOffset, normalized.dataColOffset)
        .getResizedRange(bodyRowCount - 1, bodyColCount - 1);
    } else {
      clearTargetRange = selectedRange;
    }

    clearTargetRange.load(["values", "numberFormat"]);

    await context.sync();

    const targetValues = clearTargetRange.values;
    const targetNumberFormats = clearTargetRange.numberFormat;

    const nextValues = [];
    const nextNumberFormats = [];

    for (let rowIndex = 0; rowIndex < targetValues.length; rowIndex++) {
      const valueRow = [];
      const formatRow = [];

      for (let columnIndex = 0; columnIndex < targetValues[rowIndex].length; columnIndex++) {
        const rawValue = targetValues[rowIndex][columnIndex];

        if (typeof rawValue === "number") {
          valueRow.push(rawValue);
          formatRow.push(targetNumberFormats[rowIndex][columnIndex]);
          continue;
        }

        const cleanedText = removeSignificanceMarkersFromText(rawValue);
        const resolved = resolveNumericOutput(cleanedText);

        if (resolved !== null) {
          valueRow.push(resolved.value);
          formatRow.push(resolved.format);
        } else {
          valueRow.push(cleanedText);
          formatRow.push("@");
        }
      }

      nextValues.push(valueRow);
      nextNumberFormats.push(formatRow);
    }

    clearTargetRange.numberFormat = nextNumberFormats;
    clearTargetRange.values = nextValues;

    clearTargetRange.format.font.bold = false;
    clearTargetRange.format.fill.clear();

    await context.sync();

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
 * Reads selected range and calls buildTablePreviewModel to display a short summary.
 *
 * Read-only: does not write to Excel, does not calculate significance.
 *
 * Three normalizer states:
 *   1. pass-through  — selection is numeric-only; existing flow runs unchanged.
 *   2. normalized    — broad/full-table selection decomposed; normalized values used.
 *   3. blocked       — broad selection but decomposition failed; early return with message.
 */
async function runCheckTable() {
  await Excel.run(async (context) => {
    const calculationSettings = readCalculationSettingsFromPanel();
    const selectedRange = context.workbook.getSelectedRange();

    selectedRange.load(["values", "text", "rowIndex", "columnIndex", "rowCount", "columnCount"]);

    await context.sync();

    const selectedValues = selectedRange.values;
    const selectedText = selectedRange.text;

    if (!selectedValues || selectedValues.length < 1 || !selectedValues[0] || selectedValues[0].length < 1) {
      setCheckMessage("Нет данных в выделенном диапазоне.");
      return;
    }

    const cleanedValues = removeSignificanceMarkersFromMatrix(selectedValues);

    const normalized = normalizeSelectedRange(cleanedValues, selectedText);

    // State 3: normalization needed but blocked — stop and show reason.
    if (normalized.normalizationNeeded && !normalized.normalizationApplied) {
      const codes = normalized.blockingReasons.length > 0
        ? ` [${normalized.blockingReasons.join(", ")}]`
        : "";
      setCheckMessage(`${normalized.blockingMessage}${codes}`);
      return;
    }

    let modelInput;
    const normalizationLines = [];

    if (normalized.normalizationNeeded && normalized.normalizationApplied) {
      // State 2: broad selection successfully decomposed — use normalized partitions.
      modelInput = {
        values: normalized.valuesForCalculation,
        leftLabelValues: normalized.leftLabelValues,
        bannerContext: normalized.bannerContext,
        settings: calculationSettings,
      };

      normalizationLines.push("Диапазон нормализован: заголовки/лейблы/баннер отделены от данных.");

      const parts = [];
      if (normalized.titleRows.length > 0) parts.push(`заголовков: ${normalized.titleRows.length}`);
      if (normalized.subtitleRows.length > 0) parts.push(`подзаголовков: ${normalized.subtitleRows.length}`);
      if (normalized.bannerRows.length > 0) parts.push(`строк баннера: ${normalized.bannerRows.length}`);
      if (normalized.labelColumns.length > 0) parts.push(`колонок меток: ${normalized.labelColumns.length}`);
      if (parts.length > 0) normalizationLines.push(`Отделено: ${parts.join(", ")}.`);
    } else {
      // State 1: numeric-only selection — existing flow unchanged.
      const leftLabelValues = await loadLabelValuesForSelectedRange(
        context,
        selectedRange,
        calculationSettings
      );
      modelInput = { values: cleanedValues, leftLabelValues };
    }

    const model = buildTablePreviewModel(modelInput);
    const { summary, qualitySummary, warnings } = model;

    if (summary.rowCount === 0) {
      setCheckMessage("Выделенный диапазон пуст.");
      return;
    }

    const lines = [
      `Проверка завершена. Строк: ${summary.rowCount}. Блоков: ${summary.detectedBlocks}. Баз: ${summary.baseRows}. Предупреждений: ${qualitySummary.warningCount}. Критических: ${qualitySummary.criticalCount}.`,
    ];

    if (normalizationLines.length > 0) {
      lines.push("");
      lines.push(...normalizationLines);
    }

    if (warnings && warnings.length > 0) {
      lines.push("");
      lines.push("Предупреждения:");
      for (const warning of warnings) {
        lines.push(`- [${warning.severity}] ${warning.text}`);
      }
    }

    setCheckMessage(lines.join("\n"));
  });
}

function setCheckMessage(message) {
  const checkPanel = document.getElementById("check-panel");
  const checkResult = document.getElementById("check-result");

  if (checkPanel) {
    checkPanel.style.display = "block";
  }

  if (checkResult) {
    checkResult.textContent = message || "";
  }
}

async function runTableInventory() {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    worksheet.load("name");

    const usedRange = worksheet.getUsedRangeOrNullObject();
    usedRange.load(["isNullObject", "rowIndex", "columnIndex", "rowCount", "columnCount"]);

    await context.sync();

    if (usedRange.isNullObject) {
      setInventoryMessage("Лист пуст. Таблицы не найдены.");
      return;
    }

    const cellCount = usedRange.rowCount * usedRange.columnCount;
    if (cellCount > SCAN_CELL_LIMIT) {
      setInventoryMessage(
        `Лист слишком большой для сканирования (${usedRange.rowCount} стр. × ${usedRange.columnCount} кол. = ${cellCount} ячеек, лимит: ${SCAN_CELL_LIMIT}).\nИспользуйте «Проверить таблицу» с выделенным диапазоном.`
      );
      return;
    }

    usedRange.load("values");
    await context.sync();

    const items = scanWorksheetForTables({
      values: usedRange.values,
      usedRangeRowOffset: usedRange.rowIndex,
      usedRangeColOffset: usedRange.columnIndex,
      sheetName: worksheet.name,
    });

    if (items.length === 0) {
      setInventoryMessage("Таблицы не найдены. RIT не обнаружил блоков с данными на этом листе.");
      return;
    }

    const lines = [`Найдено таблиц: ${items.length}.`, ""];

    items.forEach((item, idx) => {
      const header = item.title
        ? `${idx + 1}. ${item.title} — ${item.rangeAddress}`
        : `${idx + 1}. ${item.rangeAddress}`;
      lines.push(header);
      lines.push(`   ${item.rowCount} строк, ${item.columnCount} колонок.`);

      if (item.previewSummary) {
        lines.push(`   ${item.previewSummary}.`);
      }

      const warnParts = [];
      if (item.criticalCount > 0) warnParts.push(`Критических: ${item.criticalCount}`);
      if (item.warningsCount > 0) warnParts.push(`Предупреждений: ${item.warningsCount}`);
      if (warnParts.length > 0) lines.push(`   ${warnParts.join(". ")}.`);

      if (!item.isLikelyTable) {
        lines.push("   Не опознана как таблица RIT.");
      }

      lines.push(
        `   Проверка: ${item.canRunCheckTable ? "да" : "нет"}. Значимость: ${item.canRunSignificance ? "да" : "нет"}.`
      );

      if (item.reasonsIfNotRunnable.length > 0) {
        lines.push(`   [${item.reasonsIfNotRunnable.join("; ")}]`);
      }

      lines.push("");
    });

    setInventoryMessage(lines.join("\n").trimEnd());
  });
}

function setInventoryMessage(message) {
  const panel = document.getElementById("inventory-panel");
  const result = document.getElementById("inventory-result");
  if (panel) panel.style.display = "block";
  if (result) result.textContent = message || "";
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
 * Reads banner rows above selected range.
 *
 * PURPOSE:
 * Detection-only banner engine stage.
 * Reads:
 * - lower banner row directly above selection;
 * - up to maxBannerScanRows rows above it.
 *
 * This function does not read merge metadata yet.
 */
async function loadBannerContextForSelectedRange(context, selectedRange, calculationSettings) {
  const maxBannerScanRows = 5;

  selectedRange.load(["rowIndex", "columnIndex", "columnCount"]);

  await context.sync();

  const selectedStartRowIndex = selectedRange.rowIndex;
  const selectedStartColumnIndex = selectedRange.columnIndex;
  const selectedColumnCount = selectedRange.columnCount;

  if (selectedStartRowIndex === 0) {
    return {
      selectedColumnCount,
      lowerBannerRow: [],
      upperScanRows: [],
      messages: [
        {
          severity: MESSAGE_SEVERITY.WARNING,
          code: "BANNER_NO_ROWS_ABOVE_SELECTION",
          text: "Баннер: над выделенным диапазоном нет строк для анализа.",
        },
      ],
    };
  }

  const worksheet = selectedRange.worksheet;

  const lowerBannerRange = worksheet.getRangeByIndexes(
    selectedStartRowIndex - 1,
    selectedStartColumnIndex,
    1,
    selectedColumnCount
  );

  lowerBannerRange.load("text");

  const availableUpperRowCount = Math.min(maxBannerScanRows, selectedStartRowIndex - 1);

  let upperScanRows = [];

  if (availableUpperRowCount > 0) {
    const upperScanRange = worksheet.getRangeByIndexes(
      selectedStartRowIndex - 1 - availableUpperRowCount,
      selectedStartColumnIndex,
      availableUpperRowCount,
      selectedColumnCount
    );

    upperScanRange.load("text");

    await context.sync();

    upperScanRows = upperScanRange.text.slice().reverse();

    return {
      selectedColumnCount,
      lowerBannerRow: lowerBannerRange.text[0],
      upperScanRows,
      messages: [],
    };
  }

  await context.sync();

  return {
    selectedColumnCount,
    lowerBannerRow: lowerBannerRange.text[0],
    upperScanRows: [],
    messages: [],
  };
}

/**
 * Reads expanded banner rows and reconstructs possible horizontal spans.
 *
 * PURPOSE:
 * Diagnostic-only fallback for merged banner headers.
 *
 * WHY:
 * Office.js may not expose merged areas reliably for continuation cells.
 * In practice, merged banner rows often look like:
 *   Age | "" | ""
 * So we reconstruct spans as:
 *   non-empty cell + following empty cells until next non-empty cell.
 */
async function loadBannerSpanDiagnosticsForSelectedRange(context, selectedRange) {
  const maxBannerScanRows = 5;
  const maxColumnsToScanLeft = 10;
  const maxColumnsToScanRight = 10;

  selectedRange.load(["rowIndex", "columnIndex", "columnCount"]);

  await context.sync();

  const selectedStartRowIndex = selectedRange.rowIndex;
  const selectedStartColumnIndex = selectedRange.columnIndex;
  const selectedColumnCount = selectedRange.columnCount;
  const selectedEndColumnIndex = selectedStartColumnIndex + selectedColumnCount - 1;

  if (selectedStartRowIndex === 0) {
    return [
      "Banner span diagnostics:",
      "- Над выделенным диапазоном нет строк для анализа span-структуры.",
    ].join("\n");
  }

  const availableRowCount = Math.min(maxBannerScanRows + 1, selectedStartRowIndex);
  const firstBannerRowIndex = selectedStartRowIndex - availableRowCount;

  const scanStartColumnIndex = Math.max(0, selectedStartColumnIndex - maxColumnsToScanLeft);

  const scanEndColumnIndex = selectedEndColumnIndex + maxColumnsToScanRight;

  const scanColumnCount = scanEndColumnIndex - scanStartColumnIndex + 1;

  const bannerScanRange = selectedRange.worksheet.getRangeByIndexes(
    firstBannerRowIndex,
    scanStartColumnIndex,
    availableRowCount,
    scanColumnCount
  );

  bannerScanRange.load("text");

  await context.sync();

  const lines = [];

  lines.push("Banner span diagnostics:");
  lines.push(
    `- Диапазон проверки: строки ${firstBannerRowIndex + 1}:${selectedStartRowIndex}, колонки ${getExcelColumnLetter(scanStartColumnIndex)}:${getExcelColumnLetter(scanEndColumnIndex)}.`
  );
  lines.push(
    `- Выделение по колонкам: ${getExcelColumnLetter(selectedStartColumnIndex)}:${getExcelColumnLetter(selectedEndColumnIndex)}.`
  );

  const lowerBannerLocalRowIndex = availableRowCount - 1;
  const lowerBannerRowText = bannerScanRange.text[lowerBannerLocalRowIndex] || [];

  for (let localRowIndex = 0; localRowIndex < availableRowCount; localRowIndex++) {
    const sheetRowIndex = firstBannerRowIndex + localRowIndex;
    const rowOffsetFromSelection = sheetRowIndex - selectedStartRowIndex;
    const rowText = bannerScanRange.text[localRowIndex] || [];

    const spans = reconstructHorizontalSpansFromRowText(
      rowText,
      scanStartColumnIndex,
      selectedStartColumnIndex,
      selectedEndColumnIndex,
      lowerBannerRowText,
      localRowIndex === lowerBannerLocalRowIndex
    );

    lines.push(`- Row offset ${rowOffsetFromSelection}:`);

    if (spans.length === 0) {
      lines.push("  - spans: none");
      continue;
    }

    for (const span of spans) {
      const selectedPartText =
        span.selectedStartColumnIndex !== null
          ? `, selected cols ${getExcelColumnLetter(span.selectedStartColumnIndex)}:${getExcelColumnLetter(span.selectedEndColumnIndex)}`
          : ", outside selection";

      lines.push(
        `  - "${span.label}": sheet cols ${getExcelColumnLetter(span.startColumnIndex)}:${getExcelColumnLetter(span.endColumnIndex)}, length=${span.length}${selectedPartText}`
      );
    }
  }

  return lines.join("\n");
}

/**
 * Reconstructs possible horizontal spans from row text.
 *
 * RULE:
 * A non-empty cell starts a span.
 * Empty cells to the right are treated as continuation until next non-empty cell.
 *
 * For upper banner rows, spans are additionally constrained by the lower banner row:
 * if lower level has a continuous non-empty area starting at the span start,
 * upper span should not extend beyond that lower-level area.
 */
function reconstructHorizontalSpansFromRowText(
  rowText,
  scanStartColumnIndex,
  selectedStartColumnIndex,
  selectedEndColumnIndex,
  lowerBannerRowText = [],
  isLowerBannerRow = false
) {
  const spans = [];
  let currentSpan = null;

  for (let localColumnIndex = 0; localColumnIndex < rowText.length; localColumnIndex++) {
    const sheetColumnIndex = scanStartColumnIndex + localColumnIndex;
    const cellText = normalizeBannerDiagnosticCellText(rowText[localColumnIndex]);

    if (cellText) {
      if (currentSpan) {
        currentSpan.endColumnIndex = sheetColumnIndex - 1;
        refineDiagnosticSpanRightBoundaryByLowerBannerRow(
          currentSpan,
          lowerBannerRowText,
          scanStartColumnIndex,
          isLowerBannerRow
        );
        finalizeDiagnosticSpanSelection(
          currentSpan,
          selectedStartColumnIndex,
          selectedEndColumnIndex
        );
        spans.push(currentSpan);
      }

      currentSpan = {
        label: cellText,
        startColumnIndex: sheetColumnIndex,
        endColumnIndex: sheetColumnIndex,
        length: 1,
        selectedStartColumnIndex: null,
        selectedEndColumnIndex: null,
      };

      continue;
    }

    if (currentSpan) {
      currentSpan.endColumnIndex = sheetColumnIndex;
    }
  }

  if (currentSpan) {
    refineDiagnosticSpanRightBoundaryByLowerBannerRow(
      currentSpan,
      lowerBannerRowText,
      scanStartColumnIndex,
      isLowerBannerRow
    );
    finalizeDiagnosticSpanSelection(currentSpan, selectedStartColumnIndex, selectedEndColumnIndex);
    spans.push(currentSpan);
  }

  return spans
    .map((span) => ({
      ...span,
      length: span.endColumnIndex - span.startColumnIndex + 1,
    }))
    .filter((span) => span.label && span.length > 0);
}

/**
 * Adds selected-range intersection metadata to reconstructed span.
 */
function finalizeDiagnosticSpanSelection(span, selectedStartColumnIndex, selectedEndColumnIndex) {
  const intersectionStart = Math.max(span.startColumnIndex, selectedStartColumnIndex);
  const intersectionEnd = Math.min(span.endColumnIndex, selectedEndColumnIndex);

  if (intersectionStart <= intersectionEnd) {
    span.selectedStartColumnIndex = intersectionStart;
    span.selectedEndColumnIndex = intersectionEnd;
  }
}

/**
 * Refines upper banner span right boundary using lower banner row.
 *
 * PURPOSE:
 * Prevents a merged-like upper label from stretching to the end of the scan range.
 *
 * Example:
 * Upper row:
 *   Age | "" | "" | "" | ""
 * Lower row:
 *   Total | 18-24 | 25-34 | "" | ""
 *
 * Without refinement:
 *   Age spans all scanned empty cells.
 *
 * With refinement:
 *   Age spans only the continuous non-empty lower-level area: Total..25-34.
 */
function refineDiagnosticSpanRightBoundaryByLowerBannerRow(
  span,
  lowerBannerRowText,
  scanStartColumnIndex,
  isLowerBannerRow
) {
  if (isLowerBannerRow) {
    return;
  }

  if (!lowerBannerRowText || lowerBannerRowText.length === 0) {
    return;
  }

  const spanStartLocalColumnIndex = span.startColumnIndex - scanStartColumnIndex;

  if (spanStartLocalColumnIndex < 0 || spanStartLocalColumnIndex >= lowerBannerRowText.length) {
    return;
  }

  const lowerStartText = normalizeBannerDiagnosticCellText(
    lowerBannerRowText[spanStartLocalColumnIndex]
  );

  if (!lowerStartText) {
    return;
  }

  let lowerAreaEndLocalColumnIndex = spanStartLocalColumnIndex;

  for (
    let localColumnIndex = spanStartLocalColumnIndex + 1;
    localColumnIndex < lowerBannerRowText.length;
    localColumnIndex++
  ) {
    const lowerCellText = normalizeBannerDiagnosticCellText(lowerBannerRowText[localColumnIndex]);

    if (!lowerCellText) {
      break;
    }

    lowerAreaEndLocalColumnIndex = localColumnIndex;
  }

  const lowerAreaEndSheetColumnIndex = scanStartColumnIndex + lowerAreaEndLocalColumnIndex;

  span.endColumnIndex = Math.min(span.endColumnIndex, lowerAreaEndSheetColumnIndex);
}

/**
 * Normalizes diagnostic banner cell text.
 */
function normalizeBannerDiagnosticCellText(rawValue) {
  if (rawValue === null || rawValue === undefined) {
    return "";
  }

  return String(rawValue).trim();
}

/**
 * Reads merge diagnostics for banner rows above selected range.
 *
 * PURPOSE:
 * Temporary spike for understanding how Office.js exposes merged cells.
 *
 * This function does not affect calculations.
 */
async function loadBannerMergeDiagnosticsForSelectedRange(context, selectedRange) {
  const maxBannerScanRows = 5;

  selectedRange.load(["rowIndex", "columnIndex", "columnCount"]);

  await context.sync();

  const selectedStartRowIndex = selectedRange.rowIndex;
  const selectedStartColumnIndex = selectedRange.columnIndex;
  const selectedColumnCount = selectedRange.columnCount;

  if (selectedStartRowIndex === 0) {
    return [
      "Merge diagnostics:",
      "- Над выделенным диапазоном нет строк для анализа merge-структуры.",
    ].join("\n");
  }

  const availableRowCount = Math.min(maxBannerScanRows + 1, selectedStartRowIndex);
  const firstBannerRowIndex = selectedStartRowIndex - availableRowCount;

  const bannerScanRange = selectedRange.worksheet.getRangeByIndexes(
    firstBannerRowIndex,
    selectedStartColumnIndex,
    availableRowCount,
    selectedColumnCount
  );

  bannerScanRange.load(["text", "values", "rowIndex", "columnIndex", "rowCount", "columnCount"]);

  await context.sync();

  const lines = [];

  lines.push("Merge diagnostics:");
  lines.push(
    `- Диапазон проверки: ${availableRowCount} строк над выделением, ${selectedColumnCount} колонок.`
  );

  for (let localRowIndex = 0; localRowIndex < availableRowCount; localRowIndex++) {
    const sheetRowIndex = firstBannerRowIndex + localRowIndex;
    const rowOffsetFromSelection = sheetRowIndex - selectedStartRowIndex;

    lines.push(`- Row offset ${rowOffsetFromSelection}:`);

    for (let localColumnIndex = 0; localColumnIndex < selectedColumnCount; localColumnIndex++) {
      const sheetColumnIndex = selectedStartColumnIndex + localColumnIndex;

      const cellText =
        bannerScanRange.text &&
        bannerScanRange.text[localRowIndex] &&
        bannerScanRange.text[localRowIndex][localColumnIndex] !== undefined
          ? bannerScanRange.text[localRowIndex][localColumnIndex]
          : "";

      const mergeInfo = await getCellMergeDiagnosticInfo(
        context,
        selectedRange.worksheet,
        sheetRowIndex,
        sheetColumnIndex
      );

      lines.push(`  - col ${localColumnIndex + 1}: text="${cellText}", ${mergeInfo}`);
    }
  }

  return lines.join("\n");
}

/**
 * Attempts to get merge diagnostic info for one worksheet cell.
 *
 * IMPORTANT:
 * Office.js merged-cell support can vary by runtime/API set.
 * This helper is intentionally defensive.
 */
async function getCellMergeDiagnosticInfo(context, worksheet, rowIndex, columnIndex) {
  const cell = worksheet.getRangeByIndexes(rowIndex, columnIndex, 1, 1);

  try {
    const mergedArea = cell.getMergedAreasOrNullObject();

    mergedArea.load([
      "isNullObject",
      "address",
      "rowIndex",
      "columnIndex",
      "rowCount",
      "columnCount",
    ]);

    await context.sync();

    if (mergedArea.isNullObject) {
      return "merged=false";
    }

    return [
      "merged=true",
      `area="${mergedArea.address}"`,
      `topRow=${mergedArea.rowIndex + 1}`,
      `leftCol=${mergedArea.columnIndex + 1}`,
      `rows=${mergedArea.rowCount}`,
      `cols=${mergedArea.columnCount}`,
    ].join(", ");
  } catch (error) {
    return `mergeInfo=unavailable (${error && error.message ? error.message : "unknown error"})`;
  }
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

  const oneTailedTestCheckbox = document.getElementById("one-tailed-test");

  return {
    confidenceLevel: confidenceLevelElement ? confidenceLevelElement.value : "95",
    oneTailedTest: oneTailedTestCheckbox ? oneTailedTestCheckbox.checked : false,

    roundCellValues: getCheckboxValue("round-cell-values"),

    compareWithPreviousColumn: getCheckboxValue("compare-with-previous-column"),
    applyPreviousColumnFill: getCheckboxValue("apply-previous-column-fill"),

    writeBannerLetters: getCheckboxValue("write-banner-letters"),
    respectBannerStructure: getCheckboxValue("respect-banner-structure"),
    autoDetectWaveBanners: getCheckboxValue("auto-detect-wave-banners"),
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
  initializeSettingsTooltips();

  bindMutuallyExclusiveCheckboxes("compare-only-with-total", "exclude-total-from-comparisons");

  bindMutuallyExclusiveCheckboxes("first-column-is-total", "total-in-each-banner");

  initializePreviousColumnComparisonSettings();
  initializeSettingsResetButton();
  initializeSettingsToggle();
  initializeBannerStructureSettings();

  const helpLink = document.getElementById("help-link");

  if (helpLink) {
    helpLink.addEventListener("click", (event) => {
      event.preventDefault();
      window.open(
        "https://hedgef0g.github.io/research-insights-toolkit/assets/rit-help-ru.html",
        "_blank"
      );
    });
  }
}

/**
 * Adds Russian tooltips to settings controls.
 *
 * PURPOSE:
 * Keep taskpane.html clean and define all setting explanations in one place.
 */
function initializeSettingsTooltips() {
  for (const [elementId, tooltipText] of Object.entries(SETTINGS_TOOLTIPS)) {
    const element = document.getElementById(elementId);

    if (!element) {
      continue;
    }

    element.title = tooltipText;
    element.setAttribute("aria-label", tooltipText);

    const explicitLabel = document.querySelector(`label[for="${elementId}"]`);

    if (explicitLabel) {
      explicitLabel.title = tooltipText;
      continue;
    }

    const wrappingLabel = element.closest("label");

    if (wrappingLabel) {
      wrappingLabel.title = tooltipText;
      continue;
    }

    const parent = element.parentElement;

    if (parent) {
      const labelLikeText = parent.querySelector("span, .checkbox-text, .label-text");

      if (labelLikeText) {
        labelLikeText.title = tooltipText;
      }
    }
  }
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

  const BANNER_UPPER_SCAN_LIMIT = 5;

  selectedRange.load(["rowIndex", "columnIndex", "columnCount"]);

  await context.sync();

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

  bannerRange.load("text");

  await context.sync();

  const bannerTexts = bannerRange.text[0] || [];
  const markerByColumnIndex = new Map();
  const clearMarkerColumnIndexes = new Set();

  const updatedBannerTexts = bannerTexts.map((currentText, columnIndex) => {
    if (calculationSettings.firstColumnIsTotal && columnIndex === 0) {
      clearMarkerColumnIndexes.add(columnIndex);
      return removeTrailingBannerMarker(currentText);
    }

    const markerIndex = calculationSettings.firstColumnIsTotal ? columnIndex - 1 : columnIndex;

    const marker = significanceLabels[markerIndex];

    if (!marker) {
      return currentText;
    }

    markerByColumnIndex.set(columnIndex, marker);

    return appendOrReplaceTrailingBannerMarker(currentText, marker);
  });

  const upperScanRowCount = Math.min(BANNER_UPPER_SCAN_LIMIT, selectedStartRowIndex - 1);
  const needsUpperScan =
    upperScanRowCount > 0 &&
    bannerTexts.some(
      (text, columnIndex) =>
        (text || "") === "" &&
        (markerByColumnIndex.has(columnIndex) || clearMarkerColumnIndexes.has(columnIndex))
    );

  let upperScanTexts = [];

  if (needsUpperScan) {
    const upperScanRange = selectedRange.worksheet.getRangeByIndexes(
      selectedStartRowIndex - 1 - upperScanRowCount,
      selectedStartColumnIndex,
      upperScanRowCount,
      selectedColumnCount
    );

    upperScanRange.load("text");

    await context.sync();

    upperScanTexts = upperScanRange.text.slice().reverse();
  }

  const cellWriteQueue = [];

  for (let columnIndex = 0; columnIndex < selectedColumnCount; columnIndex++) {
    const currentText = bannerTexts[columnIndex] || "";
    const nextText = updatedBannerTexts[columnIndex] || "";
    const marker = markerByColumnIndex.get(columnIndex);
    const shouldClearMarker = clearMarkerColumnIndexes.has(columnIndex);

    if (currentText === "" && (marker || shouldClearMarker)) {
      let queued = false;

      for (let rowOffset = 0; rowOffset < upperScanTexts.length; rowOffset++) {
        const upperCellText =
          (upperScanTexts[rowOffset] && upperScanTexts[rowOffset][columnIndex]) || "";

        if (upperCellText !== "") {
          const updatedUpperCellText = marker
            ? appendOrReplaceTrailingBannerMarker(upperCellText, marker)
            : getTrailingBannerMarker(upperCellText)
              ? removeTrailingBannerMarker(upperCellText)
              : upperCellText;

          if (updatedUpperCellText === upperCellText) {
            queued = true;
            break;
          }

          cellWriteQueue.push({
            rowIndex: selectedStartRowIndex - 2 - rowOffset,
            colIndex: selectedStartColumnIndex + columnIndex,
            text: updatedUpperCellText,
          });
          queued = true;
          break;
        }
      }

      if (queued) {
        continue;
      }
    }

    if (nextText === "" && currentText === "") {
      continue;
    }

    if (nextText === currentText) {
      continue;
    }

    cellWriteQueue.push({
      rowIndex: selectedStartRowIndex - 1,
      colIndex: selectedStartColumnIndex + columnIndex,
      text: nextText,
    });
  }

  for (const { rowIndex, colIndex, text } of cellWriteQueue) {
    const cell = selectedRange.worksheet.getRangeByIndexes(rowIndex, colIndex, 1, 1);

    cell.values = [[text]];
  }

  await context.sync();
}

/**
 * Writes group-local significance letters into the lowest banner level.
 *
 * The lowest banner level is the row immediately above selected range.
 *
 * RULES:
 * - upper banner levels are not modified;
 * - labels are local to banner groups;
 * - Total columns are skipped if excluded from comparisons;
 * - global Total column is skipped;
 * - existing trailing marker like "(a)" is replaced;
 * - same trailing marker is not duplicated.
 */
async function writeBannerMarkersAboveSelectedRangeUsingBannerStructure(
  context,
  selectedRange,
  bannerStructure,
  calculationSettings
) {
  const BANNER_UPPER_SCAN_LIMIT = 5;

  selectedRange.load(["rowIndex", "columnIndex", "columnCount"]);

  await context.sync();

  const selectedStartRowIndex = selectedRange.rowIndex;
  const selectedStartColumnIndex = selectedRange.columnIndex;
  const selectedColumnCount = selectedRange.columnCount;

  if (selectedStartRowIndex === 0) {
    setStatusMessage(
      "Данные расположены в первой строке. Добавьте строку над выделенным массивом для подстановки букв в баннер."
    );

    return;
  }

  const labelMap = buildBannerLocalSignificanceLabelMap(bannerStructure, calculationSettings);

  const bannerRange = selectedRange.worksheet.getRangeByIndexes(
    selectedStartRowIndex - 1,
    selectedStartColumnIndex,
    1,
    selectedColumnCount
  );

  bannerRange.load("text");

  await context.sync();

  const currentBannerTexts = bannerRange.text[0] || [];
  const nextBannerTexts = [];

  for (let columnIndex = 0; columnIndex < selectedColumnCount; columnIndex++) {
    const currentText = currentBannerTexts[columnIndex] || "";
    const label = labelMap.get(columnIndex);

    if (!label) {
      nextBannerTexts.push(removeTrailingBannerMarker(currentText));
      continue;
    }

    nextBannerTexts.push(appendOrReplaceTrailingBannerMarker(currentText, label));
  }

  // When a lower banner cell is blank but carries a label, the visible banner
  // header lives in a row above (multi-row banner layout). Load upper rows so
  // we can redirect the marker write to the nearest non-empty cell above.
  const upperScanRowCount = Math.min(BANNER_UPPER_SCAN_LIMIT, selectedStartRowIndex - 1);
  const needsUpperScan =
    upperScanRowCount > 0 &&
    currentBannerTexts.some((text, i) => (text || "") === "" && labelMap.has(i));

  let upperScanTexts = [];

  if (needsUpperScan) {
    const upperScanRange = selectedRange.worksheet.getRangeByIndexes(
      selectedStartRowIndex - 1 - upperScanRowCount,
      selectedStartColumnIndex,
      upperScanRowCount,
      selectedColumnCount
    );

    upperScanRange.load("text");

    await context.sync();

    // Reverse so index 0 = row immediately above the lower banner row.
    upperScanTexts = upperScanRange.text.slice().reverse();
  }

  const cellWriteQueue = [];

  for (let columnIndex = 0; columnIndex < selectedColumnCount; columnIndex++) {
    const nextText = nextBannerTexts[columnIndex] || "";
    const currentText = currentBannerTexts[columnIndex] || "";

    if (nextText === "" && currentText === "") {
      continue;
    }

    // Lower banner cell is blank but this column has a label: find the nearest
    // non-empty cell above and write the marker there instead.
    if (currentText === "" && labelMap.get(columnIndex)) {
      const label = labelMap.get(columnIndex);
      let queued = false;

      for (let rowOffset = 0; rowOffset < upperScanTexts.length; rowOffset++) {
        const upperCellText =
          (upperScanTexts[rowOffset] && upperScanTexts[rowOffset][columnIndex]) || "";

        if (upperCellText !== "") {
          cellWriteQueue.push({
            rowIndex: selectedStartRowIndex - 2 - rowOffset,
            colIndex: selectedStartColumnIndex + columnIndex,
            text: appendOrReplaceTrailingBannerMarker(upperCellText, label),
          });
          queued = true;
          break;
        }
      }

      if (!queued) {
        // No non-empty cell found above; fall back to the lower banner row.
        cellWriteQueue.push({
          rowIndex: selectedStartRowIndex - 1,
          colIndex: selectedStartColumnIndex + columnIndex,
          text: nextText,
        });
      }

      continue;
    }

    // Normal case: lower banner cell is non-empty; write in place.
    cellWriteQueue.push({
      rowIndex: selectedStartRowIndex - 1,
      colIndex: selectedStartColumnIndex + columnIndex,
      text: nextText,
    });
  }

  for (const { rowIndex, colIndex, text } of cellWriteQueue) {
    const cell = selectedRange.worksheet.getRangeByIndexes(rowIndex, colIndex, 1, 1);

    cell.numberFormat = [["@"]];
    cell.values = [[text]];
  }

  await context.sync();
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
  const previousColumnFillCheckbox = document.getElementById("apply-previous-column-fill");

  const firstColumnIsTotalCheckbox = document.getElementById("first-column-is-total");
  const excludeTotalCheckbox = document.getElementById("exclude-total-from-comparisons");

  if (previousColumnCheckbox) {
    previousColumnCheckbox.addEventListener("change", () => {
      if (previousColumnCheckbox.checked && previousColumnFillCheckbox) {
        previousColumnFillCheckbox.checked = true;
      }

      refreshSettingsPanelState();
      handleSettingsPersistenceAfterChange();
    });
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
  refreshBannerStructureSettingsState();
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
  const respectBannerStructureCheckbox = document.getElementById("respect-banner-structure");

  const fillOnlyTotalComparisonsCheckbox = document.getElementById("fill-only-total-comparisons");

  const warningElement = document.getElementById("previous-column-total-warning");

  if (!previousColumnCheckbox) {
    return;
  }

  const isPreviousColumnMode = previousColumnCheckbox.checked;
  const firstColumnIsTotal = firstColumnIsTotalCheckbox
    ? firstColumnIsTotalCheckbox.checked
    : false;

  const respectBannerStructure = respectBannerStructureCheckbox
    ? respectBannerStructureCheckbox.checked
    : false;

  const hasValidTotalSource = firstColumnIsTotal || respectBannerStructure;

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
    if (!hasValidTotalSource) {
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
      !respectBannerStructure &&
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

/**
 * Initializes Reset settings button.
 */
function initializeSettingsResetButton() {
  const resetButton = document.getElementById("reset-settings");

  if (!resetButton) {
    return;
  }

  resetButton.addEventListener("click", () => {
    resetSettingsToDefaults();
  });
}

/**
 * Resets UI settings to defaults and clears saved local settings.
 */
function resetSettingsToDefaults() {
  applySettingsToPanel(DEFAULT_CALCULATION_SETTINGS);
  clearSavedLocalSettings();
  refreshSettingsPanelState();

  setStatusMessage("Настройки сброшены к значениям по умолчанию.");
}

/**
 * Applies UI rules for banner structure mode.
 *
 * RULE:
 * If respect-banner-structure is enabled, manual Total placement settings
 * are disabled because Total placement will be detected by banner engine.
 */
function refreshBannerStructureSettingsState() {
  const respectBannerStructureCheckbox = document.getElementById("respect-banner-structure");
  if (!respectBannerStructureCheckbox) {
    return;
  }
  const respectBannerStructure = respectBannerStructureCheckbox.checked;

  const autoDetectWaveBannersCheckbox = document.getElementById("auto-detect-wave-banners");
  const autoDetectWaveBannersWrapper = document.getElementById("auto-detect-wave-banners-wrapper");

  const firstColumnIsTotalCheckbox = document.getElementById("first-column-is-total");
  const totalInEachBannerCheckbox = document.getElementById("total-in-each-banner");

  if (autoDetectWaveBannersWrapper) {
    autoDetectWaveBannersWrapper.style.display = respectBannerStructure ? "" : "none";
  }

  if (autoDetectWaveBannersCheckbox && !respectBannerStructure) {
    autoDetectWaveBannersCheckbox.checked = false;
  }

  if (firstColumnIsTotalCheckbox) {
    if (respectBannerStructure) {
      firstColumnIsTotalCheckbox.checked = false;
      firstColumnIsTotalCheckbox.disabled = true;
    } else {
      firstColumnIsTotalCheckbox.disabled = false;
    }
  }

  if (totalInEachBannerCheckbox) {
    if (respectBannerStructure) {
      totalInEachBannerCheckbox.checked = false;
      totalInEachBannerCheckbox.disabled = true;
    } else {
      totalInEachBannerCheckbox.disabled = false;
    }
  }
}

function initializeBannerStructureSettings() {
  const respectBannerStructureCheckbox = document.getElementById("respect-banner-structure");

  if (respectBannerStructureCheckbox) {
    respectBannerStructureCheckbox.addEventListener("change", refreshSettingsPanelState);
  }

  refreshSettingsPanelState();
}

/**
 * Converts zero-based column index to Excel column letter.
 */
function getExcelColumnLetter(columnIndex) {
  let dividend = columnIndex + 1;
  let columnName = "";

  while (dividend > 0) {
    const modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }

  return columnName;
}

function getFirstBannerStructureError(bannerStructure) {
  if (!bannerStructure || !bannerStructure.messages) {
    return null;
  }

  return bannerStructure.messages.find((message) => message.severity === "error") || null;
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
  const markerMatch = text.match(/\s*\(([^()]*)\)\s*$/);

  if (!markerMatch) {
    return null;
  }

  const markerLabel = markerMatch[1];

  if (!generateSignificanceLabels().includes(markerLabel)) {
    return null;
  }

  return {
    label: markerLabel,
    start: markerMatch.index,
  };
}


