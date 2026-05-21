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

import {
  interpretSelectedRange,
  loadLabelValuesForSelectedRange,
  detectLeadingEmptyColumns,
} from "./selected-range-interpreter";

const USER_VISIBLE_BANNER_MESSAGE_CODES = new Set([
  "GLOBAL_TOTAL_USED",
  "BANNER_AUTO_PREVIOUS_COLUMN_APPLIED",
  "BANNER_TOTAL_ONLY_NO_TOTAL_PAIRS",
  "BANNER_MULTIPLE_LOCAL_TOTALS",
  "BANNER_TOTAL_OUTSIDE_SELECTION",
  "BANNER_MALFORMED_STRUCTURE",
  "BANNER_NO_ROWS_ABOVE_SELECTION",
]);

const SCAN_CELL_LIMIT = 250000;
const INVENTORY_CONTENT_SHEET_NAME = "Content";
const INVENTORY_CONTENT_COLUMNS = [
  "#",
  "Sheet",
  "Title",
  "Range",
  "Rows",
  "Columns",
  "Status",
  "Summary",
  "Notes",
  "Warnings",
  "Critical",
];

const INVENTORY_CLIENT_COLUMNS = ["#", "Название таблицы", "Подзаголовок", "Лист"];

const INVENTORY_FULL_CHECK_COLUMNS = [
  "#",
  "Sheet",
  "Title",
  "Range",
  "Status",
  "Summary",
  "Rows",
  "Columns",
  "Metric rows",
  "Base rows",
  "Blocks",
  "Selected base",
  "Metric types",
  "Warnings",
  "Critical",
  "Issue codes",
  "Notes",
  "Label split",
  "Label cols",
];

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

  { id: "preferred-base", type: "value", settingName: "preferredBase" },

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

  preferredBase: "auto",

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

  "preferred-base":
    "Выберите тип базы для расчёта значимости. «Авто» использует приоритет: Effective → Unweighted → Base → Weighted. Если выбранный тип базы не найден в таблице, используется автоматический приоритет.",

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
  const runAllTablesButton = document.getElementById("run-all-tables"); // Auto-runner button.
  const clearAllTablesButton = document.getElementById("clear-all-tables"); // Auto-clear button.

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

  if (runAllTablesButton) {
    runAllTablesButton.addEventListener("click", runAutoSignificance);
  }

  if (clearAllTablesButton) {
    clearAllTablesButton.addEventListener("click", clearAutoSignificance);
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

    selectedRange.load(["values", "text", "rowIndex", "columnIndex", "rowCount", "columnCount"]);

    await context.sync();

    const selectedValues = selectedRange.values;
    const selectedText = selectedRange.text;

    if (!selectedValues || selectedValues.length < 2 || selectedValues[0].length < 2) {
      setStatusMessage("Please select at least 2 columns and 2 rows.");
      return;
    }

    const interpretation = await interpretSelectedRange(
      context,
      selectedRange,
      selectedValues,
      selectedText,
      calculationSettings
    );

    if (interpretation.state === "blocked") {
      const codes =
        interpretation.blockingReasons && interpretation.blockingReasons.length > 0
          ? ` [${interpretation.blockingReasons.join(", ")}]`
          : "";
      setStatusMessage(`${interpretation.blockingMessage}${codes}`);
      return;
    }

    const {
      valuesForCalculation,
      textForCalculation,
      leftLabelValues,
      normalizationStatusLines,
      bannerContext: interpretedBannerContext,
    } = interpretation;

    const selectedRangeGuardrailWarnings = interpretation.selectedRangeGuardrailWarnings;
    let { writeTargetRange } = interpretation;

    if (
      !valuesForCalculation ||
      valuesForCalculation.length < 2 ||
      !valuesForCalculation[0] ||
      valuesForCalculation[0].length < 2
    ) {
      setStatusMessage("Please select at least 2 columns and 2 rows.");
      return;
    }

    writeTargetRange.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);

    await context.sync();

    const targetStartRowIndex = writeTargetRange.rowIndex;

    if (calculationSettings.writeBannerLetters && targetStartRowIndex === 0) {
      outputElement.textContent =
        "Данные расположены в первой строке. Добавьте строку над выделенным массивом и запустите расчёт повторно.";

      return;
    }

    writeTargetRange.values = valuesForCalculation;

    writeTargetRange.format.font.bold = false;
    writeTargetRange.format.fill.clear();
    writeTargetRange.format.horizontalAlignment = "Center";
    writeTargetRange.format.verticalAlignment = "Center";

    await context.sync();

    const detectionResult = detectMetricRowsFromLeftLabels(
      valuesForCalculation,
      leftLabelValues
    ); // Row type diagnostics based on left-side labels.

    const calculationBlocks = buildCalculationBlocks(detectionResult, { preferredBase: calculationSettings.preferredBase }); // List of metric blocks to calculate.

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
      // interpretedBannerContext is already sanitized by selected-range-interpreter
      // (all RIT markers stripped from banner rows) so detection is idempotent
      // across repeated Runs and Checks.
      const bannerContext = interpretedBannerContext;

      bannerStructure = detectBannerStructure(bannerContext, calculationSettings);

      if (bannerContext.messages && bannerContext.messages.length > 0) {
        bannerStructure.messages = [...bannerContext.messages, ...(bannerStructure.messages || [])];
      }
    }

    const fullCellResultMatrix = createEmptyCellResultMatrix(
      valuesForCalculation.length,
      valuesForCalculation[0].length
    ); // Full-size marker storage matching the data body being calculated.

    for (const calculationBlock of calculationBlocks) {
      const smallBaseResult = applySmallBaseRulesForCalculationBlock(
        valuesForCalculation,
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
        valuesForCalculation,
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
          appendSelectedRangeGuardrailMessages(statusMessages, selectedRangeGuardrailWarnings).join(
            "\n"
          )
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
      writeTargetRange,
      textForCalculation,
      fullCellResultMatrix,
      detectionResult,
      calculationSettings
    );

    if (calculationSettings.writeBannerLetters) {
      // Remove any stale RIT markers that a previous run may have written at
      // label-column positions (left of writeTargetRange).  These linger when
      // the user re-runs without clearing first after the dataColOffset fix.
      await clearStaleBannerMarkersLeftOfWriteRange(
        context,
        writeTargetRange.worksheet,
        selectedRange.columnIndex,
        writeTargetRange.rowIndex,
        writeTargetRange.columnIndex
      );

      if (calculationSettings.respectBannerStructure && bannerStructure) {
        await writeBannerMarkersAboveSelectedRangeUsingBannerStructure(
          context,
          writeTargetRange,
          bannerStructure,
          calculationSettings
        );
      } else {
        await writeBannerMarkersAboveSelectedRange(
          context,
          writeTargetRange,
          calculationSettings
        );
      }
    }

    await context.sync();

    const statusMessages = [`Расчёт выполнен. Обработано блоков: ${calculationBlocks.length}.`];

    if (normalizationStatusLines.length > 0) {
      statusMessages.push("");
      statusMessages.push(...normalizationStatusLines);
    }

    const bannerUserMessages = formatBannerUserMessages(bannerStructure);

    if (bannerUserMessages) {
      statusMessages.push("");
      statusMessages.push(bannerUserMessages);
    }

    setStatusMessage(
      appendSelectedRangeGuardrailMessages(statusMessages, selectedRangeGuardrailWarnings).join(
        "\n"
      )
    );
  });
}


/**
 * Runs the full significance pipeline for a single named range on a named sheet.
 *
 * Mirrors the core calculation path of runSignificanceFromSelection() without
 * touching the selected-range UI state. Returns a compact result object so the
 * caller can aggregate per-table outcomes without crashing on the first failure.
 *
 * Returns { status, blocksProcessed, message, rangeAddress }.
 * status: "processed" | "skipped" | "blocked" | "error"
 *
 * NOTE: runSignificanceFromSelection() and this helper share the same pipeline
 * logic. A future refactor could unify them once the auto-runner pattern is
 * proven stable. Tracked as a follow-up tech-debt item.
 */
async function runSignificanceForRange(sheetName, rangeAddress, calculationSettings) {
  return await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const sourceRange = worksheet.getRange(rangeAddress);

    sourceRange.load(["values", "text", "rowIndex", "columnIndex", "rowCount", "columnCount"]);

    await context.sync();

    const selectedValues = sourceRange.values;
    const selectedText = sourceRange.text;

    if (!selectedValues || selectedValues.length < 2 || selectedValues[0].length < 2) {
      return { status: "skipped", message: "слишком мало данных", rangeAddress };
    }

    const interpretation = await interpretSelectedRange(
      context,
      sourceRange,
      selectedValues,
      selectedText,
      calculationSettings
    );

    if (interpretation.state === "blocked") {
      const codes =
        interpretation.blockingReasons && interpretation.blockingReasons.length > 0
          ? ` [${interpretation.blockingReasons.join(", ")}]`
          : "";
      return {
        status: "blocked",
        message: `${interpretation.blockingMessage}${codes}`,
        rangeAddress,
      };
    }

    const {
      valuesForCalculation,
      textForCalculation,
      leftLabelValues,
      bannerContext: interpretedBannerContext,
    } = interpretation;

    let { writeTargetRange } = interpretation;

    if (
      !valuesForCalculation ||
      valuesForCalculation.length < 2 ||
      !valuesForCalculation[0] ||
      valuesForCalculation[0].length < 2
    ) {
      return { status: "skipped", message: "нет данных для расчёта", rangeAddress };
    }

    writeTargetRange.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);

    await context.sync();

    const targetStartRowIndex = writeTargetRange.rowIndex;

    if (calculationSettings.writeBannerLetters && targetStartRowIndex === 0) {
      return {
        status: "skipped",
        message: "данные в первой строке — баннер недоступен",
        rangeAddress,
      };
    }

    writeTargetRange.values = valuesForCalculation;
    writeTargetRange.format.font.bold = false;
    writeTargetRange.format.fill.clear();
    writeTargetRange.format.horizontalAlignment = "Center";
    writeTargetRange.format.verticalAlignment = "Center";

    await context.sync();

    const detectionResult = detectMetricRowsFromLeftLabels(valuesForCalculation, leftLabelValues);
    const calculationBlocks = buildCalculationBlocks(detectionResult, { preferredBase: calculationSettings.preferredBase });

    if (!calculationBlocks || calculationBlocks.length === 0) {
      return { status: "skipped", message: "нет блоков расчёта", rangeAddress };
    }

    let bannerStructure = null;

    if (calculationSettings.respectBannerStructure) {
      const bannerContext = interpretedBannerContext;
      bannerStructure = detectBannerStructure(bannerContext, calculationSettings);
      if (bannerContext && bannerContext.messages && bannerContext.messages.length > 0) {
        bannerStructure.messages = [...bannerContext.messages, ...(bannerStructure.messages || [])];
      }
    }

    const fullCellResultMatrix = createEmptyCellResultMatrix(
      valuesForCalculation.length,
      valuesForCalculation[0].length
    );

    for (const calculationBlock of calculationBlocks) {
      const smallBaseResult = applySmallBaseRulesForCalculationBlock(
        valuesForCalculation,
        calculationBlock,
        fullCellResultMatrix,
        calculationSettings
      );

      if (smallBaseResult.errorMessage) {
        return { status: "error", message: smallBaseResult.errorMessage, rangeAddress };
      }

      const blockCalculationSettings = {
        ...calculationSettings,
        excludedColumnIndexes: smallBaseResult.excludedColumnIndexes,
        bannerStructure,
      };

      const blockResults = calculateBlockResults(
        valuesForCalculation,
        calculationBlock,
        blockCalculationSettings
      );

      const bannerStructureError = getFirstBannerStructureError(bannerStructure);

      if (bannerStructureError) {
        return { status: "error", message: bannerStructureError.text, rangeAddress };
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
      writeTargetRange,
      textForCalculation,
      fullCellResultMatrix,
      detectionResult,
      calculationSettings
    );

    if (calculationSettings.writeBannerLetters) {
      await clearStaleBannerMarkersLeftOfWriteRange(
        context,
        writeTargetRange.worksheet,
        sourceRange.columnIndex,
        writeTargetRange.rowIndex,
        writeTargetRange.columnIndex
      );

      if (calculationSettings.respectBannerStructure && bannerStructure) {
        await writeBannerMarkersAboveSelectedRangeUsingBannerStructure(
          context,
          writeTargetRange,
          bannerStructure,
          calculationSettings
        );
      } else {
        await writeBannerMarkersAboveSelectedRange(context, writeTargetRange, calculationSettings);
      }
    }

    await context.sync();

    return {
      status: "processed",
      blocksProcessed: calculationBlocks.length,
      message: `обработано блоков: ${calculationBlocks.length}`,
      rangeAddress,
    };
  });
}

/**
 * Auto-runner: processes all "available" inventory candidates in the workbook.
 *
 * Collects the workbook inventory, filters to candidates with
 * candidateStatus === "available" that have a usable range address, then
 * calls runSignificanceForRange for each one. Skips the Content sheet and
 * any candidate that is uncertain, rejected, or has no range.
 *
 * Shows a compact summary (processed / skipped / error counts) in the
 * significance status panel when done.
 */
async function runAutoSignificance() {
  const calculationSettings = readCalculationSettingsFromPanel();

  if (calculationSettings.compareWithPreviousColumn && calculationSettings.compareOnlyWithTotal) {
    setStatusMessage(
      // eslint-disable-next-line quotes
      'Режим “Сравнение с предыдущей колонкой” несовместим с режимом “Сравнивать только с Тотал”.'
    );
    return;
  }

  if (
    calculationSettings.excludeTotalFromComparisons &&
    !calculationSettings.firstColumnIsTotal &&
    !calculationSettings.respectBannerStructure
  ) {
    setStatusMessage(
      'Для режима “Не сравнивать с Тотал” нужно указать расположение Тотала. Сейчас поддерживается вариант “Первая колонка — Тотал” или режим “Учитывать структуру баннера”.'
    );
    return;
  }

  if (
    calculationSettings.compareOnlyWithTotal &&
    !calculationSettings.firstColumnIsTotal &&
    !calculationSettings.respectBannerStructure
  ) {
    setStatusMessage(
      'Для режима “Сравнивать только с Тотал” нужно указать расположение Тотала. Сейчас поддерживается вариант “Первая колонка — Тотал” или режим “Учитывать структуру баннера”.'
    );
    return;
  }

  if (
    calculationSettings.compareWithPreviousColumn &&
    calculationSettings.fillOnlyTotalComparisons
  ) {
    setStatusMessage(
      'Режим “Сравнение с предыдущей колонкой” несовместим с настройкой “Заливка только для Тотала”.'
    );
    return;
  }

  // Collect inventory to identify eligible candidates.
  let inventoryResults;
  try {
    await Excel.run(async (context) => {
      inventoryResults = await collectWorkbookInventoryResults(context, calculationSettings);
      // Normalize without inserting backlinks — only needed for resolvedRangeAddress.
      normalizeBacklinkItems(inventoryResults.sheetResults, false);
    });
  } catch (err) {
    setStatusMessage(`Автозапуск: ошибка при сканировании книги — ${err.message || err}`);
    return;
  }

  // Partition candidates: eligible to process vs. pre-skipped due to status/range.
  // The Content sheet is excluded entirely and does not count toward skipped.
  const eligible = [];
  let skipped = 0;
  const detailLines = [];

  for (const sheetResult of inventoryResults.sheetResults) {
    if (sheetResult.sheetName === INVENTORY_CONTENT_SHEET_NAME) {
      continue;
    }
    for (const item of sheetResult.items) {
      const rangeAddr = item.resolvedRangeAddress || item.rangeAddress;
      const label = `- ${sheetResult.sheetName} ${rangeAddr || item.rangeAddress || "?"}`;

      if (!rangeAddr) {
        skipped++;
        detailLines.push(`${label}: пропущено — нет диапазона`);
      } else if (item.candidateStatus === "uncertain") {
        skipped++;
        detailLines.push(`${label}: пропущено — кандидат неопределён`);
      } else if (item.candidateStatus === "rejected") {
        skipped++;
        detailLines.push(`${label}: пропущено — не опознан как таблица ResearchSignal`);
      } else if (item.candidateStatus === "available" && item.canRunCheckTable) {
        eligible.push({
          sheetName: sheetResult.sheetName,
          rangeAddress: rangeAddr,
          title: item.resolvedTitle || (item.title || ""),
        });
      } else {
        // Catch-all for unknown future statuses.
        skipped++;
        detailLines.push(`${label}: пропущено — статус «${item.candidateStatus || "unknown"}»`);
      }
    }
  }

  if (eligible.length === 0) {
    const noEligibleLines = [
      "Автозапуск: доступных кандидатов не найдено.",
      'Проверьте статусы таблиц через «Найти таблицы» / «С полной проверкой».',
    ];
    if (skipped > 0) {
      noEligibleLines.push("", `Пропущено: ${skipped}.`, ...detailLines);
    }
    setStatusMessage(noEligibleLines.join("\n"));
    return;
  }

  let processed = 0;
  let errors = 0;

  for (const candidate of eligible) {
    try {
      const result = await runSignificanceForRange(
        candidate.sheetName,
        candidate.rangeAddress,
        calculationSettings
      );

      if (result.status === "processed") {
        processed++;
      } else if (result.status === "skipped" || result.status === "blocked") {
        skipped++;
        detailLines.push(
          `- ${candidate.sheetName} ${candidate.rangeAddress}: пропущено — ${result.message}`
        );
      } else {
        errors++;
        detailLines.push(
          `- ${candidate.sheetName} ${candidate.rangeAddress}: ошибка — ${result.message}`
        );
      }
    } catch (err) {
      errors++;
      detailLines.push(
        `- ${candidate.sheetName} ${candidate.rangeAddress}: ошибка — ${err.message || "неизвестная ошибка"}`
      );
    }
  }

  const summaryLines = [
    "Автозапуск завершён.",
    `Обработано таблиц: ${processed}.`,
    `Пропущено: ${skipped}.`,
    `Ошибок: ${errors}.`,
  ];

  if (detailLines.length > 0) {
    summaryLines.push("", ...detailLines);
  }

  setStatusMessage(summaryLines.join("\n"));
}

/**
 * Clears significance markers for a single named range on a named sheet.
 *
 * Mirrors the clear path of clearSignificanceFromSelection() without touching
 * the Excel selection UI state. Returns a compact result object so the caller
 * can aggregate per-table outcomes.
 *
 * Returns { status, message }.
 * status: "cleared" | "skipped" | "error"
 */
async function clearSignificanceForRange(sheetName, rangeAddress) {
  return await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const sourceRange = worksheet.getRange(rangeAddress);

    sourceRange.load(["values", "text"]);

    await context.sync();

    const selectedValues = sourceRange.values;
    const selectedText = sourceRange.text;

    if (!selectedValues || selectedValues.length < 1 || !selectedValues[0] || selectedValues[0].length < 1) {
      return { status: "skipped", message: "нет данных в диапазоне" };
    }

    const cleanedValues = removeSignificanceMarkersFromMatrix(selectedValues);
    const normalized = normalizeSelectedRange(cleanedValues, selectedText);

    if (normalized.normalizationNeeded && !normalized.normalizationApplied) {
      const codes =
        normalized.blockingReasons && normalized.blockingReasons.length > 0
          ? ` [${normalized.blockingReasons.join(", ")}]`
          : "";
      return { status: "skipped", message: `${normalized.blockingMessage}${codes}` };
    }

    let clearTargetRange;

    if (normalized.normalizationNeeded && normalized.normalizationApplied) {
      const bodyRowCount = normalized.valuesForCalculation.length;
      let bodyColCount = normalized.valuesForCalculation[0].length;
      let effectiveClearColOffset = normalized.dataColOffset;

      const clearLeadingEmptyCols = detectLeadingEmptyColumns(normalized.textForCalculation);
      if (clearLeadingEmptyCols > 0) {
        bodyColCount -= clearLeadingEmptyCols;
        effectiveClearColOffset += clearLeadingEmptyCols;
      }

      if (bodyRowCount < 1 || bodyColCount < 1) {
        return { status: "skipped", message: "нет данных после нормализации" };
      }

      clearTargetRange = sourceRange
        .getCell(normalized.dataRowOffset, effectiveClearColOffset)
        .getResizedRange(bodyRowCount - 1, bodyColCount - 1);
    } else {
      const leadingBlankColsForClear = detectLeadingEmptyColumns(selectedText);

      if (leadingBlankColsForClear > 0) {
        clearTargetRange = sourceRange
          .getCell(0, leadingBlankColsForClear)
          .getResizedRange(
            sourceRange.rowCount - 1,
            sourceRange.columnCount - leadingBlankColsForClear - 1
          );
      } else {
        clearTargetRange = sourceRange;
      }
    }

    clearTargetRange.load(["values", "numberFormat", "rowCount", "columnCount"]);

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

    await clearBannerMarkersAboveRange(context, clearTargetRange);

    return { status: "cleared", message: "очищено" };
  });
}

/**
 * Auto-clear: removes significance markers from all "available" inventory
 * candidates in the workbook.
 *
 * Mirrors runAutoSignificance() but calls clearSignificanceForRange() instead
 * of running the significance pipeline.
 */
async function clearAutoSignificance() {
  let inventoryResults;
  try {
    await Excel.run(async (context) => {
      inventoryResults = await collectWorkbookInventoryResults(context, readCalculationSettingsFromPanel());
      normalizeBacklinkItems(inventoryResults.sheetResults, false);
    });
  } catch (err) {
    setStatusMessage(`Автоочистка: ошибка при сканировании книги — ${err.message || err}`);
    return;
  }

  const eligible = [];
  let skipped = 0;
  const detailLines = [];

  for (const sheetResult of inventoryResults.sheetResults) {
    if (sheetResult.sheetName === INVENTORY_CONTENT_SHEET_NAME) {
      continue;
    }
    for (const item of sheetResult.items) {
      const rangeAddr = item.resolvedRangeAddress || item.rangeAddress;
      const label = `- ${sheetResult.sheetName} ${rangeAddr || item.rangeAddress || "?"}`;

      if (!rangeAddr) {
        skipped++;
        detailLines.push(`${label}: пропущено — нет диапазона`);
      } else if (item.candidateStatus === "uncertain") {
        skipped++;
        detailLines.push(`${label}: пропущено — кандидат неопределён`);
      } else if (item.candidateStatus === "rejected") {
        skipped++;
        detailLines.push(`${label}: пропущено — не опознан как таблица ResearchSignal`);
      } else if (item.candidateStatus === "available" && item.canRunCheckTable) {
        eligible.push({ sheetName: sheetResult.sheetName, rangeAddress: rangeAddr });
      } else {
        skipped++;
        detailLines.push(`${label}: пропущено — статус «${item.candidateStatus || "unknown"}»`);
      }
    }
  }

  if (eligible.length === 0) {
    const noEligibleLines = [
      "Автоочистка: доступных кандидатов не найдено.",
      'Проверьте статусы таблиц через «Найти таблицы» / «С полной проверкой».',
    ];
    if (skipped > 0) {
      noEligibleLines.push("", `Пропущено: ${skipped}.`, ...detailLines);
    }
    setStatusMessage(noEligibleLines.join("\n"));
    return;
  }

  let cleared = 0;
  let errors = 0;

  for (const candidate of eligible) {
    try {
      const result = await clearSignificanceForRange(candidate.sheetName, candidate.rangeAddress);

      if (result.status === "cleared") {
        cleared++;
      } else if (result.status === "skipped") {
        skipped++;
        detailLines.push(
          `- ${candidate.sheetName} ${candidate.rangeAddress}: пропущено — ${result.message}`
        );
      } else {
        errors++;
        detailLines.push(
          `- ${candidate.sheetName} ${candidate.rangeAddress}: ошибка — ${result.message}`
        );
      }
    } catch (err) {
      errors++;
      detailLines.push(
        `- ${candidate.sheetName} ${candidate.rangeAddress}: ошибка — ${err.message || "неизвестная ошибка"}`
      );
    }
  }

  const summaryLines = [
    "Автоочистка завершена.",
    `Очищено таблиц: ${cleared}.`,
    `Пропущено: ${skipped}.`,
    `Ошибок: ${errors}.`,
  ];

  if (detailLines.length > 0) {
    summaryLines.push("", ...detailLines);
  }

  setStatusMessage(summaryLines.join("\n"));
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
      let bodyColCount = normalized.valuesForCalculation[0].length;
      let effectiveClearColOffset = normalized.dataColOffset;

      // Mirror the tertiary strip in interpretSelectedRange State 2: the
      // normalizer may leave leading all-blank helper columns (e.g. column B
      // in a mean-only table) inside the normalized body.  Clear must exclude
      // those same columns so it does not widen the clear target beyond the
      // real data body.
      const clearLeadingEmptyCols = detectLeadingEmptyColumns(
        normalized.textForCalculation
      );
      if (clearLeadingEmptyCols > 0) {
        bodyColCount -= clearLeadingEmptyCols;
        effectiveClearColOffset += clearLeadingEmptyCols;
      }

      if (bodyRowCount < 1 || bodyColCount < 1) {
        setStatusMessage("Нет данных в выделенном диапазоне.");
        return;
      }

      clearTargetRange = selectedRange
        .getCell(normalized.dataRowOffset, effectiveClearColOffset)
        .getResizedRange(bodyRowCount - 1, bodyColCount - 1);
    } else {
      // Pass-through: the selection is a clean numeric-only range.
      // Detect leading all-blank helper columns and exclude them from the
      // clear target so that Clear does not remove fill/formatting from helper
      // cells.  Uses the same detectLeadingEmptyColumns function as Run's
      // interpretSelectedRange passThrough path.
      const leadingBlankColsForClear = detectLeadingEmptyColumns(selectedText);

      if (leadingBlankColsForClear > 0) {
        clearTargetRange = selectedRange
          .getCell(0, leadingBlankColsForClear)
          .getResizedRange(
            selectedRange.rowCount - 1,
            selectedRange.columnCount - leadingBlankColsForClear - 1
          );
      } else {
        clearTargetRange = selectedRange;
      }
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

    await clearBannerMarkersAboveRange(context, clearTargetRange);

    setStatusMessage("Significance markers removed.");
  });
}

/**
 * Removes RIT-generated trailing banner significance markers from the visible
 * banner/header cells above the data body that was just cleared.
 *
 * PURPOSE:
 * Mirrors the Run "mark letters in banner" placement so Clear undoes the same
 * cells Run wrote into, including sparse / vertically merged banner layouts
 * where the nearest non-empty cell above receives the marker.
 *
 * RULES:
 * - Scans the row immediately above the data body plus up to
 *   BANNER_UPPER_SCAN_LIMIT additional rows above, matching Run.
 * - Removes only trailing markers recognized by getTrailingBannerMarker,
 *   which restricts matches to single-character significance labels.
 *   Ordinary parenthesized header text such as "Wave (quarter)",
 *   "Brand (new)", or "Волна (квартал)" is preserved.
 * - Never writes to the data body or to the label column.
 */
async function clearBannerMarkersAboveRange(context, targetRange) {
  const BANNER_UPPER_SCAN_LIMIT = 5;

  targetRange.load(["rowIndex", "columnIndex", "columnCount"]);

  await context.sync();

  const targetStartRowIndex = targetRange.rowIndex;
  const targetStartColumnIndex = targetRange.columnIndex;
  const targetColumnCount = targetRange.columnCount;

  if (targetStartRowIndex === 0 || targetColumnCount < 1) {
    return;
  }

  const totalScanRowCount = Math.min(BANNER_UPPER_SCAN_LIMIT + 1, targetStartRowIndex);

  if (totalScanRowCount < 1) {
    return;
  }

  const bannerScanRange = targetRange.worksheet.getRangeByIndexes(
    targetStartRowIndex - totalScanRowCount,
    targetStartColumnIndex,
    totalScanRowCount,
    targetColumnCount
  );

  bannerScanRange.load("text");

  await context.sync();

  const bannerTexts = bannerScanRange.text;
  const cellWriteQueue = [];

  for (let rowOffset = 0; rowOffset < totalScanRowCount; rowOffset++) {
    const rowTexts = bannerTexts[rowOffset] || [];

    for (let columnIndex = 0; columnIndex < targetColumnCount; columnIndex++) {
      const currentText = rowTexts[columnIndex];

      if (currentText === null || currentText === undefined || currentText === "") {
        continue;
      }

      if (!getTrailingBannerMarker(currentText)) {
        continue;
      }

      cellWriteQueue.push({
        rowIndex: targetStartRowIndex - totalScanRowCount + rowOffset,
        colIndex: targetStartColumnIndex + columnIndex,
        text: removeTrailingBannerMarker(currentText),
      });
    }
  }

  if (cellWriteQueue.length === 0) {
    return;
  }

  for (const { rowIndex, colIndex, text } of cellWriteQueue) {
    const cell = targetRange.worksheet.getRangeByIndexes(rowIndex, colIndex, 1, 1);

    cell.values = [[text]];
  }

  await context.sync();
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

    if (
      !selectedValues ||
      selectedValues.length < 1 ||
      !selectedValues[0] ||
      selectedValues[0].length < 1
    ) {
      setCheckMessage("Нет данных в выделенном диапазоне.");
      return;
    }

    const interpretation = await interpretSelectedRange(
      context,
      selectedRange,
      selectedValues,
      selectedText,
      calculationSettings
    );

    if (interpretation.state === "blocked") {
      const codes =
        interpretation.blockingReasons.length > 0
          ? ` [${interpretation.blockingReasons.join(", ")}]`
          : "";
      setCheckMessage(`${interpretation.blockingMessage}${codes}`);
      return;
    }

    const { valuesForCalculation, leftLabelValues, bannerContext, normalized } = interpretation;

    const normalizationLines = [];

    if (interpretation.state === "normalized") {
      normalizationLines.push("Диапазон нормализован: заголовки/лейблы/баннер отделены от данных.");

      const parts = [];
      if (normalized.titleRows.length > 0) parts.push(`заголовков: ${normalized.titleRows.length}`);
      if (normalized.subtitleRows.length > 0)
        parts.push(`подзаголовков: ${normalized.subtitleRows.length}`);
      if (normalized.bannerRows.length > 0)
        parts.push(`строк баннера: ${normalized.bannerRows.length}`);
      if (normalized.labelColumns.length > 0)
        parts.push(`колонок меток: ${normalized.labelColumns.length}`);
      if (parts.length > 0) normalizationLines.push(`Отделено: ${parts.join(", ")}.`);
    }

    const modelInput = {
      values: valuesForCalculation,
      leftLabelValues,
      bannerContext,
      settings: calculationSettings,
      trailingBodyRows: normalized?.trailingBodyRows,
    };

    const model = buildTablePreviewModel(modelInput);
    const {
      summary,
      qualitySummary,
      userVisibleIssues,
      bannerStructure,
      calculationBlocks,
      rowDiagnostics,
    } = model;

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

    const bannerLines = formatCheckBannerSummary(bannerStructure);
    if (bannerLines.length > 0) {
      lines.push("");
      lines.push(...bannerLines);
    }

    const issueLines = formatCheckUserVisibleIssues(userVisibleIssues);
    if (issueLines.length > 0) {
      lines.push("");
      lines.push(...issueLines);
    }

    const blockLines = formatCheckCalculationBlocks(calculationBlocks, rowDiagnostics);
    if (blockLines.length > 0) {
      lines.push("");
      lines.push(...blockLines);
    }

    setCheckMessage(lines.join("\n"));
  });
}

function formatCheckUserVisibleIssues(issues) {
  if (!Array.isArray(issues) || issues.length === 0) {
    return [];
  }

  const lines = ["Проблемы проверки:"];
  for (const issue of issues) {
    lines.push(`- [${issue.severity}] ${issue.message}`);
  }
  return lines;
}

/**
 * Builds a compact calculation-block summary for the Check table text output.
 *
 * Returns an array of text lines. One header line ("Блоки расчёта: N.") is
 * followed by one bullet per block. If no blocks are detected, returns a
 * single "Блоки расчёта: не обнаружены." line.
 */
function formatCheckCalculationBlocks(calculationBlocks, rowDiagnostics) {
  if (!Array.isArray(calculationBlocks) || calculationBlocks.length === 0) {
    return ["Блоки расчёта: не обнаружены."];
  }

  const labelMap = new Map();
  if (Array.isArray(rowDiagnostics)) {
    for (const diag of rowDiagnostics) {
      if (diag && diag.rowIndex != null) {
        labelMap.set(diag.rowIndex, diag.primaryLabel || "");
      }
    }
  }

  const rowRef = (rowIndex) => {
    if (rowIndex == null) return null;
    const label = labelMap.get(rowIndex);
    return label ? `стр. ${rowIndex + 1} «${label}»` : `стр. ${rowIndex + 1}`;
  };

  const baseSubtypeNote = (subtype) => {
    if (subtype === "effective") return " [эффективная]";
    if (subtype === "unweighted") return " [невзвешенная]";
    if (subtype === "weighted") return " [взвешенная — ПРИМЕЧАНИЕ: резервный вариант]";
    return "";
  };

  const baseRef = (block) => {
    if (block.baseRowIndex == null) return "";
    return ` База: ${rowRef(block.baseRowIndex)}${baseSubtypeNote(block.baseSubtype)}.`;
  };

  const lines = [`Блоки расчёта: ${calculationBlocks.length}.`];

  for (const block of calculationBlocks) {
    switch (block.metricType) {
      case "proportion": {
        const count = Array.isArray(block.valueRowIndexes) ? block.valueRowIndexes.length : 0;
        lines.push(`- Пропорции: строк со значениями: ${count}.${baseRef(block)}`);
        break;
      }
      case "mean": {
        const meanRef = block.valueRowIndex != null ? rowRef(block.valueRowIndex) : "нет";
        const sdPart = block.sdRowIndex != null ? ` СО: ${rowRef(block.sdRowIndex)}.` : "";
        const varPart =
          block.varianceRowIndex != null ? ` Дисперсия: ${rowRef(block.varianceRowIndex)}.` : "";
        lines.push(`- Среднее: ${meanRef}.${sdPart}${varPart}${baseRef(block)}`);
        break;
      }
      case "npsStructure": {
        const npsRef = block.valueRowIndex != null ? rowRef(block.valueRowIndex) : "нет";
        const promPart =
          block.promotersRowIndex != null ? ` Промоутеры: ${rowRef(block.promotersRowIndex)}.` : "";
        const neutPart =
          block.neutralRowIndex != null ? ` Нейтральные: ${rowRef(block.neutralRowIndex)}.` : "";
        const detPart =
          block.detractorsRowIndex != null ? ` Критики: ${rowRef(block.detractorsRowIndex)}.` : "";
        lines.push(`- NPS: ${npsRef}.${promPart}${neutPart}${detPart}${baseRef(block)}`);
        break;
      }
      case "npsSpread": {
        const npsRef = block.valueRowIndex != null ? rowRef(block.valueRowIndex) : "нет";
        const sdPart = block.sdRowIndex != null ? ` СО: ${rowRef(block.sdRowIndex)}.` : "";
        const varPart =
          block.varianceRowIndex != null ? ` Дисперсия: ${rowRef(block.varianceRowIndex)}.` : "";
        lines.push(`- NPS (разброс): ${npsRef}.${sdPart}${varPart}${baseRef(block)}`);
        break;
      }
      default:
        lines.push(`- Блок «${block.metricType || "unknown"}»`);
    }
  }

  return lines;
}

/**
 * Builds a compact banner summary for the Check table text output.
 *
 * Surfaces enough of bannerStructure for manual validation that Check
 * consumes the same banner-aware interpretation inputs as Run.
 * Returns an array of text lines (empty if banner-aware was disabled).
 */
function formatCheckBannerSummary(bannerStructure) {
  if (!bannerStructure || bannerStructure.isEnabled !== true) {
    return ["Баннер: не проверялся (учёт структуры выключен)."];
  }

  if (!bannerStructure.isDetected) {
    return ["Баннер: не обнаружен."];
  }

  const groups = Array.isArray(bannerStructure.groups) ? bannerStructure.groups : [];
  const waveGroupCount = groups.filter((group) => group && group.semanticType === "wave").length;
  const localTotalGroupCount = groups.filter((group) => group && group.hasLocalTotal).length;
  const hasGlobalTotal =
    bannerStructure.globalTotalColumnIndex !== null &&
    bannerStructure.globalTotalColumnIndex !== undefined;

  const headerParts = [`Баннер: обнаружен. Групп: ${groups.length}.`];
  headerParts.push(`Wave-групп: ${waveGroupCount}.`);
  headerParts.push(`Local Total: ${localTotalGroupCount > 0 ? "да" : "нет"}.`);
  headerParts.push(
    hasGlobalTotal
      ? `Global Total: колонка данных ${bannerStructure.globalTotalColumnIndex + 1}.`
      : "Global Total: нет."
  );
  if (bannerStructure.recommendedComparisonMode) {
    headerParts.push(`Режим сравнения: ${bannerStructure.recommendedComparisonMode}.`);
  }

  const lines = [headerParts.join(" ")];

  if (groups.length > 0) {
    lines.push("Группы:");
    for (const group of groups) {
      const label = group && group.label ? group.label : "(без названия)";
      const columnIndexes = Array.isArray(group && group.columnIndexes) ? group.columnIndexes : [];
      const cols = columnIndexes.length;
      const semantic = group && group.semanticType ? group.semanticType : "default";
      const mode =
        group && group.recommendedComparisonMode ? group.recommendedComparisonMode : "default";
      const totalNote = group && group.hasLocalTotal ? ", local Total" : "";
      lines.push(`- ${label} — колонок ${cols} / ${mode} / ${semantic}${totalNote}`);
    }
  }

  return lines;
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

function getInventoryCandidateStatusLabel(candidateStatus) {
  if (candidateStatus === "available") {
    return "Кандидат — рекомендуется «Проверить таблицу»";
  }

  if (candidateStatus === "uncertain") {
    return "Кандидат неопределён — требуется «Проверить таблицу»";
  }

  return "Не опознан как таблица ResearchSignal";
}

function formatInventoryItemLines(item, index) {
  const lines = [];
  // Prefer resolvedTitle (set after backlink normalization) over raw title.
  // Fall back to raw title only if it is not a generated backlink marker.
  const displayTitle =
    item.resolvedTitle ||
    (item.title && !isGeneratedBacklinkRow(item.title) ? item.title : null);
  // Prefer resolvedRangeAddress (adjusted for any inserted backlink rows).
  const displayRange = item.resolvedRangeAddress || item.rangeAddress;
  const header = displayTitle ? `${index}. ${displayTitle} — ${displayRange}` : `${index}. ${displayRange}`;

  lines.push(header);
  lines.push(`   ${item.rowCount} строк, ${item.columnCount} колонок.`);

  if (item.previewSummary) {
    lines.push(`   ${item.previewSummary}.`);
  }

  if (item.selectedBaseSubtypeLabel) {
    lines.push(`   База: ${item.selectedBaseSubtypeLabel}.`);
  }

  const warnParts = [];
  if (item.criticalCount > 0) warnParts.push(`Критических: ${item.criticalCount}`);
  if (item.warningsCount > 0) warnParts.push(`Предупреждений: ${item.warningsCount}`);
  if (warnParts.length > 0) lines.push(`   ${warnParts.join(". ")}.`);

  // candidateStatus replaces the former "Значимость: да/нет" line.
  // Inventory is a candidate finder only; Check Table is the authoritative step.
  lines.push(`   ${getInventoryCandidateStatusLabel(item.candidateStatus)}.`);

  if (item.candidateNotes && item.candidateNotes.length > 0) {
    lines.push(`   [${item.candidateNotes.join("; ")}]`);
  }

  return lines;
}

function formatWorkbookInventoryMessage({ scannedSheets, sheetResults, skippedSheets }) {
  const totalCandidates = sheetResults.reduce((sum, sheetResult) => sum + sheetResult.items.length, 0);
  const lines = [
    `Таблиц в книге: ${totalCandidates}.`,
    `Листов с кандидатами: ${sheetResults.length}.`,
    `Просканировано листов: ${scannedSheets}.`,
  ];

  lines.push(`Лист ${INVENTORY_CONTENT_SHEET_NAME} обновлён.`);

  if (totalCandidates === 0) {
    lines.push("");
    lines.push("RIT не обнаружил в книге блоков данных, похожих на таблицы для проверки.");
  }

  for (const sheetResult of sheetResults) {
    lines.push("");
    lines.push(`Лист: ${sheetResult.sheetName}`);

    sheetResult.items.forEach((item, index) => {
      lines.push(...formatInventoryItemLines(item, index + 1));
      lines.push("");
    });

    if (lines[lines.length - 1] === "") {
      lines.pop();
    }
  }

  if (skippedSheets.length > 0) {
    lines.push("");
    lines.push("Пропущенные листы:");

    skippedSheets.forEach((sheet) => {
      if (sheet.reason === "empty") {
        lines.push(`- ${sheet.sheetName}: пустой лист.`);
        return;
      }

      lines.push(
        `- ${sheet.sheetName}: слишком большой для сканирования (${sheet.rowCount} стр. × ${sheet.columnCount} кол. = ${sheet.cellCount} ячеек, лимит: ${SCAN_CELL_LIMIT}).`
      );
    });
  }

  lines.push("");
  lines.push("Table Inventory — это только поиск кандидатов. Для проверки используйте «Проверить таблицу».");

  return lines.join("\n").trimEnd();
}

async function collectWorkbookInventoryResults(context, settings) {
  const worksheets = context.workbook.worksheets;
  worksheets.load("items/name");

  await context.sync();

  const worksheetEntries = worksheets.items
    .filter((worksheet) => worksheet.name !== INVENTORY_CONTENT_SHEET_NAME)
    .map((worksheet) => {
      const usedRange = worksheet.getUsedRangeOrNullObject();
      usedRange.load(["isNullObject", "rowIndex", "columnIndex", "rowCount", "columnCount"]);

      return { worksheet, usedRange };
    });

  await context.sync();

  const scannedEntries = [];
  const skippedSheets = [];

  for (const entry of worksheetEntries) {
    const { worksheet, usedRange } = entry;

    if (usedRange.isNullObject) {
      skippedSheets.push({ sheetName: worksheet.name, reason: "empty" });
      continue;
    }

    const cellCount = usedRange.rowCount * usedRange.columnCount;
    if (cellCount > SCAN_CELL_LIMIT) {
      skippedSheets.push({
        sheetName: worksheet.name,
        reason: "tooLarge",
        rowCount: usedRange.rowCount,
        columnCount: usedRange.columnCount,
        cellCount,
      });
      continue;
    }

    usedRange.load("values");
    scannedEntries.push(entry);
  }

  await context.sync();

  const sheetResults = scannedEntries
    .map(({ worksheet, usedRange }) => ({
      sheetName: worksheet.name,
      usedRangeRowOffset: usedRange.rowIndex,
      usedRangeColOffset: usedRange.columnIndex,
      usedRangeValues: usedRange.values,
      items: scanWorksheetForTables({
        values: usedRange.values,
        usedRangeRowOffset: usedRange.rowIndex,
        usedRangeColOffset: usedRange.columnIndex,
        sheetName: worksheet.name,
        settings,
      }),
    }))
    .filter((sheetResult) => sheetResult.items.length > 0);

  return {
    scannedSheets: scannedEntries.length,
    sheetResults,
    skippedSheets,
  };
}

function buildInventoryContentCandidateRows(sheetResults) {
  const rows = [];
  let candidateIndex = 1;

  for (const sheetResult of sheetResults) {
    for (const item of sheetResult.items) {
      rows.push([
        candidateIndex,
        sheetResult.sheetName,
        item.resolvedTitle || (isGeneratedBacklinkRow(item.title) ? "" : (item.title || "")),
        item.resolvedRangeAddress || item.rangeAddress || "",
        item.rowCount ?? "",
        item.columnCount ?? "",
        getInventoryCandidateStatusLabel(item.candidateStatus),
        item.previewSummary || "",
        item.candidateNotes && item.candidateNotes.length > 0 ? item.candidateNotes.join("; ") : "",
        item.warningsCount ?? 0,
        item.criticalCount ?? 0,
      ]);
      candidateIndex += 1;
    }
  }

  if (rows.length === 0) {
    rows.push([
      "",
      "",
      "",
      "",
      "",
      "",
      "Нет кандидатов",
      "RIT не обнаружил в книге блоков данных, похожих на таблицы для проверки.",
      "",
      "",
      "",
    ]);
  }

  return rows;
}

function toSafeExcelCellValue(value) {
  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return value;
  }

  if (Array.isArray(value)) {
    return value.map((item) => toSafeExcelCellValue(item)).join(", ");
  }

  if (value instanceof Date) {
    return value.toISOString();
  }

  if (typeof value === "object") {
    try {
      return JSON.stringify(value);
    } catch (error) {
      return String(value);
    }
  }

  return String(value);
}

function normalizeRowsToColumnCount(rows, columnCount) {
  if (!Array.isArray(rows) || rows.length === 0) {
    return [];
  }

  return rows.map((row) => {
    const sourceRow = Array.isArray(row) ? row : [row];
    const normalizedRow = [];

    for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      normalizedRow.push(toSafeExcelCellValue(sourceRow[columnIndex]));
    }

    return normalizedRow;
  });
}

function buildInventoryContentSkippedRows(skippedSheets) {
  if (!skippedSheets || skippedSheets.length === 0) {
    return [];
  }

  return skippedSheets.map((sheet) => {
    if (sheet.reason === "empty") {
      return [sheet.sheetName, "Skipped", "Пустой лист", "", ""];
    }

    return [
      sheet.sheetName,
      "Skipped",
      "Слишком большой для сканирования",
      `${sheet.rowCount} строк, ${sheet.columnCount} колонок`,
      `${sheet.cellCount} ячеек; лимит ${SCAN_CELL_LIMIT}`,
    ];
  });
}

function readContentOutputModeFromPanel() {
  const select = document.getElementById("content-output-mode");
  return select ? select.value : "minimal-check";
}

function readBacklinkSettingFromPanel() {
  return getCheckboxValue("content-add-backlinks");
}

/**
 * Parses the first cell reference from a range address like "B3:F15" or "A1".
 * Returns { rowIndex, colIndex } as 0-based sheet indexes, or null on failure.
 */
function parseRangeStartCell(rangeAddress) {
  if (!rangeAddress) return null;
  const match = rangeAddress.match(/^([A-Z]+)(\d+)/i);
  if (!match) return null;
  const colLetters = match[1].toUpperCase();
  const rowNum = parseInt(match[2], 10);
  if (!rowNum || rowNum < 1) return null;
  let colIndex = 0;
  for (let i = 0; i < colLetters.length; i++) {
    colIndex = colIndex * 26 + (colLetters.charCodeAt(i) - 64);
  }
  return { rowIndex: rowNum - 1, colIndex: colIndex - 1 };
}

/**
 * Shifts the start row and/or end row of an A1-notation range address independently.
 *
 * Examples:
 *   adjustRangeRows("A3:F15", { startOffset: 1, endOffset: 1 }) → "A4:F16"
 *   adjustRangeRows("A3:F15", { startOffset: 1, endOffset: 0 }) → "A4:F15"
 *   adjustRangeRows("A3:F15", { startOffset: 0, endOffset: 0 }) → "A3:F15"
 */
function adjustRangeRows(rangeAddress, { startOffset = 0, endOffset = 0 } = {}) {
  if (!rangeAddress) return rangeAddress;
  let result = rangeAddress;
  if (startOffset !== 0) {
    result = result.replace(/^([A-Z]+)(\d+)/i, (_, col, row) => col.toUpperCase() + (parseInt(row, 10) + startOffset));
  }
  if (endOffset !== 0 && result.includes(":")) {
    result = result.replace(/:([A-Z]+)(\d+)/i, (_, col, row) => ":" + col.toUpperCase() + (parseInt(row, 10) + endOffset));
  }
  return result;
}

/**
 * Annotates each inventory item with backlinkState and resolvedRangeAddress.
 *
 * Uses the already-loaded usedRangeValues on each sheetResult to detect
 * whether the first cell of the item's range (or the row immediately above it)
 * is an existing "← Оглавление" row, so no extra Office.js round trips are needed.
 *
 * backlinkState values:
 *   "in-range"      — first row of detected range IS a backlink row (scanner included it)
 *   "above-range"   — row immediately above detected range is a backlink row
 *   "will-insert"   — no backlink detected; one will be inserted above
 *   "cannot-insert" — table starts at sheet row 0; cannot insert above
 *
 * resolvedRangeAddress points to the actual table body after all insertions on
 * the sheet are accounted for. Items are sorted top-to-bottom per sheet so that
 * each will-insert item contributes +1 to the cumulative shift of all items below
 * it, matching the bottom-to-top insertion order used by ensureBacklinkRows.
 *
 * backlinksEnabled:
 *   true  → will-insert items predict a +1 row insertion and increment insertionsAbove.
 *   false → no insertions predicted; only in-range stale start-rows are corrected.
 */
function normalizeBacklinkItems(sheetResults, backlinksEnabled) {
  for (const sheetResult of sheetResults) {
    const { usedRangeRowOffset, usedRangeColOffset, usedRangeValues } = sheetResult;

    const cellAt = (r, c) => {
      if (r < 0 || r >= usedRangeValues.length) return null;
      const rowArr = usedRangeValues[r];
      if (!rowArr || c < 0 || c >= rowArr.length) return null;
      return rowArr[c];
    };

    // Pass 1: classify backlinkState using original scanned values.
    // Collect parsed positions for sorting.
    const parsedItems = sheetResult.items.map((item) => {
      const parsed = parseRangeStartCell(item.rangeAddress);
      if (!parsed) {
        item.backlinkState = "cannot-insert";
        return { item, rowIndex: -1 };
      }

      const localRow = parsed.rowIndex - usedRangeRowOffset;
      const localCol = parsed.colIndex - usedRangeColOffset;

      if (isGeneratedBacklinkRow(cellAt(localRow, localCol))) {
        item.backlinkState = "in-range";
        // Backlink is row localRow; try the row after it first (title shifted into
        // the band when backlink was inserted above a firstRowOfBand title), then
        // the row before it (title was above the original band, rowAbove case).
        item.resolvedTitle =
          resolveTitleLikeText(usedRangeValues[localRow + 1]) ||
          resolveTitleLikeText(usedRangeValues[localRow - 1]);
      } else if (isGeneratedBacklinkRow(cellAt(localRow - 1, localCol))) {
        item.backlinkState = "above-range";
        // Backlink is at localRow-1; original title was the row above the backlink.
        item.resolvedTitle = resolveTitleLikeText(usedRangeValues[localRow - 2]);
      } else if (parsed.rowIndex === 0) {
        item.backlinkState = "cannot-insert";
      } else {
        item.backlinkState = "will-insert";
      }

      return { item, rowIndex: parsed.rowIndex };
    });

    // Pass 2: compute resolvedRangeAddress with cumulative insertion accounting.
    // Sort top-to-bottom so insertionsAbove accumulates as items are processed.
    const sorted = parsedItems.slice().sort((a, b) => a.rowIndex - b.rowIndex);
    let insertionsAbove = 0;

    for (const { item } of sorted) {
      const base = insertionsAbove;

      switch (item.backlinkState) {
        case "in-range":
          // Existing backlink IS the first detected row — strip it (start +1).
          // No new insertion, so end shifts only by prior insertions above.
          item.resolvedRangeAddress = adjustRangeRows(item.rangeAddress, {
            startOffset: 1 + base,
            endOffset: base,
          });
          break;

        case "above-range":
          // Existing backlink above — no new insertion.
          // Both endpoints shift only by prior insertions above.
          item.resolvedRangeAddress = adjustRangeRows(item.rangeAddress, {
            startOffset: base,
            endOffset: base,
          });
          break;

        case "will-insert":
          if (backlinksEnabled) {
            // New backlink row inserted above the table: entire table shifts +1,
            // plus any prior insertions above on this sheet.
            item.resolvedRangeAddress = adjustRangeRows(item.rangeAddress, {
              startOffset: 1 + base,
              endOffset: 1 + base,
            });
            insertionsAbove++;
          } else {
            // Backlinks OFF: no insertion planned. Apply only prior baseShift.
            item.resolvedRangeAddress = adjustRangeRows(item.rangeAddress, {
              startOffset: base,
              endOffset: base,
            });
          }
          break;

        default: // cannot-insert
          item.resolvedRangeAddress = adjustRangeRows(item.rangeAddress, {
            startOffset: base,
            endOffset: base,
          });
          break;
      }
    }
  }
}

/**
 * Builds a map from candidate key → 1-based row number on the Content sheet.
 * Key format: "<sheetName>::<rangeAddress>"
 */
function buildContentRowMap(sheetResults, mode) {
  const map = new Map();
  const firstDataRow = mode === "client" ? 3 : 8;
  let index = 0;
  for (const sheetResult of sheetResults) {
    for (const item of sheetResult.items) {
      const key = `${sheetResult.sheetName}::${item.rangeAddress}`;
      map.set(key, firstDataRow + index);
      index++;
    }
  }
  return map;
}

/**
 * Builds a hyperlink documentReference pointing to the given 1-based row
 * in the Content sheet (column A).
 */
function getContentRowReference(contentSheetName, contentRow) {
  if (!contentSheetName || !contentRow) return null;
  const escaped = contentSheetName.replace(/'/g, "''");
  const needsQuotes = /[^A-Za-z0-9_]/.test(escaped);
  const quotedSheet = needsQuotes ? `'${escaped}'` : escaped;
  return `${quotedSheet}!A${contentRow}`;
}

const BACKLINK_MARKER = "← Оглавление";

/**
 * Returns true if the cell value is a generated backlink marker.
 */
function isGeneratedBacklinkRow(cellValue) {
  if (cellValue === null || cellValue === undefined) return false;
  return String(cellValue).trim() === BACKLINK_MARKER;
}

/**
 * Returns the first title-like text from a row, or "" if the row is not
 * title-like.  A row qualifies as title-like when it has at most maxNonEmpty
 * non-empty cells (sparse, like a merged heading), contains no numeric cells,
 * and does not consist solely of the backlink marker.
 * Mirrors the sparsity logic of the scanner's detectFirstRowTitle.
 */
function resolveTitleLikeText(rowValues, maxNonEmpty = 3) {
  if (!rowValues) return "";
  let nonEmpty = 0;
  let title = "";
  for (const cell of rowValues) {
    if (cell === null || cell === undefined || cell === "") continue;
    const s = String(cell).trim();
    if (!s) continue;
    nonEmpty++;
    if (nonEmpty > maxNonEmpty) return "";
    if (typeof cell === "number") return "";
    if (/^-?[\d.,]+%?$/.test(s)) return "";
    if (isGeneratedBacklinkRow(s)) return "";
    if (!title) title = s;
  }
  return title;
}

/**
 * Writes or updates a single backlink row for one candidate table.
 *
 * backlinkState drives the exact action:
 *   "in-range"      — backlink is already at detectedRowIndex; update in place.
 *   "above-range"   — backlink is at detectedRowIndex-1; update in place.
 *   "will-insert"   — insert a new row at detectedRowIndex, write backlink there.
 *   "cannot-insert" — table starts at sheet row 0; skip.
 *
 * detectedRowIndex is always the 0-based row index of the START of the
 * scanner-detected range (i.e., item.rangeAddress start), NOT the resolved
 * table body. This lets the function locate the correct row to write into
 * without depending on resolvedRangeAddress.
 */
async function writeOrUpdateBacklink(context, worksheet, detectedRowIndex, detectedColIndex, contentRow, backlinkState) {
  const contentRef = getContentRowReference(INVENTORY_CONTENT_SHEET_NAME, contentRow);

  const writeBacklinkToCell = (rowIndex, colIndex) => {
    const cell = worksheet.getRangeByIndexes(rowIndex, colIndex, 1, 1);
    cell.values = [[BACKLINK_MARKER]];
    if (contentRef) {
      try {
        cell.hyperlink = {
          documentReference: contentRef,
          textToDisplay: BACKLINK_MARKER,
          screenTip: `Оглавление, строка ${contentRow}`,
        };
      } catch (_) {
        // Non-fatal: leave plain text.
      }
    }
    cell.format.fill.clear();
    cell.format.font.bold = false;
  };

  if (backlinkState === "in-range") {
    // Backlink row IS the first row of the detected range — update it in place.
    writeBacklinkToCell(detectedRowIndex, detectedColIndex);
    await context.sync();
    return;
  }

  if (backlinkState === "above-range") {
    // Backlink row is immediately above the detected range — update it in place.
    writeBacklinkToCell(detectedRowIndex - 1, detectedColIndex);
    await context.sync();
    return;
  }

  if (backlinkState === "cannot-insert") {
    return;
  }

  // "will-insert": insert a blank row at detectedRowIndex, then write the backlink.
  worksheet.getRangeByIndexes(detectedRowIndex, 0, 1, 1).getEntireRow().insert(Excel.InsertShiftDirection.down);
  await context.sync();

  writeBacklinkToCell(detectedRowIndex, detectedColIndex);
  await context.sync();
}

/**
 * Ensures backlink rows exist above every detected candidate table.
 *
 * Uses item.backlinkState (set by normalizeBacklinkItems) to decide whether
 * to insert a new row or update an existing one.
 *
 * Candidates within each sheet are sorted bottom-to-top (by detected range
 * start) so that row insertions for lower tables do not shift the 0-based row
 * indexes of candidates higher up on the same sheet.
 */
async function ensureBacklinkRows(context, sheetResults, contentRowMap) {
  const worksheetCandidates = new Map();

  for (const sheetResult of sheetResults) {
    const candidates = [];
    for (const item of sheetResult.items) {
      const parsed = parseRangeStartCell(item.rangeAddress);
      if (!parsed) continue;
      const key = `${sheetResult.sheetName}::${item.rangeAddress}`;
      const contentRow = contentRowMap.get(key);
      if (contentRow == null) continue;
      candidates.push({
        detectedRowIndex: parsed.rowIndex,
        detectedColIndex: parsed.colIndex,
        contentRow,
        backlinkState: item.backlinkState || "will-insert",
      });
    }
    if (candidates.length > 0) {
      worksheetCandidates.set(sheetResult.sheetName, candidates);
    }
  }

  for (const [sheetName, candidates] of worksheetCandidates) {
    const sorted = candidates.slice().sort((a, b) => b.detectedRowIndex - a.detectedRowIndex);
    const worksheet = context.workbook.worksheets.getItem(sheetName);

    for (const candidate of sorted) {
      await writeOrUpdateBacklink(
        context,
        worksheet,
        candidate.detectedRowIndex,
        candidate.detectedColIndex,
        candidate.contentRow,
        candidate.backlinkState
      );
    }
  }
}

function getContentTableHyperlinkTarget(sheetName, rangeAddress) {
  if (!sheetName || !rangeAddress) return null;
  const escaped = sheetName.replace(/'/g, "''");
  const needsQuotes = /[^A-Za-z0-9_]/.test(escaped);
  const quotedSheet = needsQuotes ? `'${escaped}'` : escaped;
  return `${quotedSheet}!${rangeAddress}`;
}

function buildClientContentRows(sheetResults) {
  const rows = [];
  let index = 1;
  for (const sheetResult of sheetResults) {
    for (const item of sheetResult.items) {
      rows.push([index, item.resolvedTitle || (isGeneratedBacklinkRow(item.title) ? "" : (item.title || "")), "", item.sheetName || sheetResult.sheetName]);
      index++;
    }
  }
  if (rows.length === 0) {
    rows.push(["", "Кандидаты не обнаружены", "", ""]);
  }
  return rows;
}

async function ensureInventoryContentWorksheet(context) {
  const worksheets = context.workbook.worksheets;
  const worksheet = worksheets.getItemOrNullObject(INVENTORY_CONTENT_SHEET_NAME);
  worksheet.load("isNullObject");

  await context.sync();

  if (!worksheet.isNullObject) {
    return worksheet;
  }

  return worksheets.add(INVENTORY_CONTENT_SHEET_NAME);
}

function writeMinimalCheckContent(worksheet, inventoryResults) {
  const allItems = [];
  for (const sheetResult of inventoryResults.sheetResults) {
    for (const item of sheetResult.items) {
      allItems.push(item);
    }
  }

  const candidateRows = normalizeRowsToColumnCount(
    buildInventoryContentCandidateRows(inventoryResults.sheetResults),
    INVENTORY_CONTENT_COLUMNS.length
  );
  const skippedRows = normalizeRowsToColumnCount(
    buildInventoryContentSkippedRows(inventoryResults.skippedSheets),
    5
  );
  const totalCandidates = inventoryResults.sheetResults.reduce(
    (sum, sheetResult) => sum + sheetResult.items.length,
    0
  );

  const titleRange = worksheet.getRange("A1:K1");
  titleRange.values = normalizeRowsToColumnCount([["Table Inventory Content"]], 11);
  titleRange.merge();
  titleRange.format.font.bold = true;
  titleRange.format.font.size = 14;

  const metadataRows = normalizeRowsToColumnCount(
    [
      ["Generated sheet", INVENTORY_CONTENT_SHEET_NAME],
      ["Scanned sheets", inventoryResults.scannedSheets],
      ["Candidate sheets", inventoryResults.sheetResults.length],
      ["Detected candidates", totalCandidates],
      ["Reminder", "Inventory is a candidate finder only. Use «Проверить таблицу» for interpretation."],
    ],
    2
  );

  const metadataRange = worksheet.getRangeByIndexes(1, 0, metadataRows.length, 2);
  metadataRange.values = metadataRows;
  metadataRange.format.font.bold = false;
  worksheet.getRange("A2:A6").format.font.bold = true;

  const headerRowIndex = 7;
  const headerRange = worksheet.getRangeByIndexes(
    headerRowIndex - 1,
    0,
    1,
    INVENTORY_CONTENT_COLUMNS.length
  );
  headerRange.values = normalizeRowsToColumnCount([INVENTORY_CONTENT_COLUMNS], INVENTORY_CONTENT_COLUMNS.length);
  headerRange.format.font.bold = true;

  const candidateRange = worksheet.getRangeByIndexes(
    headerRowIndex,
    0,
    candidateRows.length,
    INVENTORY_CONTENT_COLUMNS.length
  );
  candidateRange.values = candidateRows;
  candidateRange.format.wrapText = true;

  const tableRange = worksheet.getRangeByIndexes(
    headerRowIndex - 1,
    0,
    candidateRows.length + 1,
    INVENTORY_CONTENT_COLUMNS.length
  );
  tableRange.format.borders.getItem("EdgeBottom").style = "Continuous";
  tableRange.format.borders.getItem("EdgeTop").style = "Continuous";
  tableRange.format.borders.getItem("EdgeLeft").style = "Continuous";
  tableRange.format.borders.getItem("EdgeRight").style = "Continuous";
  tableRange.format.borders.getItem("InsideHorizontal").style = "Continuous";
  tableRange.format.borders.getItem("InsideVertical").style = "Continuous";

  if (skippedRows.length > 0) {
    const skippedSectionStartRowIndex = headerRowIndex + candidateRows.length + 2;
    const skippedTitleRange = worksheet.getRangeByIndexes(skippedSectionStartRowIndex - 1, 0, 1, 5);
    skippedTitleRange.values = normalizeRowsToColumnCount([["Skipped sheets"]], 5);
    skippedTitleRange.merge();
    skippedTitleRange.format.font.bold = true;

    const skippedHeaderRange = worksheet.getRangeByIndexes(skippedSectionStartRowIndex, 0, 1, 5);
    skippedHeaderRange.values = normalizeRowsToColumnCount(
      [["Sheet", "Status", "Summary", "Notes", "Warnings"]],
      5
    );
    skippedHeaderRange.format.font.bold = true;

    const skippedDataRange = worksheet.getRangeByIndexes(
      skippedSectionStartRowIndex + 1,
      0,
      skippedRows.length,
      5
    );
    skippedDataRange.values = skippedRows;
    skippedDataRange.format.wrapText = true;
  }

  const columnWidths = [42, 120, 220, 92, 52, 70, 260, 190, 240, 72, 72];
  columnWidths.forEach((width, index) => {
    worksheet.getRangeByIndexes(0, index, 1, 1).format.columnWidth = width;
  });

  worksheet.getRange("A:K").format.verticalAlignment = "Top";

  const RANGE_COL_INDEX = 3;
  for (let i = 0; i < allItems.length; i++) {
    const item = allItems[i];
    const effectiveRange = item.resolvedRangeAddress || item.rangeAddress;
    const hyperlinkTarget = getContentTableHyperlinkTarget(item.sheetName, effectiveRange);
    if (hyperlinkTarget) {
      const cell = worksheet.getRangeByIndexes(headerRowIndex + i, RANGE_COL_INDEX, 1, 1);
      cell.hyperlink = {
        documentReference: hyperlinkTarget,
        screenTip: `${item.sheetName}!${effectiveRange}`,
      };
    }
  }
}

function writeClientFacingContent(worksheet, inventoryResults) {
  const allItems = [];
  for (const sheetResult of inventoryResults.sheetResults) {
    for (const item of sheetResult.items) {
      allItems.push(item);
    }
  }

  const clientRows = buildClientContentRows(inventoryResults.sheetResults);
  const colCount = INVENTORY_CLIENT_COLUMNS.length;

  const titleRange = worksheet.getRangeByIndexes(0, 0, 1, colCount);
  titleRange.values = normalizeRowsToColumnCount([["Инвентарь таблиц"]], colCount);
  titleRange.merge();
  titleRange.format.font.bold = true;
  titleRange.format.font.size = 14;

  const headerRowIndex = 2;
  const headerRange = worksheet.getRangeByIndexes(headerRowIndex - 1, 0, 1, colCount);
  headerRange.values = normalizeRowsToColumnCount([INVENTORY_CLIENT_COLUMNS], colCount);
  headerRange.format.font.bold = true;

  const dataRows = normalizeRowsToColumnCount(clientRows, colCount);
  const dataRange = worksheet.getRangeByIndexes(headerRowIndex, 0, dataRows.length, colCount);
  dataRange.values = dataRows;
  dataRange.format.wrapText = false;

  const tableRange = worksheet.getRangeByIndexes(headerRowIndex - 1, 0, dataRows.length + 1, colCount);
  tableRange.format.borders.getItem("EdgeBottom").style = "Continuous";
  tableRange.format.borders.getItem("EdgeTop").style = "Continuous";
  tableRange.format.borders.getItem("EdgeLeft").style = "Continuous";
  tableRange.format.borders.getItem("EdgeRight").style = "Continuous";
  tableRange.format.borders.getItem("InsideHorizontal").style = "Continuous";
  tableRange.format.borders.getItem("InsideVertical").style = "Continuous";

  const columnWidths = [42, 260, 180, 120];
  columnWidths.forEach((width, index) => {
    worksheet.getRangeByIndexes(0, index, 1, 1).format.columnWidth = width;
  });

  worksheet.getRangeByIndexes(0, 0, dataRows.length + headerRowIndex, colCount).format.verticalAlignment = "Top";

  const TITLE_COL_INDEX = 1;
  for (let i = 0; i < allItems.length; i++) {
    const item = allItems[i];
    const effectiveRange = item.resolvedRangeAddress || item.rangeAddress;
    const hyperlinkTarget = getContentTableHyperlinkTarget(item.sheetName, effectiveRange);
    if (hyperlinkTarget) {
      const cell = worksheet.getRangeByIndexes(headerRowIndex + i, TITLE_COL_INDEX, 1, 1);
      const displayTitle = item.resolvedTitle || (isGeneratedBacklinkRow(item.title) ? "" : (item.title || ""));
      cell.hyperlink = {
        documentReference: hyperlinkTarget,
        textToDisplay: displayTitle || `Таблица ${i + 1}`,
        screenTip: `${item.sheetName}!${effectiveRange}`,
      };
    }
  }
}

function buildFullCheckCandidateRows(sheetResults) {
  const rows = [];
  let candidateIndex = 1;

  for (const sheetResult of sheetResults) {
    for (const item of sheetResult.items) {
      const metricTypes = [];
      if (item.hasProportions) metricTypes.push("Пропорции");
      if (item.hasMeans) metricTypes.push("Средние");
      if (item.hasNps) metricTypes.push("NPS");

      rows.push([
        candidateIndex,
        sheetResult.sheetName,
        item.resolvedTitle || (isGeneratedBacklinkRow(item.title) ? "" : (item.title || "")),
        item.resolvedRangeAddress || item.rangeAddress || "",
        getInventoryCandidateStatusLabel(item.candidateStatus),
        item.previewSummary || "",
        item.rowCount ?? "",
        item.columnCount ?? "",
        item.detectedMetricRows ?? "",
        item.detectedBaseRows ?? "",
        item.detectedBlocks ?? "",
        item.selectedBaseSubtypeLabel || "",
        metricTypes.join(", "),
        item.warningsCount ?? 0,
        item.criticalCount ?? 0,
        (item.qualityIssueCodes || []).map((q) => q.code).join(", "),
        item.candidateNotes && item.candidateNotes.length > 0 ? item.candidateNotes.join("; ") : "",
        item.labelSplitConfidence || "",
        item.labelColCount ?? "",
      ]);
      candidateIndex += 1;
    }
  }

  if (rows.length === 0) {
    const emptyRow = new Array(INVENTORY_FULL_CHECK_COLUMNS.length).fill("");
    emptyRow[5] = "Нет кандидатов";
    rows.push(emptyRow);
  }

  return rows;
}

function writeFullCheckContent(worksheet, inventoryResults) {
  const allItems = [];
  for (const sheetResult of inventoryResults.sheetResults) {
    for (const item of sheetResult.items) {
      allItems.push(item);
    }
  }

  const colCount = INVENTORY_FULL_CHECK_COLUMNS.length;
  const candidateRows = normalizeRowsToColumnCount(
    buildFullCheckCandidateRows(inventoryResults.sheetResults),
    colCount
  );
  const totalCandidates = inventoryResults.sheetResults.reduce(
    (sum, sheetResult) => sum + sheetResult.items.length,
    0
  );

  const titleRange = worksheet.getRangeByIndexes(0, 0, 1, colCount);
  titleRange.values = normalizeRowsToColumnCount([["Table Full Check"]], colCount);
  titleRange.merge();
  titleRange.format.font.bold = true;
  titleRange.format.font.size = 14;

  const metadataRows = normalizeRowsToColumnCount(
    [
      ["Generated sheet", INVENTORY_CONTENT_SHEET_NAME],
      ["Scanned sheets", inventoryResults.scannedSheets],
      ["Candidate sheets", inventoryResults.sheetResults.length],
      ["Detected candidates", totalCandidates],
      ["Mode", "Full Check — expanded diagnostics for all detected candidates"],
    ],
    2
  );

  const metadataRange = worksheet.getRangeByIndexes(1, 0, metadataRows.length, 2);
  metadataRange.values = metadataRows;
  worksheet.getRange("A2:A6").format.font.bold = true;

  const headerRowIndex = 7;
  const headerRange = worksheet.getRangeByIndexes(headerRowIndex - 1, 0, 1, colCount);
  headerRange.values = normalizeRowsToColumnCount([INVENTORY_FULL_CHECK_COLUMNS], colCount);
  headerRange.format.font.bold = true;

  const candidateRange = worksheet.getRangeByIndexes(headerRowIndex, 0, candidateRows.length, colCount);
  candidateRange.values = candidateRows;
  candidateRange.format.wrapText = true;

  const tableRange = worksheet.getRangeByIndexes(
    headerRowIndex - 1,
    0,
    candidateRows.length + 1,
    colCount
  );
  tableRange.format.borders.getItem("EdgeBottom").style = "Continuous";
  tableRange.format.borders.getItem("EdgeTop").style = "Continuous";
  tableRange.format.borders.getItem("EdgeLeft").style = "Continuous";
  tableRange.format.borders.getItem("EdgeRight").style = "Continuous";
  tableRange.format.borders.getItem("InsideHorizontal").style = "Continuous";
  tableRange.format.borders.getItem("InsideVertical").style = "Continuous";

  const columnWidths = [42, 100, 160, 92, 200, 160, 50, 60, 72, 72, 50, 130, 110, 72, 72, 200, 170, 85, 70];
  columnWidths.forEach((width, index) => {
    worksheet.getRangeByIndexes(0, index, 1, 1).format.columnWidth = width;
  });

  worksheet.getRangeByIndexes(0, 0, candidateRows.length + headerRowIndex, colCount).format.verticalAlignment = "Top";

  const RANGE_COL_INDEX = 3;
  for (let i = 0; i < allItems.length; i++) {
    const item = allItems[i];
    const effectiveRange = item.resolvedRangeAddress || item.rangeAddress;
    const hyperlinkTarget = getContentTableHyperlinkTarget(item.sheetName, effectiveRange);
    if (hyperlinkTarget) {
      const cell = worksheet.getRangeByIndexes(headerRowIndex + i, RANGE_COL_INDEX, 1, 1);
      cell.hyperlink = {
        documentReference: hyperlinkTarget,
        screenTip: `${item.sheetName}!${effectiveRange}`,
      };
    }
  }
}

async function writeInventoryContentSheet(context, inventoryResults) {
  const worksheet = await ensureInventoryContentWorksheet(context);
  const existingUsedRange = worksheet.getUsedRangeOrNullObject();
  existingUsedRange.load("isNullObject");

  await context.sync();

  if (!existingUsedRange.isNullObject) {
    existingUsedRange.unmerge();
    existingUsedRange.clear();
  }

  const mode = readContentOutputModeFromPanel();

  if (mode === "client") {
    writeClientFacingContent(worksheet, inventoryResults);
  } else if (mode === "full-check") {
    writeFullCheckContent(worksheet, inventoryResults);
  } else {
    writeMinimalCheckContent(worksheet, inventoryResults);
  }

  worksheet.position = 0;

  await context.sync();
}

async function runTableInventory() {
  await Excel.run(async (context) => {
    const inventoryResults = await collectWorkbookInventoryResults(context, readCalculationSettingsFromPanel());
    const addBacklinks = readBacklinkSettingFromPanel();
    const contentMode = readContentOutputModeFromPanel();

    // Always normalize: fixes resolvedRangeAddress for items whose detected range
    // starts with a generated backlink row (in-range), even when backlinks are OFF.
    // When backlinks are ON, also predicts the +1 end-row shift for will-insert.
    normalizeBacklinkItems(inventoryResults.sheetResults, addBacklinks);

    const contentRowMap = addBacklinks
      ? buildContentRowMap(inventoryResults.sheetResults, contentMode)
      : null;

    // Write Content first (hyperlinks use resolvedRangeAddress when available).
    await writeInventoryContentSheet(context, inventoryResults);

    // Insert/update backlink rows after Content is written so the Content sheet
    // is not affected by row insertions in source sheets.
    if (addBacklinks && contentRowMap) {
      await ensureBacklinkRows(context, inventoryResults.sheetResults, contentRowMap);
    }

    setInventoryMessage(
      formatWorkbookInventoryMessage({
        scannedSheets: inventoryResults.scannedSheets,
        sheetResults: inventoryResults.sheetResults,
        skippedSheets: inventoryResults.skippedSheets,
      })
    );
  });
}

function setInventoryMessage(message) {
  const panel = document.getElementById("inventory-panel");
  const result = document.getElementById("inventory-result");
  if (panel) panel.style.display = "block";
  if (result) result.textContent = message || "";
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

    preferredBase: getInputValue("preferred-base", "auto"),

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
  initializePreviousColumnTotalWarningPlacement();

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

function initializePreviousColumnTotalWarningPlacement() {
  const warningElement = document.getElementById("previous-column-total-warning");
  const totalInEachBannerCheckbox = document.getElementById("total-in-each-banner");

  if (!warningElement || !totalInEachBannerCheckbox || !totalInEachBannerCheckbox.parentElement) {
    return;
  }

  totalInEachBannerCheckbox.parentElement.insertAdjacentElement("afterend", warningElement);
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
  const totalInEachBannerCheckbox = document.getElementById("total-in-each-banner");
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
  const totalInEachBanner = totalInEachBannerCheckbox ? totalInEachBannerCheckbox.checked : false;

  const respectBannerStructure = respectBannerStructureCheckbox
    ? respectBannerStructureCheckbox.checked
    : false;

  const hasManualTotalPlacement = firstColumnIsTotal || totalInEachBanner;
  const hasValidTotalSource = hasManualTotalPlacement || respectBannerStructure;

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
      hasManualTotalPlacement &&
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
 * Clears RIT-generated banner markers from banner cells that are within the
 * full selected range but LEFT of writeTargetRange (i.e., label-column
 * positions that were stripped by detectEmbeddedLabelColumns).
 *
 * This removes stale markers written by pre-fix runs so they do not appear
 * alongside the correctly-placed new markers after an in-place Run without
 * an explicit Clear between runs.
 *
 * Only acts when writeTargetRange is narrower than the original selection
 * (i.e., at least one label column was stripped on the left).
 *
 * @param {Excel.RequestContext} context
 * @param {Excel.Worksheet} worksheet
 * @param {number} selectedRangeColIndex  - column index of the original selectedRange
 * @param {number} writeTargetRowIndex    - rowIndex of writeTargetRange (first data row)
 * @param {number} writeTargetColIndex    - columnIndex of writeTargetRange (first data col)
 */
async function clearStaleBannerMarkersLeftOfWriteRange(
  context,
  worksheet,
  selectedRangeColIndex,
  writeTargetRowIndex,
  writeTargetColIndex
) {
  const leftColumnCount = writeTargetColIndex - selectedRangeColIndex;

  // Nothing to clean: write target starts at the selection boundary, or data
  // starts in row 0 (no banner rows above it).
  if (leftColumnCount <= 0 || writeTargetRowIndex === 0) {
    return;
  }

  const BANNER_SCAN_LIMIT = 6;
  const totalScanRows = Math.min(BANNER_SCAN_LIMIT, writeTargetRowIndex);

  const bannerScanRange = worksheet.getRangeByIndexes(
    writeTargetRowIndex - totalScanRows,
    selectedRangeColIndex,
    totalScanRows,
    leftColumnCount
  );
  bannerScanRange.load("text");
  await context.sync();

  const writeQueue = [];

  for (let rowOffset = 0; rowOffset < totalScanRows; rowOffset++) {
    const rowTexts = bannerScanRange.text[rowOffset] || [];

    for (let colOffset = 0; colOffset < leftColumnCount; colOffset++) {
      const cellText = rowTexts[colOffset] || "";

      if (!getTrailingBannerMarker(cellText)) {
        continue;
      }

      writeQueue.push({
        rowIndex: writeTargetRowIndex - totalScanRows + rowOffset,
        colIndex: selectedRangeColIndex + colOffset,
        text: removeTrailingBannerMarker(cellText),
      });
    }
  }

  if (writeQueue.length === 0) {
    return;
  }

  for (const { rowIndex, colIndex, text } of writeQueue) {
    worksheet.getRangeByIndexes(rowIndex, colIndex, 1, 1).values = [[text]];
  }

  await context.sync();
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

  if (!generateSignificanceLabels().includes(markerLabel)) {
    return null;
  }

  return {
    label: markerLabel,
    start: markerMatch.index,
  };
}
