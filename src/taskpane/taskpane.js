/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, window, Excel, Office */

import {
  createEmptyCellResultMatrix,
  applyComparisonResultsToFullCellResultMatrix,
  applySmallBaseRulesForCalculationBlock,
  keepMarkersOnlyInAllowedRows,
  compareProportionRowsUsingBaseRow,
  compareMeanBlockByRowIndexes,
  compareNpsStructureBlockByRowIndexes,
  compareNpsSpreadBlockByRowIndexes,
  generateSignificanceLabels,
  buildBannerLocalSignificanceLabelMap,
  isSignificanceMarkerLabel,
  detectSignificanceMarkerOverflow,
  detectBatchMarkerOverflow,
} from "../core/significance";

import {
  detectMetricRowsFromLeftLabels,
  buildCalculationBlocks,
  getAllowedMarkerRowIndexes,
  LABEL_SCAN_COLUMNS_LEFT,
} from "../core/metric-detector";

import { writeCellResultsToSelectedRange } from "../core/excel-writer";

import { buildTablePreviewModel } from "../core/table-preview-model";
import { scanWorksheetForTables } from "../core/table-inventory-scanner";

import {
  collectStatisticTypeLabels,
  buildSignificanceFootnoteCellValue,
  buildProcessedRangeFootnoteSuffix,
  resolveFootnoteSpan,
  resolveFootnotePlacement,
  FOOTNOTE_SCAN_WINDOW_ROWS,
} from "../core/significance-footnote";

import { detectBannerStructure } from "../core/banner-detector";

import { buildDesignRecolorJob } from "../core/design-recolor";

import { hasEmptyDataRowGap, selectionHasMultiTableGap } from "../core/range-normalizer";

import { filterWorkbookCandidates } from "../core/batch-candidate-filter";

import { interpretSelectedRange } from "./selected-range-interpreter";

import { resolveCurrentTableFromActiveCell } from "./active-cell-resolver";

import { t, setLanguage, loadSavedLanguage, applyI18n } from "./localization";

import { createMarkerOverflowDecider } from "./taskpane-dialogs";

import {
  setStatusMessage,
  setCheckMessage,
  setInventoryMessage,
  nonContiguousSelectionMessage,
  formatBannerUserMessages,
  formatBannerUserMessagesExcludingCodes,
  appendSelectedRangeGuardrailMessages,
  formatStatusWithSelectedRangeGuardrails,
  buildCheckResolverMessage,
  runningStatusMessage,
  buildBatchProgressStatus,
} from "./taskpane-status";

// Minimal UI yield so the browser can repaint the live progress status between
// table-level batch updates. Intentionally a 0ms timeout — no artificial delay.
function yieldToUi() {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

function getClearPipelineOptions() {
  return {
    setStatusMessage,
    runningStatusMessage,
    nonContiguousSelectionMessage,
    t,
    perfLog,
  };
}

async function clearSignificanceFromSelection() {
  return clearSignificanceFromSelectionPipeline(getClearPipelineOptions());
}

import {
  SETTINGS_CONTROL_CONFIG,
  DEFAULT_CALCULATION_SETTINGS,
  saveSettingsToLocalStorage,
  clearSavedLocalSettings,
  applySettingsToPanel,
  loadSavedSettingsIntoPanel,
  initializeSettingsPanel,
} from "./taskpane-settings";

import {
  formatSkippedCandidateDetail,
  checkMetricTypesFromBlocks,
  formatCheckUserVisibleIssues,
  formatCheckCalculationBlocks,
  formatCheckBannerSummary,
} from "./taskpane-check-formatters";

import {
  runReportSkipReasonLabel,
  runReportStatusLabel,
  runReportMetricTypes,
  formatIssueDetailsForReport,
  runReportWarningDetails,
} from "./taskpane-run-report-formatters";

import {
  BACKLINK_MARKER,
  isGeneratedBacklinkRow,
  getInventoryCandidateStatusLabel,
  getContentCandidateStatusLabel,
  formatInventoryItemLines,
  resolveContentDisplayTitle,
  buildClientContentRows,
  getContentTableHyperlinkTarget,
  getContentRowReference,
} from "./taskpane-inventory-formatters";

import {
  INVENTORY_SCAN_EMERGENCY_CELL_LIMIT,
  inventoryScanErrorMessage,
  formatInventorySkippedSheetLine,
  buildInventoryContentSkippedRow,
} from "./taskpane-inventory-scan";

import {
  BANNER_SCAN_AREA_STATS,
  perfNow,
  perfElapsed,
  perfLog,
  perfEnabled,
  perfBannerWriteProfileEnabled,
  createAggregatedWriterPerfDetails,
  mergeWriterPerfDetails,
  createAggregatedBannerPerfDetails,
  roundBannerDiagnosticRatio,
  mergeBannerPerfDetails,
} from "./taskpane-performance";

import {
  clearSignificanceFromSelection as clearSignificanceFromSelectionPipeline,
  clearSignificanceForRange,
  clearSignificanceForSheetBatched,
  clearSignificanceForWorkbookBatched,
  groupCandidatesBySheet,
  accumulateClearDetails,
  applySignificanceFootnoteRemovalJobsInOwnContext,
} from "./clear-pipeline";

const INVENTORY_CONTENT_SHEET_NAME = "Content";
const RUN_REPORT_SHEET_NAME = "Run report";
const GENERATED_SHEET_NAMES = new Set([INVENTORY_CONTENT_SHEET_NAME, RUN_REPORT_SHEET_NAME]);

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
  "Issue details",
  "Notes",
  "Label split",
  "Label cols",
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
  const checkTableButton = document.getElementById("check-table"); // Read-only table check button.
  const findTablesButton = document.getElementById("find-tables"); // Table inventory button.
  const runAllTablesButton = document.getElementById("run-all-tables"); // Auto-runner button.
  const clearAllTablesButton = document.getElementById("clear-all-tables"); // Auto-clear button.

  loadSavedLanguage();
  applyI18n();

  initializeSettingsPanel({
    initializePreviousColumnComparisonSettings,
    initializeSettingsResetButton,
    initializeBannerStructureSettings,
  });
  loadSavedSettingsIntoPanel();
  refreshSettingsPanelState();
  initializeSettingsPersistence();

  if (calculateButton) {
    calculateButton.addEventListener("click", runSignificanceFromSelection);
  }

  if (clearButton) {
    clearButton.addEventListener("click", clearSignificanceFromSelection);
  }

  if (checkTableButton) {
    checkTableButton.addEventListener("click", runCheckTable);
  }

  if (findTablesButton) {
    findTablesButton.addEventListener("click", runWorkbookCheck);
  }

  if (runAllTablesButton) {
    runAllTablesButton.addEventListener("click", runAutoSignificance);
  }

  if (clearAllTablesButton) {
    clearAllTablesButton.addEventListener("click", clearAutoSignificance);
  }

  const autorunCurrentTableButton = document.getElementById("autorun-current-table");
  const clearCurrentTableButton = document.getElementById("clear-current-table");
  const runSheetTablesButton = document.getElementById("run-sheet-tables");
  const clearSheetTablesButton = document.getElementById("clear-sheet-tables");

  if (autorunCurrentTableButton) {
    autorunCurrentTableButton.addEventListener("click", runAutoCurrentTableSignificance);
  }

  if (clearCurrentTableButton) {
    clearCurrentTableButton.addEventListener("click", clearAutoCurrentTableSignificance);
  }

  if (runSheetTablesButton) {
    runSheetTablesButton.addEventListener("click", runCurrentSheetSignificance);
  }

  if (clearSheetTablesButton) {
    clearSheetTablesButton.addEventListener("click", clearCurrentSheetSignificance);
  }

  const checkSheetTablesButton = document.getElementById("check-sheet-tables");

  if (checkSheetTablesButton) {
    checkSheetTablesButton.addEventListener("click", runCurrentSheetCheck);
  }

  const contentFindTablesButton = document.getElementById("content-find-tables");

  if (contentFindTablesButton) {
    contentFindTablesButton.addEventListener("click", runTableInventory);
  }

  document.getElementById("check-add-report")?.addEventListener("change", updateCheckHints);

  const checkSelectedRangeButton = document.getElementById("check-selected-range");
  if (checkSelectedRangeButton) {
    checkSelectedRangeButton.addEventListener("click", runCheckSelectedRange);
  }

  initActionScopeShell();
  initPanelDismiss();
  initLanguageSelector();
});

function initPanelDismiss() {
  document.querySelectorAll(".panel-dismiss").forEach((btn) => {
    btn.addEventListener("click", () => {
      const panel = document.getElementById(btn.dataset.dismiss);
      if (panel) panel.style.display = "none";
    });
  });
}

function initLanguageSelector() {
  document.querySelectorAll(".lang-btn[data-lang]").forEach((btn) => {
    btn.addEventListener("click", () => {
      // setLanguage() calls applyI18n() internally; updateCheckHints() then
      // re-applies the mode-suffix in the new language.
      setLanguage(btn.dataset.lang);
      updateCheckHints();
      refreshActionWarnings();
    });
  });
}

// ─── Action + Scope shell (issue #167 PR1) ───────────────────────────────────

let _currentAction = "run";
let _currentScope = "current_table";
function isCheckReportEnabled() {
  const cb = document.getElementById("check-add-report");
  return cb ? cb.checked : false;
}

function updateCheckHints() {
  const modeText = isCheckReportEnabled() ? t("hint.checkWriteMode") : t("hint.checkReadOnly");
  document.querySelectorAll(".check-workspace-hint[data-hint-base]").forEach((el) => {
    el.textContent = el.dataset.hintBase + modeText;
  });
}

function collectActionWarningKeys(action) {
  if (action !== "run" && action !== "autorun") {
    return [];
  }

  const labelsOnLeftSide = getCheckboxValue("labels-on-left-side");
  const addTableFootnoteRequested = getCheckboxValue("add-table-footnote");
  const recolorRequested = getCheckboxValue("recolor-banner-and-labels");
  const warningKeys = [];

  if (labelsOnLeftSide && addTableFootnoteRequested) {
    warningKeys.push("footnote-labels-left");
  }

  if (labelsOnLeftSide && recolorRequested) {
    warningKeys.push("recolor-labels-left");
  } else if (action === "run" && recolorRequested) {
    warningKeys.push("manual-run-recolor");
  }

  return warningKeys;
}

function refreshActionWarnings() {
  const warningBlock = document.getElementById("action-settings-warnings");
  if (!warningBlock) {
    return;
  }

  const activeWarnings = new Set(collectActionWarningKeys(_currentAction));
  let hasWarnings = false;

  warningBlock.querySelectorAll("[data-action-warning]").forEach((line) => {
    const show = activeWarnings.has(line.dataset.actionWarning);
    line.style.display = show ? "" : "none";
    hasWarnings = hasWarnings || show;
  });

  warningBlock.style.display = hasWarnings ? "" : "none";
}

function updateActionScopeShell(action, scope) {
  // Update action tab active states
  document.querySelectorAll(".action-tab").forEach((tab) => {
    const active = tab.dataset.action === action;
    tab.classList.toggle("is-active", active);
    tab.setAttribute("aria-selected", active ? "true" : "false");
  });

  // Расчёт and Оглавление hide the scope selector; Автозапуск and Проверка show it.
  const runSelected = action === "run";
  const contentSelected = action === "content";
  const autorunSelected = action === "autorun";
  const scopeSelector = document.getElementById("scope-selector");
  if (scopeSelector) scopeSelector.style.display = (runSelected || contentSelected) ? "none" : "";

  document.querySelectorAll(".scope-btn").forEach((btn) => {
    btn.style.display = "";
    btn.classList.toggle("is-active", btn.dataset.scope === scope);
  });

  // Determine the effective workspace key.
  // Расчёт is always selected-range (current_table); Оглавление always workbook.
  let effectiveScope;
  if (runSelected) {
    effectiveScope = "current_table";
  } else if (contentSelected) {
    effectiveScope = "whole_workbook";
  } else {
    effectiveScope = scope;
  }
  const workspaceKey = `${action}-${effectiveScope}`;
  document.querySelectorAll(".action-workspace").forEach((ws) => {
    ws.style.display = ws.dataset.workspace === workspaceKey ? "" : "none";
  });

  // Show run-add-report for Автозапуск (sheet/workbook detected-table flows)
  const runReportControl = document.getElementById("run-report-control");
  if (runReportControl) {
    runReportControl.style.display = autorunSelected ? "" : "none";
  }

  // Show check-add-report for all Check scopes
  const checkReportControl = document.getElementById("check-report-control");
  if (checkReportControl) {
    checkReportControl.style.display = action === "check" ? "" : "none";
  }

  updateCheckHints();
  refreshActionWarnings();
}

function initActionScopeShell() {
  const tabsContainer = document.getElementById("action-tabs");
  const scopeContainer = document.getElementById("scope-selector");

  if (tabsContainer) {
    tabsContainer.addEventListener("click", (e) => {
      const tab = e.target.closest("[data-action]");
      if (!tab) return;
      const action = tab.dataset.action;
      if (action === _currentAction) return;
      _currentAction = action;
      updateActionScopeShell(_currentAction, _currentScope);
    });
  }

  if (scopeContainer) {
    scopeContainer.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-scope]");
      if (!btn || btn.disabled) return;
      const scope = btn.dataset.scope;
      if (scope === _currentScope) return;
      _currentScope = scope;
      updateActionScopeShell(_currentAction, _currentScope);
    });
  }

  updateActionScopeShell(_currentAction, _currentScope);
}

// ─────────────────────────────────────────────────────────────────────────────

/**
 * Read-only marker-overflow check for a single range, mirroring the front half
 * of runSignificanceForRangeInContext (load → interpret → detect banner) but
 * WITHOUT writing anything. Returns true when the table would need more
 * single-character markers than the alphabet provides.
 *
 * Used by the banner-aware batch preflight so overflow is judged on the exact
 * group-local required count rather than raw table width.
 */
async function computeRangeMarkerOverflowInContext(
  context,
  sheetName,
  rangeAddress,
  calculationSettings
) {
  const worksheet = context.workbook.worksheets.getItem(sheetName);
  const sourceRange = worksheet.getRange(rangeAddress);

  sourceRange.load(["values", "text", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
  await context.sync();

  const selectedValues = sourceRange.values;
  const selectedText = sourceRange.text;

  if (!selectedValues || selectedValues.length < 2 || selectedValues[0].length < 2) {
    return false;
  }

  const interpretation = await interpretSelectedRange(
    context,
    sourceRange,
    selectedValues,
    selectedText,
    calculationSettings
  );

  if (interpretation.state === "blocked") {
    return false;
  }

  const { valuesForCalculation, bannerContext } = interpretation;

  if (
    !valuesForCalculation ||
    valuesForCalculation.length < 2 ||
    !valuesForCalculation[0] ||
    valuesForCalculation[0].length < 2
  ) {
    return false;
  }

  let bannerStructure = null;
  if (calculationSettings.respectBannerStructure) {
    bannerStructure = detectBannerStructure(bannerContext, calculationSettings);
  }

  return detectSignificanceMarkerOverflow(
    valuesForCalculation[0].length,
    calculationSettings,
    bannerStructure
  );
}

/**
 * Exact, read-only batch overflow pass for banner-aware runs.
 *
 * Interprets and banner-detects each eligible table (no writes) and returns
 * true as soon as one would overflow on its group-local required count. Errors
 * during preflight are ignored here — the per-table run reports them — so a
 * single bad table never blocks the decision.
 */
async function detectBatchMarkerOverflowExact(eligible, calculationSettings) {
  let overflow = false;

  await Excel.run(async (context) => {
    for (const candidate of eligible) {
      try {
        if (
          await computeRangeMarkerOverflowInContext(
            context,
            candidate.sheetName,
            candidate.rangeAddress,
            calculationSettings
          )
        ) {
          overflow = true;
          return;
        }
      } catch (_) {
        /* preflight read failure is non-fatal; per-table run handles it */
      }
    }
  });

  return overflow;
}

/**
 * Operation-level marker-overflow preflight for batch runs (current-sheet /
 * workbook). Decides BEFORE any table is written, so Stop leaves the whole
 * operation without partial results.
 *
 * Comparison-mode aware:
 * - previous-column mode uses arrow markers, never letter labels → no dialog;
 * - non-banner modes → conservative raw column-count gate;
 * - banner-aware mode → raw width is not meaningful (Total/label columns and
 *   per-group widths), so when the cheap gate says overflow is *possible* an
 *   exact read-only pass confirms it on the group-local required count.
 *
 * Returns true when the user chose to stop (caller must abort before writing
 * any table). On Continue it sets allowMultiCharacterMarkers on the shared
 * calculationSettings so every table uses multi-character markers.
 *
 * The per-table preflight inside runSignificanceForRangeInContext stays as a
 * safety net; because the shared decider's choice is cached here, it never
 * re-prompts.
 */
async function preflightBatchMarkerOverflow(eligible, itemMap, calculationSettings, decider) {
  // Previous-column mode never uses letter labels — capacity is irrelevant.
  if (calculationSettings.compareWithPreviousColumn) {
    return false;
  }

  const columnCounts = eligible.map((candidate) => {
    const item = itemMap.get(`${candidate.sheetName}!${candidate.rangeAddress}`);
    return item ? item.columnCount : undefined;
  });

  // Cheap gate first: if no table is even wider than the alphabet, nothing can
  // overflow (the real required count is always <= raw column count).
  if (!detectBatchMarkerOverflow(columnCounts, calculationSettings)) {
    return false;
  }

  // A table is wide enough that overflow is possible. In banner-aware mode raw
  // width over-counts (Total columns, per-group widths), so confirm with an
  // exact read-only pass before prompting. Non-banner modes use the raw gate.
  if (calculationSettings.respectBannerStructure) {
    const exactOverflow = await detectBatchMarkerOverflowExact(eligible, calculationSettings);
    if (!exactOverflow) {
      return false;
    }
  }

  if ((await decider.resolve()) === "stop") {
    return true;
  }

  calculationSettings.allowMultiCharacterMarkers = true;
  return false;
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
  const _t0 = perfNow();
  setStatusMessage(runningStatusMessage("run", "table"));
  const markerOverflowDecider = createMarkerOverflowDecider();
  await Excel.run(async (context) => {
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

    // getSelectedRange() throws a RichApi.Error for non-contiguous (Ctrl+Click
    // multi-area) selections. Catch that error here and surface a user-facing
    // message rather than letting the runtime error propagate.
    let selectedRange;
    try {
      selectedRange = context.workbook.getSelectedRange();
      selectedRange.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
      await context.sync();
    } catch (_selectionErr) {
      setStatusMessage(nonContiguousSelectionMessage());
      return;
    }

    if (
      calculationSettings.compareWithPreviousColumn &&
      calculationSettings.fillOnlyTotalComparisons
    ) {
      setStatusMessage(
        // eslint-disable-next-line quotes
        'Режим "Сравнение с предыдущей колонкой" несовместим с настройкой "Заливка только для Тотала".'
      );

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

    const _tInterp = perfNow();
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

    // "address" is loaded here (Manual Run only) so the footnote can report the
    // actual processed/write-target range — which may differ from the original
    // selection after selected-range normalization.
    writeTargetRange.load(["rowIndex", "columnIndex", "rowCount", "columnCount", "address"]);

    await context.sync();

    const targetStartRowIndex = writeTargetRange.rowIndex;

    if (calculationSettings.writeBannerLetters && targetStartRowIndex === 0) {
      setStatusMessage(
        "Данные расположены в первой строке. Добавьте строку над выделенным массивом и запустите расчёт повторно."
      );

      return;
    }

    // Detect banner structure before any write so marker-overflow can be
    // resolved up front. Detection is pure (interpretedBannerContext is already
    // sanitized by the interpreter, so it is idempotent across repeated Runs).
    let bannerStructure = null;

    if (calculationSettings.respectBannerStructure) {
      const bannerContext = interpretedBannerContext;

      bannerStructure = detectBannerStructure(bannerContext, calculationSettings);

      if (bannerContext.messages && bannerContext.messages.length > 0) {
        bannerStructure.messages = [...bannerContext.messages, ...(bannerStructure.messages || [])];
      }
    }

    // Marker-overflow preflight: when there are more comparable columns than
    // single-character markers, ask the user before writing anything. Stopping
    // here leaves the selection untouched (no partial results).
    if (
      detectSignificanceMarkerOverflow(
        valuesForCalculation[0].length,
        calculationSettings,
        bannerStructure
      )
    ) {
      if ((await markerOverflowDecider.resolve()) === "stop") {
        setStatusMessage(t("markerOverflow.stopped"));
        return;
      }
      calculationSettings.allowMultiCharacterMarkers = true;
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

    const _tWrite = perfNow();
    // Pass the sheet-absolute anchor of the write target so the writer can
    // build A1 addresses for the chunked RangeAreas formatting path
    // (ExcelApi 1.9). writeTargetRange.rowIndex / .columnIndex were loaded
    // and synced above.
    const writerDetails = writeCellResultsToSelectedRange(
      writeTargetRange,
      textForCalculation,
      fullCellResultMatrix,
      detectionResult,
      calculationSettings,
      {
        captureWriterDetails: perfEnabled(),
        anchorRowIndex: writeTargetRange.rowIndex,
        anchorColumnIndex: writeTargetRange.columnIndex,
      }
    );

    const _knownDims = {
      rowIndex: writeTargetRange.rowIndex,
      columnIndex: writeTargetRange.columnIndex,
      columnCount: writeTargetRange.columnCount,
    };

    if (calculationSettings.writeBannerLetters) {
      // Combined banner clear + write: one read sync (which also flushes the
      // queued data writes above) followed by queued marker writes that are
      // flushed by the final context.sync() below.
      //
      // sourceStartColIndex widens the scan to cover any columns between the
      // selection start and the write target start, absorbing the stale-left
      // marker cleanup.
      await applyBannerMarkerUpdatesForRange(
        context,
        writeTargetRange,
        calculationSettings,
        bannerStructure,
        _knownDims,
        selectedRange.columnIndex
      );
    }

    await context.sync();

    // Design recolor (issue #306) is intentionally NOT applied in Manual Run.
    // A manual selected-range Run may cover only part of a table, so its
    // geometry need not include the full banner/label structure; recoloring it
    // could produce a misleading partial design result. Design recolor is
    // therefore limited to the detected-table autorun flows (current-table,
    // current-sheet, workbook), and a Manual Run workspace warning tells the
    // user to use those when the setting is enabled.

    // Footnote: a single processed table, so it is applied right after the
    // calculation writes complete. Insertion happens after all calculation
    // reads/writes so it cannot shift the range being calculated.
    //
    // Manual Run only (issue #308): append the actual processed/write-target
    // range so the user can see what was actually processed. This may be
    // narrower than the original selection after selected-range normalization.
    const processedScopeSuffix = buildProcessedRangeFootnoteSuffix(writeTargetRange.address);

    const footnoteJob = buildSignificanceFootnoteJob({
      sheetName: "",
      dataStartRowIndex: writeTargetRange.rowIndex,
      dataStartColIndex: writeTargetRange.columnIndex,
      dataRowCount: writeTargetRange.rowCount,
      dataColCount: writeTargetRange.columnCount,
      leftLabelValues,
      adjacentLabelColumnCount: interpretation.adjacentLabelColumnCount,
      calculationBlocks,
      calculationSettings,
      processedScopeSuffix,
    });
    if (footnoteJob) {
      try {
        await writeOrInsertSignificanceFootnoteRow(
          context,
          writeTargetRange.worksheet,
          footnoteJob
        );
      } catch (footnoteErr) {
        console.warn("RIT: не удалось записать подпись под таблицей.", footnoteErr);
      }
    }

    const statusMessages = [t("status.runDone", { count: calculationBlocks.length })];

    if (normalizationStatusLines.length > 0) {
      statusMessages.push("");
      statusMessages.push(...normalizationStatusLines);
    }

    const bannerUserMessages = formatBannerUserMessages(bannerStructure);

    if (bannerUserMessages) {
      statusMessages.push("");
      statusMessages.push(bannerUserMessages);
    }

    perfLog("runSignificanceFromSelection", {
      interpretMs: _tWrite - _tInterp,
      writeMs: perfElapsed(_tWrite),
      ...(writerDetails ? { writerDetails } : {}),
      totalMs: perfElapsed(_t0),
    });
    setStatusMessage(
      appendSelectedRangeGuardrailMessages(statusMessages, selectedRangeGuardrailWarnings).join(
        "\n"
      )
    );
  });
}


/**
 * Core significance pipeline for a single named range, using a caller-supplied
 * Office.js RequestContext.
 *
 * Extracted so sheet/workbook autorun loops can share one Excel.run context
 * across all tables, amortising per-context initialisation overhead. Callers
 * that need a self-contained execution unit use runSignificanceForRange below,
 * which wraps this function in its own Excel.run.
 *
 * Returns { status, blocksProcessed, message, rangeAddress }.
 * status: "processed" | "skipped" | "blocked" | "error"
 */
async function runSignificanceForRangeInContext(
  context,
  sheetName,
  rangeAddress,
  calculationSettings,
  markerOverflowDecider = null
) {
  const _p0 = perfNow();
  const worksheet = context.workbook.worksheets.getItem(sheetName);
  const sourceRange = worksheet.getRange(rangeAddress);

  sourceRange.load(["values", "text", "rowIndex", "columnIndex", "rowCount", "columnCount"]);

  await context.sync();
  const _pLoad = perfNow();

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

  const { writeTargetRange } = interpretation;

  if (
    !valuesForCalculation ||
    valuesForCalculation.length < 2 ||
    !valuesForCalculation[0] ||
    valuesForCalculation[0].length < 2
  ) {
    return { status: "skipped", message: "нет данных для расчёта", rangeAddress };
  }

  const _pInterp = perfNow();

  // Compute write-target row/column indices from the already-loaded sourceRange
  // properties rather than issuing a separate load+sync on writeTargetRange.
  // sourceRange.rowIndex and .columnIndex are loaded in the sync above;
  // interpretation.dataRowOffset / dataColOffset are pure-JS values from the
  // normalization path, so no additional round-trip is required.
  const targetStartRowIndex = sourceRange.rowIndex + interpretation.dataRowOffset;
  const targetStartColIndex = sourceRange.columnIndex + interpretation.dataColOffset;

  if (calculationSettings.writeBannerLetters && targetStartRowIndex === 0) {
    return {
      status: "skipped",
      message: "данные в первой строке — баннер недоступен",
      rangeAddress,
    };
  }

  // Detect banner structure before any write so marker-overflow can be resolved
  // up front (the dialog must appear before results are written).
  let bannerStructure = null;

  if (calculationSettings.respectBannerStructure) {
    const bannerContext = interpretedBannerContext;
    bannerStructure = detectBannerStructure(bannerContext, calculationSettings);
    if (bannerContext && bannerContext.messages && bannerContext.messages.length > 0) {
      bannerStructure.messages = [...bannerContext.messages, ...(bannerStructure.messages || [])];
    }
  }

  // Marker-overflow preflight. When more comparable columns exist than
  // single-character markers, ask once per operation. Stopping returns before
  // any write so no partial results are written for this table.
  if (
    detectSignificanceMarkerOverflow(
      valuesForCalculation[0].length,
      calculationSettings,
      bannerStructure
    )
  ) {
    const activeDecider = markerOverflowDecider || createMarkerOverflowDecider();
    if ((await activeDecider.resolve()) === "stop") {
      return { status: "stopped", message: "расчёт остановлен — превышен лимит маркеров", rangeAddress };
    }
    // Mutate the (per-operation) settings so every table in the batch uses
    // multi-character markers consistently after the user opts to continue.
    calculationSettings.allowMultiCharacterMarkers = true;
  }

  writeTargetRange.values = valuesForCalculation;
  writeTargetRange.format.font.bold = false;
  writeTargetRange.format.fill.clear();
  writeTargetRange.format.horizontalAlignment = "Center";
  writeTargetRange.format.verticalAlignment = "Center";

  // No intermediate sync here. The pre-write format ops and
  // writeCellResultsToSelectedRange are all writes with no intervening Excel
  // reads, so they can share the final context.sync below. Removing this
  // round-trip saves one Office.js sync per table in the autorun batch loop.
  const detectionResult = detectMetricRowsFromLeftLabels(valuesForCalculation, leftLabelValues);
  const calculationBlocks = buildCalculationBlocks(detectionResult, { preferredBase: calculationSettings.preferredBase });

  if (!calculationBlocks || calculationBlocks.length === 0) {
    return { status: "skipped", message: "нет блоков расчёта", rangeAddress };
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

  const _pCalc = perfNow();

  // Autorun never loads writeTargetRange.rowIndex / .columnIndex on the proxy
  // (they're computed from the already-loaded sourceRange + offsets to save a
  // round-trip). Pass them explicitly so the writer can use the chunked
  // RangeAreas formatting path (ExcelApi 1.9).
  const writerDetails = writeCellResultsToSelectedRange(
    writeTargetRange,
    textForCalculation,
    fullCellResultMatrix,
    detectionResult,
    calculationSettings,
    {
      captureWriterDetails: perfEnabled(),
      anchorRowIndex: targetStartRowIndex,
      anchorColumnIndex: targetStartColIndex,
    }
  );

  const _knownDims = {
    rowIndex: targetStartRowIndex,
    columnIndex: targetStartColIndex,
    columnCount: valuesForCalculation[0].length,
  };
  const _pValueWrite = perfNow();
  let _pStaleLeftClear = _pValueWrite;
  let _pBannerClear = _pValueWrite;
  let _pBannerWrite = _pValueWrite;
  let bannerDetails = null;

  if (calculationSettings.writeBannerLetters) {
    // staleLeftClear phase: no separate sync needed.  The combined banner
    // updater below receives sourceRange.columnIndex so its scan also covers
    // any columns between the source start and the write target start,
    // absorbing the former clearStaleBannerMarkersLeftOfWriteRange work.
    _pStaleLeftClear = _pValueWrite;

    // Combined clear + write in a single read/plan/write phase.  The first
    // sync inside applyBannerMarkerUpdatesForRange reads the banner area and
    // also flushes pending data-body writes; the queued marker writes are
    // flushed by the final context.sync() below.  bannerClearMs measures the
    // read+plan portion and bannerWriteMs measures the write portion so the
    // perf log shape is preserved.
    bannerDetails = await applyBannerMarkerUpdatesForRange(
      context,
      writeTargetRange,
      calculationSettings,
      bannerStructure,
      _knownDims,
      sourceRange.columnIndex
    );
    _pBannerClear = perfNow();
    // No second sync inside the combined function; the final sync flushes
    // queued writes.  bannerWriteMs is captured around that final sync below.
    _pBannerWrite = _pBannerClear;
  }

  await context.sync();
  const _pWrite = perfNow();
  // After the final sync, attribute the time it took to flush queued banner
  // writes (if any) to bannerWriteMs so the legacy log shape continues to
  // separate "load / plan" cost from "write" cost.  When banner letters are
  // disabled the final sync only flushes data writes and we leave
  // bannerWriteMs at 0.
  if (calculationSettings.writeBannerLetters) {
    _pBannerWrite = _pWrite;
  }

  // Footnote job (collected, not applied here). Applying it now would insert a
  // worksheet row that shifts every candidate range still queued in the batch.
  const footnoteJob = buildSignificanceFootnoteJob({
    sheetName,
    dataStartRowIndex: targetStartRowIndex,
    dataStartColIndex: targetStartColIndex,
    dataRowCount: valuesForCalculation.length,
    dataColCount: valuesForCalculation[0].length,
    leftLabelValues,
    adjacentLabelColumnCount: interpretation.adjacentLabelColumnCount,
    calculationBlocks,
    calculationSettings,
  });

  // Design recolor job (collected, not applied here). Like the footnote job it is
  // pure geometry computed from the already-known interpretation; the run flow
  // applies it after calculations (before footnote row insertions).
  const recolorJob = buildDesignRecolorJob({
    sheetName,
    dataStartRowIndex: targetStartRowIndex,
    dataStartColIndex: targetStartColIndex,
    dataRowCount: valuesForCalculation.length,
    dataColCount: valuesForCalculation[0].length,
    adjacentLabelColumnCount: interpretation.adjacentLabelColumnCount,
    bannerRowCount: resolveBannerRecolorRowCount(interpretation, targetStartRowIndex),
    calculationSettings,
  });

  return {
    status: "processed",
    blocksProcessed: calculationBlocks.length,
    message: `обработано блоков: ${calculationBlocks.length}`,
    rangeAddress,
    footnoteJob,
    recolorJob,
    _phasesMs: _p0 !== 0 ? {
      loadMs: _pLoad - _p0,
      interpMs: _pInterp - _pLoad,
      calcMs: _pCalc - _pInterp,
      writeMs: _pWrite - _pCalc,
      writeDetails: {
        valueWriteMs: _pValueWrite - _pCalc,
        staleLeftClearMs: _pStaleLeftClear - _pValueWrite,
        bannerClearMs: _pBannerClear - _pStaleLeftClear,
        bannerWriteMs: _pBannerWrite - _pBannerClear,
        finalSyncMs: calculationSettings.writeBannerLetters ? 0 : (_pWrite - _pBannerWrite),
        ...(writerDetails ? { writerDetails } : {}),
        ...(bannerDetails ? { bannerDetails } : {}),
      },
    } : null,
  };
}


/**
 * Runs the full significance pipeline for a single named range on a named sheet.
 *
 * Thin wrapper around runSignificanceForRangeInContext that provides its own
 * Excel.run context. Used by current-table autorun and as a per-table fallback
 * when the shared-context batch in sheet/workbook autorun encounters an
 * Office.js error that corrupts the shared context.
 *
 * Returns { status, blocksProcessed, message, rangeAddress }.
 * status: "processed" | "skipped" | "blocked" | "error"
 */
async function runSignificanceForRange(
  sheetName,
  rangeAddress,
  calculationSettings,
  markerOverflowDecider = null
) {
  const result = await Excel.run((context) =>
    runSignificanceForRangeInContext(
      context,
      sheetName,
      rangeAddress,
      calculationSettings,
      markerOverflowDecider
    )
  );
  if (result._phasesMs) {
    perfLog("runSignificanceForRange", { ...result._phasesMs, rangeAddress });
  }
  return result;
}


// ─── Run report sheet helpers ─────────────────────────────────────────────────

/**
 * Builds a Map from `sheetName!rangeAddress` -> item for quick metadata lookup.
 * Keys both resolvedRangeAddress and rangeAddress so candidates can be matched
 * regardless of which address variant filterWorkbookCandidates resolved.
 */
function buildItemMetadataMap(inventoryResults) {
  const map = new Map();
  for (const sheetResult of inventoryResults.sheetResults) {
    for (const item of sheetResult.items) {
      if (item.resolvedRangeAddress) {
        map.set(`${sheetResult.sheetName}!${item.resolvedRangeAddress}`, item);
      }
      const fallback = item.rangeAddress;
      if (fallback && !map.has(`${sheetResult.sheetName}!${fallback}`)) {
        map.set(`${sheetResult.sheetName}!${fallback}`, item);
      }
    }
  }
  return map;
}

const RUN_REPORT_COLUMNS = [
  "Лист",
  "Таблица",
  "Диапазон",
  "Статус",
  "Причина / Сообщение",
  "База",
  "Типы метрик",
  "Предупреждений",
  "Критических",
  "Детали",
  "Блоков обработано",
];

async function ensureRunReportWorksheet(context) {
  const worksheets = context.workbook.worksheets;
  const existing = worksheets.getItemOrNullObject(RUN_REPORT_SHEET_NAME);
  existing.load("isNullObject");
  await context.sync();
  return existing.isNullObject ? worksheets.add(RUN_REPORT_SHEET_NAME) : existing;
}

function writeRunReportContent(worksheet, reportRows, runLabel) {
  const colCount = RUN_REPORT_COLUMNS.length;

  // Title row
  const titleRange = worksheet.getRangeByIndexes(0, 0, 1, colCount);
  titleRange.values = normalizeRowsToColumnCount([[runLabel]], colCount);
  titleRange.merge();
  titleRange.format.font.bold = true;
  titleRange.format.font.size = 13;

  // Metadata
  const now = new Date().toLocaleString("ru-RU");
  const metaRows = normalizeRowsToColumnCount(
    [
      ["Создан", now],
      ["Строк", reportRows.length],
    ],
    2
  );
  const metaRange = worksheet.getRangeByIndexes(1, 0, metaRows.length, 2);
  metaRange.values = metaRows;
  worksheet.getRangeByIndexes(1, 0, metaRows.length, 1).format.font.bold = true;

  // Header
  const headerRowIndex = 4;
  const headerRange = worksheet.getRangeByIndexes(headerRowIndex - 1, 0, 1, colCount);
  headerRange.values = normalizeRowsToColumnCount([RUN_REPORT_COLUMNS], colCount);
  headerRange.format.font.bold = true;

  // Data rows
  const dataRows = normalizeRowsToColumnCount(
    reportRows.map((r) => [
      r.sheetName,
      r.title,
      r.rangeAddress,
      runReportStatusLabel(r.status),
      r.message,
      r.selectedBase,
      r.metricTypes,
      r.warnings,
      r.critical,
      r.warningDetails,
      r.blocksProcessed,
    ]),
    colCount
  );

  if (dataRows.length > 0) {
    const dataRange = worksheet.getRangeByIndexes(headerRowIndex, 0, dataRows.length, colCount);
    dataRange.values = dataRows;

    // Color-code status column (index 3)
    for (let i = 0; i < dataRows.length; i++) {
      const status = reportRows[i].status;
      const cell = worksheet.getRangeByIndexes(headerRowIndex + i, 3, 1, 1);
      if (status === "processed") {
        cell.format.fill.color = "#E2F0D9";
      } else if (status === "error") {
        cell.format.fill.color = "#FCE4D6";
      }
    }
  }

  // Table borders
  const tableRange = worksheet.getRangeByIndexes(
    headerRowIndex - 1,
    0,
    Math.max(dataRows.length, 1) + 1,
    colCount
  );
  tableRange.format.borders.getItem("EdgeBottom").style = "Continuous";
  tableRange.format.borders.getItem("EdgeTop").style = "Continuous";
  tableRange.format.borders.getItem("EdgeLeft").style = "Continuous";
  tableRange.format.borders.getItem("EdgeRight").style = "Continuous";
  tableRange.format.borders.getItem("InsideHorizontal").style = "Continuous";
  tableRange.format.borders.getItem("InsideVertical").style = "Continuous";

  // Column widths: Sheet, Title, Range, Status, Message, Base, Metrics, Warn, Critical, Details, Blocks
  [100, 160, 100, 85, 200, 120, 100, 75, 75, 200, 85].forEach((width, i) => {
    worksheet.getRangeByIndexes(0, i, 1, 1).format.columnWidth = width;
  });

  worksheet
    .getRangeByIndexes(0, 0, headerRowIndex + Math.max(dataRows.length, 1), colCount)
    .format.verticalAlignment = "Top";
}

/**
 * Creates or updates the Run report sheet, writes report rows, moves the sheet
 * to the last position, and returns the worksheet object.
 *
 * Placement uses a two-tier strategy:
 *
 * Tier 1 — direct position assignment.
 *   Content writes and position assignment are in separate sync batches so the
 *   heavy write batch does not interfere with the position batch. After the
 *   assignment a verification pass re-reads the actual server-side position.
 *   If the sheet is already last, we return early.
 *
 * Tier 2 — copy / delete / rename fallback.
 *   Setting worksheet.position on an existing sheet is silently ignored by
 *   some Office.js host builds (confirmed in smoke testing). When Tier 1
 *   verification shows the sheet is still not last, we fall back to:
 *     1. Load the worksheets collection to find the actual current last sheet.
 *     2. worksheet.copy("After", lastSheet) — places a copy at the true end.
 *     3. copy.activate() — make the copy active before deleting the original,
 *        to avoid the "cannot delete the active/only sheet" host error.
 *     4. worksheet.delete() — remove the original (now-not-active) sheet.
 *     5. copy.name = RUN_REPORT_SHEET_NAME — restore the canonical name.
 *   The copy operation is layout-level only; all content was already written
 *   to the original sheet before Tier 1 was attempted, and the copy carries
 *   that content forward.
 */
async function writeRunReportSheet(context, reportRows, runLabel) {
  const _t0 = perfNow();
  const worksheet = await ensureRunReportWorksheet(context);

  const existingUsed = worksheet.getUsedRangeOrNullObject();
  existingUsed.load("isNullObject");
  await context.sync();

  if (!existingUsed.isNullObject) {
    existingUsed.unmerge();
    existingUsed.clear();
  }

  writeRunReportContent(worksheet, reportRows, runLabel);

  // Tier 1 — flush content writes first, then attempt direct position assignment.
  // Keeping these in separate batches prevents the heavy write operations from
  // interfering with the position mutation on the server side.
  await context.sync();

  const worksheets = context.workbook.worksheets;
  worksheets.load("count");
  await context.sync();
  worksheet.position = worksheets.count - 1;
  await context.sync();

  // Verification: reload the sheet's actual position from the host.
  worksheet.load("position");
  worksheets.load("count");
  await context.sync();

  if (worksheet.position === worksheets.count - 1) {
    perfLog("writeRunReportSheet", { totalMs: perfElapsed(_t0), tier: 1 });
    return worksheet; // Tier 1 succeeded — sheet is already last.
  }

  // Tier 2 — copy/delete/rename fallback.
  // Direct position assignment did not take effect; use Worksheet.copy() to
  // physically place the sheet at the end of the tab bar.

  // Reload worksheet items in tab order to find the current last sheet.
  worksheets.load("items/name");
  await context.sync();

  const allSheets = worksheets.items; // Ordered by tab position (0-based).
  let lastSheet = null;
  for (let i = allSheets.length - 1; i >= 0; i--) {
    if (allSheets[i].name !== RUN_REPORT_SHEET_NAME) {
      lastSheet = allSheets[i];
      break;
    }
  }

  if (lastSheet === null) {
    // The Run report is the only sheet; it is trivially last.
    perfLog("writeRunReportSheet", { totalMs: perfElapsed(_t0), tier: "only-sheet" });
    return worksheet;
  }

  // Copy the sheet to immediately after lastSheet (the true last position).
  // The copy receives a temporary system name, e.g. "Run report (2)".
  const copy = worksheet.copy("After", lastSheet);

  // Activate the copy before deleting the original to satisfy the host
  // constraint that the active sheet cannot be deleted.
  copy.activate();
  await context.sync();

  // Remove the original (now not-active) sheet and rename the copy.
  worksheet.delete();
  await context.sync();

  copy.name = RUN_REPORT_SHEET_NAME;
  await context.sync();

  perfLog("writeRunReportSheet", { totalMs: perfElapsed(_t0), tier: 2 });
  return copy;
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
  const _t0 = perfNow();
  const calculationSettings = readCalculationSettingsFromPanel();
  // One overflow decision shared across every table processed in this run.
  const markerOverflowDecider = createMarkerOverflowDecider();

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

  setStatusMessage(runningStatusMessage("run", "workbook"));
  const _progressScope = "workbook";

  // Collect inventory to identify eligible candidates.
  const _tScan = perfNow();
  let inventoryResults;
  try {
    inventoryResults = await collectWorkbookInventoryResults(calculationSettings);
    // Normalize without inserting backlinks — only needed for resolvedRangeAddress.
    normalizeBacklinkItems(inventoryResults.sheetResults, false);
  } catch (err) {
    setStatusMessage(`Автозапуск: ошибка при сканировании книги — ${err.message || err}`);
    return;
  }

  const addReport = getCheckboxValue("run-add-report");
  const itemMap = buildItemMetadataMap(inventoryResults);

  // Partition candidates: eligible to process vs. pre-skipped due to status/range.
  // Generated sheets (Content, Run report) are excluded entirely (not counted toward skipped).
  const { eligible, skipped: preSkipped } = filterWorkbookCandidates(inventoryResults, {
    generatedSheetNames: GENERATED_SHEET_NAMES,
  });
  let skipped = preSkipped.length;
  const detailLines = preSkipped.map(formatSkippedCandidateDetail);

  // Seed report rows with pre-execution skips.
  const reportRows = preSkipped.map((candidate) => {
    const item = itemMap.get(`${candidate.sheetName}!${candidate.rangeAddress}`);
    return {
      sheetName: candidate.sheetName,
      title: item ? (item.resolvedTitle || (isGeneratedBacklinkRow(item.title) ? "" : item.title) || "") : "",
      rangeAddress: candidate.rangeAddress || "",
      status: "skipped",
      message: runReportSkipReasonLabel(candidate.reason),
      selectedBase: item ? (item.selectedBaseSubtypeLabel || "") : "",
      metricTypes: runReportMetricTypes(item),
      warnings: item ? (item.warningsCount ?? "") : "",
      critical: item ? (item.criticalCount ?? "") : "",
      warningDetails: runReportWarningDetails(item),
      blocksProcessed: "",
    };
  });

  if (eligible.length === 0) {
    const noEligibleLines = [
      "Автозапуск: доступных кандидатов не найдено.",
      'Проверьте статусы таблиц через «Проверить книгу» / «Оглавление → С полной проверкой».',
    ];
    if (skipped > 0) {
      noEligibleLines.push("", `Пропущено: ${skipped}.`, ...detailLines);
    }
    setStatusMessage(noEligibleLines.join("\n"));

    if (addReport && reportRows.length > 0) {
      try {
        await Excel.run(async (context) => {
          const sheet = await writeRunReportSheet(context, reportRows, "Автозапуск — диагностический отчёт");
          sheet.activate();
          await context.sync();
        });
      } catch (reportErr) {
        setStatusMessage(
          (document.getElementById("significance-result")?.textContent || "") +
          `\n[Отчёт: ошибка записи — ${reportErr.message || reportErr}]`
        );
      }
    }
    return;
  }

  // Operation-level marker-overflow preflight: decide once, before any table is
  // written. Stop aborts the whole run with no partial results.
  if (await preflightBatchMarkerOverflow(eligible, itemMap, calculationSettings, markerOverflowDecider)) {
    setStatusMessage(t("markerOverflow.stopped"));
    return;
  }

  let processed = 0;
  let errors = 0;
  // Set when the user chooses Stop at the marker-overflow dialog: halts further
  // table processing for the rest of this operation.
  let markerOverflowStopped = false;
  // Footnote jobs collected during the run and applied bottom-to-top per sheet
  // AFTER all calculations finish (insertions must not shift pending ranges).
  const footnoteJobs = [];
  // Design recolor jobs (issue #306) collected during the run and applied after
  // all calculations finish, before footnote insertions (recolor never shifts
  // rows, so applying it first keeps both jobs' captured geometry valid).
  const recolorJobs = [];

  // Process all eligible tables in a single Excel.run to amortise per-context
  // overhead. If any table causes an Office.js sync error (corrupting the shared
  // context), we record that table as an error, exit the shared context, and
  // fall back to per-table Excel.run for any remaining candidates.
  const _tLoop = perfNow();
  const _perfPhases = {
    loadMs: 0,
    interpMs: 0,
    calcMs: 0,
    writeMs: 0,
    writeDetails: {
      valueWriteMs: 0,
      staleLeftClearMs: 0,
      bannerClearMs: 0,
      bannerWriteMs: 0,
      finalSyncMs: 0,
      writerDetails: createAggregatedWriterPerfDetails(),
      bannerDetails: createAggregatedBannerPerfDetails(),
    },
  };
  let _batchEndedAt = 0;
  try {
    await Excel.run(async (context) => {
      for (let _bi = 0; _bi < eligible.length; _bi++) {
        const candidate = eligible[_bi];
        const item = itemMap.get(`${candidate.sheetName}!${candidate.rangeAddress}`);
        try {
          const result = await runSignificanceForRangeInContext(
            context,
            candidate.sheetName,
            candidate.rangeAddress,
            calculationSettings,
            markerOverflowDecider
          );
          if (result.status === "stopped") {
            // User chose to stop at the marker-overflow dialog. Stop processing
            // further tables; tables already processed keep their results.
            skipped++;
            detailLines.push(`- ${candidate.sheetName} ${candidate.rangeAddress}: ${result.message}`);
            markerOverflowStopped = true;
            break;
          }
          if (result.status === "processed") {
            processed++;
            if (result.footnoteJob) footnoteJobs.push(result.footnoteJob);
            if (result.recolorJob) recolorJobs.push(result.recolorJob);
          } else if (result.status === "skipped" || result.status === "blocked") {
            skipped++;
            detailLines.push(`- ${candidate.sheetName} ${candidate.rangeAddress}: пропущено — ${result.message}`);
          } else {
            errors++;
            detailLines.push(`- ${candidate.sheetName} ${candidate.rangeAddress}: ошибка — ${result.message}`);
          }
          if (result._phasesMs) {
            _perfPhases.loadMs += result._phasesMs.loadMs;
            _perfPhases.interpMs += result._phasesMs.interpMs;
            _perfPhases.calcMs += result._phasesMs.calcMs;
            _perfPhases.writeMs += result._phasesMs.writeMs;
            if (result._phasesMs.writeDetails) {
              _perfPhases.writeDetails.valueWriteMs += result._phasesMs.writeDetails.valueWriteMs;
              _perfPhases.writeDetails.staleLeftClearMs += result._phasesMs.writeDetails.staleLeftClearMs;
              _perfPhases.writeDetails.bannerClearMs += result._phasesMs.writeDetails.bannerClearMs;
              _perfPhases.writeDetails.bannerWriteMs += result._phasesMs.writeDetails.bannerWriteMs;
              _perfPhases.writeDetails.finalSyncMs += result._phasesMs.writeDetails.finalSyncMs;
              mergeWriterPerfDetails(
                _perfPhases.writeDetails.writerDetails,
                result._phasesMs.writeDetails.writerDetails
              );
              mergeBannerPerfDetails(
                _perfPhases.writeDetails.bannerDetails,
                result._phasesMs.writeDetails.bannerDetails,
                candidate.sheetName
              );
            }
          }
          reportRows.push({
            sheetName: candidate.sheetName,
            title: candidate.title,
            rangeAddress: candidate.rangeAddress,
            status: result.status,
            message: result.message || "",
            selectedBase: item ? (item.selectedBaseSubtypeLabel || "") : "",
            metricTypes: runReportMetricTypes(item),
            warnings: item ? (item.warningsCount ?? "") : "",
            critical: item ? (item.criticalCount ?? "") : "",
            warningDetails: runReportWarningDetails(item),
            blocksProcessed: result.blocksProcessed != null ? result.blocksProcessed : "",
          });
          // Live table-level progress (after each processed table).
          setStatusMessage(
            buildBatchProgressStatus({
              action: "run",
              scope: _progressScope,
              currentIndex: _bi + 1,
              total: eligible.length,
              sheetName: candidate.sheetName,
            })
          );
          await yieldToUi();
        } catch (err) {
          errors++;
          const errMsg = err.message || "неизвестная ошибка";
          detailLines.push(`- ${candidate.sheetName} ${candidate.rangeAddress}: ошибка — ${errMsg}`);
          reportRows.push({
            sheetName: candidate.sheetName,
            title: candidate.title,
            rangeAddress: candidate.rangeAddress,
            status: "error",
            message: errMsg,
            selectedBase: item ? (item.selectedBaseSubtypeLabel || "") : "",
            metricTypes: runReportMetricTypes(item),
            warnings: item ? (item.warningsCount ?? "") : "",
            critical: item ? (item.criticalCount ?? "") : "",
            warningDetails: runReportWarningDetails(item),
            blocksProcessed: "",
          });
          _batchEndedAt = _bi + 1;
          throw err; // shared context may be corrupted; exit batch
        }
      }
      _batchEndedAt = eligible.length;
    });
  } catch (_batchErr) {
    // shared context aborted; fall back to per-table for any remaining candidates
  }
  for (let _fi = _batchEndedAt; _fi < eligible.length; _fi++) {
    const candidate = eligible[_fi];
    const item = itemMap.get(`${candidate.sheetName}!${candidate.rangeAddress}`);
    try {
      const result = await runSignificanceForRange(
        candidate.sheetName,
        candidate.rangeAddress,
        calculationSettings,
        markerOverflowDecider
      );
      if (result.status === "stopped") {
        skipped++;
        detailLines.push(`- ${candidate.sheetName} ${candidate.rangeAddress}: ${result.message}`);
        markerOverflowStopped = true;
        break;
      }
      if (result.status === "processed") {
        processed++;
        if (result.footnoteJob) footnoteJobs.push(result.footnoteJob);
        if (result.recolorJob) recolorJobs.push(result.recolorJob);
      } else if (result.status === "skipped" || result.status === "blocked") {
        skipped++;
        detailLines.push(`- ${candidate.sheetName} ${candidate.rangeAddress}: пропущено — ${result.message}`);
      } else {
        errors++;
        detailLines.push(`- ${candidate.sheetName} ${candidate.rangeAddress}: ошибка — ${result.message}`);
      }
      if (result._phasesMs) {
        _perfPhases.loadMs += result._phasesMs.loadMs;
        _perfPhases.interpMs += result._phasesMs.interpMs;
        _perfPhases.calcMs += result._phasesMs.calcMs;
        _perfPhases.writeMs += result._phasesMs.writeMs;
        if (result._phasesMs.writeDetails) {
          _perfPhases.writeDetails.valueWriteMs += result._phasesMs.writeDetails.valueWriteMs;
          _perfPhases.writeDetails.staleLeftClearMs += result._phasesMs.writeDetails.staleLeftClearMs;
          _perfPhases.writeDetails.bannerClearMs += result._phasesMs.writeDetails.bannerClearMs;
          _perfPhases.writeDetails.bannerWriteMs += result._phasesMs.writeDetails.bannerWriteMs;
          _perfPhases.writeDetails.finalSyncMs += result._phasesMs.writeDetails.finalSyncMs;
          mergeWriterPerfDetails(
            _perfPhases.writeDetails.writerDetails,
            result._phasesMs.writeDetails.writerDetails
          );
          mergeBannerPerfDetails(
            _perfPhases.writeDetails.bannerDetails,
            result._phasesMs.writeDetails.bannerDetails,
            candidate.sheetName
          );
        }
      }
      reportRows.push({
        sheetName: candidate.sheetName,
        title: candidate.title,
        rangeAddress: candidate.rangeAddress,
        status: result.status,
        message: result.message || "",
        selectedBase: item ? (item.selectedBaseSubtypeLabel || "") : "",
        metricTypes: runReportMetricTypes(item),
        warnings: item ? (item.warningsCount ?? "") : "",
        critical: item ? (item.criticalCount ?? "") : "",
        warningDetails: runReportWarningDetails(item),
        blocksProcessed: result.blocksProcessed != null ? result.blocksProcessed : "",
      });
      // Live table-level progress (after each processed table).
      setStatusMessage(
        buildBatchProgressStatus({
          action: "run",
          scope: _progressScope,
          currentIndex: _fi + 1,
          total: eligible.length,
          sheetName: candidate.sheetName,
        })
      );
      await yieldToUi();
    } catch (err) {
      errors++;
      const errMsg = err.message || "неизвестная ошибка";
      detailLines.push(`- ${candidate.sheetName} ${candidate.rangeAddress}: ошибка — ${errMsg}`);
      reportRows.push({
        sheetName: candidate.sheetName,
        title: candidate.title,
        rangeAddress: candidate.rangeAddress,
        status: "error",
        message: errMsg,
        selectedBase: item ? (item.selectedBaseSubtypeLabel || "") : "",
        metricTypes: runReportMetricTypes(item),
        warnings: item ? (item.warningsCount ?? "") : "",
        critical: item ? (item.criticalCount ?? "") : "",
        warningDetails: runReportWarningDetails(item),
        blocksProcessed: "",
      });
    }
  }

  // Apply design recolor first (no row shifts) so its rectangles match the
  // unshifted geometry, then footnotes (which insert rows, bottom-to-top).
  await applyDesignRecolorJobsInOwnContext(recolorJobs);

  // Apply collected footnotes after ALL calculations (bottom-to-top per sheet).
  await applySignificanceFootnoteJobsInOwnContext(footnoteJobs);

  const _tLoopDone = perfNow();
  const summaryLines = [
    "Автозапуск завершён.",
    `Обработано таблиц: ${processed}.`,
    `Пропущено: ${skipped}.`,
    `Ошибок: ${errors}.`,
  ];

  if (markerOverflowStopped) {
    summaryLines.push("", t("markerOverflow.stopped"));
  }

  if (detailLines.length > 0) {
    summaryLines.push("", ...detailLines);
  }

  perfLog("runAutoSignificance", {
    scanMs: _tLoop - _tScan,
    loopMs: _tLoopDone - _tLoop,
    tablesProcessed: processed,
    ...(processed > 0 ? { perTablePhaseMs: _perfPhases } : {}),
    totalMs: perfElapsed(_t0),
  });
  setStatusMessage(summaryLines.join("\n"));

  if (addReport) {
    try {
      await Excel.run(async (context) => {
        const sheet = await writeRunReportSheet(context, reportRows, "Автозапуск — диагностический отчёт");
        sheet.activate();
        await context.sync();
      });
    } catch (reportErr) {
      setStatusMessage(
        summaryLines.join("\n") +
        `\n[Отчёт: ошибка записи — ${reportErr.message || reportErr}]`
      );
    }
  }
}

/**
 * Autorun current-table: resolves the detected table under the active cell
 * and runs the full significance pipeline on it.
 *
 * Uses the same pre-resolver selection guard as runCheckTable (#210) to block
 * broad multi-table selections before the active-cell resolver is called.
 * Delegates the significance pipeline to runSignificanceForRange.
 */
async function runAutoCurrentTableSignificance() {
  const _t0 = perfNow();
  const calculationSettings = readCalculationSettingsFromPanel();

  if (calculationSettings.compareWithPreviousColumn && calculationSettings.compareOnlyWithTotal) {
    setStatusMessage(
      // eslint-disable-next-line quotes
      'Режим "Сравнение с предыдущей колонкой" несовместим с режимом "Сравнивать только с Тотал".'
    );
    return;
  }

  if (
    calculationSettings.excludeTotalFromComparisons &&
    !calculationSettings.firstColumnIsTotal &&
    !calculationSettings.respectBannerStructure
  ) {
    setStatusMessage(
      'Для режима "Не сравнивать с Тотал" нужно указать расположение Тотала. Сейчас поддерживается вариант "Первая колонка — Тотал" или режим "Учитывать структуру баннера".'
    );
    return;
  }

  if (
    calculationSettings.compareOnlyWithTotal &&
    !calculationSettings.firstColumnIsTotal &&
    !calculationSettings.respectBannerStructure
  ) {
    setStatusMessage(
      'Для режима "Сравнивать только с Тотал" нужно указать расположение Тотала. Сейчас поддерживается вариант "Первая колонка — Тотал" или режим "Учитывать структуру баннера".'
    );
    return;
  }

  if (
    calculationSettings.compareWithPreviousColumn &&
    calculationSettings.fillOnlyTotalComparisons
  ) {
    setStatusMessage(
      'Режим "Сравнение с предыдущей колонкой" несовместим с настройкой "Заливка только для Тотала".'
    );
    return;
  }

  setStatusMessage(runningStatusMessage("run", "table"));

  let resolverResult = null;

  try {
    await Excel.run(async (context) => {
      // Pre-resolver selection guard (mirrors runCheckTable from #210).
      // A broad selection spanning multiple table-like blocks (separated by
      // empty rows) is blocked before the active-cell resolver runs.
      const selectionForGuard = context.workbook.getSelectedRange();
      selectionForGuard.load(["values"]);
      await context.sync();

      if (selectionHasMultiTableGap(selectionForGuard.values)) {
        resolverResult = {
          status: "blocked",
          sheetName: "",
          message: t("status.multiTableGapAutorun"),
        };
        return;
      }

      resolverResult = await resolveCurrentTableFromActiveCell(context, calculationSettings);
    });
  } catch (err) {
    setStatusMessage(t("status.autorunTableDetectError", { msg: err.message || err }));
    return;
  }

  if (!resolverResult) {
    setStatusMessage(t("status.autorunTableNotResolved"));
    return;
  }

  const addReport = getCheckboxValue("run-add-report");

  if (resolverResult.status !== "ok") {
    const msg = resolverResult.message || "Не удалось определить таблицу под активной ячейкой.";
    setStatusMessage(msg);

    if (addReport) {
      try {
        await Excel.run(async (context) => {
          const sheet = await writeRunReportSheet(context, [{
            sheetName: resolverResult.sheetName || "",
            title: "",
            rangeAddress: "",
            status: "skipped",
            message: msg,
            selectedBase: "",
            metricTypes: "",
            warnings: 0,
            critical: 0,
            warningDetails: "",
            blocksProcessed: 0,
          }], "Автозапуск — Текущая таблица");
          sheet.activate();
          await context.sync();
        });
      } catch (reportErr) {
        setStatusMessage(msg + `\n[Отчёт: ошибка записи — ${reportErr.message || reportErr}]`);
      }
    }
    return;
  }

  const { sheetName, rangeAddress } = resolverResult;
  const resolvedTableTitle = resolveContentDisplayTitle(resolverResult.candidateMeta, 1);

  const _tRun = perfNow();
  let result;
  try {
    result = await runSignificanceForRange(sheetName, rangeAddress, calculationSettings);
  } catch (err) {
    const errMsg = err.message || "неизвестная ошибка";
    setStatusMessage(t("status.autorunCalcError", { msg: errMsg }));

    if (addReport) {
      try {
        await Excel.run(async (context) => {
          const sheet = await writeRunReportSheet(context, [{
            sheetName,
            title: resolvedTableTitle,
            rangeAddress,
            status: "error",
            message: errMsg,
            selectedBase: resolverResult.candidateMeta ? (resolverResult.candidateMeta.selectedBaseSubtypeLabel || "") : "",
            metricTypes: runReportMetricTypes(resolverResult.candidateMeta || null),
            warnings: 0,
            critical: 0,
            warningDetails: "",
            blocksProcessed: 0,
          }], "Автозапуск — Текущая таблица");
          sheet.activate();
          await context.sync();
        });
      } catch (_) { /* non-fatal — primary error already shown */ }
    }
    return;
  }

  const statusMsg =
    result.status === "processed"
      ? t("status.autorunProcessed", { sheet: sheetName, range: rangeAddress, count: result.blocksProcessed })
      : result.status === "stopped"
        ? t("markerOverflow.stopped")
        : t("status.autorunSkipped", { msg: result.message || t("status.resolverFallback") });

  perfLog("runAutoCurrentTableSignificance", {
    resolveMs: _tRun - _t0,
    runMs: perfElapsed(_tRun),
    totalMs: perfElapsed(_t0),
  });
  setStatusMessage(statusMsg);

  // Design recolor applied before the footnote (recolor never shifts rows).
  if (result.status === "processed" && result.recolorJob) {
    await applyDesignRecolorJobsInOwnContext([result.recolorJob]);
  }

  // Footnote applied after calculation (separate context). Single table, so
  // bottom-to-top ordering is trivially satisfied.
  if (result.status === "processed" && result.footnoteJob) {
    await applySignificanceFootnoteJobsInOwnContext([result.footnoteJob]);
  }

  if (addReport) {
    try {
      await Excel.run(async (context) => {
        const sheet = await writeRunReportSheet(context, [{
          sheetName,
          title: resolvedTableTitle,
          rangeAddress,
          status: result.status,
          message: result.message || "",
          selectedBase: resolverResult.candidateMeta ? (resolverResult.candidateMeta.selectedBaseSubtypeLabel || "") : "",
          metricTypes: runReportMetricTypes(resolverResult.candidateMeta || null),
          warnings: 0,
          critical: 0,
          warningDetails: "",
          blocksProcessed: result.blocksProcessed != null ? result.blocksProcessed : "",
        }], "Автозапуск — Текущая таблица");
        sheet.activate();
        await context.sync();
      });
    } catch (reportErr) {
      setStatusMessage(statusMsg + "\n" + t("status.autorunReportWriteError", { msg: reportErr.message || reportErr }));
    }
  }
}

/**
 * Autorun current-table Clear: resolves the table under the active cell and
 * clears significance markers only from that resolved range.
 *
 * Applies the same pre-resolver selection guard as runAutoCurrentTableSignificance
 * to block broad multi-table selections before the resolver runs. Does not touch
 * the arbitrary selected range — the clear target is always the resolver result.
 */
async function clearAutoCurrentTableSignificance() {
  setStatusMessage(runningStatusMessage("clear", "table"));
  let resolverResult = null;

  try {
    await Excel.run(async (context) => {
      const selectionForGuard = context.workbook.getSelectedRange();
      selectionForGuard.load(["values"]);
      await context.sync();

      if (selectionHasMultiTableGap(selectionForGuard.values)) {
        resolverResult = {
          status: "blocked",
          sheetName: "",
          message: t("status.multiTableGapAutorun"),
        };
        return;
      }

      resolverResult = await resolveCurrentTableFromActiveCell(context, readCalculationSettingsFromPanel());
    });
  } catch (err) {
    setStatusMessage(t("status.autorunTableDetectError", { msg: err.message || err }));
    return;
  }

  if (!resolverResult) {
    setStatusMessage(t("status.autorunTableNotResolved"));
    return;
  }

  if (resolverResult.status !== "ok") {
    setStatusMessage(resolverResult.message || t("status.resolverFallback"));
    return;
  }

  const { sheetName, rangeAddress } = resolverResult;

  let result;
  try {
    result = await clearSignificanceForRange(sheetName, rangeAddress, getClearPipelineOptions());
  } catch (err) {
    setStatusMessage(t("status.autorunClearError", { msg: err.message || err }));
    return;
  }

  if (result.status === "cleared") {
    if (result.footnoteRemovalJob) {
      await applySignificanceFootnoteRemovalJobsInOwnContext([result.footnoteRemovalJob]);
    }
    setStatusMessage(t("status.autorunCleared", { sheet: sheetName, range: rangeAddress }));
  } else {
    setStatusMessage(t("status.autorunClearSkipped", { msg: result.message || t("status.noDataInRange") }));
  }
}


/**
 * Auto-clear: removes significance markers from all "available" inventory
 * candidates in the workbook.
 *
 * Attempts a workbook-level staged batch first (1 Excel.run, 4 sync phases
 * for the whole workbook).  If that fails, falls back to the per-sheet staged
 * batch (1 Excel.run per sheet, 4 sync phases per sheet).  If a per-sheet
 * batch also fails, remaining tables on that sheet fall back to per-table
 * clearSignificanceForRange.
 *
 * batchMode in RIT_PERF diagnostics reflects which path actually ran:
 *   "workbook" — workbook-level batch succeeded (contextRuns: 1, syncPhases: 4)
 *   "sheet"    — workbook batch failed; per-sheet batches ran
 *   "fallback" — all batches failed; per-table fallback ran
 */
async function clearAutoSignificance() {
  const _t0 = perfNow();
  setStatusMessage(runningStatusMessage("clear", "workbook"));
  let inventoryResults;
  try {
    inventoryResults = await collectWorkbookInventoryResults(readCalculationSettingsFromPanel());
    normalizeBacklinkItems(inventoryResults.sheetResults, false);
  } catch (err) {
    setStatusMessage(`Автоочистка: ошибка при сканировании книги — ${err.message || err}`);
    return;
  }

  const { eligible, skipped: preSkipped } = filterWorkbookCandidates(inventoryResults, {
    generatedSheetNames: GENERATED_SHEET_NAMES,
  });
  let skipped = preSkipped.length;
  const detailLines = preSkipped.map(formatSkippedCandidateDetail);

  if (eligible.length === 0) {
    const noEligibleLines = [
      "Автоочистка: доступных кандидатов не найдено.",
      'Проверьте статусы таблиц через «Проверить книгу» / «Оглавление → С полной проверкой».',
    ];
    if (skipped > 0) {
      noEligibleLines.push("", `Пропущено: ${skipped}.`, ...detailLines);
    }
    setStatusMessage(noEligibleLines.join("\n"));
    return;
  }

  let cleared = 0;
  let errors = 0;
  let contextRuns = 0;
  let sheetsCleared = 0;
  let batchMode = "workbook";
  const _clearAggregate = {
    syncPhases: 0,
    sourceRangesLoaded: 0,
    targetRangesLoaded: 0,
    bannerRangesLoaded: 0,
    bodyCellsRead: 0,
    bodyCellsChanged: 0,
    bannerCellsRead: 0,
    bannerCellsChanged: 0,
    bannerWriteCommands: 0,
  };

  // Always build the by-sheet map so the per-sheet fallback loop has it ready.
  const bySheet = groupCandidatesBySheet(eligible);

  // Generated footnote rows to silently delete after all clearing completes.
  const footnoteRemovalJobs = [];

  // ── ATTEMPT 1: workbook-level staged batch (1 context, 4 syncs total) ──────
  let _wbResults = null;
  let _wbBatchDiags = null;

  try {
    await Excel.run(async (context) => {
      const r = await clearSignificanceForWorkbookBatched(context, eligible);
      _wbResults = r.results;
      _wbBatchDiags = r.batchDiags;
      if (Array.isArray(r.footnoteRemovalJobs)) footnoteRemovalJobs.push(...r.footnoteRemovalJobs);
    });
    contextRuns = 1;
  } catch (_wbErr) {
    // Workbook-level batch failed; fall through to per-sheet staged batching.
    batchMode = "sheet";
  }

  if (_wbResults !== null) {
    // Workbook batch succeeded — process results.
    _clearAggregate.syncPhases += _wbBatchDiags.syncPhases;
    _clearAggregate.sourceRangesLoaded += _wbBatchDiags.sourceRangesLoaded;
    _clearAggregate.targetRangesLoaded += _wbBatchDiags.targetRangesLoaded;
    _clearAggregate.bannerRangesLoaded += _wbBatchDiags.bannerRangesLoaded;

    const sheetClearedSet = new Set();
    for (let _i = 0; _i < _wbResults.length; _i++) {
      const result = _wbResults[_i];
      const candidate = eligible[_i];
      if (result.status === "cleared") {
        cleared++;
        sheetClearedSet.add(candidate.sheetName);
        accumulateClearDetails(_clearAggregate, result._clearDetails);
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
    }
    sheetsCleared = sheetClearedSet.size;
  } else {
    // ── FALLBACK: per-sheet staged batching ────────────────────────────────
    for (const [sheetName, sheetCandidates] of bySheet) {
      let batchEndedAt = 0;
      let sheetClearedCount = 0;
      contextRuns++;

      // Staged 4-phase batch: all N tables on this sheet in 4 syncs total.
      try {
        await Excel.run(async (context) => {
          const worksheetRef = context.workbook.worksheets.getItem(sheetName);
          const { results: batchResults, sheetDiags, footnoteRemovalRows } = await clearSignificanceForSheetBatched(
            context,
            worksheetRef,
            sheetCandidates
          );
          _clearAggregate.syncPhases += sheetDiags.syncPhases;
          _clearAggregate.sourceRangesLoaded += sheetDiags.sourceRangesLoaded;
          _clearAggregate.targetRangesLoaded += sheetDiags.targetRangesLoaded;
          _clearAggregate.bannerRangesLoaded += sheetDiags.bannerRangesLoaded;
          if (Array.isArray(footnoteRemovalRows)) {
            for (const rowIndex of footnoteRemovalRows) {
              footnoteRemovalJobs.push({ sheetName, footnoteRowIndex: rowIndex });
            }
          }
          for (let _si = 0; _si < batchResults.length; _si++) {
            const result = batchResults[_si];
            const candidate = sheetCandidates[_si];
            if (result.status === "cleared") {
              cleared++;
              sheetClearedCount++;
              accumulateClearDetails(_clearAggregate, result._clearDetails);
            } else if (result.status === "skipped") {
              skipped++;
              detailLines.push(
                `- ${sheetName} ${candidate.rangeAddress}: пропущено — ${result.message}`
              );
            } else {
              errors++;
              detailLines.push(
                `- ${sheetName} ${candidate.rangeAddress}: ошибка — ${result.message}`
              );
            }
          }
          batchEndedAt = sheetCandidates.length;
        });
      } catch (_batchErr) {
        // Per-sheet batch failed; fall back to per-table for all candidates on this sheet.
        if (batchMode === "sheet") batchMode = "fallback";
      }

      for (let _fi = batchEndedAt; _fi < sheetCandidates.length; _fi++) {
        const candidate = sheetCandidates[_fi];
        contextRuns++;
        _clearAggregate.syncPhases += 4; // per-table fallback: 4 syncs each
        try {
          const result = await clearSignificanceForRange(
            sheetName,
            candidate.rangeAddress,
            getClearPipelineOptions()
          );
          if (result.status === "cleared") {
            cleared++;
            sheetClearedCount++;
            accumulateClearDetails(_clearAggregate, result._clearDetails);
            if (result.footnoteRemovalJob) footnoteRemovalJobs.push(result.footnoteRemovalJob);
          } else if (result.status === "skipped") {
            skipped++;
            detailLines.push(
              `- ${sheetName} ${candidate.rangeAddress}: пропущено — ${result.message}`
            );
          } else {
            errors++;
            detailLines.push(
              `- ${sheetName} ${candidate.rangeAddress}: ошибка — ${result.message}`
            );
          }
        } catch (err) {
          errors++;
          detailLines.push(
            `- ${sheetName} ${candidate.rangeAddress}: ошибка — ${err.message || "неизвестная ошибка"}`
          );
        }
      }

      if (sheetClearedCount > 0) sheetsCleared++;
    }
  }

  // Silently delete generated footnote rows for cleared tables (bottom-to-top).
  await applySignificanceFootnoteRemovalJobsInOwnContext(footnoteRemovalJobs);

  const summaryLines = [
    "Автоочистка завершена.",
    `Очищено таблиц: ${cleared}.`,
    `Пропущено: ${skipped}.`,
    `Ошибок: ${errors}.`,
  ];

  if (detailLines.length > 0) {
    summaryLines.push("", ...detailLines);
  }

  perfLog("clearAutoSignificance", {
    batchMode,
    tablesCleared: cleared,
    sheetsCleared,
    contextRuns,
    sheetCount: bySheet.size,
    syncPhases: _clearAggregate.syncPhases,
    sourceRangesLoaded: _clearAggregate.sourceRangesLoaded,
    targetRangesLoaded: _clearAggregate.targetRangesLoaded,
    bannerRangesLoaded: _clearAggregate.bannerRangesLoaded,
    bodyCellsRead: _clearAggregate.bodyCellsRead,
    bodyCellsChanged: _clearAggregate.bodyCellsChanged,
    bannerCellsRead: _clearAggregate.bannerCellsRead,
    bannerCellsChanged: _clearAggregate.bannerCellsChanged,
    bannerWriteCommands: _clearAggregate.bannerWriteCommands,
    totalMs: perfElapsed(_t0),
  });
  setStatusMessage(summaryLines.join("\n"));
}

/**
 * Current-sheet run: processes all "available" inventory candidates on the
 * active worksheet only.
 *
 * Mirrors runAutoSignificance() but scans only the active sheet via
 * collectActiveSheetInventoryResults(). Content sheet is silently ignored.
 * Settings validation is identical to runAutoSignificance().
 */
async function runCurrentSheetSignificance() {
  const _t0 = perfNow();
  const calculationSettings = readCalculationSettingsFromPanel();
  // One overflow decision shared across every table processed in this run.
  const markerOverflowDecider = createMarkerOverflowDecider();

  if (calculationSettings.compareWithPreviousColumn && calculationSettings.compareOnlyWithTotal) {
    setStatusMessage(
      // eslint-disable-next-line quotes
      'Режим "Сравнение с предыдущей колонкой" несовместим с режимом "Сравнивать только с Тотал".'
    );
    return;
  }

  if (
    calculationSettings.excludeTotalFromComparisons &&
    !calculationSettings.firstColumnIsTotal &&
    !calculationSettings.respectBannerStructure
  ) {
    setStatusMessage(
      'Для режима "Не сравнивать с Тотал" нужно указать расположение Тотала. Сейчас поддерживается вариант "Первая колонка — Тотал" или режим "Учитывать структуру баннера".'
    );
    return;
  }

  if (
    calculationSettings.compareOnlyWithTotal &&
    !calculationSettings.firstColumnIsTotal &&
    !calculationSettings.respectBannerStructure
  ) {
    setStatusMessage(
      'Для режима "Сравнивать только с Тотал" нужно указать расположение Тотала. Сейчас поддерживается вариант "Первая колонка — Тотал" или режим "Учитывать структуру баннера".'
    );
    return;
  }

  if (
    calculationSettings.compareWithPreviousColumn &&
    calculationSettings.fillOnlyTotalComparisons
  ) {
    setStatusMessage(
      'Режим "Сравнение с предыдущей колонкой" несовместим с настройкой "Заливка только для Тотала".'
    );
    return;
  }

  setStatusMessage(runningStatusMessage("run", "sheet"));
  const _progressScope = "sheet";

  const _tScan = perfNow();
  let inventoryResults;
  try {
    inventoryResults = await collectActiveSheetInventoryResults(calculationSettings);
    normalizeBacklinkItems(inventoryResults.sheetResults, false);
  } catch (err) {
    setStatusMessage(`Лист — запуск: ошибка при сканировании листа — ${err.message || err}`);
    return;
  }

  const addReport = getCheckboxValue("run-add-report");
  const itemMap = buildItemMetadataMap(inventoryResults);

  const { eligible, skipped: preSkipped } = filterWorkbookCandidates(inventoryResults, {
    generatedSheetNames: GENERATED_SHEET_NAMES,
  });
  let skipped = preSkipped.length;
  const detailLines = preSkipped.map(formatSkippedCandidateDetail);

  const reportRows = preSkipped.map((candidate) => {
    const item = itemMap.get(`${candidate.sheetName}!${candidate.rangeAddress}`);
    return {
      sheetName: candidate.sheetName,
      title: item ? (item.resolvedTitle || (isGeneratedBacklinkRow(item.title) ? "" : item.title) || "") : "",
      rangeAddress: candidate.rangeAddress || "",
      status: "skipped",
      message: runReportSkipReasonLabel(candidate.reason),
      selectedBase: item ? (item.selectedBaseSubtypeLabel || "") : "",
      metricTypes: runReportMetricTypes(item),
      warnings: item ? (item.warningsCount ?? "") : "",
      critical: item ? (item.criticalCount ?? "") : "",
      warningDetails: runReportWarningDetails(item),
      blocksProcessed: "",
    };
  });

  if (eligible.length === 0) {
    const noEligibleLines = [
      "Лист — запуск: доступных кандидатов не найдено.",
      'Проверьте статусы таблиц через «Проверить книгу» / «Оглавление → С полной проверкой».',
    ];
    if (skipped > 0) {
      noEligibleLines.push("", `Пропущено: ${skipped}.`, ...detailLines);
    }
    setStatusMessage(noEligibleLines.join("\n"));

    if (addReport && reportRows.length > 0) {
      try {
        await Excel.run(async (context) => {
          const sheet = await writeRunReportSheet(context, reportRows, "Лист: запуск — диагностический отчёт");
          sheet.activate();
          await context.sync();
        });
      } catch (reportErr) {
        setStatusMessage(
          (document.getElementById("significance-result")?.textContent || "") +
          `\n[Отчёт: ошибка записи — ${reportErr.message || reportErr}]`
        );
      }
    }
    return;
  }

  // Operation-level marker-overflow preflight: decide once, before any table is
  // written. Stop aborts the whole run with no partial results.
  if (await preflightBatchMarkerOverflow(eligible, itemMap, calculationSettings, markerOverflowDecider)) {
    setStatusMessage(t("markerOverflow.stopped"));
    return;
  }

  let processed = 0;
  let errors = 0;
  // Set when the user chooses Stop at the marker-overflow dialog: halts further
  // table processing for the rest of this operation.
  let markerOverflowStopped = false;
  // Footnote jobs collected during the run and applied bottom-to-top per sheet
  // AFTER all calculations finish (insertions must not shift pending ranges).
  const footnoteJobs = [];
  // Design recolor jobs (issue #306) collected during the run and applied after
  // all calculations finish, before footnote insertions (recolor never shifts
  // rows, so applying it first keeps both jobs' captured geometry valid).
  const recolorJobs = [];

  // Process all eligible tables in a single Excel.run to amortise per-context
  // overhead. Same fallback strategy as runAutoSignificance: on any Office.js
  // sync error we record that table as an error, exit the shared context, and
  // continue with per-table Excel.run for any remaining candidates.
  const _tLoop = perfNow();
  const _perfPhases = {
    loadMs: 0,
    interpMs: 0,
    calcMs: 0,
    writeMs: 0,
    writeDetails: {
      valueWriteMs: 0,
      staleLeftClearMs: 0,
      bannerClearMs: 0,
      bannerWriteMs: 0,
      finalSyncMs: 0,
      writerDetails: createAggregatedWriterPerfDetails(),
      bannerDetails: createAggregatedBannerPerfDetails(),
    },
  };
  let _batchEndedAt = 0;
  try {
    await Excel.run(async (context) => {
      for (let _bi = 0; _bi < eligible.length; _bi++) {
        const candidate = eligible[_bi];
        const item = itemMap.get(`${candidate.sheetName}!${candidate.rangeAddress}`);
        try {
          const result = await runSignificanceForRangeInContext(
            context,
            candidate.sheetName,
            candidate.rangeAddress,
            calculationSettings,
            markerOverflowDecider
          );
          if (result.status === "stopped") {
            // User chose to stop at the marker-overflow dialog. Stop processing
            // further tables; tables already processed keep their results.
            skipped++;
            detailLines.push(`- ${candidate.sheetName} ${candidate.rangeAddress}: ${result.message}`);
            markerOverflowStopped = true;
            break;
          }
          if (result.status === "processed") {
            processed++;
            if (result.footnoteJob) footnoteJobs.push(result.footnoteJob);
            if (result.recolorJob) recolorJobs.push(result.recolorJob);
          } else if (result.status === "skipped" || result.status === "blocked") {
            skipped++;
            detailLines.push(`- ${candidate.sheetName} ${candidate.rangeAddress}: пропущено — ${result.message}`);
          } else {
            errors++;
            detailLines.push(`- ${candidate.sheetName} ${candidate.rangeAddress}: ошибка — ${result.message}`);
          }
          if (result._phasesMs) {
            _perfPhases.loadMs += result._phasesMs.loadMs;
            _perfPhases.interpMs += result._phasesMs.interpMs;
            _perfPhases.calcMs += result._phasesMs.calcMs;
            _perfPhases.writeMs += result._phasesMs.writeMs;
            if (result._phasesMs.writeDetails) {
              _perfPhases.writeDetails.valueWriteMs += result._phasesMs.writeDetails.valueWriteMs;
              _perfPhases.writeDetails.staleLeftClearMs += result._phasesMs.writeDetails.staleLeftClearMs;
              _perfPhases.writeDetails.bannerClearMs += result._phasesMs.writeDetails.bannerClearMs;
              _perfPhases.writeDetails.bannerWriteMs += result._phasesMs.writeDetails.bannerWriteMs;
              _perfPhases.writeDetails.finalSyncMs += result._phasesMs.writeDetails.finalSyncMs;
              mergeWriterPerfDetails(
                _perfPhases.writeDetails.writerDetails,
                result._phasesMs.writeDetails.writerDetails
              );
              mergeBannerPerfDetails(
                _perfPhases.writeDetails.bannerDetails,
                result._phasesMs.writeDetails.bannerDetails,
                candidate.sheetName
              );
            }
          }
          reportRows.push({
            sheetName: candidate.sheetName,
            title: candidate.title,
            rangeAddress: candidate.rangeAddress,
            status: result.status,
            message: result.message || "",
            selectedBase: item ? (item.selectedBaseSubtypeLabel || "") : "",
            metricTypes: runReportMetricTypes(item),
            warnings: item ? (item.warningsCount ?? "") : "",
            critical: item ? (item.criticalCount ?? "") : "",
            warningDetails: runReportWarningDetails(item),
            blocksProcessed: result.blocksProcessed != null ? result.blocksProcessed : "",
          });
          // Live table-level progress (after each processed table).
          setStatusMessage(
            buildBatchProgressStatus({
              action: "run",
              scope: _progressScope,
              currentIndex: _bi + 1,
              total: eligible.length,
              sheetName: candidate.sheetName,
            })
          );
          await yieldToUi();
        } catch (err) {
          errors++;
          const errMsg = err.message || "неизвестная ошибка";
          detailLines.push(`- ${candidate.sheetName} ${candidate.rangeAddress}: ошибка — ${errMsg}`);
          reportRows.push({
            sheetName: candidate.sheetName,
            title: candidate.title,
            rangeAddress: candidate.rangeAddress,
            status: "error",
            message: errMsg,
            selectedBase: item ? (item.selectedBaseSubtypeLabel || "") : "",
            metricTypes: runReportMetricTypes(item),
            warnings: item ? (item.warningsCount ?? "") : "",
            critical: item ? (item.criticalCount ?? "") : "",
            warningDetails: runReportWarningDetails(item),
            blocksProcessed: "",
          });
          _batchEndedAt = _bi + 1;
          throw err; // shared context may be corrupted; exit batch
        }
      }
      _batchEndedAt = eligible.length;
    });
  } catch (_batchErr) {
    // shared context aborted; fall back to per-table for any remaining candidates
  }
  for (let _fi = _batchEndedAt; _fi < eligible.length; _fi++) {
    const candidate = eligible[_fi];
    const item = itemMap.get(`${candidate.sheetName}!${candidate.rangeAddress}`);
    try {
      const result = await runSignificanceForRange(
        candidate.sheetName,
        candidate.rangeAddress,
        calculationSettings,
        markerOverflowDecider
      );
      if (result.status === "stopped") {
        skipped++;
        detailLines.push(`- ${candidate.sheetName} ${candidate.rangeAddress}: ${result.message}`);
        markerOverflowStopped = true;
        break;
      }
      if (result.status === "processed") {
        processed++;
        if (result.footnoteJob) footnoteJobs.push(result.footnoteJob);
        if (result.recolorJob) recolorJobs.push(result.recolorJob);
      } else if (result.status === "skipped" || result.status === "blocked") {
        skipped++;
        detailLines.push(`- ${candidate.sheetName} ${candidate.rangeAddress}: пропущено — ${result.message}`);
      } else {
        errors++;
        detailLines.push(`- ${candidate.sheetName} ${candidate.rangeAddress}: ошибка — ${result.message}`);
      }
      if (result._phasesMs) {
        _perfPhases.loadMs += result._phasesMs.loadMs;
        _perfPhases.interpMs += result._phasesMs.interpMs;
        _perfPhases.calcMs += result._phasesMs.calcMs;
        _perfPhases.writeMs += result._phasesMs.writeMs;
        if (result._phasesMs.writeDetails) {
          _perfPhases.writeDetails.valueWriteMs += result._phasesMs.writeDetails.valueWriteMs;
          _perfPhases.writeDetails.staleLeftClearMs += result._phasesMs.writeDetails.staleLeftClearMs;
          _perfPhases.writeDetails.bannerClearMs += result._phasesMs.writeDetails.bannerClearMs;
          _perfPhases.writeDetails.bannerWriteMs += result._phasesMs.writeDetails.bannerWriteMs;
          _perfPhases.writeDetails.finalSyncMs += result._phasesMs.writeDetails.finalSyncMs;
          mergeWriterPerfDetails(
            _perfPhases.writeDetails.writerDetails,
            result._phasesMs.writeDetails.writerDetails
          );
          mergeBannerPerfDetails(
            _perfPhases.writeDetails.bannerDetails,
            result._phasesMs.writeDetails.bannerDetails,
            candidate.sheetName
          );
        }
      }
      reportRows.push({
        sheetName: candidate.sheetName,
        title: candidate.title,
        rangeAddress: candidate.rangeAddress,
        status: result.status,
        message: result.message || "",
        selectedBase: item ? (item.selectedBaseSubtypeLabel || "") : "",
        metricTypes: runReportMetricTypes(item),
        warnings: item ? (item.warningsCount ?? "") : "",
        critical: item ? (item.criticalCount ?? "") : "",
        warningDetails: runReportWarningDetails(item),
        blocksProcessed: result.blocksProcessed != null ? result.blocksProcessed : "",
      });
      // Live table-level progress (after each processed table).
      setStatusMessage(
        buildBatchProgressStatus({
          action: "run",
          scope: _progressScope,
          currentIndex: _fi + 1,
          total: eligible.length,
          sheetName: candidate.sheetName,
        })
      );
      await yieldToUi();
    } catch (err) {
      errors++;
      const errMsg = err.message || "неизвестная ошибка";
      detailLines.push(`- ${candidate.sheetName} ${candidate.rangeAddress}: ошибка — ${errMsg}`);
      reportRows.push({
        sheetName: candidate.sheetName,
        title: candidate.title,
        rangeAddress: candidate.rangeAddress,
        status: "error",
        message: errMsg,
        selectedBase: item ? (item.selectedBaseSubtypeLabel || "") : "",
        metricTypes: runReportMetricTypes(item),
        warnings: item ? (item.warningsCount ?? "") : "",
        critical: item ? (item.criticalCount ?? "") : "",
        warningDetails: runReportWarningDetails(item),
        blocksProcessed: "",
      });
    }
  }

  // Apply design recolor first (no row shifts) so its rectangles match the
  // unshifted geometry, then footnotes (which insert rows, bottom-to-top).
  await applyDesignRecolorJobsInOwnContext(recolorJobs);

  // Apply collected footnotes after ALL calculations (bottom-to-top per sheet).
  await applySignificanceFootnoteJobsInOwnContext(footnoteJobs);

  const _tLoopDone = perfNow();
  const summaryLines = [
    "Лист — запуск завершён.",
    `Обработано таблиц: ${processed}.`,
    `Пропущено: ${skipped}.`,
    `Ошибок: ${errors}.`,
  ];

  if (markerOverflowStopped) {
    summaryLines.push("", t("markerOverflow.stopped"));
  }

  if (detailLines.length > 0) {
    summaryLines.push("", ...detailLines);
  }

  perfLog("runCurrentSheetSignificance", {
    scanMs: _tLoop - _tScan,
    loopMs: _tLoopDone - _tLoop,
    tablesProcessed: processed,
    ...(processed > 0 ? { perTablePhaseMs: _perfPhases } : {}),
    totalMs: perfElapsed(_t0),
  });
  setStatusMessage(summaryLines.join("\n"));

  if (addReport) {
    try {
      await Excel.run(async (context) => {
        const sheet = await writeRunReportSheet(context, reportRows, "Лист: запуск — диагностический отчёт");
        sheet.activate();
        await context.sync();
      });
    } catch (reportErr) {
      setStatusMessage(
        summaryLines.join("\n") +
        `\n[Отчёт: ошибка записи — ${reportErr.message || reportErr}]`
      );
    }
  }
}

/**
 * Current-sheet clear: removes significance markers from all "available"
 * inventory candidates on the active worksheet only.
 *
 * Mirrors clearAutoSignificance() — all tables on the same sheet share one
 * Excel.run context, with per-table fallback on context corruption.
 * Content sheet is silently ignored.
 */
async function clearCurrentSheetSignificance() {
  const _t0 = perfNow();
  setStatusMessage(runningStatusMessage("clear", "sheet"));
  let inventoryResults;
  try {
    inventoryResults = await collectActiveSheetInventoryResults(readCalculationSettingsFromPanel());
    normalizeBacklinkItems(inventoryResults.sheetResults, false);
  } catch (err) {
    setStatusMessage(`Лист — очистка: ошибка при сканировании листа — ${err.message || err}`);
    return;
  }

  const { eligible, skipped: preSkipped } = filterWorkbookCandidates(inventoryResults, {
    generatedSheetNames: GENERATED_SHEET_NAMES,
  });
  let skipped = preSkipped.length;
  const detailLines = preSkipped.map(formatSkippedCandidateDetail);

  if (eligible.length === 0) {
    const noEligibleLines = [
      "Лист — очистка: доступных кандидатов не найдено.",
      'Проверьте статусы таблиц через «Проверить книгу» / «Оглавление → С полной проверкой».',
    ];
    if (skipped > 0) {
      noEligibleLines.push("", `Пропущено: ${skipped}.`, ...detailLines);
    }
    setStatusMessage(noEligibleLines.join("\n"));
    return;
  }

  let cleared = 0;
  let errors = 0;
  let contextRuns = 0;
  let sheetsCleared = 0;
  const _clearAggregate = {
    syncPhases: 0,
    sourceRangesLoaded: 0,
    targetRangesLoaded: 0,
    bannerRangesLoaded: 0,
    bodyCellsRead: 0,
    bodyCellsChanged: 0,
    bannerCellsRead: 0,
    bannerCellsChanged: 0,
    bannerWriteCommands: 0,
  };

  // Group by sheet; each sheet is cleared in one staged 4-sync batch.
  const bySheet = groupCandidatesBySheet(eligible);

  // Generated footnote rows to silently delete after all clearing completes.
  const footnoteRemovalJobs = [];

  for (const [sheetName, sheetCandidates] of bySheet) {
    let batchEndedAt = 0;
    let sheetClearedCount = 0;
    contextRuns++;

    // Staged 4-phase batch: all N tables on this sheet in 4 syncs total.
    try {
      await Excel.run(async (context) => {
        const worksheetRef = context.workbook.worksheets.getItem(sheetName);
        const { results: batchResults, sheetDiags, footnoteRemovalRows } = await clearSignificanceForSheetBatched(
          context,
          worksheetRef,
          sheetCandidates
        );
        _clearAggregate.syncPhases += sheetDiags.syncPhases;
        _clearAggregate.sourceRangesLoaded += sheetDiags.sourceRangesLoaded;
        _clearAggregate.targetRangesLoaded += sheetDiags.targetRangesLoaded;
        _clearAggregate.bannerRangesLoaded += sheetDiags.bannerRangesLoaded;
        if (Array.isArray(footnoteRemovalRows)) {
          for (const rowIndex of footnoteRemovalRows) {
            footnoteRemovalJobs.push({ sheetName, footnoteRowIndex: rowIndex });
          }
        }
        for (let _si = 0; _si < batchResults.length; _si++) {
          const result = batchResults[_si];
          const candidate = sheetCandidates[_si];
          if (result.status === "cleared") {
            cleared++;
            sheetClearedCount++;
            accumulateClearDetails(_clearAggregate, result._clearDetails);
          } else if (result.status === "skipped") {
            skipped++;
            detailLines.push(
              `- ${sheetName} ${candidate.rangeAddress}: пропущено — ${result.message}`
            );
          } else {
            errors++;
            detailLines.push(
              `- ${sheetName} ${candidate.rangeAddress}: ошибка — ${result.message}`
            );
          }
        }
        batchEndedAt = sheetCandidates.length;
      });
    } catch (_batchErr) {
      // Staged batch failed; fall back to per-table for all candidates on this sheet.
    }

    for (let _fi = batchEndedAt; _fi < sheetCandidates.length; _fi++) {
      const candidate = sheetCandidates[_fi];
      contextRuns++;
      _clearAggregate.syncPhases += 4; // per-table fallback: 4 syncs each
      try {
        const result = await clearSignificanceForRange(
          sheetName,
          candidate.rangeAddress,
          getClearPipelineOptions()
        );
        if (result.status === "cleared") {
          cleared++;
          sheetClearedCount++;
          accumulateClearDetails(_clearAggregate, result._clearDetails);
          if (result.footnoteRemovalJob) footnoteRemovalJobs.push(result.footnoteRemovalJob);
        } else if (result.status === "skipped") {
          skipped++;
          detailLines.push(
            `- ${sheetName} ${candidate.rangeAddress}: пропущено — ${result.message}`
          );
        } else {
          errors++;
          detailLines.push(
            `- ${sheetName} ${candidate.rangeAddress}: ошибка — ${result.message}`
          );
        }
      } catch (err) {
        errors++;
        detailLines.push(
          `- ${sheetName} ${candidate.rangeAddress}: ошибка — ${err.message || "неизвестная ошибка"}`
        );
      }
    }

    if (sheetClearedCount > 0) sheetsCleared++;
  }

  // Silently delete generated footnote rows for cleared tables (bottom-to-top).
  await applySignificanceFootnoteRemovalJobsInOwnContext(footnoteRemovalJobs);

  const summaryLines = [
    "Лист — очистка завершена.",
    `Очищено таблиц: ${cleared}.`,
    `Пропущено: ${skipped}.`,
    `Ошибок: ${errors}.`,
  ];

  if (detailLines.length > 0) {
    summaryLines.push("", ...detailLines);
  }

  perfLog("clearCurrentSheetSignificance", {
    tablesCleared: cleared,
    sheetsCleared,
    contextRuns,
    sheetCount: bySheet.size,
    syncPhases: _clearAggregate.syncPhases,
    sourceRangesLoaded: _clearAggregate.sourceRangesLoaded,
    targetRangesLoaded: _clearAggregate.targetRangesLoaded,
    bannerRangesLoaded: _clearAggregate.bannerRangesLoaded,
    bodyCellsRead: _clearAggregate.bodyCellsRead,
    bodyCellsChanged: _clearAggregate.bodyCellsChanged,
    bannerCellsRead: _clearAggregate.bannerCellsRead,
    bannerCellsChanged: _clearAggregate.bannerCellsChanged,
    bannerWriteCommands: _clearAggregate.bannerWriteCommands,
    totalMs: perfElapsed(_t0),
  });
  setStatusMessage(summaryLines.join("\n"));
}

/**
 * Read-only check for a single detected candidate range.
 *
 * Loads the named range from the given sheet, runs the full interpretation +
 * preview-model pipeline, and returns a structured result. Does NOT write
 * anything back to Excel.
 *
 * @param {string} sheetName
 * @param {string} rangeAddress
 * @param {object} calculationSettings
 * @returns {Promise<{ status: "checked"|"blocked"|"skipped", model?: object, message?: string }>}
 */
async function checkTableForRange(sheetName, rangeAddress, calculationSettings) {
  return await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const sourceRange = worksheet.getRange(rangeAddress);

    sourceRange.load(["values", "text", "rowIndex", "columnIndex", "rowCount", "columnCount"]);

    await context.sync();

    const selectedValues = sourceRange.values;
    const selectedText = sourceRange.text;

    if (
      !selectedValues ||
      selectedValues.length < 1 ||
      !selectedValues[0] ||
      selectedValues[0].length < 1
    ) {
      return { status: "skipped", message: "нет данных в диапазоне" };
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
      return { status: "blocked", message: `${interpretation.blockingMessage}${codes}` };
    }

    const { valuesForCalculation, leftLabelValues, bannerContext, normalized } = interpretation;

    const modelInput = {
      values: valuesForCalculation,
      leftLabelValues,
      bannerContext,
      settings: calculationSettings,
      trailingBodyRows: normalized?.trailingBodyRows,
    };

    const model = buildTablePreviewModel(modelInput);
    return { status: "checked", model };
  });
}

/**
 * Current-sheet check: reports all detected inventory candidates on the
 * active worksheet only.
 *
 * Mirrors runCheckTable() for the single-range case but iterates over all
 * inventory candidates on the active sheet. For available candidates it runs
 * the full read-only check pipeline (interpretation + preview model). For
 * uncertain, rejected, and missing-range candidates it records the status
 * without attempting execution. Nothing is written to Excel.
 *
 * Content sheet is silently ignored (consistent with current-sheet Run/Clear).
 * Output goes to the existing check panel so wording clearly states it covers
 * the active sheet only.
 */
async function runCurrentSheetCheck() {
  const calculationSettings = readCalculationSettingsFromPanel();
  setCheckMessage(runningStatusMessage("check", "sheet"));

  let inventoryResults;
  let activeSheetName = "";

  try {
    inventoryResults = await collectActiveSheetInventoryResults(calculationSettings);
    normalizeBacklinkItems(inventoryResults.sheetResults, false);
  } catch (err) {
    setCheckMessage(
      `Лист — проверка: ошибка при сканировании листа — ${err.message || err}`
    );
    return;
  }

  // Determine the active sheet name for display.
  if (inventoryResults.sheetResults.length > 0) {
    activeSheetName = inventoryResults.sheetResults[0].sheetName;
  } else if (inventoryResults.skippedSheets && inventoryResults.skippedSheets.length > 0) {
    activeSheetName = inventoryResults.skippedSheets[0].sheetName;
  }

  // Handle empty / skipped active sheet (no candidates found).
  if (inventoryResults.sheetResults.length === 0) {
    const skipped = inventoryResults.skippedSheets || [];
    if (skipped.length > 0 && skipped[0].reason === "empty") {
      setCheckMessage(
        `Лист — проверка (только активный лист): лист «${activeSheetName || "?"}» пустой.`
      );
    } else if (skipped.length > 0) {
      setCheckMessage(
        `Лист — проверка (только активный лист): ${formatInventorySkippedSheetLine(skipped[0])}`
      );
    } else if (inventoryResults.scannedSheets === 0 && activeSheetName === "") {
      // Active sheet is the Content sheet — silently ignored.
      setCheckMessage(
        "Лист — проверка (только активный лист): активный лист является сгенерированным листом и пропущен."
      );
    } else {
      setCheckMessage(
        `Лист — проверка (только активный лист): на листе «${activeSheetName || "?"}» кандидатов не найдено.`
      );
    }
    return;
  }

  const sheetResult = inventoryResults.sheetResults[0];
  activeSheetName = sheetResult.sheetName;
  const allItems = sheetResult.items;

  let availableCount = 0;
  let uncertainCount = 0;
  let rejectedCount = 0;
  let missingCount = 0;

  const candidateLines = [];
  const checkReportRows = [];

  for (let i = 0; i < allItems.length; i++) {
    const item = allItems[i];
    const rangeAddr = item.resolvedRangeAddress || item.rangeAddress || null;
    const reportTitle = resolveContentDisplayTitle(item, i + 1);
    const header = `${i + 1}. ${reportTitle} — ${rangeAddr || "?"}`;

    candidateLines.push("");
    candidateLines.push(header);

    if (!rangeAddr) {
      missingCount++;
      candidateLines.push("   Пропущено — нет диапазона.");
      checkReportRows.push({ sheetName: activeSheetName, title: reportTitle, rangeAddress: "", status: "skipped", message: "Нет диапазона", selectedBase: item.selectedBaseSubtypeLabel || "", metricTypes: runReportMetricTypes(item), warnings: item.warningsCount ?? "", critical: item.criticalCount ?? "", warningDetails: runReportWarningDetails(item), blocksProcessed: "" });
      continue;
    }

    if (item.candidateStatus === "rejected") {
      rejectedCount++;
      candidateLines.push("   Отклонён — не опознан как таблица ResearchSignal.");
      if (item.previewSummary) candidateLines.push(`   ${item.previewSummary}.`);
      if (item.candidateNotes && item.candidateNotes.length > 0) {
        candidateLines.push(`   [${item.candidateNotes.join("; ")}]`);
      }
      checkReportRows.push({ sheetName: activeSheetName, title: reportTitle, rangeAddress: rangeAddr, status: "skipped", message: "Не опознан как таблица", selectedBase: item.selectedBaseSubtypeLabel || "", metricTypes: runReportMetricTypes(item), warnings: item.warningsCount ?? "", critical: item.criticalCount ?? "", warningDetails: runReportWarningDetails(item), blocksProcessed: "" });
      continue;
    }

    if (item.candidateStatus === "uncertain") {
      uncertainCount++;
      candidateLines.push("   Неопределён — граница данных неоднозначна.");
      if (item.previewSummary) candidateLines.push(`   ${item.previewSummary}.`);
      if (item.candidateNotes && item.candidateNotes.length > 0) {
        candidateLines.push(`   [${item.candidateNotes.join("; ")}]`);
      }
      checkReportRows.push({ sheetName: activeSheetName, title: reportTitle, rangeAddress: rangeAddr, status: "skipped", message: "Граница данных неоднозначна", selectedBase: item.selectedBaseSubtypeLabel || "", metricTypes: runReportMetricTypes(item), warnings: item.warningsCount ?? "", critical: item.criticalCount ?? "", warningDetails: runReportWarningDetails(item), blocksProcessed: "" });
      continue;
    }

    if (item.candidateStatus === "available") {
      availableCount++;

      try {
        const checkResult = await checkTableForRange(
          activeSheetName,
          rangeAddr,
          calculationSettings
        );

        if (checkResult.status === "checked") {
          const { summary, qualitySummary, userVisibleIssues, calculationBlocks } = checkResult.model;
          candidateLines.push("   Доступен.");
          candidateLines.push(
            `   ${item.rowCount ?? summary.rowCount} строк, ${item.columnCount ?? ""} колонок.` +
              ` Блоков: ${summary.detectedBlocks}. Баз: ${summary.baseRows}.`
          );
          const warnParts = [];
          if (qualitySummary.criticalCount > 0)
            warnParts.push(`Критических: ${qualitySummary.criticalCount}`);
          if (qualitySummary.warningCount > 0)
            warnParts.push(`Предупреждений: ${qualitySummary.warningCount}`);
          if (warnParts.length > 0) candidateLines.push(`   ${warnParts.join(". ")}.`);
          const issueDetails = (userVisibleIssues || []).map((iss) => `[${iss.severity}] ${iss.message}`).join("; ");
          checkReportRows.push({ sheetName: activeSheetName, title: reportTitle, rangeAddress: rangeAddr, status: "checked", message: `Строк: ${summary.rowCount}. Блоков: ${summary.detectedBlocks}. Баз: ${summary.baseRows}.`, selectedBase: item.selectedBaseSubtypeLabel || "", metricTypes: checkMetricTypesFromBlocks(calculationBlocks), warnings: qualitySummary.warningCount, critical: qualitySummary.criticalCount, warningDetails: issueDetails, blocksProcessed: summary.detectedBlocks });
        } else if (checkResult.status === "blocked") {
          candidateLines.push(`   Доступен — проверка заблокирована: ${checkResult.message}`);
          checkReportRows.push({ sheetName: activeSheetName, title: reportTitle, rangeAddress: rangeAddr, status: "blocked", message: checkResult.message || "", selectedBase: item.selectedBaseSubtypeLabel || "", metricTypes: runReportMetricTypes(item), warnings: item.warningsCount ?? "", critical: item.criticalCount ?? "", warningDetails: runReportWarningDetails(item), blocksProcessed: 0 });
        } else {
          candidateLines.push(
            `   Доступен — проверка пропущена: ${checkResult.message || "неизвестная причина"}`
          );
          checkReportRows.push({ sheetName: activeSheetName, title: reportTitle, rangeAddress: rangeAddr, status: "skipped", message: checkResult.message || "Пропущено", selectedBase: item.selectedBaseSubtypeLabel || "", metricTypes: runReportMetricTypes(item), warnings: item.warningsCount ?? "", critical: item.criticalCount ?? "", warningDetails: runReportWarningDetails(item), blocksProcessed: "" });
        }
      } catch (err) {
        candidateLines.push(
          `   Доступен — ошибка при проверке: ${err.message || "неизвестная ошибка"}`
        );
        checkReportRows.push({ sheetName: activeSheetName, title: reportTitle, rangeAddress: rangeAddr, status: "error", message: err.message || "неизвестная ошибка", selectedBase: item.selectedBaseSubtypeLabel || "", metricTypes: runReportMetricTypes(item), warnings: item.warningsCount ?? "", critical: item.criticalCount ?? "", warningDetails: runReportWarningDetails(item), blocksProcessed: "" });
      }

      if (item.previewSummary) candidateLines.push(`   ${item.previewSummary}.`);
      if (item.candidateNotes && item.candidateNotes.length > 0) {
        candidateLines.push(`   [${item.candidateNotes.join("; ")}]`);
      }
    } else {
      // Unknown / future status — report without attempting check.
      missingCount++;
      candidateLines.push(`   Пропущено — статус «${item.candidateStatus || "unknown"}».`);
      checkReportRows.push({ sheetName: activeSheetName, title: reportTitle, rangeAddress: rangeAddr, status: "skipped", message: `Статус «${item.candidateStatus || "unknown"}»`, selectedBase: item.selectedBaseSubtypeLabel || "", metricTypes: runReportMetricTypes(item), warnings: item.warningsCount ?? "", critical: item.criticalCount ?? "", warningDetails: runReportWarningDetails(item), blocksProcessed: "" });
    }
  }

  const summaryLines = [
    `Лист — проверка (только активный лист): «${activeSheetName}».`,
    `Кандидатов: ${allItems.length}. Доступно: ${availableCount}. Неопределённых: ${uncertainCount}. Отклонено: ${rejectedCount}. Пропущено: ${missingCount}.`,
  ];

  summaryLines.push(...candidateLines);
  summaryLines.push("");
  summaryLines.push("Данные не изменены.");

  setCheckMessage(summaryLines.join("\n"));

  if (isCheckReportEnabled() && checkReportRows.length > 0) {
    try {
      await Excel.run(async (context) => {
        await writeRunReportSheet(context, checkReportRows, "Проверка — Текущий лист");
      });
    } catch (reportErr) {
      const el = document.getElementById("check-result");
      if (el) el.textContent += `\n\n[Run report: ошибка записи — ${reportErr.message || reportErr}]`;
    }
  }
}

/**
 * Workbook-wide check: scans all sheets and reports found table candidates.
 *
 * Read-only — does NOT write the Content sheet or any other Excel output.
 * Reports the workbook scan summary to the check panel so the result is
 * clearly a check/report rather than Content generation.
 *
 * Content / Оглавление remains the dedicated action for creating/updating
 * the Content sheet.
 */
async function runWorkbookCheck() {
  const calculationSettings = readCalculationSettingsFromPanel();
  setCheckMessage(runningStatusMessage("check", "workbook"));

  let inventoryResults;
  try {
    inventoryResults = await collectWorkbookInventoryResults(calculationSettings);
    normalizeBacklinkItems(inventoryResults.sheetResults, false);
  } catch (err) {
    setCheckMessage(`Книга — проверка: ошибка при сканировании — ${err.message || err}`);
    return;
  }

  const { scannedSheets, sheetResults, skippedSheets } = inventoryResults;
  const totalCandidates = sheetResults.reduce((sum, s) => sum + s.items.length, 0);

  const summaryLines = [
    `Книга — проверка: просканировано листов: ${scannedSheets}.`,
    `Кандидатов найдено: ${totalCandidates}.`,
  ];

  const checkReportRows = [];
  let globalItemIndex = 0;

  for (const sheetResult of sheetResults) {
    summaryLines.push("");
    summaryLines.push(`Лист: ${sheetResult.sheetName}`);

    for (let i = 0; i < sheetResult.items.length; i++) {
      globalItemIndex++;
      const item = sheetResult.items[i];
      const rangeAddr = item.resolvedRangeAddress || item.rangeAddress || null;
      const reportTitle = resolveContentDisplayTitle(item, globalItemIndex);
      const header = `  ${i + 1}. ${reportTitle} — ${rangeAddr || "?"}`;
      summaryLines.push(header);

      if (item.candidateStatus === "available") {
        const warnParts = [];
        if (item.criticalCount > 0) warnParts.push(`Критических: ${item.criticalCount}`);
        if (item.warningsCount > 0) warnParts.push(`Предупреждений: ${item.warningsCount}`);
        const warnStr = warnParts.length > 0 ? ` ${warnParts.join(". ")}.` : "";
        summaryLines.push(`     Доступен.${warnStr}`);
        checkReportRows.push({ sheetName: sheetResult.sheetName, title: reportTitle, rangeAddress: rangeAddr || "", status: "checked", message: "Кандидат найден (сканирование книги).", selectedBase: item.selectedBaseSubtypeLabel || "", metricTypes: runReportMetricTypes(item), warnings: item.warningsCount ?? "", critical: item.criticalCount ?? "", warningDetails: runReportWarningDetails(item), blocksProcessed: "" });
      } else if (item.candidateStatus === "uncertain") {
        summaryLines.push("     Неопределён — граница данных неоднозначна.");
        checkReportRows.push({ sheetName: sheetResult.sheetName, title: reportTitle, rangeAddress: rangeAddr || "", status: "skipped", message: "Граница данных неоднозначна", selectedBase: item.selectedBaseSubtypeLabel || "", metricTypes: runReportMetricTypes(item), warnings: item.warningsCount ?? "", critical: item.criticalCount ?? "", warningDetails: runReportWarningDetails(item), blocksProcessed: "" });
      } else if (item.candidateStatus === "rejected") {
        summaryLines.push("     Отклонён — не опознан как таблица ResearchSignal.");
        checkReportRows.push({ sheetName: sheetResult.sheetName, title: reportTitle, rangeAddress: rangeAddr || "", status: "skipped", message: "Не опознан как таблица", selectedBase: item.selectedBaseSubtypeLabel || "", metricTypes: runReportMetricTypes(item), warnings: item.warningsCount ?? "", critical: item.criticalCount ?? "", warningDetails: runReportWarningDetails(item), blocksProcessed: "" });
      } else {
        summaryLines.push(`     Пропущено — статус «${item.candidateStatus || "unknown"}».`);
        checkReportRows.push({ sheetName: sheetResult.sheetName, title: reportTitle, rangeAddress: rangeAddr || "", status: "skipped", message: `Статус «${item.candidateStatus || "unknown"}»`, selectedBase: item.selectedBaseSubtypeLabel || "", metricTypes: runReportMetricTypes(item), warnings: item.warningsCount ?? "", critical: item.criticalCount ?? "", warningDetails: runReportWarningDetails(item), blocksProcessed: "" });
      }

      if (item.previewSummary) summaryLines.push(`     ${item.previewSummary}.`);
      if (item.candidateNotes && item.candidateNotes.length > 0) {
        summaryLines.push(`     [${item.candidateNotes.join("; ")}]`);
      }
    }
  }

  if (skippedSheets && skippedSheets.length > 0) {
    summaryLines.push("");
    summaryLines.push("Пропущенные листы:");
    for (const sheet of skippedSheets) {
      summaryLines.push(`- ${formatInventorySkippedSheetLine(sheet)}`);
    }
  }

  summaryLines.push("");
  summaryLines.push("Данные не изменены. Для детальной проверки используйте «Проверка → Текущий лист».");

  setCheckMessage(summaryLines.join("\n"));

  if (isCheckReportEnabled() && checkReportRows.length > 0) {
    try {
      await Excel.run(async (context) => {
        await writeRunReportSheet(context, checkReportRows, "Проверка — Вся книга");
      });
    } catch (reportErr) {
      const el = document.getElementById("check-result");
      if (el) el.textContent += `\n\n[Run report: ошибка записи — ${reportErr.message || reportErr}]`;
    }
  }
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
 * Read-only check pipeline for a named range inside an existing Excel.run context.
 *
 * Loads the range, runs interpretSelectedRange + buildTablePreviewModel, and
 * returns a structured result. Callers handle display and Run report writing.
 *
 * Must be called inside an Excel.run callback — uses the shared context directly
 * so no nested Excel.run is created.
 *
 * Reserved for future manual pre-run check (Расчёт → Проверить выделение).
 * Not wired into any public UI path yet.
 *
 * @param {Excel.RequestContext} context
 * @param {string} sheetName
 * @param {string} rangeAddress  Local A1 address, e.g. "B3:F15".
 * @param {object} settings      Calculation settings (same shape as the rest of the pipeline).
 * @returns {Promise<{
 *   status: "checked" | "blocked" | "empty" | "no-data",
 *   sheetName: string,
 *   rangeAddress: string,
 *   model?: object,
 *   normalizationLines?: string[],
 *   message?: string
 * }>}
 */
async function checkSelectedRangePreview(context, sheetName, rangeAddress, settings) {
  const worksheet = context.workbook.worksheets.getItem(sheetName);
  const sourceRange = worksheet.getRange(rangeAddress);
  sourceRange.load(["values", "text", "rowIndex", "columnIndex", "rowCount", "columnCount"]);

  await context.sync();

  const selectedValues = sourceRange.values;
  const selectedText = sourceRange.text;

  if (
    !selectedValues ||
    selectedValues.length < 1 ||
    !selectedValues[0] ||
    selectedValues[0].length < 1
  ) {
    return { status: "no-data", sheetName, rangeAddress, message: t("status.noDataInRange") };
  }

  // Pre-normalization guard: check the raw loaded range for all-empty row gaps
  // BEFORE interpretSelectedRange is called. Must run here because normalization
  // can reduce a multi-table range to a single-table body, silently masking the
  // presence of additional tables and producing a misleading one-table diagnostic.
  if (hasEmptyDataRowGap(selectedValues)) {
    return {
      status: "blocked",
      sheetName,
      rangeAddress,
      message:
        t("status.multiTableGapCheckSelection"),
    };
  }

  const interpretation = await interpretSelectedRange(
    context,
    sourceRange,
    selectedValues,
    selectedText,
    settings
  );

  if (interpretation.state === "blocked") {
    const codes =
      interpretation.blockingReasons.length > 0
        ? ` [${interpretation.blockingReasons.join(", ")}]`
        : "";
    return {
      status: "blocked",
      sheetName,
      rangeAddress,
      message: `${interpretation.blockingMessage}${codes}`,
    };
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
    // Report the adjacent label AREA width (shared source of truth with design
    // recolor), not just the normalizer's text-only label columns: a merged or
    // left-stored two-column label header counts as two label columns.
    if (interpretation.adjacentLabelColumnCount > 0)
      parts.push(`колонок меток: ${interpretation.adjacentLabelColumnCount}`);
    if (parts.length > 0) normalizationLines.push(`Отделено: ${parts.join(", ")}.`);
  }

  const model = buildTablePreviewModel({
    values: valuesForCalculation,
    leftLabelValues,
    bannerContext,
    settings,
    trailingBodyRows: normalized?.trailingBodyRows,
  });

  if (model.summary.rowCount === 0) {
    return { status: "empty", sheetName, rangeAddress, normalizationLines, model, message: "Диапазон пуст." };
  }

  return { status: "checked", sheetName, rangeAddress, normalizationLines, model };
}

/**
 * Check Текущая таблица: resolves the detected table under the active cell
 * and runs the read-only check pipeline on it.
 *
 * Uses resolveCurrentTableFromActiveCell (#208) to find the candidate range,
 * then delegates the data pipeline to checkSelectedRangePreview.
 * For non-ok resolver status: shows a user-facing message.
 *
 * Pre-resolver selection guard: if the user currently has a broad multi-table
 * selection active (non-empty row gap between non-empty row groups), we block
 * before the resolver runs to avoid a misleading one-table diagnostic. This is
 * NOT a return to selected-range Check semantics — a normal single-cell or
 * single-table selection proceeds to the active-cell resolver as usual.
 *
 * Does not modify the Excel workbook unless "Записать результат" is enabled,
 * in which case it writes a single row to the Run report sheet.
 */
async function runCheckTable() {
  setCheckMessage(runningStatusMessage("check"));
  await Excel.run(async (context) => {
    const calculationSettings = readCalculationSettingsFromPanel();

    // Pre-resolver selection sanity guard.
    // Load the current selected range and check for a multi-table gap BEFORE
    // calling the active-cell resolver. If the selection spans multiple table-like
    // blocks separated by empty rows, block immediately with a clear message.
    // A single-cell selection or a normal single-table selection passes through.
    //
    // getSelectedRange() throws a RichApi.Error for non-contiguous (Ctrl+Click
    // multi-area) selections. Catch that error here and surface a user-facing
    // message rather than letting the runtime error propagate.
    let guardValues;
    try {
      const selectionForGuard = context.workbook.getSelectedRange();
      selectionForGuard.load(["values"]);
      await context.sync();
      guardValues = selectionForGuard.values;
    } catch (_selectionErr) {
      setCheckMessage(nonContiguousSelectionMessage());
      return;
    }

    if (selectionHasMultiTableGap(guardValues)) {
      const msg = t("status.multiTableGapCheck");
      setCheckMessage(msg);
      if (isCheckReportEnabled()) {
        await writeRunReportSheet(context, [{
          sheetName: "",
          title: "",
          rangeAddress: "",
          status: "blocked",
          message: msg,
          selectedBase: "",
          metricTypes: "",
          warnings: 0,
          critical: 0,
          warningDetails: "",
          blocksProcessed: 0,
        }], "Проверка — Текущая таблица");
      }
      return;
    }

    // Resolve active cell → detected table (read-only, no selected-range dependency).
    const resolverResult = await resolveCurrentTableFromActiveCell(context, calculationSettings);

    if (resolverResult.status !== "ok") {
      const msg = buildCheckResolverMessage(resolverResult);
      setCheckMessage(msg);
      if (isCheckReportEnabled()) {
        await writeRunReportSheet(context, [{
          sheetName: resolverResult.sheetName || "",
          title: "",
          rangeAddress: "",
          status: "skipped",
          message: msg,
          selectedBase: "",
          metricTypes: "",
          warnings: 0,
          critical: 0,
          warningDetails: "",
          blocksProcessed: 0,
        }], "Проверка — Текущая таблица");
      }
      return;
    }

    const { sheetName, rangeAddress } = resolverResult;
    const resolvedTableTitle = resolveContentDisplayTitle(resolverResult.candidateMeta, 1);

    const checkResult = await checkSelectedRangePreview(
      context,
      sheetName,
      rangeAddress,
      calculationSettings
    );

    if (checkResult.status === "no-data") {
      setCheckMessage(checkResult.message);
      return;
    }

    if (checkResult.status === "blocked") {
      const msg = checkResult.message;
      setCheckMessage(msg);
      if (isCheckReportEnabled()) {
        await writeRunReportSheet(context, [{
          sheetName,
          title: resolvedTableTitle,
          rangeAddress,
          status: "blocked",
          message: msg,
          selectedBase: "",
          metricTypes: "",
          warnings: 0,
          critical: 0,
          warningDetails: "",
          blocksProcessed: 0,
        }], "Проверка — Текущая таблица");
      }
      return;
    }

    if (checkResult.status === "empty") {
      setCheckMessage(checkResult.message);
      return;
    }

    // status === "checked"
    const { model, normalizationLines } = checkResult;
    const {
      summary,
      qualitySummary,
      userVisibleIssues,
      bannerStructure,
      calculationBlocks,
      rowDiagnostics,
    } = model;

    const lines = [
      t("status.checkDone", {
        sheet: sheetName,
        range: rangeAddress,
        rows: summary.rowCount,
        blocks: summary.detectedBlocks,
        bases: summary.baseRows,
        warnings: qualitySummary.warningCount,
        critical: qualitySummary.criticalCount,
      }),
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

    if (isCheckReportEnabled()) {
      const warningDetails = (userVisibleIssues || [])
        .map((iss) => `[${iss.severity}] ${iss.message}`)
        .join("; ");
      await writeRunReportSheet(context, [{
        sheetName,
        title: resolvedTableTitle,
        rangeAddress,
        status: "checked",
        message: `Строк: ${summary.rowCount}. Блоков: ${summary.detectedBlocks}. Баз: ${summary.baseRows}.`,
        selectedBase: "",
        metricTypes: checkMetricTypesFromBlocks(calculationBlocks),
        warnings: qualitySummary.warningCount,
        critical: qualitySummary.criticalCount,
        warningDetails,
        blocksProcessed: summary.detectedBlocks,
      }], "Проверка — Текущая таблица");
    }
  });
}

/**
 * Расчёт → Проверить выделение: runs a read-only selected-range check directly
 * on the current Excel selection — does NOT use the active-cell resolver.
 *
 * Reads the selected range address, calls checkSelectedRangePreview with the
 * exact selection, and renders the result in the check panel. Nothing is written
 * to the workbook.
 */
async function runCheckSelectedRange() {
  setCheckMessage(runningStatusMessage("check"));
  await Excel.run(async (context) => {
    const calculationSettings = readCalculationSettingsFromPanel();

    // getSelectedRange() throws a RichApi.Error for non-contiguous (Ctrl+Click
    // multi-area) selections. Catch the error here and show a user-facing
    // message rather than letting the runtime error propagate.
    let sheetName;
    let rangeAddress;
    try {
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load(["address"]);
      await context.sync();
      const worksheet = selectedRange.worksheet;
      worksheet.load(["name"]);
      await context.sync();

      sheetName = worksheet.name;
      const fullAddress = selectedRange.address;
      const exclamationIndex = fullAddress.lastIndexOf("!");
      rangeAddress =
        exclamationIndex >= 0 ? fullAddress.substring(exclamationIndex + 1) : fullAddress;
    } catch (_selectionErr) {
      setCheckMessage(nonContiguousSelectionMessage());
      return;
    }

    const checkResult = await checkSelectedRangePreview(
      context,
      sheetName,
      rangeAddress,
      calculationSettings
    );

    if (checkResult.status === "no-data") {
      setCheckMessage(checkResult.message);
      return;
    }

    if (checkResult.status === "blocked") {
      setCheckMessage(checkResult.message);
      return;
    }

    if (checkResult.status === "empty") {
      setCheckMessage(checkResult.message);
      return;
    }

    // status === "checked"
    const { model, normalizationLines } = checkResult;
    const { summary, qualitySummary, userVisibleIssues, bannerStructure, calculationBlocks, rowDiagnostics } = model;

    const lines = [
      t("status.checkSelectionDone", {
        sheet: sheetName,
        range: rangeAddress,
        rows: summary.rowCount,
        blocks: summary.detectedBlocks,
        bases: summary.baseRows,
        warnings: qualitySummary.warningCount,
        critical: qualitySummary.criticalCount,
      }),
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
      lines.push(`- ${formatInventorySkippedSheetLine(sheet)}`);
    });
  }

  lines.push("");
  lines.push("Оглавление — только поиск кандидатов. Для детальной проверки используйте «Проверить таблицу».");

  return lines.join("\n").trimEnd();
}

async function listWorkbookInventoryWorksheetNames() {
  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");

    await context.sync();

    return worksheets.items
      .map((worksheet) => worksheet.name)
      .filter((sheetName) => !GENERATED_SHEET_NAMES.has(sheetName));
  });
}

async function collectWorksheetInventoryResultWithContext(context, worksheet, sheetName, settings) {
  const usedRange = worksheet.getUsedRangeOrNullObject();
  usedRange.load(["isNullObject", "rowIndex", "columnIndex", "rowCount", "columnCount"]);

  await context.sync();

  if (usedRange.isNullObject) {
    return {
      scanned: false,
      skippedSheet: { sheetName, reason: "empty" },
    };
  }

  const cellCount = usedRange.rowCount * usedRange.columnCount;
  if (cellCount > INVENTORY_SCAN_EMERGENCY_CELL_LIMIT) {
    return {
      scanned: false,
      skippedSheet: {
        sheetName,
        reason: "emergencyLimit",
        rowCount: usedRange.rowCount,
        columnCount: usedRange.columnCount,
        cellCount,
      },
    };
  }

  usedRange.load("values");

  try {
    await context.sync();
  } catch (error) {
    return {
      scanned: false,
      skippedSheet: {
        sheetName,
        reason: "scanError",
        rowCount: usedRange.rowCount,
        columnCount: usedRange.columnCount,
        cellCount,
        message: inventoryScanErrorMessage(error),
      },
    };
  }

  try {
    const items = scanWorksheetForTables({
      values: usedRange.values,
      usedRangeRowOffset: usedRange.rowIndex,
      usedRangeColOffset: usedRange.columnIndex,
      sheetName,
      settings: { ...settings, backlinkMarker: BACKLINK_MARKER },
    });

    return {
      scanned: true,
      sheetResult: items.length > 0
        ? {
            sheetName,
            usedRangeRowOffset: usedRange.rowIndex,
            usedRangeColOffset: usedRange.columnIndex,
            usedRangeValues: usedRange.values,
            items,
          }
        : null,
    };
  } catch (error) {
    return {
      scanned: false,
      skippedSheet: {
        sheetName,
        reason: "scanError",
        rowCount: usedRange.rowCount,
        columnCount: usedRange.columnCount,
        cellCount,
        message: inventoryScanErrorMessage(error),
      },
    };
  }
}

async function collectWorksheetInventoryResultByName(sheetName, settings) {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    return collectWorksheetInventoryResultWithContext(context, worksheet, sheetName, settings);
  });
}

async function collectWorkbookInventoryResults(settings) {
  const worksheetNames = await listWorkbookInventoryWorksheetNames();
  const sheetResults = [];
  const skippedSheets = [];
  let scannedSheets = 0;

  for (const sheetName of worksheetNames) {
    const result = await collectWorksheetInventoryResultByName(sheetName, settings);
    if (result.scanned) {
      scannedSheets += 1;
      if (result.sheetResult) {
        sheetResults.push(result.sheetResult);
      }
    } else if (result.skippedSheet) {
      skippedSheets.push(result.skippedSheet);
    }
  }

  return {
    scannedSheets,
    sheetResults,
    skippedSheets,
  };
}

/**
 * Collects inventory results for the active worksheet only.
 *
 * Mirrors collectWorkbookInventoryResults() but scans only the active sheet,
 * returning the same { scannedSheets, sheetResults, skippedSheets } shape so
 * that filterWorkbookCandidates() and normalizeBacklinkItems() work unchanged.
 */
async function collectActiveSheetInventoryResults(settings) {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    worksheet.load("name");

    await context.sync();

    if (GENERATED_SHEET_NAMES.has(worksheet.name)) {
      return { scannedSheets: 0, sheetResults: [], skippedSheets: [] };
    }

    const result = await collectWorksheetInventoryResultWithContext(context, worksheet, worksheet.name, settings);
    if (!result.scanned) {
      return {
        scannedSheets: 0,
        sheetResults: [],
        skippedSheets: result.skippedSheet ? [result.skippedSheet] : [],
      };
    }

    return {
      scannedSheets: 1,
      sheetResults: result.sheetResult ? [result.sheetResult] : [],
      skippedSheets: [],
    };
  });
}

function buildInventoryContentCandidateRows(sheetResults) {
  const rows = [];
  let candidateIndex = 1;

  for (const sheetResult of sheetResults) {
    for (const item of sheetResult.items) {
      rows.push([
        candidateIndex,
        sheetResult.sheetName,
        resolveContentDisplayTitle(item, candidateIndex),
        item.resolvedRangeAddress || item.rangeAddress || "",
        item.rowCount ?? "",
        item.columnCount ?? "",
        getContentCandidateStatusLabel(item),
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

  return skippedSheets.map((sheet) => buildInventoryContentSkippedRow(sheet));
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

// ─── Significance settings footnote ────────────────────────────────────────────
//
// When the "Добавлять подпись под таблицей" setting is on, each PROCESSED table
// gets one footnote row inserted directly below it, merged across the full table
// width (label columns + data columns). Skipped/blocked/error tables get nothing.
//
// IMPORTANT ordering rule (see issue #281): footnotes must NOT be inserted while
// calculations are still running, because inserting a worksheet row shifts every
// candidate range below it and corrupts the ranges still queued for calculation.
// Instead, each run flow collects footnote "jobs" (pure geometry + text, computed
// from the already-known ranges) and applies them AFTER all calculations finish,
// bottom-to-top per worksheet so earlier insertions never shift a not-yet-applied
// job's row index.

/**
 * Builds a footnote job for a single processed table, or null when the footnote
 * setting is off / geometry is unusable.
 *
 * Geometry is sheet-absolute and 0-based:
 *   - dataStartRowIndex / dataStartColIndex: top-left of the data body.
 *   - dataRowCount / dataColCount: data body dimensions.
 *   - leftLabelValues: label columns immediately left of the data body. Their
 *     width sets how far left the merged footnote extends (the "full table width
 *     including label columns").
 *
 * Returns null when the footnote setting is off. The feature is explicitly
 * incompatible with "labels on left side" mode (labels are not adjacent to the
 * data, so a footnote cannot span the table correctly): readCalculationSettings-
 * FromPanel already forces addTableFootnote=false in that mode, and this extra
 * guard makes the rule independent of the caller.
 *
 * processedScopeSuffix (issue #308) is an optional, already-formatted detail
 * (e.g. " Обработано: B12:F34.") appended to the visible footnote text. It is
 * supplied only by Manual Run; auto-run callers omit it so their footnote text
 * is unchanged.
 */
function buildSignificanceFootnoteJob({
  sheetName,
  dataStartRowIndex,
  dataStartColIndex,
  dataRowCount,
  dataColCount,
  leftLabelValues,
  adjacentLabelColumnCount,
  calculationBlocks,
  calculationSettings,
  processedScopeSuffix,
}) {
  if (!calculationSettings.addTableFootnote) return null;
  if (calculationSettings.labelsOnLeftSide) return null;
  if (!Number.isFinite(dataStartRowIndex) || !Number.isFinite(dataStartColIndex)) return null;
  if (!(dataRowCount > 0) || !(dataColCount > 0)) return null;

  // Visual table left edge: the wider of the loaded label width and the
  // interpreter's adjacent label-area width, so a stripped structural gap column
  // between labels and data (normalized autorun) is still spanned. See
  // resolveFootnoteSpan.
  const labelColumns =
    Array.isArray(leftLabelValues) && Array.isArray(leftLabelValues[0]) ? leftLabelValues[0].length : 0;

  const { tableLeftColIndex, tableRightColIndex } = resolveFootnoteSpan({
    dataStartColIndex,
    dataColCount,
    labelColumns,
    adjacentLabelColumnCount,
  });
  const tableBottomRowIndex = dataStartRowIndex + dataRowCount - 1;

  const footnoteCellValue = buildSignificanceFootnoteCellValue({
    confidenceLevel: calculationSettings.confidenceLevel,
    oneTailedTest: calculationSettings.oneTailedTest,
    statisticLabels: collectStatisticTypeLabels(calculationBlocks),
    scopeDetail: processedScopeSuffix,
  });

  return {
    sheetName,
    tableBottomRowIndex,
    tableLeftColIndex,
    tableRightColIndex,
    dataStartColIndex,
    footnoteCellValue,
  };
}

/**
 * Inserts or updates a single footnote row for one table.
 *
 * A small bounded window of rows below the table body is scanned (values loaded
 * across a span that also covers a couple of columns to the LEFT, so the marker
 * is seen even when a label/data blank-gap column shifts the merged footnote's
 * anchor). Placement is decided by the pure resolveFootnotePlacement helper:
 *
 *  - if a generated RIT footnote already exists for the table (marker-based), it
 *    is updated in place — no new row, no duplicate;
 *  - ordinary user note rows directly below the table are skipped, and the
 *    generated footnote is inserted BELOW them (never overwriting them);
 *  - otherwise a brand-new worksheet row is inserted so nothing is overwritten.
 *
 * The footnote cells are merged across exactly [tableLeftColIndex .. tableRightColIndex].
 */
async function writeOrInsertSignificanceFootnoteRow(context, worksheet, job) {
  const { tableBottomRowIndex, tableLeftColIndex, tableRightColIndex, dataStartColIndex, footnoteCellValue } = job;
  const width = tableRightColIndex - tableLeftColIndex + 1;
  if (width < 1) return;

  const firstRowBelowTable = tableBottomRowIndex + 1;

  // Scan a span starting a few columns LEFT of the table so a left-offset marker
  // (from a run whose label/gap geometry differed) is still detected.
  const scanStartCol = Math.max(0, tableLeftColIndex - LABEL_SCAN_COLUMNS_LEFT);
  const scanColCount = tableRightColIndex - scanStartCol + 1;
  const dataColRef = Number.isFinite(dataStartColIndex) ? dataStartColIndex : tableLeftColIndex;
  const dataColStartOffset = dataColRef - scanStartCol;

  // Load the bounded window of rows below the table to decide placement.
  const scanRange = worksheet.getRangeByIndexes(
    firstRowBelowTable,
    scanStartCol,
    FOOTNOTE_SCAN_WINDOW_ROWS,
    scanColCount
  );
  scanRange.load("values");
  await context.sync();

  const placement = resolveFootnotePlacement(
    scanRange.values,
    firstRowBelowTable,
    dataColStartOffset
  );
  const footnoteRowIndex = placement.rowIndex;

  if (placement.mode === "insert") {
    // Insert a new blank row so nothing existing below the table is overwritten.
    worksheet
      .getRangeByIndexes(footnoteRowIndex, 0, 1, 1)
      .getEntireRow()
      .insert(Excel.InsertShiftDirection.down);
    await context.sync();
  } else {
    // Updating an existing generated footnote row in place: clear the full scan
    // span first so a prior merge and any left-offset marker cell are removed
    // before the row is rewritten at the current geometry. The row is RIT's own
    // generated footnote row, so clearing its span never touches user content.
    const priorRange = worksheet.getRangeByIndexes(footnoteRowIndex, scanStartCol, 1, scanColCount);
    priorRange.unmerge();
    priorRange.clear(Excel.ClearApplyTo.contents);
    await context.sync();
  }

  const footnoteRange = worksheet.getRangeByIndexes(footnoteRowIndex, tableLeftColIndex, 1, width);

  // Clear any prior merge first: a merged range rejects multi-cell writes, and an
  // existing footnote from a previous run may already be merged across this span.
  footnoteRange.unmerge();
  await context.sync();

  const leftCell = worksheet.getRangeByIndexes(footnoteRowIndex, tableLeftColIndex, 1, 1);
  leftCell.values = [[footnoteCellValue]];

  if (width > 1) {
    footnoteRange.merge();
  }

  footnoteRange.format.font.italic = true;
  footnoteRange.format.font.bold = false;
  footnoteRange.format.fill.clear();
  footnoteRange.format.horizontalAlignment = "Left";
  footnoteRange.format.verticalAlignment = "Center";
  await context.sync();
}

/**
 * Applies a list of footnote jobs. Jobs are grouped per worksheet and applied
 * bottom-to-top (descending table bottom row) so that inserting a footnote row
 * for a lower table never shifts the row index of a higher table still pending.
 *
 * Tolerant: a failure on one footnote is logged and skipped; it never aborts the
 * Run itself (calculations have already completed and been written at this point).
 */
async function applySignificanceFootnoteJobs(context, jobs) {
  if (!Array.isArray(jobs) || jobs.length === 0) return;

  const jobsBySheet = new Map();
  for (const job of jobs) {
    if (!job || !job.sheetName) continue;
    if (!jobsBySheet.has(job.sheetName)) jobsBySheet.set(job.sheetName, []);
    jobsBySheet.get(job.sheetName).push(job);
  }

  for (const [sheetName, sheetJobs] of jobsBySheet) {
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const sorted = sheetJobs
      .slice()
      .sort((a, b) => b.tableBottomRowIndex - a.tableBottomRowIndex);
    for (const job of sorted) {
      try {
        await writeOrInsertSignificanceFootnoteRow(context, worksheet, job);
      } catch (footnoteErr) {
        console.warn("RIT: не удалось записать подпись под таблицей.", footnoteErr);
      }
    }
  }
}

/**
 * Convenience wrapper that applies footnote jobs inside their own Excel.run.
 * Used by run flows whose calculations completed in a different context.
 */
async function applySignificanceFootnoteJobsInOwnContext(jobs) {
  if (!Array.isArray(jobs) || jobs.length === 0) return;
  try {
    await Excel.run(async (context) => {
      await applySignificanceFootnoteJobs(context, jobs);
    });
  } catch (footnoteErr) {
    console.warn("RIT: не удалось записать подписи под таблицами.", footnoteErr);
  }
}

// ─── Design recolor (issue #306) ────────────────────────────────────────────────
//
// Optional design feature: after a successful Run, recolor the banner/header rows
// above the data body and the adjacent row-label columns with one shared color.
//
// Geometry comes from the selected-range interpretation (banner band + adjacent
// label column count), so the writer never re-guesses label width. Each job is a
// small set of sheet-absolute rectangles; applying a job sets ONE fill per
// rectangle (never per cell, never per row), keeping the operation range-level.
//
// Recolor sets only fill color: it inserts no rows and changes no values, so it
// never shifts ranges. Recolor jobs are therefore applied BEFORE footnote jobs
// (which DO insert rows) so the recolor rectangles still match the unshifted
// sheet geometry captured during calculation.

/**
 * Queues the fill writes for one design recolor job onto a worksheet proxy.
 * The caller is responsible for the context.sync() that flushes them, so several
 * jobs/rectangles can share a single sync.
 */
function queueDesignRecolorJob(worksheet, job) {
  if (!job || !Array.isArray(job.rects)) return;
  for (const rect of job.rects) {
    const range = worksheet.getRangeByIndexes(
      rect.rowIndex,
      rect.columnIndex,
      rect.rowCount,
      rect.columnCount
    );
    range.format.fill.color = job.color;
  }
}

/**
 * Applies a list of design recolor jobs grouped per worksheet. All fill writes
 * for a worksheet are queued and flushed with a single context.sync(), so the
 * cost is one sync per worksheet regardless of table count.
 *
 * Tolerant: a failure on one worksheet is logged and skipped; it never aborts the
 * Run itself (calculations have already completed and been written at this point).
 */
async function applyDesignRecolorJobs(context, jobs) {
  if (!Array.isArray(jobs) || jobs.length === 0) return;

  const jobsBySheet = new Map();
  for (const job of jobs) {
    if (!job || !job.sheetName) continue;
    if (!jobsBySheet.has(job.sheetName)) jobsBySheet.set(job.sheetName, []);
    jobsBySheet.get(job.sheetName).push(job);
  }

  for (const [sheetName, sheetJobs] of jobsBySheet) {
    try {
      const worksheet = context.workbook.worksheets.getItem(sheetName);
      for (const job of sheetJobs) {
        queueDesignRecolorJob(worksheet, job);
      }
      await context.sync();
    } catch (recolorErr) {
      console.warn("RIT: не удалось перекрасить баннер и лейблы.", recolorErr);
    }
  }
}

/**
 * Convenience wrapper that applies design recolor jobs inside their own Excel.run.
 * Used by run flows whose calculations completed in a different context.
 */
async function applyDesignRecolorJobsInOwnContext(jobs) {
  if (!Array.isArray(jobs) || jobs.length === 0) return;
  try {
    await Excel.run(async (context) => {
      await applyDesignRecolorJobs(context, jobs);
    });
  } catch (recolorErr) {
    console.warn("RIT: не удалось перекрасить баннер и лейблы.", recolorErr);
  }
}

/**
 * Computes the effective banner-row count to recolor above a data body.
 *
 * - In a normalized selection the interpreter reports the exact banner band it
 *   detected inside the selection (excluding title/subtitle rows).
 * - In a pass-through selection (the user selected only the numeric data body)
 *   any banner sits in the sheet directly above the selection, so the single
 *   header row immediately above the data body is recolored when one exists.
 */
function resolveBannerRecolorRowCount(interpretation, dataStartRowIndex) {
  const interpreted = interpretation.bannerRowsAboveData || 0;
  if (interpreted > 0) return interpreted;
  if (interpretation.state === "passThrough" && dataStartRowIndex > 0) return 1;
  return 0;
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
  titleRange.values = normalizeRowsToColumnCount([["Оглавление таблиц"]], 11);
  titleRange.merge();
  titleRange.format.font.bold = true;
  titleRange.format.font.size = 14;

  const metadataRows = normalizeRowsToColumnCount(
    [
      ["Generated sheet", INVENTORY_CONTENT_SHEET_NAME],
      ["Scanned sheets", inventoryResults.scannedSheets],
      ["Candidate sheets", inventoryResults.sheetResults.length],
      ["Detected candidates", totalCandidates],
      ["Примечание", "Оглавление — поиск кандидатов. Для детальной проверки используйте «Проверить таблицу»."],
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

  const TITLE_COL_INDEX = 2;
  for (let i = 0; i < allItems.length; i++) {
    const item = allItems[i];
    const effectiveRange = item.resolvedRangeAddress || item.rangeAddress;
    const hyperlinkTarget = getContentTableHyperlinkTarget(item.sheetName, effectiveRange);
    if (hyperlinkTarget) {
      const cell = worksheet.getRangeByIndexes(headerRowIndex + i, TITLE_COL_INDEX, 1, 1);
      cell.hyperlink = {
        documentReference: hyperlinkTarget,
        textToDisplay: resolveContentDisplayTitle(item, i + 1),
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
  titleRange.values = normalizeRowsToColumnCount([["Оглавление таблиц"]], colCount);
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
      cell.hyperlink = {
        documentReference: hyperlinkTarget,
        textToDisplay: resolveContentDisplayTitle(item, i + 1),
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
        resolveContentDisplayTitle(item, candidateIndex),
        item.resolvedRangeAddress || item.rangeAddress || "",
        getContentCandidateStatusLabel(item),
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
        formatIssueDetailsForReport(item),
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

  const TITLE_COL_INDEX = 2;
  for (let i = 0; i < allItems.length; i++) {
    const item = allItems[i];
    const effectiveRange = item.resolvedRangeAddress || item.rangeAddress;
    const hyperlinkTarget = getContentTableHyperlinkTarget(item.sheetName, effectiveRange);
    if (hyperlinkTarget) {
      const cell = worksheet.getRangeByIndexes(headerRowIndex + i, TITLE_COL_INDEX, 1, 1);
      cell.hyperlink = {
        documentReference: hyperlinkTarget,
        textToDisplay: resolveContentDisplayTitle(item, i + 1),
        screenTip: `${item.sheetName}!${effectiveRange}`,
      };
    }
  }
}

/**
 * Creates or updates the Content sheet, writes inventory content, moves the
 * sheet to the first position (tab index 0), and returns the worksheet object.
 *
 * Placement uses the same two-tier strategy as writeRunReportSheet():
 *
 * Tier 1 — direct position assignment.
 *   Content writes and position assignment are in separate sync batches.
 *   A verification pass re-reads the actual server-side position. If the
 *   sheet is already first, we return early.
 *
 * Tier 2 — copy / delete / rename fallback.
 *   When Tier 1 verification shows the sheet is still not first, we fall
 *   back to: find the first non-Content sheet, copy Content before it,
 *   activate the copy, delete the original, rename the copy.
 */
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

  // Tier 1 — flush content writes first, then attempt direct position assignment.
  await context.sync();

  const worksheets = context.workbook.worksheets;
  worksheet.position = 0;
  await context.sync();

  // Verification: reload the sheet's actual position from the host.
  worksheet.load("position");
  await context.sync();

  if (worksheet.position === 0) {
    return worksheet; // Tier 1 succeeded — sheet is already first.
  }

  // Tier 2 — copy/delete/rename fallback.
  // Direct position assignment did not take effect; use Worksheet.copy() to
  // physically place the sheet at the start of the tab bar.

  // Reload worksheet items in tab order to find the first non-Content sheet.
  worksheets.load("items/name");
  await context.sync();

  const allSheets = worksheets.items;
  let firstSheet = null;
  for (let i = 0; i < allSheets.length; i++) {
    if (allSheets[i].name !== INVENTORY_CONTENT_SHEET_NAME) {
      firstSheet = allSheets[i];
      break;
    }
  }

  if (firstSheet === null) {
    // Content is the only sheet — trivially at position 0.
    return worksheet;
  }

  // Copy the sheet immediately before firstSheet (the true first position).
  const copy = worksheet.copy("Before", firstSheet);

  // Activate the copy before deleting the original to satisfy the host
  // constraint that the active sheet cannot be deleted.
  copy.activate();
  await context.sync();

  // Remove the original (now not-active) sheet and rename the copy.
  worksheet.delete();
  await context.sync();

  copy.name = INVENTORY_CONTENT_SHEET_NAME;
  await context.sync();

  return copy;
}

async function runTableInventory() {
  setInventoryMessage(runningStatusMessage("content"));
  const _t0 = perfNow();
  const inventoryResults = await collectWorkbookInventoryResults(readCalculationSettingsFromPanel());
  const _tScan = perfNow();
  const addBacklinks = readBacklinkSettingFromPanel();
  const contentMode = readContentOutputModeFromPanel();

  // Always normalize: fixes resolvedRangeAddress for items whose detected range
  // starts with a generated backlink row (in-range), even when backlinks are OFF.
  // When backlinks are ON, also predicts the +1 end-row shift for will-insert.
  normalizeBacklinkItems(inventoryResults.sheetResults, addBacklinks);

  const contentRowMap = addBacklinks
    ? buildContentRowMap(inventoryResults.sheetResults, contentMode)
    : null;

  await Excel.run(async (context) => {
    // Write Content first (hyperlinks use resolvedRangeAddress when available).
    const contentWorksheet = await writeInventoryContentSheet(context, inventoryResults);

    // Insert/update backlink rows after Content is written so the Content sheet
    // is not affected by row insertions in source sheets.
    if (addBacklinks && contentRowMap) {
      await ensureBacklinkRows(context, inventoryResults.sheetResults, contentRowMap);
    }

    // Switch to the Content sheet after all writes are done.
    contentWorksheet.activate();
    await context.sync();
  });

  perfLog("runTableInventory", {
    scanMs: _tScan - _t0,
    contentWriteMs: perfElapsed(_tScan),
    totalMs: perfElapsed(_t0),
  });
  setInventoryMessage(
    formatWorkbookInventoryMessage({
      scannedSheets: inventoryResults.scannedSheets,
      sheetResults: inventoryResults.sheetResults,
      skippedSheets: inventoryResults.skippedSheets,
    })
  );
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

  // Footnotes require label columns adjacent to the data, so they are
  // incompatible with "labels on left side" mode. Force the feature off here so
  // a stale saved setting (addTableFootnote: true) can never create a footnote
  // while labelsOnLeftSide is true, independent of the disabled-UI affordance.
  const labelsOnLeftSide = getCheckboxValue("labels-on-left-side");
  const addTableFootnote = labelsOnLeftSide ? false : getCheckboxValue("add-table-footnote");

  // Design recolor (issue #306) requires label columns adjacent to the data, so
  // it is incompatible with "labels on left side" mode. Force the feature off
  // here so a stale saved setting can never recolor while labelsOnLeftSide is
  // true, independent of the disabled-UI affordance.
  const recolorBannerAndLabels = labelsOnLeftSide
    ? false
    : getCheckboxValue("recolor-banner-and-labels");

  return {
    confidenceLevel: confidenceLevelElement ? confidenceLevelElement.value : "95",
    oneTailedTest: oneTailedTestCheckbox ? oneTailedTestCheckbox.checked : false,

    roundCellValues: getCheckboxValue("round-cell-values"),
    resultFormattingLevel: getInputValue("result-formatting-level", "full"),
    addTableFootnote,

    compareWithPreviousColumn: getCheckboxValue("compare-with-previous-column"),
    applyPreviousColumnFill: getCheckboxValue("apply-previous-column-fill"),

    writeBannerLetters: getCheckboxValue("write-banner-letters"),
    respectBannerStructure: getCheckboxValue("respect-banner-structure"),
    autoDetectWaveBanners: getCheckboxValue("auto-detect-wave-banners"),
    labelsOnLeftSide,

    // Whether allowed Cyrillic letters may be used as significance markers.
    // Default-off for global release safety (issue #312). Independent of the
    // task pane UI language.
    useCyrillicMarkers: getCheckboxValue("use-cyrillic-markers"),

    // Runtime-only decision (not persisted). Set to true after the user opts to
    // continue with multi-character overflow markers in the overflow dialog.
    allowMultiCharacterMarkers: false,

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

    recolorBannerAndLabels,
    bannerLabelFillColor: getInputValue("banner-label-fill-color", "#FFF2CC"),

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
 * Combined banner-marker clear + write for a Run pass.
 *
 * Uses a single read/plan/write phase.  Reads the full banner scan area
 * once, strips all RIT trailing markers in-memory (the clear phase), then
 * applies the desired markers (lower banner row when non-empty, otherwise
 * nearest non-empty cell above; uses `bannerStructure` labels when respect-
 * banner-structure is on, otherwise sequential significance labels with the
 * first-column-is-total adjustment).  Diffs the desired matrix against the
 * read texts and writes back only cells that actually changed, batched into
 * contiguous same-row runs and split by markered/clear-only flag so the
 * "@" number-format guard is preserved on marker writes.
 *
 * Sync count in normal mode: 1 for the banner-area read (also flushes any
 * pending data writes from the caller) and 0 for the writes — the caller is
 * expected to issue its own final `context.sync()` to flush queued banner
 * writes.
 *
 * Development-only banner write profiling can split numberFormat and values
 * into separate syncs to measure host flush cost. It is disabled by default.
 *
 * Returns null when no banner work was needed.  Otherwise returns a perf
 * details object describing the work performed.
 */
async function applyBannerMarkerUpdatesForRange(
  context,
  writeTargetRange,
  calculationSettings,
  bannerStructure,
  knownDimensions,
  sourceStartColIndex
) {
  if (!calculationSettings.writeBannerLetters) {
    return null;
  }

  const BANNER_UPPER_SCAN_LIMIT = 5;
  const BANNER_SCAN_ROW_COUNT = BANNER_UPPER_SCAN_LIMIT + 1;

  let targetStartRowIndex, targetStartColumnIndex, targetColumnCount;
  if (knownDimensions) {
    targetStartRowIndex = knownDimensions.rowIndex;
    targetStartColumnIndex = knownDimensions.columnIndex;
    targetColumnCount = knownDimensions.columnCount;
  } else {
    writeTargetRange.load(["rowIndex", "columnIndex", "columnCount"]);
    await context.sync();
    targetStartRowIndex = writeTargetRange.rowIndex;
    targetStartColumnIndex = writeTargetRange.columnIndex;
    targetColumnCount = writeTargetRange.columnCount;
  }

  if (targetStartRowIndex === 0 || targetColumnCount < 1) {
    return null;
  }

  const totalScanRowCount = Math.min(BANNER_SCAN_ROW_COUNT, targetStartRowIndex);
  if (totalScanRowCount < 1) {
    return null;
  }

  // Widen the scan left to absorb stale-left marker cleanup when the source
  // selection extends further left than the actual write target (#272).
  const scanStartColIndex =
    sourceStartColIndex !== undefined && sourceStartColIndex < targetStartColumnIndex
      ? sourceStartColIndex
      : targetStartColumnIndex;
  const scanColCount = targetStartColumnIndex + targetColumnCount - scanStartColIndex;

  if (scanColCount < 1) {
    return null;
  }

  const bannerScanRange = writeTargetRange.worksheet.getRangeByIndexes(
    targetStartRowIndex - totalScanRowCount,
    scanStartColIndex,
    totalScanRowCount,
    scanColCount
  );
  const readSyncStartMs = perfNow();
  bannerScanRange.load("text");
  await context.sync();
  const readSyncEndMs = perfNow();
  let syncCount = 1;

  const bannerTexts = bannerScanRange.text;

  // Desired-text matrix.  Initialised from current texts, then mutated in
  // place by step 1 (strip RIT markers) and step 3 (place markers).
  const desiredTexts = new Array(totalScanRowCount);
  for (let r = 0; r < totalScanRowCount; r++) {
    const row = bannerTexts[r] || [];
    const destRow = new Array(scanColCount);
    for (let c = 0; c < scanColCount; c++) {
      destRow[c] = row[c] || "";
    }
    desiredTexts[r] = destRow;
  }

  // Step 1: strip RIT trailing markers from every cell in the scan area.
  // Keeps non-RIT parenthesised text (e.g. "Wave (quarter)") intact because
  // getTrailingBannerMarker only matches single-character significance labels.
  for (let r = 0; r < totalScanRowCount; r++) {
    const row = desiredTexts[r];
    for (let c = 0; c < scanColCount; c++) {
      const t = row[c];
      if (t && getTrailingBannerMarker(t)) {
        row[c] = removeTrailingBannerMarker(t);
      }
    }
  }

  // Step 2: compute the label for each data column.
  const useStructure = !!(calculationSettings.respectBannerStructure && bannerStructure);
  let labelMap;
  if (useStructure) {
    labelMap = buildBannerLocalSignificanceLabelMap(bannerStructure, calculationSettings);
  } else {
    const significanceLabels = generateSignificanceLabels({
      useCyrillicMarkers: Boolean(calculationSettings.useCyrillicMarkers),
      allowMultiCharacterMarkers: Boolean(calculationSettings.allowMultiCharacterMarkers),
      minimumCount: targetColumnCount,
    });
    labelMap = new Map();
    for (let dataCol = 0; dataCol < targetColumnCount; dataCol++) {
      if (calculationSettings.firstColumnIsTotal && dataCol === 0) {
        // Total column 0: step 1 above already stripped any RIT marker from
        // the lower row and the upper rows, so no additional label work is
        // required for this column.
        continue;
      }
      const markerIndex = calculationSettings.firstColumnIsTotal ? dataCol - 1 : dataCol;
      const marker = significanceLabels[markerIndex];
      if (marker) {
        labelMap.set(dataCol, marker);
      }
    }
  }

  // Step 3: place each marker into the lower banner row (immediately above
  // the data) or, if that cell is blank, the nearest non-empty cell above.
  const lowerRowIdx = totalScanRowCount - 1;
  const dataColOffsetInScan = targetStartColumnIndex - scanStartColIndex;
  const markedCellKeys = new Set();
  let cellsPlanned = 0;

  for (let dataCol = 0; dataCol < targetColumnCount; dataCol++) {
    const label = labelMap.get(dataCol);
    if (!label) continue;
    cellsPlanned++;

    const scanCol = dataColOffsetInScan + dataCol;

    // After step 1 the lower-cell text has any RIT marker stripped.  Use that
    // post-strip text to decide whether the cell is "blank" — matches the
    // legacy writers, which read the same pre-existing text and decided via
    // appendOrReplaceTrailingBannerMarker (which also operates on a stripped
    // basis when a marker is already present).
    const lowerStripped = desiredTexts[lowerRowIdx][scanCol] || "";
    if (lowerStripped !== "") {
      desiredTexts[lowerRowIdx][scanCol] = appendOrReplaceTrailingBannerMarker(lowerStripped, label);
      markedCellKeys.add(`${lowerRowIdx},${scanCol}`);
      continue;
    }

    // Lower banner cell is blank — walk upward to the nearest non-empty cell
    // and place the marker there.  This is the legacy "vertically merged /
    // multi-row banner" behaviour: when the lower row is empty because the
    // visible header lives in a higher row, the marker follows the visible
    // text.
    let placed = false;
    for (let r = lowerRowIdx - 1; r >= 0; r--) {
      const upperStripped = desiredTexts[r][scanCol] || "";
      if (upperStripped !== "") {
        desiredTexts[r][scanCol] = appendOrReplaceTrailingBannerMarker(upperStripped, label);
        markedCellKeys.add(`${r},${scanCol}`);
        placed = true;
        break;
      }
    }

    if (!placed) {
      // Fallback when no non-empty upper cell exists: write the bare marker
      // text into the lower banner row.  Both legacy writers reach this
      // branch — structure-mode via its explicit "fall back to the lower
      // banner row" arm, non-structure via fall-through past the upper-scan
      // loop into the bottom "queue write to lower banner row" branch.  The
      // @ number-format distinction between the two paths is preserved by
      // the use of `useStructure` later when batching writes.
      desiredTexts[lowerRowIdx][scanCol] = appendOrReplaceTrailingBannerMarker("", label);
      markedCellKeys.add(`${lowerRowIdx},${scanCol}`);
    }
  }

  // Step 4: diff against the original texts and queue only changed cells.
  const writesByRow = new Map();
  let cellsChanged = 0;
  let skippedNoOpWrites = 0;

  for (let r = 0; r < totalScanRowCount; r++) {
    for (let c = 0; c < scanColCount; c++) {
      const cur = (bannerTexts[r] && bannerTexts[r][c]) || "";
      const nxt = desiredTexts[r][c] || "";
      const key = `${r},${c}`;

      if (cur === nxt) {
        if (markedCellKeys.has(key)) {
          skippedNoOpWrites++;
        }
        continue;
      }

      const absRow = targetStartRowIndex - totalScanRowCount + r;
      const absCol = scanStartColIndex + c;
      if (!writesByRow.has(absRow)) writesByRow.set(absRow, []);
      const isMarkered = markedCellKeys.has(key);
      writesByRow.get(absRow).push({
        colIndex: absCol,
        text: nxt,
        markered: isMarkered,
        needsNF: isMarkered && needsNumberFormatForBannerMarker(nxt),
      });
      cellsChanged++;
    }
  }

  const changedRows = writesByRow.size;
  let maxChangedCellsInRow = 0;
  for (const items of writesByRow.values()) {
    if (items.length > maxChangedCellsInRow) {
      maxChangedCellsInRow = items.length;
    }
  }

  const queueWriteStartMs = perfNow();
  let writeCommands = 0;
  let changedCellRuns = 0;
  let oneCellWriteCommands = 0;
  let multiCellWriteCommands = 0;
  let markeredWriteCommands = 0;
  let clearOnlyWriteCommands = 0;
  let numberFormatCommands = 0;
  let numberFormatCells = 0;
  let valueWriteCells = 0;
  let markerRowsUsed = 0;
  let maxRunLength = 0;
  let numberFormatSyncMs = 0;
  let valueWriteSyncMs = 0;
  let profileSyncCount = 0;
  const writeProfileEnabled = perfBannerWriteProfileEnabled();

  const profiledValueWrites = [];
  const markerRowIndexes = new Set();
  let queueCommandEndMs = 0;
  if (cellsChanged > 0) {
    const worksheet = writeTargetRange.worksheet;
    for (const [rowIndex, items] of writesByRow) {
      items.sort((a, b) => a.colIndex - b.colIndex);
      let i = 0;
      while (i < items.length) {
        let j = i;
        // Group adjacent columns with the same markered and needsNF flags so
        // the "@" number format is applied only to bare-marker cells —
        // non-bare marker text like "Wave 1 (a)" does not need "@".
        while (
          j + 1 < items.length &&
          items[j + 1].colIndex === items[j].colIndex + 1 &&
          items[j + 1].markered === items[j].markered &&
          items[j + 1].needsNF === items[j].needsNF
        ) {
          j++;
        }
        const slice = items.slice(i, j + 1);
        const texts = slice.map((x) => x.text);
        const range = worksheet.getRangeByIndexes(rowIndex, items[i].colIndex, 1, j - i + 1);
        const runLength = texts.length;

        if (useStructure && slice[0].markered && slice[0].needsNF) {
          range.numberFormat = [texts.map(() => "@")];
          numberFormatCommands++;
          numberFormatCells += runLength;
        }

        if (writeProfileEnabled) {
          profiledValueWrites.push({ range, texts });
        } else {
          range.values = [texts];
        }

        writeCommands++;
        changedCellRuns++;
        valueWriteCells += runLength;
        if (runLength === 1) {
          oneCellWriteCommands++;
        } else {
          multiCellWriteCommands++;
        }
        if (slice[0].markered) {
          markeredWriteCommands++;
          markerRowIndexes.add(rowIndex);
        } else {
          clearOnlyWriteCommands++;
        }
        if (runLength > maxRunLength) {
          maxRunLength = runLength;
        }
        i = j + 1;
      }
    }
    markerRowsUsed = markerRowIndexes.size;
    queueCommandEndMs = perfNow();

    if (writeProfileEnabled) {
      if (numberFormatCommands > 0) {
        const numberFormatSyncStartMs = perfNow();
        await context.sync();
        numberFormatSyncMs = perfNow() - numberFormatSyncStartMs;
        syncCount++;
        profileSyncCount++;
      }

      const valueWriteSyncStartMs = perfNow();
      for (const pending of profiledValueWrites) {
        pending.range.values = [pending.texts];
      }
      await context.sync();
      valueWriteSyncMs = perfNow() - valueWriteSyncStartMs;
      syncCount++;
      profileSyncCount++;
    }
    // In normal mode, intentionally NO context.sync() here.  Caller's final
    // context.sync() flushes queued banner writes alongside other pending ops.
  }

  const queueWriteEndMs = queueCommandEndMs || perfNow();
  const details = {
    rowsScanned: totalScanRowCount,
    cellsRead: totalScanRowCount * scanColCount,
    cellsPlanned,
    cellsChanged,
    writeCommands,
    changedCellRuns,
    avgRunLength: changedCellRuns
      ? roundBannerDiagnosticRatio(cellsChanged / changedCellRuns)
      : 0,
    maxRunLength,
    oneCellWriteCommands,
    multiCellWriteCommands,
    markeredWriteCommands,
    clearOnlyWriteCommands,
    numberFormatCommands,
    changedRows,
    avgChangedCellsPerChangedRow: changedRows
      ? roundBannerDiagnosticRatio(cellsChanged / changedRows)
      : 0,
    maxChangedCellsInRow,
    skippedNoOpWrites,
    readSyncMs: readSyncEndMs - readSyncStartMs,
    planMs: queueWriteStartMs - readSyncEndMs,
    queueWriteMs: queueWriteEndMs - queueWriteStartMs,
    syncCount,
    writeProfileEnabled,
    numberFormatSyncMs,
    valueWriteSyncMs,
    profileSyncCount,
    numberFormatCells,
    valueWriteCells,
    markerRowsUsed,
  };

  Object.defineProperty(details, BANNER_SCAN_AREA_STATS, {
    value: {
      rowIndex: targetStartRowIndex - totalScanRowCount,
      columnIndex: scanStartColIndex,
      rowCount: totalScanRowCount,
      columnCount: scanColCount,
    },
    enumerable: false,
  });

  return details;
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

/**
 * Refreshes dependent settings UI state after user changes or saved settings load.
 */
function refreshSettingsPanelState() {
  refreshPreviousColumnComparisonState();
  refreshBannerStructureSettingsState();
  refreshTableFootnoteState();
  refreshDesignRecolorState();
  refreshActionWarnings();
}

/**
 * Applies UI rules for the "Перекрашивать баннер и лейблы" (design recolor)
 * option (issue #306).
 *
 * Recolor relies on the row-label columns being adjacent to the data body. When
 * "labels on left side" is enabled the labels are read from the far-left of the
 * sheet and are NOT adjacent to the data, so the recolor span could cross
 * unrelated columns. In that mode the recolor checkbox and color input are
 * disabled and a warning is shown. This is a UI affordance only —
 * readCalculationSettingsFromPanel and buildDesignRecolorJob also force the
 * feature off so a stale saved setting can never recolor while labelsOnLeftSide
 * is true. The checkbox value itself is preserved (not unchecked) so the user's
 * preference is restored if they turn labels-on-left-side off again.
 *
 * Action-area warnings are rendered separately so they can stay contextual to
 * the current Run/Autorun workspace.
 */
function refreshDesignRecolorState() {
  const recolorCheckbox = document.getElementById("recolor-banner-and-labels");
  const colorInput = document.getElementById("banner-label-fill-color");
  const labelsOnLeftSideCheckbox = document.getElementById("labels-on-left-side");
  const warningElement = document.getElementById("recolor-banner-and-labels-warning");

  if (!recolorCheckbox) {
    return;
  }

  const labelsOnLeftSide = labelsOnLeftSideCheckbox ? labelsOnLeftSideCheckbox.checked : false;

  recolorCheckbox.disabled = labelsOnLeftSide;
  if (colorInput) {
    colorInput.disabled = labelsOnLeftSide;
  }

  if (warningElement) {
    warningElement.style.display = labelsOnLeftSide ? "block" : "none";
  }
}

/**
 * Applies UI rules for the "Добавлять подпись под таблицей" (table footnote)
 * option.
 *
 * Footnotes merge across the table width including the adjacent left label
 * columns. When "labels on left side" is enabled the labels are read from the
 * far-left of the sheet and are NOT adjacent to the data, so a footnote cannot
 * be placed correctly. In that mode the footnote option is disabled and a
 * warning is shown. This is a UI affordance only — readCalculationSettingsFromPanel
 * and buildSignificanceFootnoteJob also force the feature off so a stale saved
 * setting can never create a footnote while labelsOnLeftSide is true.
 */
function refreshTableFootnoteState() {
  const addTableFootnoteCheckbox = document.getElementById("add-table-footnote");
  const labelsOnLeftSideCheckbox = document.getElementById("labels-on-left-side");
  const warningElement = document.getElementById("add-table-footnote-warning");

  if (!addTableFootnoteCheckbox) {
    return;
  }

  const labelsOnLeftSide = labelsOnLeftSideCheckbox ? labelsOnLeftSideCheckbox.checked : false;

  // Disable (but do not uncheck) so the user's preference is preserved and
  // restored if they turn labels-on-left-side off again.
  addTableFootnoteCheckbox.disabled = labelsOnLeftSide;

  if (warningElement) {
    warningElement.style.display = labelsOnLeftSide ? "block" : "none";
  }
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

  setStatusMessage(t("status.settingsReset"));
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

function getFirstBannerStructureError(bannerStructure) {
  if (!bannerStructure || !bannerStructure.messages) {
    return null;
  }

  return bannerStructure.messages.find((message) => message.severity === "error") || null;
}

/**
 * Returns true when a banner marker write needs numberFormat = "@".
 *
 * Only bare-marker cells risk Excel auto-formatting and need the guard.
 * A cell is "bare" when the only content is the marker itself, e.g. "(a)".
 * Cells with actual text like "Wave 1 (a)" are already text-safe without "@".
 */
function needsNumberFormatForBannerMarker(text) {
  return removeTrailingBannerMarker(text) === "";
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

  // Recognise any token RIT may have written as a marker — single characters
  // from any historical alphabet, or multi-character overflow markers — so
  // banner markers are cleaned up on re-runs regardless of the current
  // Cyrillic/overflow settings.
  if (!isSignificanceMarkerLabel(markerLabel)) {
    return null;
  }

  return {
    label: markerLabel,
    start: markerMatch.index,
  };
}
