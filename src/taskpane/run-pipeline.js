import {
  createEmptyCellResultMatrix,
  applyComparisonResultsToFullCellResultMatrix,
  applySmallBaseRulesForCalculationBlock,
  keepMarkersOnlyInAllowedRows,
  detectSignificanceMarkerOverflow,
  compareProportionRowsUsingBaseRow,
  compareMeanBlockByRowIndexes,
  compareNpsStructureBlockByRowIndexes,
  compareNpsSpreadBlockByRowIndexes,
} from "../core/significance";

import {
  detectMetricRowsFromLeftLabels,
  buildCalculationBlocks,
  getAllowedMarkerRowIndexes,
} from "../core/metric-detector";

import { writeCellResultsToSelectedRange } from "../core/excel-writer";

import { detectBannerStructure } from "../core/banner-detector";

import { buildDesignRecolorJob } from "../core/design-recolor";

import { interpretSelectedRange } from "./selected-range-interpreter";

import { perfNow, perfEnabled } from "./taskpane-performance";

function requireRunPipelineDependency(dependencies, name) {
  const dependency = dependencies[name];
  if (typeof dependency !== "function") {
    throw new Error(`run-pipeline dependency missing: ${name}`);
  }
  return dependency;
}

function resolveBannerRecolorRowCount(interpretation, dataStartRowIndex) {
  const interpreted = interpretation.bannerRowsAboveData || 0;
  if (interpreted > 0) return interpreted;
  if (interpretation.state === "passThrough" && dataStartRowIndex > 0) return 1;
  return 0;
}

/**
 * Calculates one detected metric block.
 *
 * PURPOSE:
 * Keeps dispatcher logic close to the per-table Run executor.
 * Each block type is routed to the correct core calculation function.
 */
export function calculateBlockResults(cleanedValues, calculationBlock, calculationSettings) {
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

export function getFirstBannerStructureError(bannerStructure) {
  if (!bannerStructure || !bannerStructure.messages) {
    return null;
  }

  return bannerStructure.messages.find((message) => message.severity === "error") || null;
}

/**
 * Core significance pipeline for a single named range, using a caller-supplied
 * Office.js RequestContext.
 *
 * Returns { status, blocksProcessed, message, rangeAddress }.
 * status: "processed" | "skipped" | "blocked" | "stopped" | "error"
 */
export async function runSignificanceForRangeInContext(
  context,
  sheetName,
  rangeAddress,
  calculationSettings,
  markerOverflowDecider = null,
  dependencies = {}
) {
  const applyBannerMarkerUpdatesForRange = requireRunPipelineDependency(
    dependencies,
    "applyBannerMarkerUpdatesForRange"
  );
  const buildSignificanceFootnoteJob = requireRunPipelineDependency(
    dependencies,
    "buildSignificanceFootnoteJob"
  );
  const createMarkerOverflowDecider = requireRunPipelineDependency(
    dependencies,
    "createMarkerOverflowDecider"
  );

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
