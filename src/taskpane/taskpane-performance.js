// Development-only performance instrumentation for RIT taskpane flows.
//
// Enable from DevTools with either:
//   window.__RIT_PERF = true
//   window.__RIT_PERF = "1"
//   localStorage.setItem('RIT_PERF', '1')
//
// Optional banner write host-flush profiling:
//   window.__RIT_BANNER_WRITE_PROFILE = true
//   localStorage.setItem('RIT_BANNER_WRITE_PROFILE', '1')
//
// This intentionally changes banner write flushing while enabled, splitting
// marker numberFormat and values into separate syncs so their host cost can be
// measured. It is disabled unless RIT_PERF is also enabled.
//
// Disable with either:
//   window.__RIT_PERF = false
//   localStorage.removeItem('RIT_PERF')
//
// The runtime window/globalThis flag takes priority over storage. This lets
// DevTools disable logging for the current taskpane session even if a stored
// RIT_PERF value exists.
//
// The flag is read dynamically on every call, so no taskpane reload is needed.
//
// When enabled, each instrumented flow emits a single console.info entry:
//   [RIT perf] <flowName>  { phase1Ms, phase2Ms, ..., totalMs }
//
// All exported functions are no-ops when disabled, adding zero overhead to
// production runs.

export const BANNER_SCAN_AREA_STATS = Symbol("bannerScanAreaStats");
const BANNER_AGGREGATE_STATE = Symbol("bannerAggregateState");

function _readPerfRuntimeFlag() {
  try {
    if (typeof globalThis === "undefined") return undefined;
    return globalThis.__RIT_PERF;
  } catch (_) {
    return undefined;
  }
}

function _perfRuntimeFlagEnabled() {
  const runtimeFlag = _readPerfRuntimeFlag();

  if (runtimeFlag === true || runtimeFlag === "1") return true;
  if (runtimeFlag === false || runtimeFlag === "0") return false;

  return undefined;
}

function _perfStorageFlagEnabled() {
  try {
    return typeof localStorage !== "undefined" && localStorage.getItem("RIT_PERF") === "1";
  } catch (_) {
    return false;
  }
}

function _perfEnabled() {
  const runtimeEnabled = _perfRuntimeFlagEnabled();

  if (runtimeEnabled !== undefined) return runtimeEnabled;

  return _perfStorageFlagEnabled();
}

function _bannerWriteProfileRuntimeFlagEnabled() {
  try {
    if (typeof globalThis === "undefined") return undefined;
    const runtimeFlag = globalThis.__RIT_BANNER_WRITE_PROFILE;
    if (runtimeFlag === true || runtimeFlag === "1") return true;
    if (runtimeFlag === false || runtimeFlag === "0") return false;
  } catch (_) {
    return undefined;
  }

  return undefined;
}

function _bannerWriteProfileStorageFlagEnabled() {
  try {
    return typeof localStorage !== "undefined" && localStorage.getItem("RIT_BANNER_WRITE_PROFILE") === "1";
  } catch (_) {
    return false;
  }
}

export function perfEnabled() {
  return _perfEnabled();
}

export function perfBannerWriteProfileEnabled() {
  if (!_perfEnabled()) return false;

  const runtimeEnabled = _bannerWriteProfileRuntimeFlagEnabled();
  if (runtimeEnabled !== undefined) return runtimeEnabled;

  return _bannerWriteProfileStorageFlagEnabled();
}

// Returns Date.now() when enabled, 0 otherwise.
export function perfNow() {
  return _perfEnabled() ? Date.now() : 0;
}

// Returns milliseconds elapsed since startMs when enabled, 0 otherwise.
export function perfElapsed(startMs) {
  return _perfEnabled() && startMs ? Date.now() - startMs : 0;
}

// Emits a console.info entry when enabled. No-op otherwise.
export function perfLog(flowName, phases) {
  if (!_perfEnabled()) return;
  console.info("[RIT perf]", flowName, phases);
}

export function createAggregatedWriterPerfDetails() {
  return {
    tablesWithWriterDetails: 0,
    buildMatricesMs: 0,
    numberFormatWriteMs: 0,
    valuesWriteMs: 0,
    boldFormatMs: 0,
    fillFormatMs: 0,
    boldCommandCount: 0,
    fillCommandCount: 0,
    boldRectCommandCountEstimate: 0,
    fillRectCommandCountEstimate: 0,
    maxBoldCommandCount: 0,
    maxFillCommandCount: 0,
    // #286 — actual formatting path taken by writer per table. null until the
    // first table sets it. Stays as a single path label when every table used
    // the same branch; becomes "mixed" if both branches appeared.
    formattingPath: null,
    boldRangeAreasAppliedCommandEstimate: 0,
    fillRangeAreasAppliedCommandEstimate: 0,
    // Spike #284 — diagnostic RangeAreas projection rolled up across tables.
    // chunkCount / commandCountEstimate / areaCountTotal sum across tables so
    // the workbook log shows projected total queued ops for a chunked
    // `worksheet.getRanges(...)` writer. maxAddressLength is the per-table max
    // so a single oversized chunk is still visible. fillRangeAreas.colorCountMax
    // is the max number of distinct fill colors observed on any one table —
    // summing colors across tables has no meaningful interpretation because
    // different tables typically share the same 1–2 colors. When #286's
    // RangeAreas writer path is active, these projections match the actual
    // queued ops the writer issued.
    boldRangeAreas: {
      areaCountTotal: 0,
      chunkCount: 0,
      commandCountEstimate: 0,
      maxAddressLength: 0,
    },
    fillRangeAreas: {
      areaCountTotal: 0,
      chunkCount: 0,
      commandCountEstimate: 0,
      maxAddressLength: 0,
      colorCountMax: 0,
    },
  };
}

export function mergeWriterPerfDetails(aggregate, writerDetails) {
  if (!aggregate || !writerDetails) {
    return;
  }

  aggregate.tablesWithWriterDetails += 1;
  aggregate.buildMatricesMs += writerDetails.buildMatricesMs || 0;
  aggregate.numberFormatWriteMs += writerDetails.numberFormatWriteMs || 0;
  aggregate.valuesWriteMs += writerDetails.valuesWriteMs || 0;
  aggregate.boldFormatMs += writerDetails.boldFormatMs || 0;
  aggregate.fillFormatMs += writerDetails.fillFormatMs || 0;
  aggregate.boldCommandCount += writerDetails.boldCommandCount || 0;
  aggregate.fillCommandCount += writerDetails.fillCommandCount || 0;
  aggregate.boldRectCommandCountEstimate += writerDetails.boldRectCommandCountEstimate || 0;
  aggregate.fillRectCommandCountEstimate += writerDetails.fillRectCommandCountEstimate || 0;
  aggregate.maxBoldCommandCount = Math.max(
    aggregate.maxBoldCommandCount,
    writerDetails.boldCommandCount || 0
  );
  aggregate.maxFillCommandCount = Math.max(
    aggregate.maxFillCommandCount,
    writerDetails.fillCommandCount || 0
  );

  // #286 — fold formattingPath across tables. Skip when null (markers-only
  // mode never set a path).
  if (writerDetails.formattingPath) {
    if (aggregate.formattingPath === null) {
      aggregate.formattingPath = writerDetails.formattingPath;
    } else if (aggregate.formattingPath !== writerDetails.formattingPath) {
      aggregate.formattingPath = "mixed";
    }
  }

  aggregate.boldRangeAreasAppliedCommandEstimate +=
    writerDetails.boldRangeAreasAppliedCommandEstimate || 0;
  aggregate.fillRangeAreasAppliedCommandEstimate +=
    writerDetails.fillRangeAreasAppliedCommandEstimate || 0;

  mergeRangeAreasProjection(aggregate.boldRangeAreas, writerDetails.boldRangeAreas);
  mergeFillRangeAreasProjection(aggregate.fillRangeAreas, writerDetails.fillRangeAreas);
}

export function mergeRangeAreasProjection(target, source) {
  if (!target || !source) {
    return;
  }

  target.areaCountTotal += source.areaCountTotal || 0;
  target.chunkCount += source.chunkCount || 0;
  target.commandCountEstimate += source.commandCountEstimate || 0;
  if ((source.maxAddressLength || 0) > target.maxAddressLength) {
    target.maxAddressLength = source.maxAddressLength;
  }
}

export function mergeFillRangeAreasProjection(target, source) {
  if (!target || !source) {
    return;
  }

  mergeRangeAreasProjection(target, source);

  if ((source.colorCount || 0) > target.colorCountMax) {
    target.colorCountMax = source.colorCount;
  }
}

export function createAggregatedBannerPerfDetails() {
  const details = {
    tablesWithBannerDetails: 0,
    rowsScanned: 0,
    cellsRead: 0,
    cellsPlanned: 0,
    cellsChanged: 0,
    writeCommands: 0,
    changedCellRuns: 0,
    avgRunLength: 0,
    maxRunLength: 0,
    oneCellWriteCommands: 0,
    multiCellWriteCommands: 0,
    markeredWriteCommands: 0,
    clearOnlyWriteCommands: 0,
    numberFormatCommands: 0,
    changedRows: 0,
    avgChangedCellsPerChangedRow: 0,
    maxChangedCellsInRow: 0,
    maxChangedCellsPerTable: 0,
    maxChangedRowsPerTable: 0,
    skippedNoOpWrites: 0,
    readSyncMs: 0,
    planMs: 0,
    queueWriteMs: 0,
    syncCount: 0,
    writeProfiledTables: 0,
    numberFormatSyncMs: 0,
    valueWriteSyncMs: 0,
    profileSyncCount: 0,
    numberFormatCells: 0,
    valueWriteCells: 0,
    markerRowsUsed: 0,
    maxWriteCommands: 0,
    tablesWithOverlappingBannerAreas: 0,
    perSheetBannerTablesMax: 0,
  };

  Object.defineProperty(details, BANNER_AGGREGATE_STATE, {
    value: {
      areasBySheet: new Map(),
      overlappingTableIds: new Set(),
      nextTableId: 1,
    },
    enumerable: false,
  });

  return details;
}

export function roundBannerDiagnosticRatio(value) {
  return Math.round(value * 100) / 100;
}

export function bannerScanAreasOverlap(a, b) {
  if (!a || !b) {
    return false;
  }

  const aRowEnd = a.rowIndex + a.rowCount;
  const bRowEnd = b.rowIndex + b.rowCount;
  const aColEnd = a.columnIndex + a.columnCount;
  const bColEnd = b.columnIndex + b.columnCount;

  return a.rowIndex < bRowEnd && b.rowIndex < aRowEnd && a.columnIndex < bColEnd && b.columnIndex < aColEnd;
}

export function mergeBannerScanAreaDiagnostics(aggregate, bannerDetails, sheetName) {
  const state = aggregate ? aggregate[BANNER_AGGREGATE_STATE] : null;
  const area = bannerDetails ? bannerDetails[BANNER_SCAN_AREA_STATS] : null;
  if (!state || !area || !sheetName) {
    return;
  }

  const tableId = state.nextTableId++;
  const sheetAreas = state.areasBySheet.get(sheetName) || [];
  let overlapsExistingArea = false;

  for (const existing of sheetAreas) {
    if (bannerScanAreasOverlap(area, existing.area)) {
      overlapsExistingArea = true;
      state.overlappingTableIds.add(existing.tableId);
    }
  }

  if (overlapsExistingArea) {
    state.overlappingTableIds.add(tableId);
  }

  sheetAreas.push({ tableId, area });
  state.areasBySheet.set(sheetName, sheetAreas);
  aggregate.tablesWithOverlappingBannerAreas = state.overlappingTableIds.size;
  aggregate.perSheetBannerTablesMax = Math.max(aggregate.perSheetBannerTablesMax, sheetAreas.length);
}

export function mergeBannerPerfDetails(aggregate, bannerDetails, sheetName) {
  if (!aggregate || !bannerDetails) {
    return;
  }

  aggregate.tablesWithBannerDetails += 1;
  aggregate.rowsScanned += bannerDetails.rowsScanned || 0;
  aggregate.cellsRead += bannerDetails.cellsRead || 0;
  aggregate.cellsPlanned += bannerDetails.cellsPlanned || 0;
  aggregate.cellsChanged += bannerDetails.cellsChanged || 0;
  aggregate.writeCommands += bannerDetails.writeCommands || 0;
  aggregate.changedCellRuns += bannerDetails.changedCellRuns || 0;
  aggregate.oneCellWriteCommands += bannerDetails.oneCellWriteCommands || 0;
  aggregate.multiCellWriteCommands += bannerDetails.multiCellWriteCommands || 0;
  aggregate.markeredWriteCommands += bannerDetails.markeredWriteCommands || 0;
  aggregate.clearOnlyWriteCommands += bannerDetails.clearOnlyWriteCommands || 0;
  aggregate.numberFormatCommands += bannerDetails.numberFormatCommands || 0;
  aggregate.changedRows += bannerDetails.changedRows || 0;
  aggregate.skippedNoOpWrites += bannerDetails.skippedNoOpWrites || 0;
  aggregate.readSyncMs += bannerDetails.readSyncMs || 0;
  aggregate.planMs += bannerDetails.planMs || 0;
  aggregate.queueWriteMs += bannerDetails.queueWriteMs || 0;
  aggregate.syncCount += bannerDetails.syncCount || 0;
  aggregate.writeProfiledTables += bannerDetails.writeProfileEnabled ? 1 : 0;
  aggregate.numberFormatSyncMs += bannerDetails.numberFormatSyncMs || 0;
  aggregate.valueWriteSyncMs += bannerDetails.valueWriteSyncMs || 0;
  aggregate.profileSyncCount += bannerDetails.profileSyncCount || 0;
  aggregate.numberFormatCells += bannerDetails.numberFormatCells || 0;
  aggregate.valueWriteCells += bannerDetails.valueWriteCells || 0;
  aggregate.markerRowsUsed += bannerDetails.markerRowsUsed || 0;
  if ((bannerDetails.writeCommands || 0) > aggregate.maxWriteCommands) {
    aggregate.maxWriteCommands = bannerDetails.writeCommands;
  }
  if ((bannerDetails.maxRunLength || 0) > aggregate.maxRunLength) {
    aggregate.maxRunLength = bannerDetails.maxRunLength;
  }
  if ((bannerDetails.maxChangedCellsInRow || 0) > aggregate.maxChangedCellsInRow) {
    aggregate.maxChangedCellsInRow = bannerDetails.maxChangedCellsInRow;
  }
  if ((bannerDetails.cellsChanged || 0) > aggregate.maxChangedCellsPerTable) {
    aggregate.maxChangedCellsPerTable = bannerDetails.cellsChanged;
  }
  if ((bannerDetails.changedRows || 0) > aggregate.maxChangedRowsPerTable) {
    aggregate.maxChangedRowsPerTable = bannerDetails.changedRows;
  }
  aggregate.avgRunLength = aggregate.changedCellRuns
    ? roundBannerDiagnosticRatio(aggregate.cellsChanged / aggregate.changedCellRuns)
    : 0;
  aggregate.avgChangedCellsPerChangedRow = aggregate.changedRows
    ? roundBannerDiagnosticRatio(aggregate.cellsChanged / aggregate.changedRows)
    : 0;

  mergeBannerScanAreaDiagnostics(aggregate, bannerDetails, sheetName);
}
