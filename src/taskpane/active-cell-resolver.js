/**
 * Active-cell current-table resolver for Research Insights Toolkit.
 *
 * Provides resolveCurrentTableFromActiveCell(context, settings), which reads
 * the active cell position from Office.js and returns a structured result
 * describing which scanned candidate table (if any) contains that cell.
 *
 * PURPOSE:
 * Read-only foundation for the future "Текущая таблица" scope in Автозапуск
 * and Проверка. This resolver is NOT wired into any production flow yet.
 * A later PR will connect Проверка → Текущая таблица first.
 *
 * ACTIVE CELL CONVENTION:
 * For multi-cell selections, Office.js reports the anchor (top-left) cell via
 * selectedRange.rowIndex / selectedRange.columnIndex. This resolver always
 * uses the anchor — consistent with the principle that "Текущая таблица"
 * means table containing the cursor, not a table matching the selected range.
 *
 * GENERATED SHEETS:
 * "Content" and "Run report" are created by RIT and contain no research
 * tables. If the active sheet is one of these, the resolver returns
 * { status: "generated-sheet" } immediately without scanning.
 *
 * SETTINGS:
 * Pass the same settings object as used elsewhere in the pipeline.
 * If the caller has a backlinkMarker available, include it as
 * settings.backlinkMarker so the scanner can exclude backlink rows.
 *
 * READ-ONLY:
 * This module never writes to the workbook.
 */

import { scanWorksheetForTables } from "../core/table-inventory-scanner";
import { findCandidateForActiveCell, extractCandidateSlice } from "../core/active-cell-table-finder";
import { hasEmptyDataRowGap } from "../core/range-normalizer";
import { INVENTORY_SCAN_EMERGENCY_CELL_LIMIT } from "./taskpane-inventory-scan";

const RESOLVER_GENERATED_SHEET_NAMES = new Set(["Content", "Run report"]);

function shouldBlockResolverScan(
  rowCount,
  columnCount,
  limit = INVENTORY_SCAN_EMERGENCY_CELL_LIMIT
) {
  return rowCount * columnCount > limit;
}

function buildResolverScanBlockedResult(sheetName, rowCount, columnCount) {
  const cellCount = rowCount * columnCount;

  return {
    status: "blocked",
    sheetName,
    message: `Лист аварийно пропущен — UsedRange слишком большой (${rowCount}×${columnCount} = ${cellCount} ячеек, аварийный лимит: ${INVENTORY_SCAN_EMERGENCY_CELL_LIMIT}).`,
    details: {
      rowCount,
      columnCount,
      cellCount,
      limit: INVENTORY_SCAN_EMERGENCY_CELL_LIMIT,
    },
  };
}

/**
 * Resolves which inventory candidate contains the active cell.
 *
 * Must be called inside an Excel.run callback: the caller passes the
 * context object, and this function performs up to three context.sync()
 * calls internally.
 *
 * @param {Excel.RequestContext} context
 * @param {object} settings  - Calculation settings (same shape as the rest of the pipeline).
 * @returns {Promise<object>} Structured result:
 *
 *   OK:
 *   { status: "ok", sheetName, rangeAddress, candidateMeta, message }
 *
 *   Non-OK:
 *   { status: "no-table" | "generated-sheet" | "ambiguous-boundary" | "blocked" | "error",
 *     sheetName, message, details }
 */
async function resolveCurrentTableFromActiveCell(context, settings) {
  let sheetName = null;

  try {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const selectedRange = context.workbook.getSelectedRange();

    worksheet.load("name");
    selectedRange.load(["rowIndex", "columnIndex"]);

    await context.sync();

    sheetName = worksheet.name;

    if (RESOLVER_GENERATED_SHEET_NAMES.has(sheetName)) {
      return {
        status: "generated-sheet",
        sheetName,
        message: `Лист «${sheetName}» создан надстройкой и не содержит исследовательских таблиц.`,
        details: null,
      };
    }

    // Anchor cell of the selection (top-left, zero-based absolute sheet coordinates).
    const activeCellRow = selectedRange.rowIndex;
    const activeCellCol = selectedRange.columnIndex;

    const usedRange = worksheet.getUsedRangeOrNullObject();
    usedRange.load(["isNullObject", "rowIndex", "columnIndex", "rowCount", "columnCount"]);

    await context.sync();

    if (usedRange.isNullObject) {
      return {
        status: "no-table",
        sheetName,
        message: "Лист пуст — таблиц не обнаружено.",
        details: null,
      };
    }

    if (shouldBlockResolverScan(usedRange.rowCount, usedRange.columnCount)) {
      return buildResolverScanBlockedResult(
        sheetName,
        usedRange.rowCount,
        usedRange.columnCount
      );
    }

    usedRange.load("values");

    await context.sync();

    const items = scanWorksheetForTables({
      values: usedRange.values,
      usedRangeRowOffset: usedRange.rowIndex,
      usedRangeColOffset: usedRange.columnIndex,
      sheetName,
      settings,
    });

    if (items.length === 0) {
      return {
        status: "no-table",
        sheetName,
        message: "На листе не обнаружено кандидатов в исследовательские таблицы.",
        details: null,
      };
    }

    const findResult = findCandidateForActiveCell(items, activeCellRow, activeCellCol);

    if (findResult.status === "no-table") {
      return {
        status: "no-table",
        sheetName,
        message:
          "Активная ячейка не находится внутри ни одного кандидата. " +
          "Перейдите в ячейку внутри таблицы.",
        details: {
          activeCellRow,
          activeCellCol,
          candidatesOnSheet: items.length,
        },
      };
    }

    if (findResult.status === "ambiguous") {
      return {
        status: "ambiguous-boundary",
        sheetName,
        message:
          "Активная ячейка входит в несколько перекрывающихся кандидатов. " +
          "Уточните позицию курсора.",
        details: {
          activeCellRow,
          activeCellCol,
          candidateAddresses: findResult.candidates.map((c) => c.rangeAddress),
        },
      };
    }

    const candidate = findResult.candidate;

    // Guard: before committing to "ok", extract the candidate slice from the
    // already-loaded usedRange.values and inspect it for all-empty row gaps.
    // This must happen at the resolver layer — before checkSelectedRangePreview
    // is called — because a candidate whose rangeAddress contains multiple
    // table-like blocks can still lead to a normal one-table diagnostic once the
    // later normalization pipeline reduces it to a single body.
    const candidateSlice = extractCandidateSlice(
      usedRange.values,
      usedRange.rowIndex,
      usedRange.columnIndex,
      candidate.rangeAddress
    );
    if (candidateSlice !== null && hasEmptyDataRowGap(candidateSlice)) {
      return {
        status: "ambiguous-boundary",
        sheetName,
        message:
          "В диапазоне кандидата обнаружено несколько блоков данных, разделённых пустыми строками. " +
          "Перейдите в ячейку внутри одной таблицы или используйте «Проверить лист».",
        details: {
          activeCellRow,
          activeCellCol,
          candidateRangeAddress: candidate.rangeAddress,
        },
      };
    }

    return {
      status: "ok",
      sheetName,
      rangeAddress: candidate.rangeAddress,
      candidateMeta: candidate,
      message: `Таблица найдена: ${sheetName}!${candidate.rangeAddress}.`,
    };
  } catch (err) {
    return {
      status: "error",
      sheetName,
      message: `Ошибка при определении текущей таблицы: ${err && err.message ? err.message : String(err)}`,
      details: null,
    };
  }
}

export { resolveCurrentTableFromActiveCell, shouldBlockResolverScan };
