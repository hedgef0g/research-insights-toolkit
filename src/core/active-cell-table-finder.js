/**
 * Pure helper for active-cell current-table resolution.
 *
 * Takes pre-loaded scanner inventory items and an active cell position
 * (zero-based absolute row/col indices) and finds which candidate table
 * contains that cell.
 *
 * Office.js-free: all inputs are plain values. Designed so that the
 * Office.js-dependent caller can load data once and pass it in.
 *
 * Used by resolveCurrentTableFromActiveCell() in
 * src/taskpane/active-cell-resolver.js.
 * NOT wired into production Check or Autorun flows yet.
 */

// ─── A1 address parsing ───────────────────────────────────────────────────────

/**
 * Converts a column letter string (e.g. "A", "Z", "AA") to a zero-based index.
 * Caller must pass an already-uppercased string.
 */
function columnLetterToIndex(letters) {
  let result = 0;
  for (let i = 0; i < letters.length; i++) {
    result = result * 26 + (letters.charCodeAt(i) - 64);
  }
  return result - 1;
}

/**
 * Parses an A1-notation range address into zero-based row/col bounds.
 *
 * Accepts:
 *   "B3:H15"  → { startRow: 2, endRow: 14, startCol: 1, endCol: 7 }
 *   "A1"      → { startRow: 0, endRow: 0, startCol: 0, endCol: 0 }
 *
 * Returns null when the address does not match the expected pattern.
 *
 * @param {string} address
 * @returns {{ startRow: number, endRow: number, startCol: number, endCol: number } | null}
 */
export function parseA1Range(address) {
  if (typeof address !== "string") return null;
  const match = /^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/.exec(address.trim());
  if (!match) return null;
  const startCol = columnLetterToIndex(match[1].toUpperCase());
  const startRow = parseInt(match[2], 10) - 1;
  const endCol = match[3] ? columnLetterToIndex(match[3].toUpperCase()) : startCol;
  const endRow = match[4] ? parseInt(match[4], 10) - 1 : startRow;
  if (startRow < 0 || startCol < 0 || endRow < startRow || endCol < startCol) return null;
  return { startRow, endRow, startCol, endCol };
}

// ─── Candidate slice extraction ──────────────────────────────────────────────

/**
 * Extracts a 2D values slice from usedRange.values that covers the candidate's
 * rangeAddress bounds.
 *
 * The candidate's rangeAddress uses absolute A1 sheet coordinates, as produced
 * by scanWorksheetForTables. usedRangeRowOffset and usedRangeColOffset are the
 * absolute sheet position of the top-left cell of usedRange.values (i.e.
 * usedRange.rowIndex and usedRange.columnIndex from Office.js).
 *
 * Returns null when:
 *   - usedRangeValues is not a non-empty array;
 *   - candidateRangeAddress cannot be parsed;
 *   - the candidate's start position falls outside usedRange.values.
 *
 * When the candidate extends beyond the usedRange edge it is clamped silently,
 * because the scanner already operates on usedRange data and out-of-range
 * indices would only arise from rounding / edge-row handling.
 *
 * @param {Array}  usedRangeValues       - 2D array from usedRange.values
 * @param {number} usedRangeRowOffset    - usedRange.rowIndex (zero-based absolute)
 * @param {number} usedRangeColOffset    - usedRange.columnIndex (zero-based absolute)
 * @param {string} candidateRangeAddress - A1 notation (absolute sheet coordinates)
 * @returns {Array|null}
 */
export function extractCandidateSlice(
  usedRangeValues,
  usedRangeRowOffset,
  usedRangeColOffset,
  candidateRangeAddress
) {
  if (!Array.isArray(usedRangeValues) || usedRangeValues.length === 0) return null;
  const bounds = parseA1Range(candidateRangeAddress);
  if (!bounds) return null;

  const rowCount = usedRangeValues.length;
  const colCount = Array.isArray(usedRangeValues[0]) ? usedRangeValues[0].length : 0;

  const relStartRow = bounds.startRow - usedRangeRowOffset;
  const relEndRow   = bounds.endRow   - usedRangeRowOffset;
  const relStartCol = bounds.startCol - usedRangeColOffset;
  const relEndCol   = bounds.endCol   - usedRangeColOffset;

  // Reject when the candidate starts before or beyond the usedRange.
  if (relStartRow < 0 || relStartRow >= rowCount) return null;
  if (relStartCol < 0 || relStartCol >= colCount) return null;
  if (relEndRow < relStartRow || relEndCol < relStartCol) return null;

  const endRow = Math.min(relEndRow, rowCount - 1);
  const endCol = Math.min(relEndCol, colCount - 1);

  return usedRangeValues
    .slice(relStartRow, endRow + 1)
    .map((row) => (Array.isArray(row) ? row.slice(relStartCol, endCol + 1) : []));
}

// ─── Candidate containment check ─────────────────────────────────────────────

/**
 * Finds which scanned inventory candidate(s) contain the given active cell.
 *
 * The active cell is the anchor (top-left corner) of the Excel selection:
 * use selectedRange.rowIndex and selectedRange.columnIndex from Office.js,
 * both of which are zero-based absolute sheet coordinates.
 *
 * Candidate range addresses are absolute A1 notation as produced by
 * scanWorksheetForTables() — they already incorporate the usedRangeRowOffset
 * and usedRangeColOffset, so no further adjustment is needed.
 *
 * Items with an unparseable rangeAddress are silently skipped.
 *
 * @param {Array}  items           - TableInventoryItem[] from scanWorksheetForTables
 * @param {number} activeCellRow   - Zero-based absolute sheet row index
 * @param {number} activeCellCol   - Zero-based absolute sheet column index
 * @returns {object} Result object:
 *   { status: "found",     candidate }      — exactly one match
 *   { status: "ambiguous", candidates }     — more than one match (overlapping bands)
 *   { status: "no-table",  candidates: [] } — no match
 */
export function findCandidateForActiveCell(items, activeCellRow, activeCellCol) {
  if (!Array.isArray(items)) {
    return { status: "no-table", candidates: [] };
  }

  const matching = [];
  for (const item of items) {
    const bounds = parseA1Range(item.rangeAddress);
    if (!bounds) continue;
    if (
      activeCellRow >= bounds.startRow &&
      activeCellRow <= bounds.endRow &&
      activeCellCol >= bounds.startCol &&
      activeCellCol <= bounds.endCol
    ) {
      matching.push(item);
    }
  }

  if (matching.length === 0) {
    return { status: "no-table", candidates: [] };
  }
  if (matching.length > 1) {
    return { status: "ambiguous", candidates: matching };
  }
  return { status: "found", candidate: matching[0] };
}
