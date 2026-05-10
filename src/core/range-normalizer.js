/**
 * Selected range normalizer for Research Insights Toolkit.
 *
 * Analyzes a 2D grid (already loaded from Excel, markers already removed)
 * and returns a NormalizedSelectionModel that partitions the selection into:
 *   - title / subtitle rows   (stripped from calculation)
 *   - banner / header rows    (extracted into bannerContext)
 *   - row label columns       (extracted into leftLabelValues)
 *   - numeric data body       (returned as valuesForCalculation)
 *
 * Three output states:
 *
 *   1. normalizationNeeded: false, normalizationApplied: false
 *      The selection already looks like numeric data only.
 *      Future integrations MUST preserve the existing strict workflow unchanged.
 *
 *   2. normalizationNeeded: true, normalizationApplied: true
 *      Decomposition succeeded.
 *      Future integrations MAY use valuesForCalculation, leftLabelValues,
 *      bannerContext, dataRowOffset, and dataColOffset.
 *
 *   3. normalizationNeeded: true, normalizationApplied: false
 *      The selection appeared broad / full-table-like, but the normalizer
 *      could not safely decompose it.
 *      Future integrations MUST stop and show blockingMessage.
 *      They MUST NOT silently fall back to running the strict workflow on
 *      the original broad selection — that can write markers to wrong cells.
 *
 * Pure module constraints:
 *   - No Office.js
 *   - No Excel writes
 *   - No significance calculation
 *   - No taskpane DOM access
 *   - No imports from other RIT modules
 */

// ─── Detection thresholds ─────────────────────────────────────────────────────

// Minimum selection size for structural analysis.
// Smaller selections are always treated as pass-through.
const GATE_MIN_ROWS = 3;
const GATE_MIN_COLS = 3;

// Pass-through gate: left-column pattern.
// First column is text-heavy AND the remaining columns are numeric-heavy.
const GATE_LEFT_COL_TEXT_FRACTION    = 0.7;
// Lowered from 0.7: banner rows that contain "2025Q4"-style text labels dilute
// the numeric fraction of the rest of the grid even when the data body is purely
// numeric.  0.35 allows detection when ~4 banner rows precede ~3 data rows.
const GATE_REST_COLS_NUMERIC_FRACTION = 0.35;

// Pass-through gate: top-row pattern.
// First row is text-heavy AND the lower rows are numeric-heavy.
const GATE_TOP_ROW_TEXT_FRACTION      = 0.6;
// Lowered from 0.7 for the same banner-dilution reason as above.
const GATE_LOWER_ROWS_NUMERIC_FRACTION = 0.3;

// Pass-through gate: overall text-heavy fallback.
const GATE_OVERALL_TEXT_FRACTION    = 0.25;
// Lowered from 0.5: when both label columns and banner rows are present, the
// overall numeric fraction of the whole grid can fall below 0.5 even for a
// structurally valid research table.
const GATE_OVERALL_NUMERIC_FRACTION = 0.3;
const GATE_OVERALL_MIN_CELLS        = 12;

// Banner row count safety cap.  More than this many rows classified as banner
// before the first body row almost certainly means the selection spans multiple
// tables (or includes unrelated content), not a single wide banner header.
// Real survey tables have at most 4–5 banner rows; the threshold is set to 6.
const MAX_BANNER_ROW_COUNT = 6;

// Title row: sparse, no numeric cells, at least one text cell.
const MAX_TITLE_ROW_NON_EMPTY_CELLS = 3;

// Banner row: wide (spans ≥50% of columns) and text-heavy.
const BANNER_MIN_FILL_FRACTION     = 0.5;
const BANNER_TEXT_FRACTION         = 0.6;

// Label column: text fraction threshold for "clearly a label column."
const LABEL_TEXT_FRACTION_THRESHOLD = 0.6;
// Text fraction below this → first column is clearly not a label column.
const LABEL_NUMERIC_THRESHOLD       = 0.5;

// Blocking: label split is uncertain AND body is still text-heavy → unsafe.
const LABEL_BLOCKING_BODY_TEXT_FRACTION = 0.4;

// Confidence thresholds on body numeric fraction.
const CONFIDENCE_HIGH_NUMERIC   = 0.7;
const CONFIDENCE_MEDIUM_NUMERIC = 0.5;

// ─── Cell helpers ─────────────────────────────────────────────────────────────

function isCellEmpty(cell) {
  return cell === "" || cell === null || cell === undefined;
}

// Strict numeric pattern: the entire cell value (after trimming) must consist of
// an optional sign, digits with an optional decimal separator (dot or comma), and
// an optional trailing percent.  Mixed strings like "2025Q4", "(a)", or
// "Волна (квартал)" do NOT match — parseFloat() would accept them because it
// stops at the first non-numeric character, but these are header labels, not data.
const STRICT_NUMERIC_RE = /^[+-]?(\d+([.,]\d*)?|\d*[.,]\d+)%?$/;

function isNumericCell(cell) {
  if (typeof cell === "number") return !isNaN(cell);
  if (typeof cell !== "string") return false;
  const trimmed = cell.trim();
  if (!trimmed) return false;
  return STRICT_NUMERIC_RE.test(trimmed);
}

function isTextOnlyCell(cell) {
  if (typeof cell !== "string" || !cell.trim()) return false;
  // A non-empty string is "text-only" when it is NOT a strict numeric value.
  // Using isNumericCell here keeps both functions consistent: anything that
  // isNumericCell rejects (e.g. "2025Q4", "(a)") is correctly treated as text.
  return !isNumericCell(cell);
}

// ─── Grid analysis helpers ────────────────────────────────────────────────────

function computeTextFraction(values, startRow, endRow, startCol, endCol) {
  let textCount = 0;
  let totalNonEmpty = 0;
  for (let r = startRow; r <= endRow; r++) {
    const row = values[r];
    if (!row) continue;
    for (let c = startCol; c <= endCol; c++) {
      const cell = row[c];
      if (isCellEmpty(cell)) continue;
      totalNonEmpty++;
      if (isTextOnlyCell(cell)) textCount++;
    }
  }
  return totalNonEmpty === 0 ? 0 : textCount / totalNonEmpty;
}

function computeNumericFraction(values, startRow, endRow, startCol, endCol) {
  let numericCount = 0;
  let totalNonEmpty = 0;
  for (let r = startRow; r <= endRow; r++) {
    const row = values[r];
    if (!row) continue;
    for (let c = startCol; c <= endCol; c++) {
      const cell = row[c];
      if (isCellEmpty(cell)) continue;
      totalNonEmpty++;
      if (isNumericCell(cell)) numericCount++;
    }
  }
  return totalNonEmpty === 0 ? 0 : numericCount / totalNonEmpty;
}

function countNonEmpty(values, startRow, endRow, startCol, endCol) {
  let count = 0;
  for (let r = startRow; r <= endRow; r++) {
    const row = values[r];
    if (!row) continue;
    for (let c = startCol; c <= endCol; c++) {
      if (!isCellEmpty(row[c])) count++;
    }
  }
  return count;
}

// ─── Pass-through gate ────────────────────────────────────────────────────────

/**
 * Returns true when the selection appears to contain non-numeric structure
 * (label columns, header rows, title rows) that warrants normalization.
 *
 * Returns false when the selection looks like numeric-only data —
 * no normalization attempt should be made in that case.
 */
function isNormalizationNeeded(values, rowCount, colCount) {
  if (rowCount < GATE_MIN_ROWS || colCount < GATE_MIN_COLS) return false;

  // Left-column pattern: col[0] is text-heavy, cols[1..N] are numeric-heavy.
  const leftColNonEmpty = countNonEmpty(values, 0, rowCount - 1, 0, 0);
  if (leftColNonEmpty >= Math.max(2, Math.ceil(rowCount * 0.5))) {
    const leftColTextFrac = computeTextFraction(values, 0, rowCount - 1, 0, 0);
    const restNumericFrac = computeNumericFraction(values, 0, rowCount - 1, 1, colCount - 1);
    if (
      leftColTextFrac    >= GATE_LEFT_COL_TEXT_FRACTION &&
      restNumericFrac    >= GATE_REST_COLS_NUMERIC_FRACTION
    ) {
      return true;
    }
  }

  // Top-row pattern: row[0] is text-heavy, rows[1..N] are numeric-heavy.
  // Requires row[0] to fill at least half the columns (wide banner/header row).
  const topRowNonEmpty = countNonEmpty(values, 0, 0, 0, colCount - 1);
  if (topRowNonEmpty >= Math.max(2, Math.ceil(colCount * 0.5))) {
    const topRowTextFrac     = computeTextFraction(values, 0, 0, 0, colCount - 1);
    const lowerNumericFrac   = computeNumericFraction(values, 1, rowCount - 1, 0, colCount - 1);
    if (
      topRowTextFrac   >= GATE_TOP_ROW_TEXT_FRACTION &&
      lowerNumericFrac >= GATE_LOWER_ROWS_NUMERIC_FRACTION
    ) {
      return true;
    }
  }

  // Sparse title-row pattern: row[0] is a sparse text-only heading (≤3 cells),
  // rows[1..N] are numeric-heavy.  The wide top-row gate above misses this case
  // because a merged-like title cell fills only 1 column regardless of table width.
  if (isTitleLikeRow(values, 0, colCount)) {
    const lowerNumericFrac = computeNumericFraction(values, 1, rowCount - 1, 0, colCount - 1);
    if (lowerNumericFrac >= GATE_LOWER_ROWS_NUMERIC_FRACTION) {
      return true;
    }
  }

  // Fallback: overall text-heavy with substantial numeric content.
  const totalCells = rowCount * colCount;
  if (totalCells >= GATE_OVERALL_MIN_CELLS) {
    const allTextFrac    = computeTextFraction(values, 0, rowCount - 1, 0, colCount - 1);
    const allNumericFrac = computeNumericFraction(values, 0, rowCount - 1, 0, colCount - 1);
    if (allTextFrac >= GATE_OVERALL_TEXT_FRACTION && allNumericFrac >= GATE_OVERALL_NUMERIC_FRACTION) {
      return true;
    }
  }

  return false;
}

// ─── Title / subtitle detection ───────────────────────────────────────────────

/**
 * A row is title-like when:
 *   - it has at least one non-empty non-numeric text cell
 *   - it has no numeric cells
 *   - it is sparse: at most MAX_TITLE_ROW_NON_EMPTY_CELLS non-empty cells
 *
 * The sparseness constraint distinguishes title rows (single merged-cell heading)
 * from banner rows (which span most or all columns).
 */
function isTitleLikeRow(values, rowIndex, colCount) {
  const row = values[rowIndex];
  if (!row) return false;

  let nonEmptyCount = 0;
  let hasText = false;

  for (let c = 0; c < colCount; c++) {
    const cell = row[c];
    if (isCellEmpty(cell)) continue;
    nonEmptyCount++;
    if (isNumericCell(cell)) return false;
    if (isTextOnlyCell(cell)) hasText = true;
  }

  if (!hasText || nonEmptyCount === 0 || nonEmptyCount > MAX_TITLE_ROW_NON_EMPTY_CELLS) {
    return false;
  }

  // Title and subtitle rows originate from col[0] — the leftmost cell must be
  // non-empty. A sparse row whose content is entirely in col[1]+ (common for
  // Excel merged banner headers, where the merge value appears only in the
  // top-left data cell) is a banner row, not a title or subtitle row.
  return !isCellEmpty(row[0]);
}

/**
 * Detects up to one title row and one subtitle row at the top of the selection.
 * Both rows must be at the very beginning; any non-title-like row stops detection.
 */
function detectTitleSubtitleRows(values, rowCount, colCount) {
  const titleRows    = [];
  const subtitleRows = [];

  if (rowCount === 0) return { titleRows, subtitleRows };

  let idx = 0;

  if (isTitleLikeRow(values, idx, colCount)) {
    titleRows.push(idx);
    idx++;

    if (idx < rowCount && isTitleLikeRow(values, idx, colCount)) {
      subtitleRows.push(idx);
    }
  }

  return { titleRows, subtitleRows };
}

// ─── Banner / header row detection ───────────────────────────────────────────

/**
 * Detects consecutive banner/header rows starting from startRow.
 *
 * A banner row is:
 *   - wide: fills ≥ BANNER_MIN_FILL_FRACTION of the total column count
 *   - text-heavy: text fraction ≥ BANNER_TEXT_FRACTION
 *
 * The fill requirement distinguishes banner rows from title rows (which are sparse).
 * Detection stops at the first row that fails either criterion.
 */
function detectBannerRows(values, startRow, rowCount, colCount) {
  const bannerRows = [];

  for (let r = startRow; r < rowCount; r++) {
    const nonEmpty  = countNonEmpty(values, r, r, 0, colCount - 1);
    const fillFrac  = colCount > 0 ? nonEmpty / colCount : 0;
    const textFrac  = computeTextFraction(values, r, r, 0, colCount - 1);

    if (fillFrac >= BANNER_MIN_FILL_FRACTION && textFrac >= BANNER_TEXT_FRACTION) {
      bannerRows.push(r);
    } else {
      break;
    }
  }

  return bannerRows;
}

// ─── Body-start detection ─────────────────────────────────────────────────────

/**
 * Finds the first row that looks like a real data body row, scanning forward
 * from startRow.
 *
 * A row qualifies when:
 *   - col[0] is non-empty, AND
 *   - at least one cell in cols[1..N-1] is numeric.
 *
 * This handles merged-like / sparse banner rows that detectBannerRows misses:
 * Excel merged areas store text only in the top-left cell of the merge, so a
 * banner row spanning many columns often looks like ["", "Всего", ""] with an
 * empty col[0]. The first real data row, by contrast, always has a label in
 * col[0] alongside numeric values to its right.
 *
 * Returns startRow as a safe fallback when no qualifying row is found before
 * endRow (validation will then produce an appropriate blocking reason).
 */
function findFirstDataBodyRow(values, startRow, rowCount, colCount) {
  if (colCount < 2) return startRow;
  for (let r = startRow; r < rowCount; r++) {
    const row = values[r] || [];
    if (isCellEmpty(row[0])) continue;
    for (let c = 1; c < colCount; c++) {
      if (isNumericCell(row[c])) return r;
    }
  }
  return startRow;
}

// ─── Body-end detection ───────────────────────────────────────────────────────

/**
 * Finds the last row at or before endRow that contains at least one numeric cell
 * in the data columns [dataColStart, dataColEnd].
 *
 * Trailing rows that have only label content (col[0] text, empty data columns)
 * are treated as footer rows and excluded from the validated body range.  This
 * prevents a single-table selection with a trailing summary row (e.g. "Все
 * респонденты" with empty data columns) from falsely triggering the
 * BODY_APPEARS_MULTI_TABLE empty-gap check.
 *
 * Returns startRow as a safe fallback when no numeric row is found.
 */
function findLastDataBodyRow(values, startRow, endRow, dataColStart, dataColEnd) {
  for (let r = endRow; r >= startRow; r--) {
    const row = values[r] || [];
    for (let c = dataColStart; c <= dataColEnd; c++) {
      if (isNumericCell(row[c])) return r;
    }
  }
  return startRow;
}

// ─── Label column detection ───────────────────────────────────────────────────

/**
 * Detects up to 2 row-label columns at the left edge of the body rows.
 *
 * Rules (applied over body rows only):
 *   col[0] text fraction < 0.5          → no label column, confident
 *   col[0] text fraction in [0.5, 0.6)  → split uncertain
 *   col[0] text fraction ≥ 0.6          → label column
 *     col[1] text fraction ≥ 0.6        → second label column
 *     col[1] text fraction < 0.6        → only one label column
 *
 * Never consumes a column with < 50% text as a label column.
 */
function detectLabelColumns(values, bodyStartRow, bodyEndRow, colCount) {
  if (colCount < 2 || bodyEndRow < bodyStartRow) {
    return { labelColCount: 0, labelSplitConfidence: "uncertain" };
  }

  const col0Frac = computeTextFraction(values, bodyStartRow, bodyEndRow, 0, 0);

  if (col0Frac < LABEL_NUMERIC_THRESHOLD) {
    return { labelColCount: 0, labelSplitConfidence: "confident" };
  }

  if (col0Frac < LABEL_TEXT_FRACTION_THRESHOLD) {
    // Ambiguous: col[0] is neither clearly text nor clearly numeric.
    return { labelColCount: 0, labelSplitConfidence: "uncertain" };
  }

  // col[0] is clearly text (≥ 0.6).
  if (colCount < 3) {
    return { labelColCount: 1, labelSplitConfidence: "confident" };
  }

  const col1Frac = computeTextFraction(values, bodyStartRow, bodyEndRow, 1, 1);
  if (col1Frac >= LABEL_TEXT_FRACTION_THRESHOLD) {
    return { labelColCount: 2, labelSplitConfidence: "confident" };
  }

  return { labelColCount: 1, labelSplitConfidence: "confident" };
}

// ─── Body validation ──────────────────────────────────────────────────────────

/**
 * Validates the candidate data body.
 * Returns an array of blocking-reason objects (empty = valid).
 */
function validateBody(values, bodyStartRow, bodyEndRow, dataColStart, dataColEnd) {
  const reasons = [];
  const bodyRowCount  = bodyEndRow - bodyStartRow + 1;
  const dataColCount  = dataColEnd - dataColStart + 1;

  if (bodyRowCount < 2) {
    reasons.push({
      code: "BODY_TOO_SHORT",
      message: "Недостаточно строк данных после удаления заголовков.",
    });
  }

  if (dataColCount < 2) {
    reasons.push({
      code: "DATA_TOO_NARROW",
      message: "Недостаточно столбцов данных после удаления столбцов меток.",
    });
  }

  // Stop early — the following checks require a valid body rectangle.
  if (reasons.length > 0) return reasons;

  // All-empty row gap inside the body → likely multiple tables in one selection.
  for (let r = bodyStartRow; r <= bodyEndRow; r++) {
    const nonEmpty = countNonEmpty(values, r, r, dataColStart, dataColEnd);
    if (nonEmpty === 0) {
      reasons.push({
        code: "BODY_APPEARS_MULTI_TABLE",
        message:
          "Выделенный диапазон содержит пустые строки внутри области данных. " +
          "Возможно, выделено несколько таблиц. Выделите одну таблицу.",
      });
      break;
    }
  }

  if (reasons.length > 0) return reasons;

  // No numeric content at all → not a data table.
  let hasNumeric = false;
  outer: for (let r = bodyStartRow; r <= bodyEndRow; r++) {
    const row = values[r] || [];
    for (let c = dataColStart; c <= dataColEnd; c++) {
      if (isNumericCell(row[c])) {
        hasNumeric = true;
        break outer;
      }
    }
  }

  if (!hasNumeric) {
    reasons.push({
      code: "NO_NUMERIC_BODY",
      message: "В выделенном диапазоне не найдено числовых данных.",
    });
  }

  return reasons;
}

// ─── Confidence scoring ───────────────────────────────────────────────────────

/**
 * Scores decomposition confidence based on body numeric density
 * and label split reliability.
 *
 * Returns "high" | "medium" | "low".
 * "low" is treated as a blocking condition by the caller.
 */
function scoreConfidence(bodyNumericFrac, labelSplitConfidence) {
  if (bodyNumericFrac < CONFIDENCE_MEDIUM_NUMERIC) return "low";

  if (labelSplitConfidence === "uncertain") {
    // Uncertain split: cap at medium regardless of numeric fraction.
    return bodyNumericFrac >= CONFIDENCE_HIGH_NUMERIC ? "medium" : "low";
  }

  if (bodyNumericFrac < CONFIDENCE_HIGH_NUMERIC) return "medium";
  return "high";
}

// ─── Grid slicing ─────────────────────────────────────────────────────────────

function sliceGrid(grid, startRow, endRow, startCol, endCol) {
  const result = [];
  for (let r = startRow; r <= endRow; r++) {
    const row = grid[r] || [];
    result.push(row.slice(startCol, endCol + 1));
  }
  return result;
}

function buildIndexRange(start, end) {
  const arr = [];
  for (let i = start; i <= end; i++) arr.push(i);
  return arr;
}

// ─── Model builders ───────────────────────────────────────────────────────────

function buildPassThroughModel(rowCount, colCount) {
  return {
    normalizationNeeded:  false,
    normalizationApplied: false,
    originalRowCount:     rowCount,
    originalColumnCount:  colCount,

    titleRows:    [],
    subtitleRows: [],
    bannerRows:   [],
    labelColumns: [],
    dataColumns:  buildIndexRange(0, colCount - 1),
    bodyRows:     buildIndexRange(0, rowCount - 1),

    valuesForCalculation:        null,
    textForCalculation:          null,
    leftLabelValues:             null,
    bannerContext:               null,

    dataRowOffset: 0,
    dataColOffset: 0,

    confidence:      "high",
    warnings:        [],
    blockingReasons: [],
    blockingMessage: "",
  };
}

function buildBlockedModel(
  rowCount,
  colCount,
  titleRows,
  subtitleRows,
  bannerRows,
  labelColCount,
  confidence,
  blockingReasons
) {
  const firstReason = blockingReasons[0];
  const blockingMessage = firstReason
    ? firstReason.message
    : "Не удаётся нормализовать выделенный диапазон. Выделите только числовую область данных.";

  return {
    normalizationNeeded:  true,
    normalizationApplied: false,
    originalRowCount:     rowCount,
    originalColumnCount:  colCount,

    titleRows,
    subtitleRows,
    bannerRows,
    labelColumns: labelColCount > 0 ? buildIndexRange(0, labelColCount - 1) : [],
    dataColumns:  [],
    bodyRows:     [],

    valuesForCalculation:        null,
    textForCalculation:          null,
    leftLabelValues:             null,
    bannerContext:               null,

    dataRowOffset: 0,
    dataColOffset: 0,

    confidence,
    warnings:        [],
    blockingReasons: blockingReasons.map((r) => r.code),
    blockingMessage,
  };
}

function buildNormalizedModel(
  values,
  text,
  rowCount,
  colCount,
  titleRows,
  subtitleRows,
  bannerRows,
  labelColCount,
  dataColStart,
  dataColEnd,
  bodyStartRow,
  bodyEndRow,
  confidence,
  warnings
) {
  const labelColumns = labelColCount > 0 ? buildIndexRange(0, labelColCount - 1) : [];
  const dataColumns  = buildIndexRange(dataColStart, dataColEnd);
  const bodyRows     = buildIndexRange(bodyStartRow, bodyEndRow);

  const valuesForCalculation = sliceGrid(values, bodyStartRow, bodyEndRow, dataColStart, dataColEnd);

  const textForCalculation = text.length > 0
    ? sliceGrid(text, bodyStartRow, bodyEndRow, dataColStart, dataColEnd)
    : [];

  // Significance marker removal (applied to values before normalization) can erase
  // pure-letter label cells like "mean" or "BASE". Use text (Office.js selectedRange.text,
  // never marker-stripped) when it covers the body rows.
  const leftLabelSource = text.length > bodyEndRow ? text : values;
  const leftLabelValues = labelColCount > 0
    ? sliceGrid(leftLabelSource, bodyStartRow, bodyEndRow, 0, labelColCount - 1)
    : [];

  // Banner scan rows are sliced to data columns only so they align with
  // valuesForCalculation. detectBannerStructure expects column indices to
  // match the data grid, not the full raw selection width.
  const bannerScanRows = bannerRows.length > 0
    ? sliceGrid(values, bannerRows[0], bannerRows[bannerRows.length - 1], dataColStart, dataColEnd)
    : [];

  const bannerContext = {
    scanRows:    bannerScanRows,
    columnCount: dataColEnd - dataColStart + 1,
  };

  return {
    normalizationNeeded:  true,
    normalizationApplied: true,
    originalRowCount:     rowCount,
    originalColumnCount:  colCount,

    titleRows,
    subtitleRows,
    bannerRows,
    labelColumns,
    dataColumns,
    bodyRows,

    valuesForCalculation,
    textForCalculation,
    leftLabelValues,
    bannerContext,

    dataRowOffset: bodyStartRow,
    dataColOffset: dataColStart,

    confidence,
    warnings,
    blockingReasons: [],
    blockingMessage: "",
  };
}

// ─── Main export ──────────────────────────────────────────────────────────────

/**
 * Analyzes a selected Excel range and returns a NormalizedSelectionModel.
 *
 * @param {Array}  rawValues - 2D array from selectedRange.values (markers already removed).
 * @param {Array}  [rawText] - Optional 2D array from selectedRange.text.
 * @param {object} [options] - Reserved for future configuration options.
 * @returns {object} NormalizedSelectionModel
 *
 * Callers must check normalizationNeeded and normalizationApplied before using
 * any sliced output grids. See module-level JSDoc for the three valid states.
 */
export function normalizeSelectedRange(rawValues, rawText, options = {}) {
  const values   = Array.isArray(rawValues) ? rawValues : [];
  const text     = Array.isArray(rawText)   ? rawText   : [];
  const rowCount = values.length;
  const colCount = rowCount > 0 && Array.isArray(values[0]) ? values[0].length : 0;

  // ── Step 0: Pass-through gate ──────────────────────────────────────────────
  // If the selection looks like numeric-only data, return immediately.
  // The existing strict workflow must run unchanged for this state.
  if (!isNormalizationNeeded(values, rowCount, colCount)) {
    return buildPassThroughModel(rowCount, colCount);
  }

  // Normalization is needed. Attempt structural decomposition.

  // ── Step 1: Title / subtitle rows ─────────────────────────────────────────
  const { titleRows, subtitleRows } = detectTitleSubtitleRows(values, rowCount, colCount);

  // ── Step 2: Banner rows ────────────────────────────────────────────────────
  // Banner detection starts immediately after title/subtitle rows.
  const firstBodyCandidate = titleRows.length + subtitleRows.length;

  // Approach A: consecutive wide text-heavy rows (existing logic).
  const wideBannerRowCount = detectBannerRows(values, firstBodyCandidate, rowCount, colCount).length;

  // Approach B: scan for the first row where col[0] is non-empty AND at least
  // one cell to its right is numeric. All rows before that point are treated as
  // banner/header rows. This covers merged-like sparse header rows (Excel stores
  // merged text only in the top-left cell, leaving col[0] empty in continuation
  // rows) that approach A misses because their fill fraction is too low.
  const firstDataBodyRow = findFirstDataBodyRow(values, firstBodyCandidate, rowCount, colCount);

  // Use whichever approach identifies a later body start (more header rows).
  const bodyStartRow = Math.max(firstBodyCandidate + wideBannerRowCount, firstDataBodyRow);
  const bannerRows   = bodyStartRow > firstBodyCandidate
    ? buildIndexRange(firstBodyCandidate, bodyStartRow - 1)
    : [];

  // Safety check: a legitimate single-table banner rarely exceeds MAX_BANNER_ROW_COUNT
  // rows.  If more rows were classified as banner it almost certainly means the
  // selection spans multiple tables (the first table's data rows were consumed into
  // the header area because they looked text-heavy / non-numeric).
  if (bannerRows.length > MAX_BANNER_ROW_COUNT) {
    return buildBlockedModel(
      rowCount,
      colCount,
      titleRows,
      subtitleRows,
      bannerRows,
      0,
      "medium",
      [
        {
          code: "HEADER_AREA_TOO_LARGE",
          message:
            "Область заголовков слишком велика — возможно, выделено несколько таблиц. " +
            "Выделите одну таблицу.",
        },
      ]
    );
  }

  const bodyEndRow   = rowCount - 1;

  // ── Step 3: Label columns ──────────────────────────────────────────────────
  // Applied over body rows only (title/subtitle/banner already excluded).
  const { labelColCount, labelSplitConfidence } = detectLabelColumns(
    values,
    bodyStartRow,
    bodyEndRow,
    colCount
  );

  const dataColStart = labelColCount;
  const dataColEnd   = colCount - 1;

  // ── Step 3b: Trim trailing footer rows ────────────────────────────────────
  // Rows after the last numeric data row (e.g. a summary label with empty data
  // columns) must not participate in the empty-gap check or in slicing outputs.
  const bodyDataEndRow = findLastDataBodyRow(values, bodyStartRow, bodyEndRow, dataColStart, dataColEnd);

  // ── Step 4: Body validation ────────────────────────────────────────────────
  const blockingReasons = validateBody(values, bodyStartRow, bodyDataEndRow, dataColStart, dataColEnd);

  // Label split safety: uncertain split + body still text-heavy → unsafe to proceed.
  if (
    blockingReasons.length === 0 &&
    labelSplitConfidence === "uncertain"
  ) {
    const bodyTextFrac = computeTextFraction(
      values,
      bodyStartRow,
      bodyDataEndRow,
      dataColStart,
      dataColEnd
    );
    if (bodyTextFrac >= LABEL_BLOCKING_BODY_TEXT_FRACTION) {
      blockingReasons.push({
        code: "LABEL_SPLIT_BLOCKING",
        message:
          "Не удаётся надёжно определить границу между метками строк и данными. " +
          "Выделите только числовую область данных.",
      });
    }
  }

  // ── Step 5: Confidence scoring ─────────────────────────────────────────────
  // Only computed when the body passed structural validation.
  let confidence = "high";

  if (blockingReasons.length === 0) {
    const bodyNumericFrac = computeNumericFraction(
      values,
      bodyStartRow,
      bodyDataEndRow,
      dataColStart,
      dataColEnd
    );
    confidence = scoreConfidence(bodyNumericFrac, labelSplitConfidence);

    if (confidence === "low") {
      blockingReasons.push({
        code: "LOW_CONFIDENCE",
        message:
          "Не удаётся надёжно определить структуру таблицы в выделенном диапазоне. " +
          "Выделите только числовую область данных.",
      });
    }
  }

  // ── Return blocked model if any reason prevents safe normalization ─────────
  if (blockingReasons.length > 0) {
    return buildBlockedModel(
      rowCount,
      colCount,
      titleRows,
      subtitleRows,
      bannerRows,
      labelColCount,
      confidence === "low" ? "low" : "medium",
      blockingReasons
    );
  }

  // ── Return normalized model ────────────────────────────────────────────────
  const warnings = [];

  if (titleRows.length > 0) {
    warnings.push({ code: "TITLE_ROWS_DETECTED",   severity: "info",    rowIndexes: titleRows });
  }
  if (subtitleRows.length > 0) {
    warnings.push({ code: "SUBTITLE_ROWS_DETECTED", severity: "info",   rowIndexes: subtitleRows });
  }
  if (bannerRows.length > 0) {
    warnings.push({ code: "BANNER_ROWS_DETECTED",  severity: "info",    rowIndexes: bannerRows });
  }
  if (labelColCount > 0) {
    warnings.push({
      code: "LABEL_COLUMNS_DETECTED",
      severity: "info",
      columnIndexes: buildIndexRange(0, labelColCount - 1),
    });
  }
  if (confidence === "medium") {
    warnings.push({ code: "NORMALIZATION_CONFIDENCE_MEDIUM", severity: "warning" });
  }

  return buildNormalizedModel(
    values,
    text,
    rowCount,
    colCount,
    titleRows,
    subtitleRows,
    bannerRows,
    labelColCount,
    dataColStart,
    dataColEnd,
    bodyStartRow,
    bodyDataEndRow,
    confidence,
    warnings
  );
}
