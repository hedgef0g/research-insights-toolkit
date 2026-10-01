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
const GATE_MIN_COLS = 2;

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

// Extended NPS tables often use numeric-looking scale labels in the first
// column ("1".."10"). Keep this high enough that ordinary numeric columns do
// not become labels unless the NPS support rows are also present.
const NPS_SCALE_LABEL_MIN_COUNT = 5;
const NPS_SCALE_RIGHT_NUMERIC_FRACTION = 0.6;

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
const NUMERIC_WITH_MARKER_SUFFIX_RE =
  /^([+-]?(\d+([.,]\d*)?|\d*[.,]\d+)%?)(\s+[\p{L}↑↓]+)+$/u;

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

function isLikelyOrdinalScaleLabelCell(cell) {
  if (typeof cell === "number") {
    return Number.isInteger(cell) && cell >= 0 && cell <= 10;
  }

  if (typeof cell !== "string") {
    return false;
  }

  const trimmed = cell.trim();
  if (!/^\d+$/.test(trimmed)) {
    return false;
  }

  const value = Number(trimmed);
  return Number.isInteger(value) && value >= 0 && value <= 10;
}

function cleanStructuralTextCell(cell) {
  if (typeof cell !== "string") {
    return cell;
  }

  const trimmed = cell.trim();
  const match = trimmed.match(NUMERIC_WITH_MARKER_SUFFIX_RE);

  return match ? match[1] : cell;
}

function cleanStructuralTextGrid(text) {
  return text.map((row) =>
    Array.isArray(row) ? row.map((cell) => cleanStructuralTextCell(cell)) : row
  );
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

  // Extended NPS pattern: col[0] is a scale-label column (numeric values 0–10
  // mixed with NPS support rows such as Detractors, Promoters, NPS, Base).
  // Left-column text fraction falls below GATE_LEFT_COL_TEXT_FRACTION because
  // the scale values "1"–"10" are numeric-looking, but the column still needs
  // to be separated from the data. Checked over the full grid — banner rows in
  // col[0] contain group names that match no NPS pattern and are silently skipped.
  if (isExtendedNpsScaleLabelColumn(values, 0, rowCount - 1, colCount)) {
    return true;
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

  // A row that fills most columns is a banner row, not a sparse title row.
  // This guards 2-column selections where any fully-filled text row would
  // otherwise satisfy the ≤3 cell count and be misclassified as a title.
  if (colCount > 0 && nonEmptyCount / colCount > BANNER_MIN_FILL_FRACTION) {
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
 * Returns true when every non-empty cell in col[1] of the body rows shares the
 * same trimmed text value AND that value is a unit-indicator pattern:
 *   "%"             — pure percent sign (text cell, no numeric value)
 *   "0%", "0.0%"   — zero-formatted percentage (value=0, format="%")
 *   "0",  "0.0"    — zero decimal (value=0, no percent format)
 *
 * A uniform unit column adjacent to an already-identified label column is
 * almost certainly a unit placeholder, not real data.  Non-zero values
 * (e.g. "5%") are treated as real data even when uniform.
 *
 * colCount must be ≥ 3 before calling (caller responsibility).
 */
function isUniformUnitColumn(values, bodyStartRow, bodyEndRow) {
  let seenValue;
  let hasNonEmpty = false;

  for (let r = bodyStartRow; r <= bodyEndRow; r++) {
    const row = values[r] || [];
    const cell = row[1];
    if (isCellEmpty(cell)) continue;
    const strVal = String(cell).trim();
    if (!strVal) continue;
    hasNonEmpty = true;
    if (seenValue === undefined) {
      seenValue = strVal;
    } else if (strVal !== seenValue) {
      return false;
    }
  }

  if (!hasNonEmpty || seenValue === undefined) return false;

  return (
    seenValue === "%" ||
    /^0(\.0+)?%$/.test(seenValue) ||
    /^0(\.0+)?$/.test(seenValue)
  );
}

/**
 * Detects up to 2 row-label columns at the left edge of the body rows.
 *
 * Rules (applied over body rows only):
 *   col[0] text fraction < 0.5          → no label column, confident
 *   col[0] text fraction in [0.5, 0.6)  → split uncertain
 *   col[0] text fraction ≥ 0.6          → label column
 *     col[1] text fraction ≥ 0.6        → second label column
 *     col[1] text fraction < 0.6        → only one label column
 *       col[1] is uniform unit value    → second label column (unit column)
 *
 * Never consumes a column with < 50% text as a label column.
 */
function detectLabelColumns(values, bodyStartRow, bodyEndRow, colCount) {
  if (colCount < 2 || bodyEndRow < bodyStartRow) {
    return { labelColCount: 0, labelSplitConfidence: "uncertain" };
  }

  const col0Frac = computeTextFraction(values, bodyStartRow, bodyEndRow, 0, 0);
  const col0LooksLikeRatingScale = isLikelyRatingScaleLabelColumn(
    values,
    bodyStartRow,
    bodyEndRow
  );

  if (col0Frac < LABEL_NUMERIC_THRESHOLD && !col0LooksLikeRatingScale) {
    if (isExtendedNpsScaleLabelColumn(values, bodyStartRow, bodyEndRow, colCount)) {
      if (colCount >= 3 && isUniformUnitColumn(values, bodyStartRow, bodyEndRow)) {
        return { labelColCount: 2, labelSplitConfidence: "confident" };
      }
      return { labelColCount: 1, labelSplitConfidence: "confident" };
    }
    return { labelColCount: 0, labelSplitConfidence: "confident" };
  }

  if (col0Frac < LABEL_TEXT_FRACTION_THRESHOLD && !col0LooksLikeRatingScale) {
    if (isExtendedNpsScaleLabelColumn(values, bodyStartRow, bodyEndRow, colCount)) {
      if (colCount >= 3 && isUniformUnitColumn(values, bodyStartRow, bodyEndRow)) {
        return { labelColCount: 2, labelSplitConfidence: "confident" };
      }
      return { labelColCount: 1, labelSplitConfidence: "confident" };
    }
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

  if (isUniformUnitColumn(values, bodyStartRow, bodyEndRow)) {
    return { labelColCount: 2, labelSplitConfidence: "confident" };
  }

  return { labelColCount: 1, labelSplitConfidence: "confident" };
}

function normalizeLabelCandidate(rawValue) {
  if (rawValue === null || rawValue === undefined) {
    return "";
  }

  return String(rawValue)
    .toLowerCase()
    .replace(/[.,:;()_-]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function isNpsScaleLabel(rawValue) {
  const label = normalizeLabelCandidate(rawValue);
  if (!label || !/^\d+$/.test(label)) {
    return false;
  }

  const value = Number(label);
  return value >= 0 && value <= 10;
}

function isExtendedNpsScaleLabelColumn(values, bodyStartRow, bodyEndRow, colCount) {
  if (colCount < 3 || bodyEndRow < bodyStartRow) {
    return false;
  }

  const rightNumericFrac = computeNumericFraction(
    values,
    bodyStartRow,
    bodyEndRow,
    1,
    colCount - 1
  );

  if (rightNumericFrac < NPS_SCALE_RIGHT_NUMERIC_FRACTION) {
    return false;
  }

  let scaleLabelCount = 0;
  let bucketOrNeutralCount = 0;
  let hasDetractors = false;
  let hasPromoters = false;
  let hasNps = false;
  let hasBase = false;

  for (let rowIndex = bodyStartRow; rowIndex <= bodyEndRow; rowIndex++) {
    const row = values[rowIndex] || [];
    const rawLabel = row[0];
    const label = normalizeLabelCandidate(rawLabel);

    if (isNpsScaleLabel(rawLabel)) {
      scaleLabelCount++;
      continue;
    }

    if (label === "detractors" || label === "detractor") {
      hasDetractors = true;
      continue;
    }

    if (label === "promoters" || label === "promoter") {
      hasPromoters = true;
      continue;
    }

    if (label === "nps" || label === "net promoter score") {
      hasNps = true;
      continue;
    }

    if (label === "base") {
      hasBase = true;
      continue;
    }

    if (
      label === "neutral" ||
      label === "neutrals" ||
      label.startsWith("bottom ") ||
      label.startsWith("mid ") ||
      label.startsWith("middle ") ||
      label.startsWith("top ")
    ) {
      bucketOrNeutralCount++;
    }
  }

  return (
    scaleLabelCount >= NPS_SCALE_LABEL_MIN_COUNT &&
    bucketOrNeutralCount > 0 &&
    hasDetractors &&
    hasPromoters &&
    hasNps &&
    hasBase
  );
}

function isLikelyRatingScaleLabelColumn(values, bodyStartRow, bodyEndRow) {
  let totalNonEmpty = 0;
  let textOnlyCount = 0;
  let ordinalScaleCount = 0;

  for (let row = bodyStartRow; row <= bodyEndRow; row++) {
    const rowArr = values[row];
    if (!rowArr) continue;

    const cell = rowArr[0];
    if (isCellEmpty(cell)) continue;

    totalNonEmpty++;

    if (isTextOnlyCell(cell)) {
      textOnlyCount++;
      continue;
    }

    if (isLikelyOrdinalScaleLabelCell(cell)) {
      ordinalScaleCount++;
    }
  }

  if (totalNonEmpty === 0) {
    return false;
  }

  const labelLikeFraction = (textOnlyCount + ordinalScaleCount) / totalNonEmpty;
  const allCellsAreLabelLike = textOnlyCount + ordinalScaleCount === totalNonEmpty;
  const minText = allCellsAreLabelLike ? 1 : 2;

  return textOnlyCount >= minText && ordinalScaleCount >= 3 && labelLikeFraction >= 0.8;
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

function hasUsableTextGrid(text, rowCount, colCount) {
  return (
    Array.isArray(text) &&
    text.length >= rowCount &&
    rowCount > 0 &&
    Array.isArray(text[0]) &&
    text[0].length >= colCount
  );
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
  warnings,
  fullBodyEndRow
) {
  // fullBodyEndRow is the pre-trim body end (before findLastDataBodyRow).
  // bodyEndRow is the trimmed end actually used for data slicing.
  const safeFullBodyEndRow = fullBodyEndRow !== undefined ? fullBodyEndRow : bodyEndRow;

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
  const leftLabelSource = text.length > safeFullBodyEndRow ? text : values;
  const leftLabelValues = labelColCount > 0
    ? sliceGrid(leftLabelSource, bodyStartRow, bodyEndRow, 0, labelColCount - 1)
    : [];

  // Rows trimmed by findLastDataBodyRow (blank/non-numeric trailing rows excluded from
  // valuesForCalculation). The Check preview uses this to detect base rows whose data
  // was stripped — without affecting Run calculations or the bodyRows slice.
  const trailingBodyRows = safeFullBodyEndRow > bodyEndRow
    ? {
        values: sliceGrid(values, bodyEndRow + 1, safeFullBodyEndRow, dataColStart, dataColEnd),
        leftLabelValues: labelColCount > 0
          ? sliceGrid(leftLabelSource, bodyEndRow + 1, safeFullBodyEndRow, 0, labelColCount - 1)
          : [],
      }
    : { values: [], leftLabelValues: [] };

  // Banner scan rows are sliced to data columns only so they align with
  // valuesForCalculation. detectBannerStructure expects column indices to
  // match the data grid, not the full raw selection width.
  const lastBannerRow = bannerRows.length > 0 ? bannerRows[bannerRows.length - 1] : -1;
  const bannerSource = text.length > lastBannerRow ? text : values;
  const bannerScanRows = bannerRows.length > 0
    ? sliceGrid(bannerSource, bannerRows[0], bannerRows[bannerRows.length - 1], dataColStart, dataColEnd)
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
    trailingBodyRows,
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
  const structureValues = hasUsableTextGrid(text, rowCount, colCount)
    ? cleanStructuralTextGrid(text)
    : values;

  // ── Step 0: Pass-through gate ──────────────────────────────────────────────
  // If the selection looks like numeric-only data, return immediately.
  // The existing strict workflow must run unchanged for this state.
  if (!isNormalizationNeeded(structureValues, rowCount, colCount)) {
    return buildPassThroughModel(rowCount, colCount);
  }

  // Normalization is needed. Attempt structural decomposition.

  // ── Step 1: Title / subtitle rows ─────────────────────────────────────────
  const { titleRows, subtitleRows } = detectTitleSubtitleRows(
    structureValues,
    rowCount,
    colCount
  );

  // ── Step 2: Banner rows ────────────────────────────────────────────────────
  // Banner detection starts immediately after title/subtitle rows.
  const firstBodyCandidate = titleRows.length + subtitleRows.length;

  // Approach A: consecutive wide text-heavy rows (existing logic).
  const wideBannerRowCount = detectBannerRows(
    structureValues,
    firstBodyCandidate,
    rowCount,
    colCount
  ).length;

  // Approach B: scan for the first row where col[0] is non-empty AND at least
  // one cell to its right is numeric. All rows before that point are treated as
  // banner/header rows. This covers merged-like sparse header rows (Excel stores
  // merged text only in the top-left cell, leaving col[0] empty in continuation
  // rows) that approach A misses because their fill fraction is too low.
  const firstDataBodyRow = findFirstDataBodyRow(
    structureValues,
    firstBodyCandidate,
    rowCount,
    colCount
  );

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
    structureValues,
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
    warnings,
    bodyEndRow   // full pre-trim end row — used only for trailingBodyRows
  );
}

/**
 * Returns true when the values grid contains at least one group of non-empty rows,
 * followed by one or more all-empty rows, followed by at least one more non-empty row.
 *
 * Leading and trailing all-empty rows are ignored, so a single-cell selection, a
 * normal single-table selection, or a table with blank edge rows does not trigger
 * the guard.
 *
 * Used as a pre-resolver sanity guard in runCheckTable() to detect a broad
 * multi-table selection before the active-cell resolver runs.
 *
 * @param {Array} values - 2D array from selectedRange.values
 * @returns {boolean}
 */
export function selectionHasMultiTableGap(values) {
  if (!Array.isArray(values) || values.length === 0) return false;

  function isRowBlank(row) {
    return (
      Array.isArray(row) &&
      row.length > 0 &&
      row.every((cell) => cell === "" || cell === null || cell === undefined)
    );
  }

  let i = 0;

  // Skip leading blank rows.
  while (i < values.length && isRowBlank(values[i])) i++;

  // Advance through the first non-blank group.
  while (i < values.length && !isRowBlank(values[i])) i++;

  if (i >= values.length) return false; // no gap follows the first group

  // Advance through the gap (one or more blank rows).
  while (i < values.length && isRowBlank(values[i])) i++;

  // Multi-table gap detected if another non-blank row exists after the gap.
  return i < values.length;
}

/**
 * Returns true when any row of the values grid is entirely empty (all cells blank).
 *
 * An all-empty row inside a data body indicates the range likely spans multiple
 * tables. Used by checkSelectedRangePreview as a guard for pass-through ranges
 * that bypass normalizeSelectedRange's validateBody (which emits
 * BODY_APPEARS_MULTI_TABLE for normalized ranges only).
 *
 * @param {Array} values  2D values grid (e.g. valuesForCalculation).
 * @returns {boolean}
 */
export function hasEmptyDataRowGap(values) {
  if (!Array.isArray(values)) return false;
  for (const row of values) {
    if (
      Array.isArray(row) &&
      row.length > 0 &&
      row.every((cell) => cell === "" || cell === null || cell === undefined)
    ) {
      return true;
    }
  }
  return false;
}
