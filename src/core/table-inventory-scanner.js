/**
 * Table Inventory Scanner for Research Insights Toolkit.
 *
 * Scans a pre-loaded worksheet used range and returns TableInventoryItem[]
 * representing candidate table regions detected on the sheet.
 *
 * Office.js-free: input is a plain 2D values array with sheet offset metadata.
 * Calls buildTablePreviewModel as the interpretation engine for each candidate.
 *
 * ROLE: read-only candidate finder only.
 *   - Returns candidates for the user to inspect via Check Table.
 *   - Does NOT claim a candidate is ready for Run significance.
 *   - Does NOT write to the workbook.
 *
 * KNOWN LIMITATIONS (intentionally preserved, not blocking):
 *   - Side-by-side tables within one contiguous row band are not detected
 *     separately; the entire band appears as a single (possibly uncertain) candidate.
 *   - Non-empty commentary rows between two tables prevent band splitting;
 *     the tables are merged into one band.
 *   - Multi-column labels beyond the two-column heuristic may degrade split quality.
 *   - Title inference from rows above the band is heuristic and may mis-assign
 *     text that belongs to an unrelated preceding section.
 *   - Candidates with no explicit Base row are not flagged by the scanner;
 *     Check Table should be used for authoritative interpretation.
 *   - Inventory candidate ranges should route to Check Table, not directly to Run.
 */

import { buildTablePreviewModel } from "./table-preview-model";

const LABEL_SCAN_COLUMNS_LEFT = 2;
const MIN_BAND_ROWS = 2;
const LABEL_TEXT_FRACTION_THRESHOLD = 0.6;
// A first-band row is title-like only when it has at most this many non-empty cells.
// This keeps wide banner rows (Total / Male / Female / …) from being mistaken for headings.
const MAX_TITLE_ROW_NON_EMPTY_CELLS = 3;
// Keep scanner-side numeric evidence aligned with the selected-range normalizer:
// mixed strings like "2025Q4" or "(a)" are header labels, not numeric cells.
const STRICT_NUMERIC_RE = /^[+-]?(\d+([.,]\d*)?|\d*[.,]\d+)%?$/;

// ─── A1 address helpers ───────────────────────────────────────────────────────

function columnIndexToLetter(index) {
  let dividend = index + 1;
  let name = "";
  while (dividend > 0) {
    const mod = (dividend - 1) % 26;
    name = String.fromCharCode(65 + mod) + name;
    dividend = Math.floor((dividend - mod) / 26);
  }
  return name;
}

function toA1Address(absRowStart, absRowEnd, absColStart, absColEnd) {
  return (
    columnIndexToLetter(absColStart) +
    (absRowStart + 1) +
    ":" +
    columnIndexToLetter(absColEnd) +
    (absRowEnd + 1)
  );
}

// ─── Cell / row utilities ─────────────────────────────────────────────────────

function isCellEmpty(cell) {
  return cell === "" || cell === null || cell === undefined;
}

function isRowEmpty(row) {
  return !row || row.every(isCellEmpty);
}

function isNumericCell(cell) {
  if (typeof cell === "number") return !isNaN(cell);
  if (typeof cell !== "string") return false;
  const trimmed = cell.trim();
  if (!trimmed) return false;
  return STRICT_NUMERIC_RE.test(trimmed);
}

function isTextOnlyCell(cell) {
  if (typeof cell !== "string" || !cell.trim()) return false;
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

function firstNonEmptyNonNumericText(row) {
  for (const cell of row) {
    if (isCellEmpty(cell)) continue;
    if (isNumericCell(cell)) continue;
    const trimmed = String(cell).trim();
    if (trimmed) return trimmed;
  }
  return null;
}

// ─── Band detection ───────────────────────────────────────────────────────────

function detectRowBands(values) {
  const bands = [];
  let inBand = false;
  let bandStart = 0;

  for (let i = 0; i < values.length; i++) {
    const empty = isRowEmpty(values[i]);
    if (!empty && !inBand) {
      bandStart = i;
      inBand = true;
    } else if (empty && inBand) {
      bands.push({ localStartRow: bandStart, localEndRow: i - 1 });
      inBand = false;
    }
  }

  if (inBand) {
    bands.push({ localStartRow: bandStart, localEndRow: values.length - 1 });
  }

  return bands;
}

// ─── Column trimming ──────────────────────────────────────────────────────────

function trimBandColumns(values, band) {
  const { localStartRow, localEndRow } = band;
  const colCount = values[0] ? values[0].length : 0;
  if (colCount === 0) return null;

  let firstCol = -1;
  let lastCol = -1;

  for (let col = 0; col < colCount; col++) {
    let hasContent = false;
    for (let row = localStartRow; row <= localEndRow; row++) {
      if (!isCellEmpty(values[row][col])) {
        hasContent = true;
        break;
      }
    }
    if (hasContent) {
      if (firstCol === -1) firstCol = col;
      lastCol = col;
    }
  }

  if (firstCol === -1) return null;
  return { ...band, localTrimmedFirstCol: firstCol, localTrimmedLastCol: lastCol };
}

// ─── Pre-filter ───────────────────────────────────────────────────────────────

function hasNumericCell(values, band) {
  const { localStartRow, localEndRow, localTrimmedFirstCol, localTrimmedLastCol } = band;
  for (let row = localStartRow; row <= localEndRow; row++) {
    for (let col = localTrimmedFirstCol; col <= localTrimmedLastCol; col++) {
      if (isNumericCell(values[row][col])) return true;
    }
  }
  return false;
}

// ─── Label / data split ───────────────────────────────────────────────────────

function computeTextFraction(values, startCol, endCol, startRow, endRow) {
  let textCount = 0;
  let totalNonEmpty = 0;
  for (let row = startRow; row <= endRow; row++) {
    const rowArr = values[row];
    if (!rowArr) continue;
    for (let col = startCol; col <= endCol; col++) {
      const cell = rowArr[col];
      if (isCellEmpty(cell)) continue;
      totalNonEmpty++;
      if (isTextOnlyCell(cell)) textCount++;
    }
  }
  if (totalNonEmpty === 0) return 0;
  return textCount / totalNonEmpty;
}

function isLikelyRatingScaleLabelColumn(values, colIndex, startRow, endRow) {
  let totalNonEmpty = 0;
  let textOnlyCount = 0;
  let ordinalScaleCount = 0;

  for (let row = startRow; row <= endRow; row++) {
    const rowArr = values[row];
    if (!rowArr) continue;

    const cell = rowArr[colIndex];
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
  // Allow a single text marker (e.g. "Base") alongside ordinal scale values when every
  // non-empty cell is label-like (no out-of-range numeric data cells).
  const allCellsAreLabelLike = textOnlyCount + ordinalScaleCount === totalNonEmpty;
  const minText = allCellsAreLabelLike ? 1 : 2;

  return textOnlyCount >= minText && ordinalScaleCount >= 3 && labelLikeFraction >= 0.8;
}

/**
 * Returns true when every body row has an empty gutter cell, where "body row" means
 * a row whose col0 (labelCol0Index) is non-empty.  Banner/header rows above the body
 * have an empty col0 and are intentionally ignored — they may carry numeric values
 * (sample sizes, years, percentages) in the gutter column that must not disqualify it.
 * Returns false if no body rows are found (conservative: do not promote to gutter).
 */
function isGutterColumnForBodyRows(values, gutterColIndex, labelCol0Index, startRow, endRow) {
  let bodyRowSeen = false;
  for (let row = startRow; row <= endRow; row++) {
    const labelCell = (values[row] || [])[labelCol0Index];
    if (isCellEmpty(labelCell)) continue;
    bodyRowSeen = true;
    if (!isCellEmpty((values[row] || [])[gutterColIndex])) return false;
  }
  return bodyRowSeen;
}

function splitLabelData(values, band) {
  const { localStartRow, localEndRow, localTrimmedFirstCol, localTrimmedLastCol } = band;
  const trimmedWidth = localTrimmedLastCol - localTrimmedFirstCol + 1;

  // Determine label column count by inspecting the character of the leftmost columns.
  // Rules:
  //   col0 text + col1 text                       → 2 label columns, confident
  //   col0 text + col1 empty gutter (body rows)   → 2 label columns, twoColumn (skip the gutter)
  //   col0 text + col1 numeric                    → 1 label column, confident
  //   col0 not text + col1 text                   → 2 label columns, twoColumn (col0 is code/gutter)
  //   col0 not text + col1 not text               → 1 label column, uncertain
  //   trimmedWidth < 2                            → 0 label columns, uncertain
  // "Empty gutter in body rows" means col1 is null in every row where col0 is non-empty.
  // Banner/header rows above the body may have numeric values in col1 and are ignored.
  let labelColCount;
  let labelSplitConfidence;

  if (trimmedWidth < 2) {
    labelColCount = 0;
    labelSplitConfidence = "uncertain";
  } else {
    const col0Fraction = computeTextFraction(
      values,
      localTrimmedFirstCol,
      localTrimmedFirstCol,
      localStartRow,
      localEndRow
    );
    const col0IsText =
      col0Fraction >= LABEL_TEXT_FRACTION_THRESHOLD ||
      isLikelyRatingScaleLabelColumn(values, localTrimmedFirstCol, localStartRow, localEndRow);

    if (!col0IsText) {
      // col0 does not look like a text label column.  Check col1: if it is strongly text
      // then col0 is a code/numeric identifier and col1 carries the real row labels.
      if (trimmedWidth >= 3) {
        const col1Fraction = computeTextFraction(
          values,
          localTrimmedFirstCol + 1,
          localTrimmedFirstCol + 1,
          localStartRow,
          localEndRow
        );
        if (col1Fraction >= LABEL_TEXT_FRACTION_THRESHOLD) {
          labelColCount = 2;
          labelSplitConfidence = "twoColumn";
        } else {
          labelColCount = 1;
          labelSplitConfidence = "uncertain";
        }
      } else {
        labelColCount = 1;
        labelSplitConfidence = "uncertain";
      }
    } else if (trimmedWidth < 3) {
      // Only 2 columns and col0 is text — use 1 label column.
      labelColCount = 1;
      labelSplitConfidence = "confident";
    } else {
      // 3+ columns, col0 is text: inspect col1.
      const col1Fraction = computeTextFraction(
        values,
        localTrimmedFirstCol + 1,
        localTrimmedFirstCol + 1,
        localStartRow,
        localEndRow
      );
      const col1IsText = col1Fraction >= LABEL_TEXT_FRACTION_THRESHOLD;

      if (col1IsText) {
        // Both col0 and col1 are text — 2 label columns.
        labelColCount = 2;
        labelSplitConfidence = "confident";
      } else if (isGutterColumnForBodyRows(values, localTrimmedFirstCol + 1, localTrimmedFirstCol, localStartRow, localEndRow)) {
        // col0 is text labels and col1 is an empty visual gutter in body rows — skip the
        // gutter so it does not land in the data matrix and trigger spurious quality warnings.
        // Banner rows above the body may have non-empty values in col1 (sample sizes, years,
        // percents) — isGutterColumnForBodyRows ignores those rows.
        labelColCount = 2;
        labelSplitConfidence = "twoColumn";
      } else {
        // col0 text, col1 has numeric / mixed content — 1 label column.
        labelColCount = 1;
        labelSplitConfidence = "confident";
      }
    }
  }

  const labelCols = [];
  const dataCols = [];

  for (let row = localStartRow; row <= localEndRow; row++) {
    const rowArr = values[row] || [];

    const labelRow = [];
    for (let c = 0; c < labelColCount; c++) {
      labelRow.push(rowArr[localTrimmedFirstCol + c] ?? "");
    }
    labelCols.push(labelRow);

    const dataRow = [];
    for (let c = labelColCount; c < trimmedWidth; c++) {
      dataRow.push(rowArr[localTrimmedFirstCol + c] ?? "");
    }
    dataCols.push(dataRow);
  }

  // Trim trailing data columns that are empty in all body rows.
  // Body rows are rows where labelCols[row][0] is non-empty (the primary label is present).
  // Banner/header rows (labelCols[row][0] empty) may add extra columns — e.g. a subgroup
  // header that has no matching data in mean/variance/base — and those trailing all-empty-
  // in-body-rows columns would otherwise produce spurious BASE_BLANK_VALUES quality warnings.
  if (labelColCount > 0 && dataCols.length > 0 && dataCols[0] && dataCols[0].length > 1) {
    const bodyRowIndexes = [];
    for (let i = 0; i < labelCols.length; i++) {
      if (!isCellEmpty(labelCols[i][0])) bodyRowIndexes.push(i);
    }
    if (bodyRowIndexes.length > 0) {
      let trimmedWidth = dataCols[0].length;
      while (trimmedWidth > 1) {
        const colIdx = trimmedWidth - 1;
        if (!bodyRowIndexes.every((ri) => isCellEmpty((dataCols[ri] || [])[colIdx]))) break;
        trimmedWidth--;
      }
      if (trimmedWidth < dataCols[0].length) {
        for (let i = 0; i < dataCols.length; i++) {
          dataCols[i] = dataCols[i].slice(0, trimmedWidth);
        }
      }
    }
  }

  return { labelCols, dataCols, labelSplitConfidence, labelColCount };
}

// ─── Title inference ──────────────────────────────────────────────────────────

/**
 * Checks whether the first row of a trimmed band is a standalone heading row.
 *
 * A row is title-like when:
 *   - it contains no numeric cells (pure text / empty)
 *   - it contains at least one non-empty non-numeric text cell
 *   - it is sparse (at most MAX_TITLE_ROW_NON_EMPTY_CELLS non-empty cells),
 *     consistent with a merged-like heading that spans the full width
 *
 * Returns { title, titleConfidence, titleSource } or null.
 */
function detectFirstRowTitle(values, band) {
  const { localStartRow, localTrimmedFirstCol, localTrimmedLastCol } = band;
  const firstRow = values[localStartRow];
  if (!firstRow) return null;

  let nonEmptyCount = 0;
  let hasText = false;

  for (let col = localTrimmedFirstCol; col <= localTrimmedLastCol; col++) {
    const cell = firstRow[col];
    if (isCellEmpty(cell)) continue;
    nonEmptyCount++;
    if (isNumericCell(cell)) return null; // any numeric cell disqualifies the row
    if (isTextOnlyCell(cell)) hasText = true;
  }

  if (!hasText || nonEmptyCount === 0 || nonEmptyCount > MAX_TITLE_ROW_NON_EMPTY_CELLS) {
    return null;
  }

  const title = firstNonEmptyNonNumericText(firstRow);
  if (!title) return null;

  return { title, titleConfidence: "high", titleSource: "firstRowOfBand" };
}

function inferTitle(values, band) {
  const { localStartRow } = band;

  if (localStartRow === 0) {
    return { title: "", titleConfidence: "none", titleSource: "sheetFallback" };
  }

  const rowAbove = values[localStartRow - 1];

  if (isRowEmpty(rowAbove)) {
    // Separator row above — look two rows up for a title.
    // Confidence is medium: the text could belong to a preceding section.
    if (localStartRow >= 2) {
      const text = firstNonEmptyNonNumericText(values[localStartRow - 2]);
      if (text) {
        return { title: text, titleConfidence: "medium", titleSource: "twoRowsAbove" };
      }
    }
  } else {
    // Row immediately above is non-empty — may be a title (medium confidence)
    const text = firstNonEmptyNonNumericText(rowAbove);
    if (text) {
      return { title: text, titleConfidence: "medium", titleSource: "rowAbove" };
    }
  }

  return { title: "", titleConfidence: "none", titleSource: "sheetFallback" };
}

// ─── Item builder ─────────────────────────────────────────────────────────────

// ─── Base subtype helpers ─────────────────────────────────────────────────────

/**
 * Maps a raw baseSubtype value to a human-readable display label.
 *
 *   "effective"  → "Effective Base"
 *   "unweighted" → "Unweighted Base"
 *   "weighted"   → "Weighted Base"
 *   undefined / null / anything else → "Base"
 */
function baseSubtypeToLabel(subtype) {
  if (subtype === "effective") return "Effective Base";
  if (subtype === "unweighted") return "Unweighted Base";
  if (subtype === "weighted") return "Weighted Base";
  return "Base";
}

/**
 * Issue codes that are informational/advisory and must not affect candidateStatus.
 *
 * These codes surface useful information in Check / Full Check / qualityIssueCodes
 * but should not downgrade an otherwise valid candidate to "uncertain".
 */
const ADVISORY_ISSUE_CODES = new Set([
  "WEIGHTED_BASE_FALLBACK",
  // BASE_BELOW_THRESHOLD reflects a runtime/data characteristic (small sample)
  // that does not indicate a table-structure problem.  It must remain visible
  // in qualityIssueCodes / Check / Full Check but must not downgrade an
  // otherwise valid candidate to "uncertain".
  "BASE_BELOW_THRESHOLD",
  // PREFERRED_BASE_NOT_FOUND fires when the user requested a specific base type
  // that was absent from the table.  The calculation still ran with the best
  // available base, so the table remains usable — the warning is informational.
  "PREFERRED_BASE_NOT_FOUND",
]);

/**
 * Derives a plain-language candidate status from model quality signals.
 *
 * "available"  — looks like a recognisable table with no blocking issues or
 *                uncertain boundaries; worth checking via Check Table.
 *                Also returned when labelSplitConfidence is "twoColumn" (two-column
 *                row labels with ordinal code + text answer are valid structures).
 * "uncertain"  — table-like but has blocking issues, quality warnings, or an
 *                ambiguous label/data boundary; Check Table may still work but
 *                results should be verified.
 * "rejected"   — no metric rows detected; unlikely to be a RIT research table.
 *
 * This replaces the former canRunSignificance flag which implied Run-readiness.
 * The scanner is a candidate finder only; Check Table is the authoritative step.
 */
function deriveCandidateStatus({ isLikelyTable, hasBlockingIssues, availabilityWarningCount, labelSplitConfidence }) {
  if (!isLikelyTable) return "rejected";
  if (hasBlockingIssues || labelSplitConfidence === "uncertain" || availabilityWarningCount > 0) return "uncertain";
  return "available";
}

function buildTableInventoryItem({ band, model, titleInfo, rangeAddress, sheetName, labelSplitConfidence, labelColCount }) {
  const { summary, qualitySummary, calculationBlocks, dataQualityIssues } = model;
  const tableId = sheetName + "!" + rangeAddress;

  const isLikelyTable = summary.detectedMetricRows > 0;
  // canRunCheckTable: true when the candidate looks table-like enough to pass to
  // Check Table. Does NOT imply the candidate is ready for Run significance.
  const canRunCheckTable = isLikelyTable;

  // Warnings that affect availability: exclude advisory codes (e.g. WEIGHTED_BASE_FALLBACK)
  // which are informational and must not downgrade an otherwise valid candidate.
  const availabilityWarningCount = (dataQualityIssues || []).filter(
    (i) => i.severity === "warning" && !ADVISORY_ISSUE_CODES.has(i.code)
  ).length;

  const candidateNotes = [];
  if (!isLikelyTable) {
    candidateNotes.push("Нет опознанных строк метрик");
  } else if (calculationBlocks.length === 0) {
    candidateNotes.push("Нет блоков расчёта");
  }
  if (isLikelyTable && qualitySummary.hasBlockingIssues) {
    candidateNotes.push("Критические проблемы качества");
  }
  if (labelSplitConfidence === "uncertain") {
    candidateNotes.push("Граница лейблов/данных не определена");
  }
  if (labelSplitConfidence === "twoColumn") {
    candidateNotes.push("Двухколоночные метки строк");
  }
  if (isLikelyTable && qualitySummary.warningCount > 0) {
    candidateNotes.push(`Предупреждений в превью: ${qualitySummary.warningCount}`);
  }

  const candidateStatus = deriveCandidateStatus({
    isLikelyTable,
    hasBlockingIssues: qualitySummary.hasBlockingIssues,
    availabilityWarningCount,
    labelSplitConfidence,
  });

  // Derive the selected base subtype label from calculationBlocks.
  // Collects the unique human-readable base type labels across all blocks that
  // have a base row, then joins them (multiple distinct subtypes are rare but
  // possible when a table has both proportion and mean blocks with different bases).
  const seenSubtypeLabels = [];
  const seenSubtypeSet = new Set();
  for (const block of calculationBlocks || []) {
    if (block.baseRowIndex != null) {
      const label = baseSubtypeToLabel(block.baseSubtype);
      if (!seenSubtypeSet.has(label)) {
        seenSubtypeSet.add(label);
        seenSubtypeLabels.push(label);
      }
    }
  }
  const selectedBaseSubtypeLabel = seenSubtypeLabels.length > 0 ? seenSubtypeLabels.join(", ") : "";

  let previewSummary = "";
  if (isLikelyTable) {
    const parts = [];
    if (summary.detectedBlocks > 0) parts.push("Блоков: " + summary.detectedBlocks);
    if (summary.baseRows > 0) parts.push("Баз: " + summary.baseRows);
    if (summary.hasNps) parts.push("NPS");
    if (summary.hasMeans) parts.push("Средние");
    previewSummary = parts.join(". ");
  }

  const columnCount = band.localTrimmedLastCol - band.localTrimmedFirstCol + 1;

  // Flat list of issue codes for diagnostic inspection.
  // Each entry: { code: string, severity: "warning"|"critical" }.
  const qualityIssueCodes = isLikelyTable
    ? (dataQualityIssues || []).map((i) => ({ code: i.code, severity: i.severity }))
    : [];

  // Full issue objects (critical + warning) sorted by severity then rowIndex.
  // Used for human-readable diagnostic details in Run report and Content check.
  const userVisibleIssues = isLikelyTable
    ? (dataQualityIssues || [])
        .filter((i) => i.severity === "critical" || i.severity === "warning")
        .sort((a, b) => {
          const rank = { critical: 0, warning: 1 };
          const sd = (rank[a.severity] ?? 99) - (rank[b.severity] ?? 99);
          if (sd !== 0) return sd;
          return (a.rowIndex ?? Number.MAX_SAFE_INTEGER) - (b.rowIndex ?? Number.MAX_SAFE_INTEGER);
        })
    : [];

  return {
    tableId,
    sheetName,
    rangeAddress,
    title: titleInfo.title,
    titleConfidence: titleInfo.titleConfidence,
    titleSource: titleInfo.titleSource,
    rowCount: summary.rowCount,
    columnCount,
    previewSummary,
    isLikelyTable,
    canRunCheckTable,
    candidateStatus,
    candidateNotes,
    labelSplitConfidence,
    labelColCount: labelColCount ?? null,
    warningsCount: isLikelyTable ? qualitySummary.warningCount : 0,
    criticalCount: isLikelyTable ? qualitySummary.criticalCount : 0,
    qualityIssueCodes,
    userVisibleIssues,
    availabilityWarningCount,
    hasBlockingIssues: qualitySummary.hasBlockingIssues,
    detectedMetricRows: summary.detectedMetricRows ?? 0,
    detectedBaseRows: summary.baseRows ?? 0,
    detectedBlocks: summary.detectedBlocks ?? 0,
    hasProportions: summary.hasProportions ?? false,
    hasNps: summary.hasNps ?? false,
    hasMeans: summary.hasMeans ?? false,
    selectedBaseSubtypeLabel,
  };
}

// ─── Main export ──────────────────────────────────────────────────────────────

/**
 * Scans a pre-loaded worksheet used range for candidate research table regions.
 *
 * @param {object} input
 * @param {Array}  input.values             - 2D array from usedRange.values
 * @param {number} input.usedRangeRowOffset - usedRange.rowIndex (zero-based sheet row)
 * @param {number} input.usedRangeColOffset - usedRange.columnIndex (zero-based sheet col)
 * @param {string} input.sheetName          - worksheet name
 * @returns {Array} TableInventoryItem[]
 */
export function scanWorksheetForTables({ values, usedRangeRowOffset, usedRangeColOffset, sheetName, settings }) {
  if (!Array.isArray(values) || values.length === 0 || !values[0]) {
    return [];
  }

  const rawBands = detectRowBands(values);
  const items = [];

  for (const rawBand of rawBands) {
    if (rawBand.localEndRow - rawBand.localStartRow < MIN_BAND_ROWS - 1) continue;

    const band = trimBandColumns(values, rawBand);
    if (!band) continue;

    // If the first row of the band is a generated backlink marker, exclude it.
    // Backlink rows are inserted by Content generation above detected tables and
    // must not be treated as title rows or body rows of the candidate table.
    // When settings.backlinkMarker is provided and matches the first content cell,
    // the active band starts one row later so rangeAddress points to the real table.
    // normalizeBacklinkItems detects the row above as "above-range" and updates
    // the backlink in-place on repeated Content generation (no duplicate rows).
    let activeBand = band;
    if (settings?.backlinkMarker) {
      const firstCell = (values[band.localStartRow] || [])[band.localTrimmedFirstCol];
      if (String(firstCell ?? "").trim() === settings.backlinkMarker) {
        activeBand = { ...band, localStartRow: band.localStartRow + 1 };
        if (activeBand.localEndRow < activeBand.localStartRow) continue;
      }
    }

    // Detect a merged-like title in the first row of the band.
    const firstRowTitleInfo = detectFirstRowTitle(values, activeBand);

    // bodyBand is what gets passed to the model: strip the title row if one was found.
    const bodyBand = firstRowTitleInfo
      ? { ...activeBand, localStartRow: activeBand.localStartRow + 1 }
      : activeBand;

    // If stripping the title row left nothing to interpret, skip.
    if (bodyBand.localEndRow < bodyBand.localStartRow) continue;

    // Pre-filter and label/data split operate on the body only.
    if (!hasNumericCell(values, bodyBand)) continue;

    const { labelCols, dataCols, labelSplitConfidence, labelColCount } = splitLabelData(values, bodyBand);

    if (!dataCols.length || !dataCols[0] || dataCols[0].length < 1) continue;

    const model = buildTablePreviewModel({ values: dataCols, leftLabelValues: labelCols, settings });

    // Range address covers the active band (backlink row excluded; title row included).
    const absRowStart = activeBand.localStartRow + usedRangeRowOffset;
    const absRowEnd = activeBand.localEndRow + usedRangeRowOffset;
    const absColStart = activeBand.localTrimmedFirstCol + usedRangeColOffset;
    const absColEnd = activeBand.localTrimmedLastCol + usedRangeColOffset;
    const rangeAddress = toA1Address(absRowStart, absRowEnd, absColStart, absColEnd);

    // Title: first-row detection takes priority; fall back to above-band lookback.
    const titleInfo = firstRowTitleInfo || inferTitle(values, activeBand);

    const item = buildTableInventoryItem({
      band: activeBand,
      model,
      titleInfo,
      rangeAddress,
      sheetName,
      labelSplitConfidence,
      labelColCount,
    });

    items.push(item);
  }

  return items;
}
