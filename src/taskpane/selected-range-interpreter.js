/* global Excel */

/**
 * Excel/taskpane adapter for selected-range interpretation.
 *
 * Shared by Run and Check. Both call interpretSelectedRange() and then
 * diverge only after interpretation: Run writes back to Excel; Check
 * builds a read-only preview model.
 *
 * Future direction: extract pure decomposition core that does not depend
 * on Office.js, then have this adapter call that core.  For now the
 * adapter remains at the Excel/taskpane level because it still needs
 * Office.js context for loading labels and banner rows from the sheet.
 */

import { removeSignificanceMarkersFromMatrix, generateSignificanceLabels } from "../core/significance";
import { LABEL_SCAN_COLUMNS_LEFT } from "../core/metric-detector";
import { normalizeSelectedRange } from "../core/range-normalizer";

const SELECTED_RANGE_GUARDRAIL_WARNING_TEXT =
  "Похоже, вы выделили лейблы строк или шапку вместе с данными. Сейчас RIT ожидает выделение только числовой части таблицы.";

// ─── Guardrail detection ───────────────────────────────────────────────────────

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
        text:
          selectedText && selectedText[rowIndex]
            ? selectedText[rowIndex][columnIndex]
            : undefined,
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

// ─── Embedded label column detection ──────────────────────────────────────────

/**
 * Returns true when every non-empty cell in `col` of `cleanedValues` is a
 * unit/indicator value.  Accepted forms:
 *   - numeric 0            (Excel .values for a 0%-formatted cell)
 *   - string "0" / "0.0"  (removeSignificanceMarkersFromText converts numbers
 *                           to strings via String(), so numeric 0 arrives as "0")
 *   - string "%"           (literal percent indicator text)
 *   - string "0%" / "0.0%" (explicit zero-percent string)
 * Any other value returns false so real data columns are not misclassified.
 */
function isEmbeddedUnitColumn(cleanedValues, col, rowCount) {
  let hasContent = false;
  for (let r = 0; r < rowCount; r++) {
    const row = cleanedValues[r] || [];
    const cell = row[col];
    if (cell === "" || cell === null || cell === undefined) continue;
    hasContent = true;
    if (typeof cell === "number" && cell === 0) continue;
    if (typeof cell === "string") {
      const t = cell.trim();
      if (t === "%" || /^0(\.0+)?%$/.test(t) || /^0(\.0+)?$/.test(t)) continue;
    }
    return false;
  }
  return hasContent;
}

/**
 * Scans the leftmost columns of cleanedValues for embedded label/unit columns.
 *
 * Col 0 must have at least one genuine-text cell (non-numeric, non-empty after
 * stripping %).  Each subsequent column is accepted only if it is a uniform
 * unit/indicator column (isEmbeddedUnitColumn).  Stops at the first column
 * that fails its test.  Always leaves at least one column as data.
 *
 * Returns 0 for strict numeric selections so the existing Run/Clear flow is
 * unchanged.
 */
export function detectEmbeddedLabelColumns(cleanedValues) {
  if (!Array.isArray(cleanedValues) || !Array.isArray(cleanedValues[0])) return 0;
  const colCount = cleanedValues[0].length;
  if (colCount < 2) return 0;

  const maxCols = Math.min(LABEL_SCAN_COLUMNS_LEFT, colCount - 1);
  const rowCount = cleanedValues.length;
  let embeddedLabelCols = 0;

  for (let col = 0; col < maxCols; col++) {
    if (col === 0) {
      const hasGenuineText = cleanedValues.some((row) => {
        const cell = row && row[col];
        if (cell === "" || cell === null || cell === undefined) return false;
        if (typeof cell === "number") return false;
        const s = String(cell).trim().replace("%", "").replace(",", ".");
        return s !== "" && Number.isNaN(Number(s));
      });
      if (hasGenuineText) {
        embeddedLabelCols++;
      } else {
        break;
      }
    } else {
      if (isEmbeddedUnitColumn(cleanedValues, col, rowCount)) {
        embeddedLabelCols++;
      } else {
        break;
      }
    }
  }

  return embeddedLabelCols;
}

/**
 * Counts consecutive all-blank columns at the start of selectedText.
 *
 * PURPOSE: Detects a leading empty/helper column that sits between the external
 * row-label column and the first real data column (e.g. a visual spacer in a
 * mean-only table).  detectEmbeddedLabelColumns only recognises columns with
 * genuine text or known unit indicators; a column whose every cell displays as
 * blank is left untouched and ends up in writeTargetRange, causing banner
 * letters to be written into the helper cell.
 *
 * Uses selectedText (the displayed cell text) rather than cleanedValues so
 * that cells which are visually empty but carry a non-empty underlying value
 * — such as 0 formatted as ";;;" — are also detected correctly.  A real data
 * column always displays something (e.g. "4.2", "150", "0%"); a helper column
 * shows nothing.
 *
 * RULE: a column qualifies when every cell's displayed text is "", null, or
 * undefined.  Always leaves at least one column as data (colCount >= 2 guard).
 * Only called when detectEmbeddedLabelColumns returns 0.
 */
export function detectLeadingEmptyColumns(selectedText) {
  if (!Array.isArray(selectedText) || !Array.isArray(selectedText[0])) return 0;
  const colCount = selectedText[0].length;
  if (colCount < 2) return 0;

  let emptyCols = 0;

  for (let col = 0; col < colCount - 1; col++) {
    const isColBlank = selectedText.every((row) => {
      const cell = row && row[col];
      return cell === "" || cell === null || cell === undefined;
    });

    if (isColBlank) {
      emptyCols++;
    } else {
      break;
    }
  }

  return emptyCols;
}

// ─── Banner context adapter ────────────────────────────────────────────────────

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

// ─── Banner context sanitization ──────────────────────────────────────────────

/**
 * Strips all trailing RIT significance markers from a single banner cell text,
 * looping until no marker remains.  For example:
 *   "2025Q4 (a) (b)"  →  "2025Q4 (a)"  →  "2025Q4"
 *   "Всё покупаю сам(а)"  →  unchanged (no whitespace before "(", so no match)
 *
 * Mirrors the getTrailingBannerMarker guard in taskpane.js: the "(X)" token
 * must be preceded by whitespace or appear at the very start of the string so
 * that mid-word parentheses such as "сам(а)" are never treated as markers.
 */
function stripAllTrailingBannerMarkersFromCell(rawText) {
  if (rawText === null || rawText === undefined) return "";
  const labels = generateSignificanceLabels();
  let result = String(rawText);
  for (;;) {
    const match = result.match(/(^|\s)\(([^()]*)\)\s*$/);
    if (!match) break;
    if (!labels.includes(match[2])) break;
    result = result.slice(0, match.index).trim();
  }
  return result;
}

/**
 * Returns a shallow copy of bannerContext where every banner cell value has
 * had all trailing RIT significance markers stripped.
 *
 * PURPOSE: prevents first-Run marker writes (e.g. "2025Q4 (a)") from
 * confusing banner detection or table-preview on subsequent Runs/Checks.
 * Without sanitization, isTechnicalWaveValueLabel returns false for marked
 * labels, mergeAdjacentWaveValueSpans no longer fires, and both data columns
 * in the first wave-value pair receive local label "a".
 *
 * Covers all banner row shapes:
 *   - lowerBannerRow / upperScanRows  (loadBannerContextForSelectedRange output)
 *   - scanRows                        (normalizeSelectedRange output, before
 *                                      buildRunBannerContext conversion)
 */
function sanitizeBannerContextForDetection(bannerContext) {
  if (!bannerContext) return bannerContext;
  const strip = stripAllTrailingBannerMarkersFromCell;
  const sanitized = { ...bannerContext };
  if (Array.isArray(sanitized.lowerBannerRow)) {
    sanitized.lowerBannerRow = sanitized.lowerBannerRow.map(strip);
  }
  if (Array.isArray(sanitized.upperScanRows)) {
    sanitized.upperScanRows = sanitized.upperScanRows.map((row) =>
      Array.isArray(row) ? row.map(strip) : row
    );
  }
  if (Array.isArray(sanitized.scanRows)) {
    sanitized.scanRows = sanitized.scanRows.map((row) =>
      Array.isArray(row) ? row.map(strip) : row
    );
  }
  return sanitized;
}

// ─── Office.js label and banner loading ───────────────────────────────────────

/**
 * Reads label values for selected range.
 *
 * By default, reads labels immediately to the left of selected data.
 * If labelsOnLeftSide is enabled, reads labels from the leftmost sheet columns.
 *
 * Exported for use by runMetricDetectionDiagnostics.
 */
export async function loadLabelValuesForSelectedRange(context, selectedRange, calculationSettings) {
  if (calculationSettings.labelsOnLeftSide) {
    return loadLabelsFromLeftSideOfSheet(context, selectedRange);
  }

  return loadLabelsImmediatelyLeftOfSelection(context, selectedRange);
}

/**
 * Reads labels immediately to the left of selected range.
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
 * Supports wide horizontal tables where metric labels stay on the far left,
 * while the user selects data columns far to the right.
 */
async function loadLabelsFromLeftSideOfSheet(context, selectedRange) {
  selectedRange.load(["rowIndex", "rowCount", "columnIndex"]);

  await context.sync();

  const selectedStartRowIndex = selectedRange.rowIndex;
  const selectedRowCount = selectedRange.rowCount;
  const selectedStartColumnIndex = selectedRange.columnIndex;

  if (selectedStartColumnIndex === 0) {
    return [];
  }

  const labelColumnCount = Math.min(LABEL_SCAN_COLUMNS_LEFT, selectedStartColumnIndex);

  const leftLabelRange = selectedRange.worksheet.getRangeByIndexes(
    selectedStartRowIndex,
    0,
    selectedRowCount,
    labelColumnCount
  );

  leftLabelRange.load("values");

  await context.sync();

  return leftLabelRange.values;
}

/**
 * Reads banner rows above selected range.
 *
 * Reads:
 * - lower banner row directly above selection;
 * - up to maxBannerScanRows rows above it.
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
          severity: "warning",
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

// ─── Main export ───────────────────────────────────────────────────────────────

/**
 * Interprets a selected Excel range and returns a stable interpretation object
 * shared by Run and Check.
 *
 * The caller must have already loaded selectedRange with at least:
 *   ["values", "text", "rowIndex", "columnIndex", "rowCount", "columnCount"]
 *
 * Return shape:
 *   state:                       "passThrough" | "normalized" | "blocked"
 *   cleanedValues:               marker-stripped values grid
 *   valuesForCalculation:        data body values (null when blocked)
 *   textForCalculation:          data body text  (null when blocked)
 *   leftLabelValues:             label column values (null when blocked)
 *   bannerContext:               banner context aligned to valuesForCalculation width
 *   writeTargetRange:            Excel Range for Run to write back into (null when blocked)
 *   dataRowOffset:               row offset into selectedRange where body starts
 *   dataColOffset:               column offset into selectedRange where body starts
 *   dataRowCount:                body row count
 *   dataColCount:                body column count
 *   normalizationStatusLines:    human-readable normalization summary for Run
 *   selectedRangeGuardrailWarnings: guardrail warning objects (empty for normalized/blocked)
 *   blockingMessage:             user-facing blocking reason (non-empty when blocked)
 *   blockingReasons:             blocking reason codes (non-empty when blocked)
 *   normalized:                  raw normalizeSelectedRange result (for Check verbose display)
 *
 * Invariants:
 *   - valuesForCalculation width === writeTargetRange column count (when non-null)
 *   - valuesForCalculation width === bannerContext.selectedColumnCount (when banner present)
 *   - label/unit/header columns are excluded from valuesForCalculation
 */
export async function interpretSelectedRange(
  context,
  selectedRange,
  selectedValues,
  selectedText,
  calculationSettings
) {
  const cleanedValues = removeSignificanceMarkersFromMatrix(selectedValues);
  const normalized = normalizeSelectedRange(cleanedValues, selectedText);

  // ── State 3: blocked ─────────────────────────────────────────────────────────
  if (normalized.normalizationNeeded && !normalized.normalizationApplied) {
    return {
      state: "blocked",
      cleanedValues,
      valuesForCalculation: null,
      textForCalculation: null,
      leftLabelValues: null,
      bannerContext: null,
      writeTargetRange: null,
      dataRowOffset: 0,
      dataColOffset: 0,
      dataRowCount: 0,
      dataColCount: 0,
      normalizationStatusLines: [],
      selectedRangeGuardrailWarnings: [],
      blockingMessage: normalized.blockingMessage,
      blockingReasons: normalized.blockingReasons,
      normalized,
    };
  }

  // ── State 2: broad selection successfully decomposed ─────────────────────────
  if (normalized.normalizationNeeded && normalized.normalizationApplied) {
    // Start with the normalizer's own partitioning.
    let valuesForCalculation = normalized.valuesForCalculation;
    let textForCalculation = normalized.textForCalculation;
    let effectiveDataColOffset = normalized.dataColOffset;
    let effectiveBannerContext = normalized.bannerContext;
    let embeddedLabelExtract = null;

    // Secondary embedded-label check — only when the normalizer left the label
    // column inside valuesForCalculation (dataColOffset === 0).  This happens
    // when the body-row text fraction for col 0 falls below the normalizer's
    // LABEL_NUMERIC_THRESHOLD (0.5), producing an "uncertain" split that the
    // normalizer conservatively resolves to labelColCount = 0.
    // detectEmbeddedLabelColumns is a lighter heuristic that catches these
    // remaining shapes by looking for at least one genuine text cell in col 0,
    // which is sufficient when the label column contains a mix of text labels
    // (e.g. "Agree", "Disagree") and numeric rows (e.g. "Base", means).
    if (normalized.dataColOffset === 0) {
      const additionalLabelCols = detectEmbeddedLabelColumns(
        normalized.valuesForCalculation
      );

      if (additionalLabelCols > 0) {
        embeddedLabelExtract = normalized.valuesForCalculation.map(
          (row) => row.slice(0, additionalLabelCols)
        );
        valuesForCalculation = normalized.valuesForCalculation.map(
          (row) => row.slice(additionalLabelCols)
        );
        textForCalculation = normalized.textForCalculation.map(
          (row) => row.slice(additionalLabelCols)
        );
        effectiveDataColOffset = additionalLabelCols;

        // Trim the normalizer's banner scan rows to the new data-only width so
        // bannerContext.selectedColumnCount === valuesForCalculation[0].length.
        if (
          normalized.bannerContext &&
          Array.isArray(normalized.bannerContext.scanRows)
        ) {
          effectiveBannerContext = {
            scanRows: normalized.bannerContext.scanRows.map(
              (row) => row.slice(additionalLabelCols)
            ),
            columnCount:
              normalized.bannerContext.columnCount - additionalLabelCols,
            messages: normalized.bannerContext.messages,
          };
        }
      }
    }

    // Tertiary check: strip leading all-visually-empty columns from the
    // effective normalized body.  The normalizer's label-column detection
    // recognises text-like leading columns (e.g. column A = "Среднее",
    // "Variance", "BASE") and sets dataColOffset accordingly, but it does
    // NOT strip helper/spacer columns that are all-blank (e.g. column B in a
    // mean-only table where the user selected A:N and the normalizer sets
    // dataColOffset=1 for column A, leaving column B as valuesForCalculation
    // column 0 with all-"" entries).
    //
    // Uses textForCalculation (displayed text) rather than valuesForCalculation
    // so that cells formatted as ";;;" (value hidden) are also treated as blank.
    // A real data column always renders something ("4.2", "150", "0%"); a helper
    // column renders nothing regardless of its underlying value.
    //
    // This strip MUST happen before dataBodyRange is built so that
    // writeTargetRange, bannerContext.selectedColumnCount, and
    // valuesForCalculation width are all aligned to real data columns.
    const additionalLeadingEmptyCols = detectLeadingEmptyColumns(textForCalculation);
    if (additionalLeadingEmptyCols > 0) {
      valuesForCalculation = valuesForCalculation.map(
        (row) => row.slice(additionalLeadingEmptyCols)
      );
      textForCalculation = textForCalculation.map(
        (row) => row.slice(additionalLeadingEmptyCols)
      );
      effectiveDataColOffset += additionalLeadingEmptyCols;

      // Keep effectiveBannerContext aligned to the new data width so
      // buildRunBannerContext produces the correct selectedColumnCount.
      // The scanRows shape is what the normalizer returns; the
      // lowerBannerRow/upperScanRows shape is produced by
      // buildRunBannerContext (should not appear here yet, but handled
      // defensively).
      if (effectiveBannerContext) {
        if (Array.isArray(effectiveBannerContext.scanRows)) {
          effectiveBannerContext = {
            ...effectiveBannerContext,
            scanRows: effectiveBannerContext.scanRows.map((row) =>
              Array.isArray(row) ? row.slice(additionalLeadingEmptyCols) : row
            ),
            columnCount:
              (effectiveBannerContext.columnCount || 0) - additionalLeadingEmptyCols,
          };
        } else if (Array.isArray(effectiveBannerContext.lowerBannerRow)) {
          effectiveBannerContext = {
            ...effectiveBannerContext,
            lowerBannerRow: effectiveBannerContext.lowerBannerRow.slice(
              additionalLeadingEmptyCols
            ),
            upperScanRows: (effectiveBannerContext.upperScanRows || []).map((row) =>
              Array.isArray(row) ? row.slice(additionalLeadingEmptyCols) : row
            ),
            selectedColumnCount:
              (effectiveBannerContext.selectedColumnCount || 0) -
              additionalLeadingEmptyCols,
          };
        }
      }
    }

    const dataBodyRange = selectedRange
      .getCell(normalized.dataRowOffset, effectiveDataColOffset)
      .getResizedRange(
        valuesForCalculation.length - 1,
        valuesForCalculation[0].length - 1
      );

    // Label values: prefer what the normalizer extracted, then the embedded-
    // label extract from the secondary check above, then load externally.
    let leftLabelValues = normalized.leftLabelValues;

    if (!Array.isArray(leftLabelValues) || leftLabelValues.length === 0) {
      if (embeddedLabelExtract !== null) {
        // Secondary detection stripped these columns from valuesForCalculation.
        leftLabelValues = embeddedLabelExtract;
      } else {
        // Fallback: normalization did not extract label columns (selection
        // excluded the label column). Load labels from the worksheet using the
        // data body range so row alignment matches the normalized values.
        leftLabelValues = await loadLabelValuesForSelectedRange(
          context,
          dataBodyRange,
          calculationSettings
        );
      }
    }

    // Banner context: prefer banner rows detected inside the selection (normalizer
    // already slices them to data columns only, and the secondary check above
    // further trims them when effectiveBannerContext differs from
    // normalized.bannerContext). When none were found — the banner is above the
    // selection in the sheet — fall back to loading from the worksheet using the
    // data body range. effectiveDataColOffset ensures selectedColumnCount aligns
    // to valuesForCalculation[0].length.
    let bannerContext = sanitizeBannerContextForDetection(
      buildRunBannerContext(effectiveBannerContext)
    );
    if (bannerContext === null && calculationSettings.respectBannerStructure) {
      bannerContext = sanitizeBannerContextForDetection(
        await loadBannerContextForSelectedRange(context, dataBodyRange, calculationSettings)
      );
    }

    return {
      state: "normalized",
      cleanedValues,
      valuesForCalculation,
      textForCalculation,
      leftLabelValues,
      bannerContext,
      writeTargetRange: dataBodyRange,
      dataRowOffset: normalized.dataRowOffset,
      dataColOffset: effectiveDataColOffset,
      dataRowCount: valuesForCalculation.length,
      dataColCount: valuesForCalculation[0].length,
      normalizationStatusLines: [
        "Диапазон нормализован: расчёт выполнен только по области данных.",
      ],
      selectedRangeGuardrailWarnings: [],
      blockingMessage: "",
      blockingReasons: [],
      normalized,
    };
  }

  // ── State 1: pass-through (numeric-only selection) ───────────────────────────

  // Check for embedded label/unit columns before loading external labels.
  const embeddedLabelCols = detectEmbeddedLabelColumns(cleanedValues);
  // When there are no embedded text/unit label columns, check for a leading
  // all-empty helper column (common in mean-only tables where a visual spacer
  // sits between the external row-label column and the first data column).
  // detectEmbeddedLabelColumns does not catch these because they contain no
  // genuine text or unit indicators, so they would otherwise be included in
  // writeTargetRange and receive a spurious banner letter.
  const leadingEmptyCols = embeddedLabelCols === 0
    ? detectLeadingEmptyColumns(selectedText)
    : 0;


  let valuesForCalculation;
  let textForCalculation;
  let leftLabelValues;
  let writeTargetRange;

  if (embeddedLabelCols > 0) {
    leftLabelValues = cleanedValues.map((row) => row.slice(0, embeddedLabelCols));
    valuesForCalculation = cleanedValues.map((row) => row.slice(embeddedLabelCols));
    textForCalculation = selectedText.map((row) => row.slice(embeddedLabelCols));
    writeTargetRange = selectedRange
      .getCell(0, embeddedLabelCols)
      .getResizedRange(
        selectedRange.rowCount - 1,
        selectedRange.columnCount - embeddedLabelCols - 1
      );
  } else if (leadingEmptyCols > 0) {
    // Strip leading empty helper columns from the write target and calculation
    // range so that banner context, significance calculation, and banner-letter
    // placement are all aligned to the real data columns.  Labels must still be
    // loaded from the worksheet because the empty columns carry no label text.
    valuesForCalculation = cleanedValues.map((row) => row.slice(leadingEmptyCols));
    textForCalculation = selectedText.map((row) => row.slice(leadingEmptyCols));
    writeTargetRange = selectedRange
      .getCell(0, leadingEmptyCols)
      .getResizedRange(
        selectedRange.rowCount - 1,
        selectedRange.columnCount - leadingEmptyCols - 1
      );
    leftLabelValues = await loadLabelValuesForSelectedRange(
      context,
      selectedRange,
      calculationSettings
    );
  } else {
    leftLabelValues = await loadLabelValuesForSelectedRange(
      context,
      selectedRange,
      calculationSettings
    );
    valuesForCalculation = cleanedValues;
    textForCalculation = selectedText;
    writeTargetRange = selectedRange;
  }

  // When banner-aware settings are on, load banner context from rows above the
  // data range. When embedded label columns were extracted, align the banner
  // range to the data columns only — otherwise selectedColumnCount includes the
  // label column(s) and detectBannerStructure misaligns groups.
  // Sanitize immediately so any RIT markers written by a previous Run are
  // stripped before the context is used by either detectBannerStructure (Run)
  // or buildTablePreviewModel (Check).
  let bannerContext = null;
  if (calculationSettings.respectBannerStructure) {
    bannerContext = sanitizeBannerContextForDetection(
      await loadBannerContextForSelectedRange(context, writeTargetRange, calculationSettings)
    );
  }

  const selectedRangeGuardrailWarnings = detectSelectedRangeGuardrails(
    selectedText,
    cleanedValues
  );

  const dataColCount =
    valuesForCalculation.length > 0 && Array.isArray(valuesForCalculation[0])
      ? valuesForCalculation[0].length
      : 0;

  return {
    state: "passThrough",
    cleanedValues,
    valuesForCalculation,
    textForCalculation,
    leftLabelValues,
    bannerContext,
    writeTargetRange,
    dataRowOffset: 0,
    dataColOffset: embeddedLabelCols > 0 ? embeddedLabelCols : leadingEmptyCols,
    dataRowCount: valuesForCalculation.length,
    dataColCount,
    normalizationStatusLines: [],
    selectedRangeGuardrailWarnings,
    blockingMessage: "",
    blockingReasons: [],
    normalized,
  };
}
