/**
 * Significance settings footnote helpers for Research Insights Toolkit.
 *
 * Pure, Office.js-free helpers shared by the taskpane (which inserts/updates
 * footnote rows after a Run) and the table inventory scanner (which must treat
 * a generated footnote row as a separator, not as table content).
 *
 * A generated footnote row is prefixed with an invisible marker so it can be
 * recognised on later scans/runs without being visible to the user. Detection
 * uses startsWith(SIGNIFICANCE_FOOTNOTE_MARKER); the visible text follows the
 * marker.
 *
 * The marker letter order is intentionally fixed and must not be changed.
 */

// Invisible generated-row prefix: U+2063 INVISIBLE SEPARATOR and U+2060 WORD
// JOINER. These render as nothing, so the footnote shows only its visible text.
export const SIGNIFICANCE_FOOTNOTE_MARKER = "⁣⁣⁠⁣⁠";

/**
 * Returns true when a cell value is a generated RIT significance footnote row.
 * Uses startsWith so the visible footnote text after the marker is irrelevant.
 */
export function isGeneratedSignificanceFootnoteRow(cellValue) {
  if (cellValue === null || cellValue === undefined) return false;
  return String(cellValue).startsWith(SIGNIFICANCE_FOOTNOTE_MARKER);
}

// Human-readable Russian statistic labels keyed by calculation-block metricType.
// Determined from the actual blocks used in a table, not a global assumption.
const STATISTIC_TYPE_LABELS = {
  proportion: "Z-критерий для долей",
  mean: "t-тест для средних",
  npsStructure: "Z-критерий для NPS",
  npsSpread: "t-тест для NPS",
};

// Stable presentation order so repeated runs produce identical footnote text.
const STATISTIC_TYPE_ORDER = ["proportion", "mean", "npsStructure", "npsSpread"];

/**
 * Collects the distinct statistic-type labels actually used by a table's
 * calculation blocks, in a stable order. Unknown metric types are ignored.
 *
 * @param {Array<{metricType?: string}>} calculationBlocks
 * @returns {string[]} ordered, de-duplicated statistic labels
 */
export function collectStatisticTypeLabels(calculationBlocks) {
  const seen = new Set();
  for (const block of calculationBlocks || []) {
    if (block && block.metricType) {
      seen.add(block.metricType);
    }
  }
  const labels = [];
  for (const type of STATISTIC_TYPE_ORDER) {
    if (seen.has(type)) labels.push(STATISTIC_TYPE_LABELS[type]);
  }
  return labels;
}

/**
 * Builds the visible footnote text (without the invisible marker prefix).
 *
 * Example: "Уровень значимости: 95%; тест: двусторонний; статистика: Z-критерий для долей."
 *
 * When a processed-scope detail is supplied (Manual Run only), it is appended
 * after the base text, e.g. "... Z-критерий для долей. Обработано: B12:F34."
 *
 * @param {object} options
 * @param {string|number} options.confidenceLevel - e.g. "95"
 * @param {boolean} options.oneTailedTest         - true → one-sided, false → two-sided
 * @param {string[]} options.statisticLabels      - labels from collectStatisticTypeLabels
 * @param {string} [options.scopeDetail]          - optional processed-scope suffix
 *                                                  (already including a leading space),
 *                                                  e.g. " Обработано: B12:F34."
 * @returns {string}
 */
export function buildSignificanceFootnoteVisibleText({
  confidenceLevel,
  oneTailedTest,
  statisticLabels,
  scopeDetail,
}) {
  const level = String(confidenceLevel ?? "").trim() || "95";
  const testLabel = oneTailedTest ? "односторонний" : "двусторонний";
  const labels = Array.isArray(statisticLabels) ? statisticLabels.filter(Boolean) : [];
  const statText = labels.length > 0 ? labels.join(", ") : "не определена";
  const base = `Уровень значимости: ${level}%; тест: ${testLabel}; статистика: ${statText}.`;
  return appendFootnoteScopeDetail(base, scopeDetail);
}

/**
 * Strips a leading sheet prefix (e.g. "Sheet1!" or "'My Sheet'!") from an A1
 * address so the footnote shows a concise local range. Returns "" for empty
 * or non-string input.
 *
 * @param {string} rangeAddress
 * @returns {string}
 */
function toLocalA1Range(rangeAddress) {
  if (typeof rangeAddress !== "string") return "";
  const trimmed = rangeAddress.trim();
  if (!trimmed) return "";
  const sep = trimmed.lastIndexOf("!");
  return sep >= 0 ? trimmed.slice(sep + 1) : trimmed;
}

/**
 * Appends a processed-scope detail suffix to existing footnote visible text.
 * A falsy/blank suffix leaves the base text unchanged, so auto-run footnotes
 * and any caller that passes no scope detail are byte-for-byte identical.
 *
 * @param {string} visibleText - base footnote text
 * @param {string} [scopeDetail] - suffix including its own leading separator
 * @returns {string}
 */
export function appendFootnoteScopeDetail(visibleText, scopeDetail) {
  const base = String(visibleText ?? "");
  if (!scopeDetail || !String(scopeDetail).trim()) return base;
  return base + scopeDetail;
}

/**
 * Builds the processed-range footnote suffix from an Excel range address.
 * The sheet prefix is dropped so the detail stays concise. Returns "" when no
 * usable range is available (caller then appends nothing).
 *
 * Example: "B12:F34" → " Обработано: B12:F34."
 *
 * @param {string} rangeAddress - actual processed/write-target range address
 * @returns {string}
 */
export function buildProcessedRangeFootnoteSuffix(rangeAddress) {
  const local = toLocalA1Range(rangeAddress);
  if (!local) return "";
  return ` Обработано: ${local}.`;
}

/**
 * Builds the full footnote cell value: invisible marker prefix + visible text.
 *
 * @param {object} options - see buildSignificanceFootnoteVisibleText
 * @returns {string}
 */
export function buildSignificanceFootnoteCellValue(options) {
  return SIGNIFICANCE_FOOTNOTE_MARKER + buildSignificanceFootnoteVisibleText(options);
}

// ─── Footnote placement (pure) ──────────────────────────────────────────────
//
// A generated footnote must be idempotent: re-running (Manual or Auto) over a
// table that already has one must REPLACE it, never duplicate it. It must also
// sit BELOW any ordinary user note rows ("Все респонденты", etc.) that belong
// to the table, and must never overwrite them.
//
// The placement decision is pure: the Office-bound caller loads a small window
// of rows below the table and hands the values matrix here. Detection of an
// existing generated footnote is strictly marker-based, so ordinary notes are
// never mistaken for a generated row.

// How many rows below the table body to scan when deciding placement. Small and
// bounded so the scan never walks unboundedly into the sheet.
export const FOOTNOTE_SCAN_WINDOW_ROWS = 4;

/**
 * True when a scanned cell value is blank (null / undefined / whitespace-only).
 */
export function isFootnoteScanCellBlank(cellValue) {
  return cellValue === null || cellValue === undefined || String(cellValue).trim() === "";
}

/**
 * Classifies one scanned row below a table. `dataColStartOffset` is the index,
 * within the scanned row array, at which the table's DATA columns begin (cells
 * left of it are the table's label/margin columns).
 *
 * - "marker": contains a generated RIT footnote (marker-based, authoritative).
 * - "blank":  no populated cells.
 * - "note":   populated only in the label/margin columns → an ordinary user
 *             note/footnote belonging to the table (e.g. "Все респонденты").
 * - "table":  populated in the data region → the start of another table.
 *
 * @param {Array<*>} rowValues
 * @param {number} dataColStartOffset
 * @returns {"marker"|"blank"|"note"|"table"}
 */
export function classifyFootnoteScanRow(rowValues, dataColStartOffset = 0) {
  if (!Array.isArray(rowValues)) return "blank";
  if (rowValues.some((cell) => isGeneratedSignificanceFootnoteRow(cell))) return "marker";

  let anyPopulated = false;
  for (let i = 0; i < rowValues.length; i++) {
    if (isFootnoteScanCellBlank(rowValues[i])) continue;
    anyPopulated = true;
    if (i >= dataColStartOffset) return "table";
  }
  return anyPopulated ? "note" : "blank";
}

/**
 * Resolves where a single table's generated footnote should be written.
 *
 * `scanRows[i]` is the loaded value row at absolute index
 * `firstRowBelowTable + i` (row immediately below the table body is i = 0).
 * `dataColStartOffset` is passed through to classifyFootnoteScanRow.
 *
 * Algorithm (top-down, stopping at the first blank or next-table boundary):
 *  - an existing generated footnote (marker row) reached before any blank/table
 *    is updated IN PLACE — no row is inserted;
 *  - consecutive ordinary note rows immediately below the table are skipped over
 *    so the generated footnote is inserted BELOW them;
 *  - otherwise the footnote is inserted at the first row below the table (after
 *    any trailing notes), never overwriting existing content.
 *
 * @param {Array<Array<*>>} scanRows
 * @param {number} firstRowBelowTable - absolute index of scanRows[0]
 * @param {number} dataColStartOffset
 * @returns {{ mode: "update"|"insert", rowIndex: number }}
 */
export function resolveFootnotePlacement(scanRows, firstRowBelowTable, dataColStartOffset = 0) {
  const rows = Array.isArray(scanRows) ? scanRows : [];
  let insertOffset = 0;

  for (let i = 0; i < rows.length; i++) {
    const kind = classifyFootnoteScanRow(rows[i], dataColStartOffset);
    if (kind === "marker") {
      return { mode: "update", rowIndex: firstRowBelowTable + i };
    }
    if (kind === "note") {
      insertOffset = i + 1;
      continue;
    }
    // "blank" or "table": the trailing note area (if any) has ended.
    break;
  }

  return { mode: "insert", rowIndex: firstRowBelowTable + insertOffset };
}
