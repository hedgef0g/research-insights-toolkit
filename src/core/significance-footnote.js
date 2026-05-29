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
 * @param {object} options
 * @param {string|number} options.confidenceLevel - e.g. "95"
 * @param {boolean} options.oneTailedTest         - true → one-sided, false → two-sided
 * @param {string[]} options.statisticLabels      - labels from collectStatisticTypeLabels
 * @returns {string}
 */
export function buildSignificanceFootnoteVisibleText({ confidenceLevel, oneTailedTest, statisticLabels }) {
  const level = String(confidenceLevel ?? "").trim() || "95";
  const testLabel = oneTailedTest ? "односторонний" : "двусторонний";
  const labels = Array.isArray(statisticLabels) ? statisticLabels.filter(Boolean) : [];
  const statText = labels.length > 0 ? labels.join(", ") : "не определена";
  return `Уровень значимости: ${level}%; тест: ${testLabel}; статистика: ${statText}.`;
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
