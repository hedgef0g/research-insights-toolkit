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

// Upper bounds that keep the processed-groups detail concise. Beyond these the
// list is considered too noisy to be useful and the caller should fall back to
// the processed range instead.
const MAX_FOOTNOTE_GROUP_LABELS = 8;
const MAX_FOOTNOTE_GROUPS_TEXT_LENGTH = 160;

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
 * Builds the processed-banner-groups footnote suffix from a list of raw group
 * labels (one per processed column, in column order). Labels are trimmed,
 * de-duplicated (preserving first-seen order) and validated:
 *
 * - empty / whitespace-only labels are dropped;
 * - at least two distinct meaningful labels are required (a single label is not
 *   informative enough to replace the raw range);
 * - if the de-duplicated list is too long or too noisy, "" is returned so the
 *   caller falls back to the processed range.
 *
 * Hierarchical labels containing " / " (a banner path) switch the wording to
 * "Обработаны группы баннера: ... ; ..." to match the nested presentation.
 *
 * Example: ["Мужской", "Женский", "Total"] → " Обработаны группы: Мужской, Женский, Total."
 *
 * @param {string[]} groupLabels
 * @returns {string} suffix with a leading space, or "" to signal fallback
 */
export function buildProcessedBannerGroupsFootnoteSuffix(groupLabels) {
  if (!Array.isArray(groupLabels)) return "";

  const seen = new Set();
  const labels = [];
  for (const raw of groupLabels) {
    const label = String(raw ?? "").trim();
    if (!label) continue;
    if (seen.has(label)) continue;
    seen.add(label);
    labels.push(label);
  }

  if (labels.length < 2) return "";
  if (labels.length > MAX_FOOTNOTE_GROUP_LABELS) return "";

  const hierarchical = labels.some((label) => label.includes(" / "));
  const joined = labels.join(hierarchical ? "; " : ", ");
  if (joined.length > MAX_FOOTNOTE_GROUPS_TEXT_LENGTH) return "";

  const lead = hierarchical ? "Обработаны группы баннера" : "Обработаны группы";
  return ` ${lead}: ${joined}.`;
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
