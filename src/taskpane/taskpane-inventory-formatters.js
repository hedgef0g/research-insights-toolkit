/**
 * Pure inventory / Content display helpers.
 *
 * No Office.js dependencies, no Excel.run / context.sync calls.
 * All functions map already-computed inventory data to display text or
 * plain data structures.
 */

// The backlink marker and its detector live in core (generated-rows.js) so
// core detection logic can treat generated backlink rows as table boundaries
// without depending on the taskpane layer. Imported (for local use inside this
// module) and re-exported to keep existing taskpane imports working from a
// single source of truth. A bare `export { ... } from` would create no local
// binding, so calls within this file would throw ReferenceError.
import { BACKLINK_MARKER, isGeneratedBacklinkRow } from "../core/generated-rows.js";

export { BACKLINK_MARKER, isGeneratedBacklinkRow };

export function getInventoryCandidateStatusLabel(candidateStatus) {
  if (candidateStatus === "available") {
    return "Кандидат — рекомендуется «Проверить таблицу»";
  }

  if (candidateStatus === "uncertain") {
    return "Кандидат неопределён — требуется «Проверить таблицу»";
  }

  return "Не опознан как таблица ResearchSignal";
}

/**
 * User-facing status label for Content sheet output.
 * Uses warningsCount / criticalCount to produce plain-language readiness labels
 * instead of internal candidate-finder wording.
 */
export function getContentCandidateStatusLabel(item) {
  const status = item.candidateStatus;
  const warnings = item.warningsCount ?? 0;
  const criticals = item.criticalCount ?? 0;

  if (status === "available") {
    return warnings > 0 ? "Есть предупреждения" : "Готово к расчёту";
  }

  if (status === "uncertain") {
    return criticals > 0 ? "Есть критические проблемы" : "Нужна проверка";
  }

  return "Пропущено";
}

export function formatInventoryItemLines(item, index) {
  const lines = [];
  // Prefer resolvedTitle (set after backlink normalization) over raw title.
  // Fall back to raw title only if it is not a generated backlink marker.
  const displayTitle =
    item.resolvedTitle ||
    (item.title && !isGeneratedBacklinkRow(item.title) ? item.title : null);
  // Prefer resolvedRangeAddress (adjusted for any inserted backlink rows).
  const displayRange = item.resolvedRangeAddress || item.rangeAddress;
  const header = displayTitle ? `${index}. ${displayTitle} — ${displayRange}` : `${index}. ${displayRange}`;

  lines.push(header);
  lines.push(`   ${item.rowCount} строк, ${item.columnCount} колонок.`);

  if (item.previewSummary) {
    lines.push(`   ${item.previewSummary}.`);
  }

  if (item.selectedBaseSubtypeLabel) {
    lines.push(`   База: ${item.selectedBaseSubtypeLabel}.`);
  }

  const warnParts = [];
  if (item.criticalCount > 0) warnParts.push(`Критических: ${item.criticalCount}`);
  if (item.warningsCount > 0) warnParts.push(`Предупреждений: ${item.warningsCount}`);
  if (warnParts.length > 0) lines.push(`   ${warnParts.join(". ")}.`);

  // candidateStatus replaces the former "Значимость: да/нет" line.
  // Inventory is a candidate finder only; Check Table is the authoritative step.
  lines.push(`   ${getInventoryCandidateStatusLabel(item.candidateStatus)}.`);

  if (item.candidateNotes && item.candidateNotes.length > 0) {
    lines.push(`   [${item.candidateNotes.join("; ")}]`);
  }

  return lines;
}

/**
 * Returns true when a candidate title is a known total/banner/header-like label
 * that should not be used as a table title in Content output.
 *
 * The scanner's detectFirstRowTitle fires on any sparse all-text first row —
 * including a single-cell "Всего" or "Total" banner — and assigns
 * titleConfidence "high".  This guard catches those labels before they reach
 * the Content display layer.
 *
 * The list is intentionally conservative: only unambiguous aggregate/total
 * words that cannot be a real research-table title on their own.
 *
 * @param {string} title - Raw title string from the scanner item.
 * @returns {boolean}
 */
export function isContentTitleFallbackLabel(title) {
  if (!title) return false;
  const normalized = String(title).trim().toLowerCase();
  // prettier-ignore
  const TOTAL_LIKE_LABELS = new Set([
    "всего",        // Russian "All / Total"
    "итого",        // Russian "Grand total / Sum"
    "total",        // English
    "grand total",  // English compound
    "overall",      // English
    "all",          // English
  ]);
  return TOTAL_LIKE_LABELS.has(normalized);
}

/**
 * Returns the Content-sheet display title for a detected table candidate.
 *
 * Only trusts a scanner/resolved title when:
 *   1. titleConfidence === "high" (first row of the detected band was a
 *      dedicated sparse heading row, not inferred from rows above), AND
 *   2. The title is not a known total/banner-like label (e.g. "Всего",
 *      "Total") that the scanner can legitimately detect as a sparse
 *      first-row title but that is not a real table heading.
 *
 * Falls back to "Таблица N" for:
 *   - titleConfidence !== "high" (medium / none)
 *   - title is a known total/banner-like label
 *   - title equals the generated backlink marker
 *   - title is empty after all checks
 *
 * @param {object} item   - TableInventoryItem (may have resolvedTitle set by normalizeBacklinkItems).
 * @param {number} index  - 1-based candidate number within the Content output.
 * @returns {string}
 */
export function resolveContentDisplayTitle(item, index) {
  const fallback = `Таблица ${index}`;

  // Medium/none confidence titles come from rows above the band and may be
  // banner or header text, not real table headings — always use the fallback.
  if (item.titleConfidence !== "high") {
    return fallback;
  }

  // High-confidence path: prefer post-backlink resolved title, then raw title.
  const title =
    item.resolvedTitle ||
    (isGeneratedBacklinkRow(item.title) ? "" : (item.title || ""));

  // Even high-confidence scanner titles can be total/banner-like labels
  // (e.g. a lone "Всего" cell at the top of a band passes detectFirstRowTitle).
  if (isContentTitleFallbackLabel(title)) {
    return fallback;
  }

  return title || fallback;
}

export function buildClientContentRows(sheetResults) {
  const rows = [];
  let index = 1;
  for (const sheetResult of sheetResults) {
    for (const item of sheetResult.items) {
      rows.push([index, resolveContentDisplayTitle(item, index), "", item.sheetName || sheetResult.sheetName]);
      index++;
    }
  }
  if (rows.length === 0) {
    rows.push(["", "Кандидаты не обнаружены", "", ""]);
  }
  return rows;
}

export function getContentTableHyperlinkTarget(sheetName, rangeAddress) {
  if (!sheetName || !rangeAddress) return null;
  const escaped = sheetName.replace(/'/g, "''");
  const needsQuotes = /[^A-Za-z0-9_]/.test(escaped);
  const quotedSheet = needsQuotes ? `'${escaped}'` : escaped;
  return `${quotedSheet}!${rangeAddress}`;
}

export function getContentRowReference(contentSheetName, contentRow) {
  if (!contentSheetName || !contentRow) return null;
  const escaped = contentSheetName.replace(/'/g, "''");
  const needsQuotes = /[^A-Za-z0-9_]/.test(escaped);
  const quotedSheet = needsQuotes ? `'${escaped}'` : escaped;
  return `${quotedSheet}!A${contentRow}`;
}
