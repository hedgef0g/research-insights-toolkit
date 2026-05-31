/**
 * Core-safe detection of GENERATED RIT rows.
 *
 * PURPOSE:
 * Some rows in a worksheet are written by RIT itself (not real table content)
 * and must be treated as hard table boundaries by detection logic — e.g. the
 * above-Base fallback must never cross them.
 *
 * This module is Office.js-free and lives in core so core detection code can
 * recognise generated rows without depending on the taskpane/UI layer. The
 * significance-footnote marker has its own module (significance-footnote.js);
 * this module owns the Content backlink marker. The taskpane re-exports the
 * backlink helpers from here so there is a single source of truth.
 */

// Visible Content-backlink marker written into the first cell of a generated
// "back to table of contents" row.
export const BACKLINK_MARKER = "← Оглавление";

/**
 * Returns true when a cell value is a generated RIT backlink row.
 * Trimmed exact match, so ordinary content is never mistaken for a backlink.
 */
export function isGeneratedBacklinkRow(cellValue) {
  if (cellValue === null || cellValue === undefined) return false;
  return String(cellValue).trim() === BACKLINK_MARKER;
}
