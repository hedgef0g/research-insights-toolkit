/**
 * Pure helpers for scanner-driven batch actions.
 *
 * Filters workbook inventory results into eligible candidates and
 * pre-execution skips, with stable reason codes for each skip.
 *
 * Office.js-free and side-effect-free. Reusable by Run, Clear, and Check.
 */

// ─── Batch skip reason codes ──────────────────────────────────────────────────

/**
 * Stable reason codes for candidates skipped before execution begins.
 * Values must not change; callers map them to display text.
 */
export const BATCH_SKIP_REASONS = Object.freeze({
  /** Candidate has no usable range address after backlink normalization. */
  MISSING_RANGE: "missing_range",
  /** Candidate status is "uncertain" — label/data boundary is ambiguous. */
  CANDIDATE_UNCERTAIN: "candidate_uncertain",
  /** Candidate status is "rejected" — not recognized as a research table. */
  CANDIDATE_REJECTED: "candidate_rejected",
  /** Candidate has an unrecognized or unsupported status. */
  UNKNOWN_STATUS: "unknown_status",
});

// ─── Candidate filter ─────────────────────────────────────────────────────────

/**
 * Partitions workbook inventory results into eligible and skipped candidates.
 *
 * Eligible:  candidateStatus === "available" && canRunCheckTable && has usable range address
 * Skipped:   uncertain, rejected, or missing-range candidates (with reason code)
 * Ignored:   generated sheets (Content, Run report, etc.) — excluded entirely, not counted toward skipped
 *
 * @param {object} inventoryResults
 * @param {Array<{ sheetName: string, items: Array }>} inventoryResults.sheetResults
 * @param {object} [options]
 * @param {string} [options.contentSheetName="Content"]  Single sheet name to exclude (legacy; ignored when generatedSheetNames is provided).
 * @param {Set<string>|string[]} [options.generatedSheetNames]  Set/array of all generated sheet names to exclude entirely.
 * @returns {{ eligible: Array, skipped: Array }}
 *
 * Each eligible entry: { sheetName, rangeAddress, title }
 * Each skipped entry:  { sheetName, rangeAddress, reason, status }
 */
export function filterWorkbookCandidates(
  inventoryResults,
  { contentSheetName = "Content", generatedSheetNames = null } = {}
) {
  const excluded =
    generatedSheetNames != null ? new Set(generatedSheetNames) : new Set([contentSheetName]);

  const eligible = [];
  const skipped = [];

  for (const sheetResult of inventoryResults.sheetResults) {
    if (excluded.has(sheetResult.sheetName)) {
      continue;
    }
    for (const item of sheetResult.items) {
      const rangeAddr = item.resolvedRangeAddress || item.rangeAddress || null;

      if (!rangeAddr) {
        skipped.push({
          sheetName: sheetResult.sheetName,
          rangeAddress: item.rangeAddress || null,
          reason: BATCH_SKIP_REASONS.MISSING_RANGE,
          status: item.candidateStatus || null,
        });
      } else if (item.candidateStatus === "uncertain") {
        skipped.push({
          sheetName: sheetResult.sheetName,
          rangeAddress: rangeAddr,
          reason: BATCH_SKIP_REASONS.CANDIDATE_UNCERTAIN,
          status: "uncertain",
        });
      } else if (item.candidateStatus === "rejected") {
        skipped.push({
          sheetName: sheetResult.sheetName,
          rangeAddress: rangeAddr,
          reason: BATCH_SKIP_REASONS.CANDIDATE_REJECTED,
          status: "rejected",
        });
      } else if (item.candidateStatus === "available" && item.canRunCheckTable) {
        eligible.push({
          sheetName: sheetResult.sheetName,
          rangeAddress: rangeAddr,
          title: item.resolvedTitle || item.title || "",
        });
      } else {
        skipped.push({
          sheetName: sheetResult.sheetName,
          rangeAddress: rangeAddr,
          reason: BATCH_SKIP_REASONS.UNKNOWN_STATUS,
          status: item.candidateStatus || null,
        });
      }
    }
  }

  return { eligible, skipped };
}
