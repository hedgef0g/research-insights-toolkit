import { BATCH_SKIP_REASONS } from "../core/batch-candidate-filter";

export function runReportSkipReasonLabel(reason) {
  switch (reason) {
    case BATCH_SKIP_REASONS.MISSING_RANGE:      return "Нет диапазона";
    case BATCH_SKIP_REASONS.CANDIDATE_UNCERTAIN: return "Кандидат неопределён";
    case BATCH_SKIP_REASONS.CANDIDATE_REJECTED:  return "Не опознан как таблица";
    default:                                      return "Неизвестный статус";
  }
}

export function runReportStatusLabel(status) {
  switch (status) {
    case "processed": return "Обработано";
    case "checked":   return "Проверено";
    case "skipped":   return "Пропущено";
    case "blocked":   return "Пропущено";
    case "error":     return "Ошибка";
    default:          return status || "";
  }
}

export function runReportMetricTypes(item) {
  if (!item) return "";
  const parts = [];
  if (item.hasProportions) parts.push("Пропорции");
  if (item.hasMeans)       parts.push("Средние");
  if (item.hasNps)         parts.push("NPS");
  return parts.join(", ");
}

/**
 * Formats issue details from an inventory item into a human-readable string.
 *
 * Prefers userVisibleIssues (full issue objects with message and location) when
 * available. Falls back to qualityIssueCodes (code-only identifiers) for
 * backward compatibility.
 *
 * Does NOT include candidateNotes — callers that need them append separately.
 */
export function formatIssueDetailsForReport(item) {
  if (!item) return "";
  if (item.userVisibleIssues && item.userVisibleIssues.length > 0) {
    return item.userVisibleIssues.map((iss) => `[${iss.severity}] ${iss.message}`).join("; ");
  }
  if (item.qualityIssueCodes && item.qualityIssueCodes.length > 0) {
    return item.qualityIssueCodes.map((q) => q.code).join(", ");
  }
  return "";
}

/**
 * Builds a human-readable warning-detail string for the Run report Details column.
 *
 * Uses formatIssueDetailsForReport() for issue messages, then appends
 * candidateNotes (structural candidate diagnostics).
 */
export function runReportWarningDetails(item) {
  if (!item) return "";
  const parts = [];
  const issueDetails = formatIssueDetailsForReport(item);
  if (issueDetails) parts.push(issueDetails);
  if (item.candidateNotes && item.candidateNotes.length > 0) {
    parts.push(...item.candidateNotes);
  }
  return parts.join("; ");
}
