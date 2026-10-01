import { BATCH_SKIP_REASONS } from "../core/batch-candidate-filter";

/**
 * Formats a pre-execution skipped candidate entry into a Russian-language
 * detail line for the status panel. Used by runAutoSignificance and
 * clearAutoSignificance when building their skipped-candidate summaries.
 */
export function formatSkippedCandidateDetail({ sheetName, rangeAddress, reason, status }) {
  const addr = rangeAddress || "?";
  const label = `- ${sheetName} ${addr}`;
  switch (reason) {
    case BATCH_SKIP_REASONS.MISSING_RANGE:
      return `${label}: пропущено — нет диапазона`;
    case BATCH_SKIP_REASONS.CANDIDATE_UNCERTAIN:
      return `${label}: пропущено — кандидат неопределён`;
    case BATCH_SKIP_REASONS.CANDIDATE_REJECTED:
      return `${label}: пропущено — не опознан как таблица ResearchSignal`;
    default:
      return `${label}: пропущено — статус «${status || "unknown"}»`;
  }
}

export function checkMetricTypesFromBlocks(calculationBlocks) {
  if (!calculationBlocks || calculationBlocks.length === 0) return "";
  let hasProportions = false,
    hasMeans = false,
    hasNps = false;
  for (const block of calculationBlocks) {
    if (block.metricType === "proportion") hasProportions = true;
    else if (block.metricType === "mean") hasMeans = true;
    else if (block.metricType === "nps") hasNps = true;
  }
  const parts = [];
  if (hasProportions) parts.push("Пропорции");
  if (hasMeans) parts.push("Средние");
  if (hasNps) parts.push("NPS");
  return parts.join(", ");
}

export function formatCheckUserVisibleIssues(issues) {
  if (!Array.isArray(issues) || issues.length === 0) {
    return [];
  }

  const lines = ["Проблемы проверки:"];
  for (const issue of issues) {
    lines.push(`- [${issue.severity}] ${issue.message}`);
  }
  return lines;
}

/**
 * Builds a compact calculation-block summary for the Check table text output.
 *
 * Returns an array of text lines. One header line ("Блоки расчёта: N.") is
 * followed by one bullet per block. If no blocks are detected, returns a
 * single "Блоки расчёта: не обнаружены." line.
 */
export function formatCheckCalculationBlocks(calculationBlocks, rowDiagnostics) {
  if (!Array.isArray(calculationBlocks) || calculationBlocks.length === 0) {
    return ["Блоки расчёта: не обнаружены."];
  }

  const labelMap = new Map();
  if (Array.isArray(rowDiagnostics)) {
    for (const diag of rowDiagnostics) {
      if (diag && diag.rowIndex != null) {
        labelMap.set(diag.rowIndex, diag.primaryLabel || "");
      }
    }
  }

  const rowRef = (rowIndex) => {
    if (rowIndex == null) return null;
    const label = labelMap.get(rowIndex);
    return label ? `стр. ${rowIndex + 1} «${label}»` : `стр. ${rowIndex + 1}`;
  };

  const baseSubtypeNote = (subtype) => {
    if (subtype === "effective") return " [эффективная]";
    if (subtype === "unweighted") return " [невзвешенная]";
    if (subtype === "weighted") return " [взвешенная — ПРИМЕЧАНИЕ: резервный вариант]";
    return "";
  };

  const baseRef = (block) => {
    if (block.baseRowIndex == null) return "";
    return ` База: ${rowRef(block.baseRowIndex)}${baseSubtypeNote(block.baseSubtype)}.`;
  };

  const lines = [`Блоки расчёта: ${calculationBlocks.length}.`];

  for (const block of calculationBlocks) {
    switch (block.metricType) {
      case "proportion": {
        const count = Array.isArray(block.valueRowIndexes) ? block.valueRowIndexes.length : 0;
        lines.push(`- Пропорции: строк со значениями: ${count}.${baseRef(block)}`);
        break;
      }
      case "mean": {
        const meanRef = block.valueRowIndex != null ? rowRef(block.valueRowIndex) : "нет";
        const sdPart = block.sdRowIndex != null ? ` СО: ${rowRef(block.sdRowIndex)}.` : "";
        const varPart =
          block.varianceRowIndex != null ? ` Дисперсия: ${rowRef(block.varianceRowIndex)}.` : "";
        lines.push(`- Среднее: ${meanRef}.${sdPart}${varPart}${baseRef(block)}`);
        break;
      }
      case "npsStructure": {
        const npsRef = block.valueRowIndex != null ? rowRef(block.valueRowIndex) : "нет";
        const promPart =
          block.promotersRowIndex != null ? ` Промоутеры: ${rowRef(block.promotersRowIndex)}.` : "";
        const neutPart =
          block.neutralRowIndex != null ? ` Нейтральные: ${rowRef(block.neutralRowIndex)}.` : "";
        const detPart =
          block.detractorsRowIndex != null ? ` Критики: ${rowRef(block.detractorsRowIndex)}.` : "";
        lines.push(`- NPS: ${npsRef}.${promPart}${neutPart}${detPart}${baseRef(block)}`);
        break;
      }
      case "npsSpread": {
        const npsRef = block.valueRowIndex != null ? rowRef(block.valueRowIndex) : "нет";
        const sdPart = block.sdRowIndex != null ? ` СО: ${rowRef(block.sdRowIndex)}.` : "";
        const varPart =
          block.varianceRowIndex != null ? ` Дисперсия: ${rowRef(block.varianceRowIndex)}.` : "";
        lines.push(`- NPS (разброс): ${npsRef}.${sdPart}${varPart}${baseRef(block)}`);
        break;
      }
      default:
        lines.push(`- Блок «${block.metricType || "unknown"}»`);
    }
  }

  return lines;
}

/**
 * Builds a compact banner summary for the Check table text output.
 *
 * Surfaces enough of bannerStructure for manual validation that Check
 * consumes the same banner-aware interpretation inputs as Run.
 * Returns an array of text lines (empty if banner-aware was disabled).
 */
export function formatCheckBannerSummary(bannerStructure) {
  if (!bannerStructure || bannerStructure.isEnabled !== true) {
    return ["Баннер: не проверялся (учёт структуры выключен)."];
  }

  if (!bannerStructure.isDetected) {
    return ["Баннер: не обнаружен."];
  }

  const groups = Array.isArray(bannerStructure.groups) ? bannerStructure.groups : [];
  const waveGroupCount = groups.filter((group) => group && group.semanticType === "wave").length;
  const localTotalGroupCount = groups.filter((group) => group && group.hasLocalTotal).length;
  const hasGlobalTotal =
    bannerStructure.globalTotalColumnIndex !== null &&
    bannerStructure.globalTotalColumnIndex !== undefined;

  const headerParts = [`Баннер: обнаружен. Групп: ${groups.length}.`];
  headerParts.push(`Wave-групп: ${waveGroupCount}.`);
  headerParts.push(`Local Total: ${localTotalGroupCount > 0 ? "да" : "нет"}.`);
  headerParts.push(
    hasGlobalTotal
      ? `Global Total: колонка данных ${bannerStructure.globalTotalColumnIndex + 1}.`
      : "Global Total: нет."
  );
  if (bannerStructure.recommendedComparisonMode) {
    headerParts.push(`Режим сравнения: ${bannerStructure.recommendedComparisonMode}.`);
  }

  const lines = [headerParts.join(" ")];

  if (groups.length > 0) {
    lines.push("Группы:");
    for (const group of groups) {
      const label = group && group.label ? group.label : "(без названия)";
      const columnIndexes = Array.isArray(group && group.columnIndexes) ? group.columnIndexes : [];
      const cols = columnIndexes.length;
      const semantic = group && group.semanticType ? group.semanticType : "default";
      const mode =
        group && group.recommendedComparisonMode ? group.recommendedComparisonMode : "default";
      const totalNote = group && group.hasLocalTotal ? ", local Total" : "";
      lines.push(`- ${label} — колонок ${cols} / ${mode} / ${semantic}${totalNote}`);
    }
  }

  return lines;
}
