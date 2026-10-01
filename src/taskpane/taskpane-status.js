/* global document */

// Banner message codes that should be surfaced to the user.
// Moved here with the formatBannerUserMessages* helpers that consume it.
const USER_VISIBLE_BANNER_MESSAGE_CODES = new Set([
  "GLOBAL_TOTAL_USED",
  "BANNER_AUTO_PREVIOUS_COLUMN_APPLIED",
  "BANNER_TOTAL_ONLY_NO_TOTAL_PAIRS",
  "BANNER_MULTIPLE_LOCAL_TOTALS",
  "BANNER_TOTAL_OUTSIDE_SELECTION",
  "BANNER_MALFORMED_STRUCTURE",
  "BANNER_NO_ROWS_ABOVE_SELECTION",
]);

export function setStatusMessage(message) {
  const statusPanel = document.getElementById("status-panel");
  const outputElement = document.getElementById("significance-result");

  if (statusPanel) {
    statusPanel.style.display = "block";
  }

  if (outputElement) {
    outputElement.textContent = message || "";
  }
}

export function setCheckMessage(message) {
  const checkPanel = document.getElementById("check-panel");
  const checkResult = document.getElementById("check-result");

  if (checkPanel) {
    checkPanel.style.display = "block";
  }

  if (checkResult) {
    checkResult.textContent = message || "";
  }
}

export function setInventoryMessage(message) {
  const panel = document.getElementById("inventory-panel");
  const result = document.getElementById("inventory-result");
  if (panel) panel.style.display = "block";
  if (result) result.textContent = message || "";
}

function resolveTranslator(translate) {
  return typeof translate === "function" ? translate : (key) => key;
}

// Shown when the user has a non-contiguous (multi-area) selection active.
// context.workbook.getSelectedRange() throws a RichApi.Error for such selections.
// Resolved via the caller-provided translator so the active UI language is used.
export function nonContiguousSelectionMessage(translate) {
  return resolveTranslator(translate)("status.nonContiguousSelection");
}

export function formatBannerUserMessages(bannerStructure) {
  if (!bannerStructure || !bannerStructure.messages) {
    return "";
  }

  const visibleMessages = bannerStructure.messages.filter((message) =>
    USER_VISIBLE_BANNER_MESSAGE_CODES.has(message.code)
  );

  if (visibleMessages.length === 0) {
    return "";
  }

  if (visibleMessages.length === 1) {
    return visibleMessages[0].text;
  }

  return ["Сообщения:", ...visibleMessages.map((message) => `- ${message.text}`)].join("\n");
}

export function formatBannerUserMessagesExcludingCodes(bannerStructure, excludedCodes = []) {
  if (!bannerStructure || !bannerStructure.messages) {
    return "";
  }

  const excludedCodeSet = new Set(excludedCodes);

  const visibleMessages = bannerStructure.messages.filter(
    (message) =>
      USER_VISIBLE_BANNER_MESSAGE_CODES.has(message.code) && !excludedCodeSet.has(message.code)
  );

  if (visibleMessages.length === 0) {
    return "";
  }

  if (visibleMessages.length === 1) {
    return visibleMessages[0].text;
  }

  return ["Сообщения:", ...visibleMessages.map((message) => `- ${message.text}`)].join("\n");
}

export function formatSelectedRangeGuardrailMessages(warnings) {
  if (!warnings || warnings.length === 0) {
    return "";
  }

  const uniqueTexts = Array.from(new Set(warnings.map((warning) => warning.text).filter(Boolean)));

  if (uniqueTexts.length === 1) {
    return uniqueTexts[0];
  }

  return ["Предупреждения:", ...uniqueTexts.map((text) => `- ${text}`)].join("\n");
}

export function appendSelectedRangeGuardrailMessages(statusMessages, warnings) {
  const guardrailMessage = formatSelectedRangeGuardrailMessages(warnings);

  if (!guardrailMessage) {
    return statusMessages;
  }

  return [...statusMessages, "", guardrailMessage];
}

export function formatStatusWithSelectedRangeGuardrails(message, warnings) {
  return appendSelectedRangeGuardrailMessages([message], warnings).join("\n");
}

export function runningStatusMessage(action, scope) {
  const prefix = {
    run:     "Расчёт начат…",
    clear:   "Очистка начата…",
    check:   "Проверка начата…",
    content: "Оглавление создаётся…",
  }[action];
  const detail = {
    table:   "Идёт обработка текущей таблицы.",
    sheet:   "Идёт обработка таблиц на листе.",
    workbook:"Идёт обработка таблиц в книге.",
  }[scope] || "";
  return detail ? `${prefix} ${detail}` : prefix;
}

// Width (in characters) of the text progress bar used by buildBatchProgressStatus.
export const BATCH_PROGRESS_BAR_WIDTH = 20;

/**
 * Build a plain-text progress bar like [████░░░░░░░░░░░░░░░░].
 *
 * Returns plain text only (no markup) because setStatusMessage writes
 * textContent. The percent is clamped to 0..100 before filling.
 */
export function buildBatchProgressBar(percent, width = BATCH_PROGRESS_BAR_WIDTH) {
  const safeWidth = Number.isFinite(width) && width > 0 ? Math.round(width) : BATCH_PROGRESS_BAR_WIDTH;
  const clamped = Math.max(0, Math.min(100, Number.isFinite(percent) ? percent : 0));
  const filled = Math.round((clamped / 100) * safeWidth);
  return `[${"█".repeat(filled)}${"░".repeat(safeWidth - filled)}]`;
}

/**
 * Build a live progress status string for heavy Run/Clear batch operations.
 *
 * Pure helper — returns plain text only. Shows the current table index, total,
 * percent complete, current sheet name, a text progress bar and the remaining
 * count. Uses a 1-based currentIndex and Math.round(currentIndex / total * 100).
 *
 * @param {object} opts
 * @param {"run"|"clear"} opts.action
 * @param {"sheet"|"workbook"} opts.scope
 * @param {number} opts.currentIndex 1-based index of the table just processed.
 * @param {number} opts.total Total number of tables in the batch.
 * @param {string} [opts.sheetName] Sheet name of the current table.
 */
export function buildBatchProgressStatus({ action, scope, currentIndex, total, sheetName } = {}) {
  const safeTotal = Number.isFinite(total) && total > 0 ? Math.floor(total) : 0;
  const rawIndex = Number.isFinite(currentIndex) && currentIndex > 0 ? Math.floor(currentIndex) : 0;
  const boundedIndex = safeTotal > 0 ? Math.min(rawIndex, safeTotal) : rawIndex;
  const percent = safeTotal > 0 ? Math.round((boundedIndex / safeTotal) * 100) : 0;
  const remaining = safeTotal > 0 ? Math.max(0, safeTotal - boundedIndex) : 0;

  const prefix =
    {
      run: "Расчёт",
      clear: "Очистка",
    }[action] || "Обработка";
  const scopeLabel =
    {
      sheet: "текущий лист",
      workbook: "вся книга",
    }[scope] || "";

  const headline = scopeLabel ? `${prefix} — ${scopeLabel}` : prefix;
  const bar = buildBatchProgressBar(percent);

  return [
    headline,
    `${bar} ${percent}%`,
    `Таблица ${boundedIndex} из ${safeTotal} · Лист: ${sheetName || "—"}`,
    `Осталось: ${remaining}`,
  ].join("\n");
}

export function buildCheckResolverMessage(resolverResult = {}, translate) {
  const t = resolveTranslator(translate);
  if (resolverResult.message) return resolverResult.message;
  switch (resolverResult.status) {
    case "no-table":
      return t("status.resolverNoTable");
    case "generated-sheet":
      return t("status.resolverGeneratedSheet");
    case "ambiguous-boundary":
      return t("status.resolverAmbiguousBoundary");
    case "blocked":
      return t("status.resolverBlocked");
    default:
      return t("status.resolverFallback");
  }
}
