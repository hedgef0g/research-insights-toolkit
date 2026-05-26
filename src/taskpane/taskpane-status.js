/* global document */

import { t } from "./localization";

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

// Shown when the user has a non-contiguous (multi-area) selection active.
// context.workbook.getSelectedRange() throws a RichApi.Error for such selections.
// Resolved via t() at call sites so the message respects the active UI language.
export function nonContiguousSelectionMessage() {
  return t("status.nonContiguousSelection");
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

export function buildCheckResolverMessage(resolverResult) {
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
