import { describe, it } from "node:test";
import assert from "node:assert/strict";

import {
  BATCH_PROGRESS_BAR_WIDTH,
  appendSelectedRangeGuardrailMessages,
  buildBatchProgressBar,
  buildBatchProgressStatus,
  buildCheckResolverMessage,
  formatBannerUserMessages,
  formatBannerUserMessagesExcludingCodes,
  formatSelectedRangeGuardrailMessages,
  formatStatusWithSelectedRangeGuardrails,
  nonContiguousSelectionMessage,
  runningStatusMessage,
} from "../../src/taskpane/taskpane-status.js";

const translate = (key) =>
  ({
    "status.nonContiguousSelection": "localized non-contiguous selection",
    "status.resolverNoTable": "localized no table",
    "status.resolverGeneratedSheet": "localized generated sheet",
    "status.resolverAmbiguousBoundary": "localized ambiguous boundary",
    "status.resolverBlocked": "localized blocked",
    "status.resolverFallback": "localized fallback",
  })[key] || `missing:${key}`;

describe("runningStatusMessage", () => {
  it("keeps current Russian run/check/content wording", () => {
    assert.strictEqual(
      runningStatusMessage("run", "table"),
      "Расчёт начат… Идёт обработка текущей таблицы."
    );
    assert.strictEqual(
      runningStatusMessage("check", "workbook"),
      "Проверка начата… Идёт обработка таблиц в книге."
    );
    assert.strictEqual(runningStatusMessage("content"), "Оглавление создаётся…");
  });
});

describe("buildBatchProgressBar", () => {
  it("builds a fixed-width progress bar and clamps percent", () => {
    assert.strictEqual(BATCH_PROGRESS_BAR_WIDTH, 20);
    assert.strictEqual(buildBatchProgressBar(50, 10), "[█████░░░░░]");
    assert.strictEqual(buildBatchProgressBar(150, 4), "[████]");
    assert.strictEqual(buildBatchProgressBar(-20, 4), "[░░░░]");
  });
});

describe("buildBatchProgressStatus", () => {
  it("formats live batch progress text", () => {
    assert.strictEqual(
      buildBatchProgressStatus({
        action: "run",
        scope: "workbook",
        currentIndex: 2,
        total: 5,
        sheetName: "Wave 1",
      }),
      [
        "Расчёт — вся книга",
        "[████████░░░░░░░░░░░░] 40%",
        "Таблица 2 из 5 · Лист: Wave 1",
        "Осталось: 3",
      ].join("\n")
    );
  });

  it("bounds the current index to the total", () => {
    assert.strictEqual(
      buildBatchProgressStatus({
        action: "clear",
        scope: "sheet",
        currentIndex: 9,
        total: 4,
      }),
      [
        "Очистка — текущий лист",
        "[████████████████████] 100%",
        "Таблица 4 из 4 · Лист: —",
        "Осталось: 0",
      ].join("\n")
    );
  });
});

describe("banner status message helpers", () => {
  const bannerStructure = {
    messages: [
      { code: "GLOBAL_TOTAL_USED", text: "Global total used." },
      { code: "INTERNAL_ONLY", text: "Internal message." },
      { code: "BANNER_MALFORMED_STRUCTURE", text: "Banner malformed." },
    ],
  };

  it("formats only user-visible banner messages", () => {
    assert.strictEqual(
      formatBannerUserMessages(bannerStructure),
      ["Сообщения:", "- Global total used.", "- Banner malformed."].join("\n")
    );
  });

  it("can exclude specific visible banner codes", () => {
    assert.strictEqual(
      formatBannerUserMessagesExcludingCodes(bannerStructure, ["GLOBAL_TOTAL_USED"]),
      "Banner malformed."
    );
  });
});

describe("selected range guardrail message helpers", () => {
  const warnings = [{ text: "Keep selected data only." }, { text: "Keep selected data only." }, { text: "" }];

  it("deduplicates warning text", () => {
    assert.strictEqual(formatSelectedRangeGuardrailMessages(warnings), "Keep selected data only.");
  });

  it("appends guardrail text after a blank separator", () => {
    assert.deepStrictEqual(appendSelectedRangeGuardrailMessages(["Done."], warnings), [
      "Done.",
      "",
      "Keep selected data only.",
    ]);
    assert.strictEqual(
      formatStatusWithSelectedRangeGuardrails("Done.", warnings),
      "Done.\n\nKeep selected data only."
    );
  });
});

describe("localized status helpers", () => {
  it("uses the caller-provided translator for non-contiguous selection messages", () => {
    assert.strictEqual(
      nonContiguousSelectionMessage(translate),
      "localized non-contiguous selection"
    );
  });

  it("uses resolver messages before translating fallback statuses", () => {
    assert.strictEqual(
      buildCheckResolverMessage({ status: "blocked", message: "Specific resolver message." }, translate),
      "Specific resolver message."
    );
  });

  it("uses the caller-provided translator for resolver status fallbacks", () => {
    assert.strictEqual(
      buildCheckResolverMessage({ status: "ambiguous-boundary" }, translate),
      "localized ambiguous boundary"
    );
    assert.strictEqual(buildCheckResolverMessage({ status: "future" }, translate), "localized fallback");
  });
});
