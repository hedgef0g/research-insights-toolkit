import { describe, it } from "node:test";
import assert from "node:assert/strict";

import {
  SCOPES,
  ACTIONS,
  SKIP_REASONS,
  normalizeAction,
  normalizeScope,
  makeProcessedResult,
  makeClearedResult,
  makeCheckedResult,
  makeSkippedResult,
  makeBlockedResult,
  makeErrorResult,
} from "../../src/core/action-constants.js";

// ─── Constants ────────────────────────────────────────────────────────────────

describe("SCOPES", () => {
  it("contains the three required scope values", () => {
    assert.strictEqual(SCOPES.CURRENT_TABLE, "current_table");
    assert.strictEqual(SCOPES.CURRENT_SHEET, "current_sheet");
    assert.strictEqual(SCOPES.WHOLE_WORKBOOK, "whole_workbook");
  });

  it("is frozen", () => {
    assert.ok(Object.isFrozen(SCOPES));
  });
});

describe("ACTIONS", () => {
  it("contains the three required action values", () => {
    assert.strictEqual(ACTIONS.RUN, "run");
    assert.strictEqual(ACTIONS.CLEAR, "clear");
    assert.strictEqual(ACTIONS.CHECK, "check");
  });

  it("is frozen", () => {
    assert.ok(Object.isFrozen(ACTIONS));
  });
});

describe("SKIP_REASONS", () => {
  it("is frozen", () => {
    assert.ok(Object.isFrozen(SKIP_REASONS));
  });

  it("contains expected reason codes", () => {
    assert.ok(typeof SKIP_REASONS.TOO_LITTLE_DATA === "string");
    assert.ok(typeof SKIP_REASONS.NO_CALCULATION_BLOCKS === "string");
    assert.ok(typeof SKIP_REASONS.BLOCKED_SELECTION === "string");
    assert.ok(typeof SKIP_REASONS.EMPTY_RANGE === "string");
  });
});

// ─── normalizeAction ──────────────────────────────────────────────────────────

describe("normalizeAction", () => {
  it("accepts lowercase valid actions", () => {
    assert.strictEqual(normalizeAction("run"), "run");
    assert.strictEqual(normalizeAction("clear"), "clear");
    assert.strictEqual(normalizeAction("check"), "check");
  });

  it("normalizes to lowercase", () => {
    assert.strictEqual(normalizeAction("RUN"), "run");
    assert.strictEqual(normalizeAction("Clear"), "clear");
    assert.strictEqual(normalizeAction("CHECK"), "check");
  });

  it("trims surrounding whitespace", () => {
    assert.strictEqual(normalizeAction("  run  "), "run");
  });

  it("returns null for unknown values", () => {
    assert.strictEqual(normalizeAction("scan"), null);
    assert.strictEqual(normalizeAction(""), null);
    assert.strictEqual(normalizeAction("write"), null);
  });

  it("returns null for null and undefined", () => {
    assert.strictEqual(normalizeAction(null), null);
    assert.strictEqual(normalizeAction(undefined), null);
  });
});

// ─── normalizeScope ───────────────────────────────────────────────────────────

describe("normalizeScope", () => {
  it("accepts lowercase valid scopes", () => {
    assert.strictEqual(normalizeScope("current_table"), "current_table");
    assert.strictEqual(normalizeScope("current_sheet"), "current_sheet");
    assert.strictEqual(normalizeScope("whole_workbook"), "whole_workbook");
  });

  it("normalizes to lowercase", () => {
    assert.strictEqual(normalizeScope("CURRENT_TABLE"), "current_table");
    assert.strictEqual(normalizeScope("Current_Sheet"), "current_sheet");
  });

  it("trims surrounding whitespace", () => {
    assert.strictEqual(normalizeScope("  whole_workbook  "), "whole_workbook");
  });

  it("returns null for unknown values", () => {
    assert.strictEqual(normalizeScope("worksheet"), null);
    assert.strictEqual(normalizeScope(""), null);
    assert.strictEqual(normalizeScope("all"), null);
  });

  it("returns null for null and undefined", () => {
    assert.strictEqual(normalizeScope(null), null);
    assert.strictEqual(normalizeScope(undefined), null);
  });
});

// ─── Result factories ─────────────────────────────────────────────────────────

describe("makeProcessedResult", () => {
  it("returns status processed with rangeAddress and blocksProcessed", () => {
    const r = makeProcessedResult("A1:D10", 3, "done");
    assert.strictEqual(r.status, "processed");
    assert.strictEqual(r.rangeAddress, "A1:D10");
    assert.strictEqual(r.blocksProcessed, 3);
    assert.strictEqual(r.message, "done");
  });

  it("defaults blocksProcessed to null and message to empty string", () => {
    const r = makeProcessedResult("A1:D10");
    assert.strictEqual(r.blocksProcessed, null);
    assert.strictEqual(r.message, "");
  });
});

describe("makeClearedResult", () => {
  it("returns status cleared", () => {
    const r = makeClearedResult("B2:E8", "cleared");
    assert.strictEqual(r.status, "cleared");
    assert.strictEqual(r.rangeAddress, "B2:E8");
    assert.strictEqual(r.message, "cleared");
  });

  it("defaults rangeAddress to null and message to empty string", () => {
    const r = makeClearedResult();
    assert.strictEqual(r.rangeAddress, null);
    assert.strictEqual(r.message, "");
  });
});

describe("makeCheckedResult", () => {
  it("returns status checked", () => {
    const r = makeCheckedResult("A1:D10", "ok");
    assert.strictEqual(r.status, "checked");
    assert.strictEqual(r.rangeAddress, "A1:D10");
    assert.strictEqual(r.message, "ok");
  });

  it("defaults rangeAddress to null and message to empty string", () => {
    const r = makeCheckedResult();
    assert.strictEqual(r.rangeAddress, null);
    assert.strictEqual(r.message, "");
  });
});

describe("makeSkippedResult", () => {
  it("returns status skipped", () => {
    const r = makeSkippedResult("A1:D10", "нет данных");
    assert.strictEqual(r.status, "skipped");
    assert.strictEqual(r.rangeAddress, "A1:D10");
    assert.strictEqual(r.message, "нет данных");
  });

  it("defaults work correctly", () => {
    const r = makeSkippedResult();
    assert.strictEqual(r.rangeAddress, null);
    assert.strictEqual(r.message, "");
  });
});

describe("makeBlockedResult", () => {
  it("returns status blocked", () => {
    const r = makeBlockedResult("A1:Z100", "unsafe selection");
    assert.strictEqual(r.status, "blocked");
    assert.strictEqual(r.rangeAddress, "A1:Z100");
    assert.strictEqual(r.message, "unsafe selection");
  });
});

describe("makeErrorResult", () => {
  it("returns status error", () => {
    const r = makeErrorResult("A1:D10", "exception");
    assert.strictEqual(r.status, "error");
    assert.strictEqual(r.rangeAddress, "A1:D10");
    assert.strictEqual(r.message, "exception");
  });
});
