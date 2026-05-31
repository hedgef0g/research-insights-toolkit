import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  BACKLINK_MARKER,
  isGeneratedBacklinkRow,
  resolveContentDisplayTitle,
  formatInventoryItemLines,
} from "../../src/taskpane/taskpane-inventory-formatters.js";

// Regression guard: these helpers call isGeneratedBacklinkRow internally. A bare
// `export { ... } from` re-export creates no local binding, so the calls would
// throw ReferenceError at runtime. The local-binding tests below would catch it.

describe("taskpane-inventory-formatters — backlink helper binding", () => {
  it("re-exports a usable backlink marker and detector", () => {
    assert.strictEqual(BACKLINK_MARKER, "← Оглавление");
    assert.strictEqual(isGeneratedBacklinkRow(BACKLINK_MARKER), true);
    assert.strictEqual(isGeneratedBacklinkRow("Agree"), false);
  });
});

describe("resolveContentDisplayTitle — generated backlink title", () => {
  it("returns the fallback and does not throw when title is the backlink marker", () => {
    const title = resolveContentDisplayTitle(
      { title: BACKLINK_MARKER, titleConfidence: "high" },
      3
    );
    assert.strictEqual(title, "Таблица 3");
  });

  it("still returns a real high-confidence title", () => {
    const title = resolveContentDisplayTitle(
      { title: "Удовлетворённость", titleConfidence: "high" },
      1
    );
    assert.strictEqual(title, "Удовлетворённость");
  });
});

describe("formatInventoryItemLines — generated backlink title", () => {
  it("does not throw and does not use the backlink marker as the display title", () => {
    const lines = formatInventoryItemLines(
      {
        title: BACKLINK_MARKER,
        rangeAddress: "Sheet1!A1:D5",
        rowCount: 5,
        columnCount: 4,
        candidateStatus: "available",
      },
      2
    );
    // Header must fall back to the range only — the marker must not appear.
    assert.strictEqual(lines[0], "2. Sheet1!A1:D5");
    assert.ok(!lines.some((line) => line.includes(BACKLINK_MARKER)));
  });
});
