import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { scanWorksheetForTables } from "../../src/core/table-inventory-scanner.js";

const SHEET = "Sheet1";
const OFFSET = { usedRangeRowOffset: 0, usedRangeColOffset: 0, sheetName: SHEET };

describe("scanWorksheetForTables", () => {
  it("two tables separated by an empty row return 2 inventory items", () => {
    const values = [
      ["Label1", 10, 20, 30],
      ["Label2", 40, 50, 60],
      ["Base", 100, 100, 100],
      [null, null, null, null],  // empty separator
      ["Label3", 70, 80, 90],
      ["Label4", 10, 20, 30],
      ["Base", 200, 200, 200],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.strictEqual(items.length, 2, `expected 2 items, got ${items.length}`);
    assert.strictEqual(items[0].sheetName, SHEET);
    assert.strictEqual(items[1].sheetName, SHEET);
  });

  it("first-row merged-like title is detected and surfaced on the item", () => {
    // Row 0: single text cell (sparse, no numeric) → title-like.
    // Rows 1-3: numeric data band.
    const values = [
      ["My Survey", null, null, null],
      ["Label1", 10, 20, 30],
      ["Label2", 40, 50, 60],
      ["Base", 100, 100, 100],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    const item = items[0];
    assert.strictEqual(item.title, "My Survey");
    assert.strictEqual(item.titleSource, "firstRowOfBand");
  });

  it("column A=label, column B=Total keeps Total as data (not a second label column)", () => {
    // Col 0 is text labels. Col 1 header is "Total" (text in row 0 only).
    // Remaining rows of col 1 are numeric → col 1 text fraction < threshold
    // → classified as data, not a second label column.
    const values = [
      ["Category", "Total", "Male", "Female"],
      ["Label1", 100, 60, 40],
      ["Label2", 200, 120, 80],
      ["Base", 300, 180, 120],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    const item = items[0];
    // Confident split means col 0 was correctly identified as the only label column.
    assert.strictEqual(
      item.labelSplitConfidence,
      "confident",
      `expected confident split; got ${item.labelSplitConfidence}`
    );
    // Total band width is 4 columns (cols 0-3).
    assert.strictEqual(item.columnCount, 4);
  });
});
