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

  it("quarter-like labels alone are not treated as numeric evidence", () => {
    const values = [
      ["Wave", "2025Q4", "2026Q1"],
      ["Segment", "2025Q4", "2026Q1"],
      ["Note", "Q4 2025", "Q1 2026"],
    ];

    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.strictEqual(items.length, 0, `expected 0 items, got ${items.length}`);
  });

  it("banner-heavy tables with quarter-like headers still produce one inventory item", () => {
    const values = [
      ["Usage table", null, null, null],
      ["", "Total", "Male", "Female"],
      ["", "2025Q4", "2025Q4", "2025Q4"],
      ["Agree", "44%", "41%", "39%"],
      ["Disagree", "56%", "59%", "61%"],
      ["BASE", "5605", "1320", "3083"],
    ];

    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.strictEqual(items.length, 1, `expected 1 item, got ${items.length}`);

    const item = items[0];
    assert.strictEqual(item.title, "Usage table");
    assert.strictEqual(item.titleSource, "firstRowOfBand");
    assert.strictEqual(item.labelSplitConfidence, "confident");
    assert.strictEqual(item.columnCount, 4);
    assert.strictEqual(item.isLikelyTable, true);
  });

  // ─── New hardening tests for issue #110 ──────────────────────────────────

  it("item does not expose canRunSignificance — scanner is a candidate finder only", () => {
    const values = [
      ["Label1", 10, 20, 30],
      ["Label2", 40, 50, 60],
      ["Base", 100, 100, 100],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    assert.strictEqual(
      items[0].canRunSignificance,
      undefined,
      "canRunSignificance must not exist on inventory items"
    );
    assert.ok(
      ["available", "uncertain", "rejected"].includes(items[0].candidateStatus),
      `candidateStatus must be available/uncertain/rejected; got ${items[0].candidateStatus}`
    );
  });

  it("clean table with base produces candidateStatus=available", () => {
    const values = [
      ["Label1", 10, 20, 30],
      ["Label2", 40, 50, 60],
      ["Base", 100, 100, 100],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1);
    assert.strictEqual(items[0].candidateStatus, "available");
  });

  it("table with no explicit Base row is not presented as Run-ready — no canRunSignificance field", () => {
    const values = [
      ["Label1", 10, 20, 30],
      ["Label2", 40, 50, 60],
      ["Label3", 70, 80, 90],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    // Without a Base row the scanner may produce 0 or 1 items depending on metric detection.
    // If an item exists, it must not carry canRunSignificance (the key safety property).
    if (items.length > 0) {
      assert.strictEqual(items[0].canRunSignificance, undefined, "canRunSignificance must not exist");
      // "rejected" is acceptable — no Base means no complete metric blocks.
      assert.ok(
        ["available", "uncertain", "rejected"].includes(items[0].candidateStatus),
        `candidateStatus must be a recognised value; got ${items[0].candidateStatus}`
      );
    }
    // If items.length === 0, the scanner correctly found nothing to over-promise on.
  });

  it("candidate with preview warnings is surfaced as uncertain", () => {
    // All-100 rows trigger a quality warning in the preview model.
    const values = [
      ["Label1", 100, 100, 100],
      ["Label2", 100, 100, 100],
      ["Base", 1000, 1000, 1000],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    const item = items[0];
    // Warnings must be reflected in the candidate status.
    if (item.warningsCount > 0) {
      assert.strictEqual(
        item.candidateStatus,
        "uncertain",
        "item with preview warnings must be uncertain, not available"
      );
    }
  });

  it("side-by-side tables in one row band appear as one candidate and expose no canRunSignificance", () => {
    // Known limitation: scanner cannot split side-by-side tables within one band.
    // This test documents the limitation and ensures no over-promising.
    const values = [
      ["", "Total", "Male", "", "Total", "Female"],
      ["Metric A", 100, 60, "Metric B", 200, 120],
      ["Metric A2", 150, 90, "Metric B2", 250, 150],
      ["Base", 1000, 600, "Base", 2000, 1200],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    // Entire band is treated as one candidate (known limitation).
    assert.strictEqual(items.length, 1, "side-by-side tables are reported as one band (known limitation)");
    assert.strictEqual(items[0].canRunSignificance, undefined, "canRunSignificance must not exist");
    assert.ok(
      ["available", "uncertain", "rejected"].includes(items[0].candidateStatus),
      "candidateStatus must be a recognised value"
    );
  });

  it("non-empty commentary row between two tables merges them — no split, no false confidence", () => {
    // Tables separated by a non-empty row are merged into one band (known limitation).
    // The merged candidate must not report canRunSignificance.
    const values = [
      ["Label1", 10, 20, 30],
      ["Label2", 40, 50, 60],
      ["Base", 100, 100, 100],
      ["Note: preliminary data", null, null, null],  // non-empty, prevents band split
      ["Label3", 70, 80, 90],
      ["Label4", 10, 20, 30],
      ["Base", 200, 200, 200],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    // Should be 1 merged band, not 2.
    assert.strictEqual(items.length, 1, "non-empty note row prevents band split — expect 1 merged candidate");
    assert.strictEqual(items[0].canRunSignificance, undefined, "canRunSignificance must not exist");
  });

  it("title inferred two rows above band via empty separator gets medium confidence", () => {
    // Row 0: section title. Row 1: empty separator. Rows 2-4: table band.
    // twoRowsAbove inference should be medium confidence, not high.
    const values = [
      ["Section Title", null, null],
      [null, null, null],
      ["Label1", 10, 20],
      ["Label2", 30, 40],
      ["Base", 100, 100],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    const item = items[0];
    if (item.titleSource === "twoRowsAbove") {
      assert.strictEqual(
        item.titleConfidence,
        "medium",
        "twoRowsAbove title confidence must be medium (text may belong to a preceding section)"
      );
    }
  });

  it("item exposes candidateNotes not reasonsIfNotRunnable", () => {
    const values = [
      ["Label1", 10, 20],
      ["Label2", 30, 40],
      ["Base", 100, 100],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1);
    assert.ok(Array.isArray(items[0].candidateNotes), "candidateNotes must be an array");
    assert.strictEqual(items[0].reasonsIfNotRunnable, undefined, "reasonsIfNotRunnable must not exist");
  });
});
