import { describe, it } from "node:test";
import assert from "node:assert/strict";

import {
  parseA1Range,
  findCandidateForActiveCell,
  extractCandidateSlice,
} from "../../src/core/active-cell-table-finder.js";
import { hasEmptyDataRowGap } from "../../src/core/range-normalizer.js";

// ─── parseA1Range ─────────────────────────────────────────────────────────────

describe("parseA1Range — valid inputs", () => {
  it("parses a simple range", () => {
    const r = parseA1Range("B3:H15");
    assert.deepStrictEqual(r, { startRow: 2, endRow: 14, startCol: 1, endCol: 7 });
  });

  it("parses A1 origin", () => {
    const r = parseA1Range("A1:A1");
    assert.deepStrictEqual(r, { startRow: 0, endRow: 0, startCol: 0, endCol: 0 });
  });

  it("parses a single-cell address", () => {
    const r = parseA1Range("C5");
    assert.deepStrictEqual(r, { startRow: 4, endRow: 4, startCol: 2, endCol: 2 });
  });

  it("handles two-letter column AA", () => {
    const r = parseA1Range("AA1:AA1");
    assert.deepStrictEqual(r, { startRow: 0, endRow: 0, startCol: 26, endCol: 26 });
  });

  it("handles two-letter column AB", () => {
    const r = parseA1Range("AB3:AC10");
    assert.deepStrictEqual(r, { startRow: 2, endRow: 9, startCol: 27, endCol: 28 });
  });

  it("is case-insensitive", () => {
    const r = parseA1Range("b3:h15");
    assert.deepStrictEqual(r, { startRow: 2, endRow: 14, startCol: 1, endCol: 7 });
  });

  it("trims surrounding whitespace", () => {
    const r = parseA1Range("  B3:H15  ");
    assert.deepStrictEqual(r, { startRow: 2, endRow: 14, startCol: 1, endCol: 7 });
  });
});

describe("parseA1Range — invalid inputs", () => {
  it("returns null for non-string input", () => {
    assert.strictEqual(parseA1Range(null), null);
    assert.strictEqual(parseA1Range(42), null);
    assert.strictEqual(parseA1Range(undefined), null);
  });

  it("returns null for empty string", () => {
    assert.strictEqual(parseA1Range(""), null);
  });

  it("returns null for address missing column letters", () => {
    assert.strictEqual(parseA1Range("3:15"), null);
  });

  it("returns null for address missing row numbers", () => {
    assert.strictEqual(parseA1Range("B:H"), null);
  });

  it("returns null for sheet-qualified address", () => {
    // Sheet!A1 notation is not supported by this parser
    assert.strictEqual(parseA1Range("Sheet1!B3:H15"), null);
  });

  it("returns null for row 0 (A1 rows are 1-based)", () => {
    assert.strictEqual(parseA1Range("A0:B5"), null);
  });
});

// ─── findCandidateForActiveCell ───────────────────────────────────────────────

function makeItem(rangeAddress, extra = {}) {
  return { rangeAddress, candidateStatus: "available", ...extra };
}

describe("findCandidateForActiveCell — found", () => {
  it("returns found when active cell is inside a candidate", () => {
    const items = [makeItem("B3:H15")];
    // row=5 (0-based), col=4 → inside B3:H15
    const result = findCandidateForActiveCell(items, 5, 4);
    assert.strictEqual(result.status, "found");
    assert.strictEqual(result.candidate.rangeAddress, "B3:H15");
  });

  it("matches the top-left corner of the candidate range", () => {
    const items = [makeItem("B3:H15")];
    // row=2 (row 3 in A1), col=1 (col B) → exactly the top-left corner
    const result = findCandidateForActiveCell(items, 2, 1);
    assert.strictEqual(result.status, "found");
  });

  it("matches the bottom-right corner of the candidate range", () => {
    const items = [makeItem("B3:H15")];
    // row=14 (row 15 in A1), col=7 (col H) → exactly the bottom-right corner
    const result = findCandidateForActiveCell(items, 14, 7);
    assert.strictEqual(result.status, "found");
  });

  it("selects the correct candidate when multiple candidates are present", () => {
    const items = [
      makeItem("B3:H15"),
      makeItem("B20:H30"),
    ];
    // row=22 (inside B20:H30, 0-based row 19..29)
    const result = findCandidateForActiveCell(items, 22, 3);
    assert.strictEqual(result.status, "found");
    assert.strictEqual(result.candidate.rangeAddress, "B20:H30");
  });

  it("carries candidateMeta fields through on the found candidate", () => {
    const items = [makeItem("B3:H15", { title: "Demo Table", candidateStatus: "available" })];
    const result = findCandidateForActiveCell(items, 5, 4);
    assert.strictEqual(result.status, "found");
    assert.strictEqual(result.candidate.title, "Demo Table");
  });
});

describe("findCandidateForActiveCell — no-table", () => {
  it("returns no-table when active cell is above the candidate", () => {
    const items = [makeItem("B3:H15")];
    // row=1 (row 2 in A1) — above the candidate which starts at row=2
    const result = findCandidateForActiveCell(items, 1, 4);
    assert.strictEqual(result.status, "no-table");
    assert.deepStrictEqual(result.candidates, []);
  });

  it("returns no-table when active cell is below the candidate", () => {
    const items = [makeItem("B3:H15")];
    const result = findCandidateForActiveCell(items, 15, 4);
    assert.strictEqual(result.status, "no-table");
  });

  it("returns no-table when active cell is left of the candidate", () => {
    const items = [makeItem("B3:H15")];
    // col=0 (col A) — left of B (col=1)
    const result = findCandidateForActiveCell(items, 5, 0);
    assert.strictEqual(result.status, "no-table");
  });

  it("returns no-table when active cell is right of the candidate", () => {
    const items = [makeItem("B3:H15")];
    // col=8 (col I) — right of H (col=7)
    const result = findCandidateForActiveCell(items, 5, 8);
    assert.strictEqual(result.status, "no-table");
  });

  it("returns no-table for an empty items array", () => {
    const result = findCandidateForActiveCell([], 0, 0);
    assert.strictEqual(result.status, "no-table");
    assert.deepStrictEqual(result.candidates, []);
  });

  it("returns no-table when items is not an array", () => {
    const result = findCandidateForActiveCell(null, 0, 0);
    assert.strictEqual(result.status, "no-table");
  });

  it("skips items with unparseable rangeAddress and returns no-table", () => {
    const items = [{ rangeAddress: "Sheet1!B3:H15" }]; // sheet-qualified not supported
    const result = findCandidateForActiveCell(items, 5, 4);
    assert.strictEqual(result.status, "no-table");
  });
});

describe("findCandidateForActiveCell — ambiguous", () => {
  it("returns ambiguous when active cell falls inside two overlapping candidates", () => {
    const items = [
      makeItem("A1:D20"),
      makeItem("A10:D30"),
    ];
    // row=12 (0-based) falls in both candidates
    const result = findCandidateForActiveCell(items, 12, 2);
    assert.strictEqual(result.status, "ambiguous");
    assert.strictEqual(result.candidates.length, 2);
  });

  it("ambiguous result includes both matching candidates", () => {
    const items = [
      makeItem("A1:D20", { title: "First" }),
      makeItem("A10:D30", { title: "Second" }),
    ];
    const result = findCandidateForActiveCell(items, 12, 2);
    assert.strictEqual(result.status, "ambiguous");
    const titles = result.candidates.map((c) => c.title);
    assert.ok(titles.includes("First"), "should include First");
    assert.ok(titles.includes("Second"), "should include Second");
  });
});

// ─── extractCandidateSlice ────────────────────────────────────────────────────

describe("extractCandidateSlice", () => {
  // usedRange starts at A1 (rowIndex=0, colIndex=0)
  const values0 = [
    [10, 20, 30, 40],
    [11, 21, 31, 41],
    [12, 22, 32, 42],
    [13, 23, 33, 43],
  ];

  it("extracts the full range when candidate matches usedRange exactly", () => {
    // A1:D4 → entire 4×4 grid
    const slice = extractCandidateSlice(values0, 0, 0, "A1:D4");
    assert.deepStrictEqual(slice, values0);
  });

  it("extracts a sub-range with zero row/col offsets", () => {
    // B2:C3 → rows 1..2 (0-based), cols 1..2 → [[21,31],[22,32]]
    const slice = extractCandidateSlice(values0, 0, 0, "B2:C3");
    assert.deepStrictEqual(slice, [
      [21, 31],
      [22, 32],
    ]);
  });

  it("applies row and column offsets correctly (usedRange starts at B3)", () => {
    // usedRange starts at B3 → rowIndex=2, colIndex=1
    // candidate C4:D5 → startRow=3,endRow=4, startCol=2,endCol=3
    // relStartRow=1, relEndRow=2, relStartCol=1, relEndCol=2
    const usedValues = [
      [10, 20, 30],   // B3:D3
      [11, 21, 31],   // B4:D4
      [12, 22, 32],   // B5:D5
    ];
    const slice = extractCandidateSlice(usedValues, 2, 1, "C4:D5");
    assert.deepStrictEqual(slice, [
      [21, 31],
      [22, 32],
    ]);
  });

  it("returns null for empty usedRangeValues", () => {
    assert.strictEqual(extractCandidateSlice([], 0, 0, "A1:B2"), null);
  });

  it("returns null for non-array usedRangeValues", () => {
    assert.strictEqual(extractCandidateSlice(null, 0, 0, "A1:B2"), null);
    assert.strictEqual(extractCandidateSlice(undefined, 0, 0, "A1:B2"), null);
  });

  it("returns null for unparseable candidateRangeAddress", () => {
    assert.strictEqual(extractCandidateSlice(values0, 0, 0, "Sheet1!A1:B2"), null);
    assert.strictEqual(extractCandidateSlice(values0, 0, 0, ""), null);
    assert.strictEqual(extractCandidateSlice(values0, 0, 0, null), null);
  });

  it("returns null when candidate starts before usedRange top edge", () => {
    // usedRange starts at row 5; candidate starts at row 3 → relStartRow = -2
    const slice = extractCandidateSlice(values0, 5, 0, "A4:D7");
    assert.strictEqual(slice, null);
  });

  it("returns null when candidate starts before usedRange left edge", () => {
    // usedRange starts at col 3; candidate starts at col 1 → relStartCol = -2
    const slice = extractCandidateSlice(values0, 0, 3, "B1:D4");
    assert.strictEqual(slice, null);
  });

  it("clamps endRow to usedRange boundary when candidate extends beyond", () => {
    // usedRange has 4 rows (0-based 0..3); candidate asks for rows 2..10 (A3:D11)
    // relEndRow = 10 → clamped to 3 → last two rows of values0
    const slice = extractCandidateSlice(values0, 0, 0, "A3:D11");
    assert.deepStrictEqual(slice, [
      [12, 22, 32, 42],
      [13, 23, 33, 43],
    ]);
  });

  it("detects empty-row gap in candidate slice (resolver smoke scenario)", () => {
    // Two numeric tables with an all-empty separator row between them,
    // all within a single 'usedRange'. The resolver extracts the candidate slice
    // and hasEmptyDataRowGap must fire.
    const usedValues = [
      [1,  2,  3],   // Table 1 base (row 0)
      [10, 20, 30],  // Table 1 data (row 1)
      ["", "", ""],  // gap row (row 2)
      [4,  5,  6],   // Table 2 base (row 3)
      [40, 50, 60],  // Table 2 data (row 4)
    ];
    // candidate spans the full usedRange: A1:C5
    const slice = extractCandidateSlice(usedValues, 0, 0, "A1:C5");
    // Verify slice is correct (all 5 rows, all 3 cols)
    assert.strictEqual(slice.length, 5);
    assert.deepStrictEqual(slice[2], ["", "", ""]);
    // Verify the gap is detected
    assert.strictEqual(hasEmptyDataRowGap(slice), true);
  });
});
