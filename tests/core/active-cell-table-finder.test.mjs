import { describe, it } from "node:test";
import assert from "node:assert/strict";

import {
  parseA1Range,
  findCandidateForActiveCell,
} from "../../src/core/active-cell-table-finder.js";

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
