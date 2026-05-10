import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { normalizeSelectedRange } from "../../src/core/range-normalizer.js";

describe("normalizeSelectedRange", () => {
  it("numeric-only selection passes through unchanged", () => {
    const values = [
      [10, 20, 30],
      [40, 50, 60],
      [70, 80, 90],
    ];
    const result = normalizeSelectedRange(values);
    assert.strictEqual(result.normalizationNeeded, false);
    assert.strictEqual(result.normalizationApplied, false);
    assert.deepStrictEqual(result.blockingReasons, []);
  });

  it("sparse title row + numeric body is normalized", () => {
    // Row 0: single text cell (sparse → title-like). Rows 1-3: all numeric.
    const values = [
      ["Survey Results", null, null, null],
      [10, 20, 30, 40],
      [50, 60, 70, 80],
      [90, 10, 20, 30],
    ];
    const result = normalizeSelectedRange(values);
    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, true);
    assert.deepStrictEqual(result.titleRows, [0]);
    assert.strictEqual(result.dataRowOffset, 1);
    assert.deepStrictEqual(result.labelColumns, []);
    assert.deepStrictEqual(result.blockingReasons, []);
  });

  it("label column + numeric data columns is normalized", () => {
    // Col 0: text labels. Cols 1-3: numeric data.
    const values = [
      ["Alpha", 10, 20, 30],
      ["Beta", 40, 50, 60],
      ["Gamma", 70, 80, 90],
    ];
    const result = normalizeSelectedRange(values);
    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, true);
    assert.deepStrictEqual(result.labelColumns, [0]);
    assert.strictEqual(result.dataColOffset, 1);
    assert.deepStrictEqual(result.titleRows, []);
    assert.deepStrictEqual(result.blockingReasons, []);
  });

  it("empty row inside body blocks normalization with BODY_APPEARS_MULTI_TABLE", () => {
    // Row 0: wide banner header. Rows 1-2 and 4-5: data. Row 3: empty gap.
    // The gap triggers the multi-table blocking reason.
    const values = [
      ["Group", "A", "B", "C"],
      ["Label1", 10, 20, 30],
      ["Label2", 40, 50, 60],
      [null, null, null, null],
      ["Label3", 70, 80, 90],
      ["Label4", 10, 20, 30],
    ];
    const result = normalizeSelectedRange(values);
    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, false);
    assert.ok(
      result.blockingReasons.includes("BODY_APPEARS_MULTI_TABLE"),
      `expected BODY_APPEARS_MULTI_TABLE in blockingReasons, got: ${JSON.stringify(result.blockingReasons)}`
    );
  });
});
