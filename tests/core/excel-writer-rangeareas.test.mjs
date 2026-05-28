import { describe, it } from "node:test";
import assert from "node:assert/strict";

import {
  columnIndexToA1,
  runSpanToA1Address,
  packAddressesIntoChunks,
  buildBoldRangeAreasDiagnostics,
  buildFillRangeAreasDiagnosticsByColor,
} from "../../src/core/excel-writer.js";

describe("columnIndexToA1", () => {
  it("maps single-letter columns", () => {
    assert.strictEqual(columnIndexToA1(0), "A");
    assert.strictEqual(columnIndexToA1(1), "B");
    assert.strictEqual(columnIndexToA1(25), "Z");
  });

  it("maps double-letter columns", () => {
    assert.strictEqual(columnIndexToA1(26), "AA");
    assert.strictEqual(columnIndexToA1(27), "AB");
    assert.strictEqual(columnIndexToA1(51), "AZ");
    assert.strictEqual(columnIndexToA1(52), "BA");
    assert.strictEqual(columnIndexToA1(701), "ZZ");
  });

  it("maps triple-letter columns up to Excel max", () => {
    assert.strictEqual(columnIndexToA1(702), "AAA");
    assert.strictEqual(columnIndexToA1(16383), "XFD");
  });
});

describe("runSpanToA1Address", () => {
  it("formats a single-cell run", () => {
    // anchor row 2, anchor col 1 → absolute row 3, column 1 + 3 = 4 ("E").
    const span = { start: 3, end: 3 };
    assert.strictEqual(runSpanToA1Address(span, 0, 2, 1), "E3");
  });

  it("formats a multi-cell run", () => {
    // anchor row 2, anchor col 1 → row 3; start col 1+1=2 ("C"), end col 1+4=5 ("F").
    const span = { start: 1, end: 4 };
    assert.strictEqual(runSpanToA1Address(span, 0, 2, 1), "C3:F3");
  });

  it("offsets row index correctly", () => {
    // anchor row 10, rowIndex 5 → absolute row 16. Anchor col 0 + span 0..2 → A..C.
    const span = { start: 0, end: 2 };
    assert.strictEqual(runSpanToA1Address(span, 5, 10, 0), "A16:C16");
  });
});

describe("packAddressesIntoChunks", () => {
  it("returns no chunks for an empty input", () => {
    assert.deepStrictEqual(packAddressesIntoChunks([]), []);
  });

  it("packs all addresses into one chunk when both caps are large enough", () => {
    const addresses = ["A1", "B1", "C1"];
    const chunks = packAddressesIntoChunks(addresses, {
      maxAreasPerChunk: 100,
      maxCharsPerChunk: 1000,
    });
    assert.strictEqual(chunks.length, 1);
    assert.strictEqual(chunks[0].address, "A1,B1,C1");
    assert.strictEqual(chunks[0].areaCount, 3);
    assert.strictEqual(chunks[0].length, "A1,B1,C1".length);
  });

  it("splits into new chunks when the area cap fires", () => {
    const addresses = ["A1", "B1", "C1", "D1", "E1"];
    const chunks = packAddressesIntoChunks(addresses, {
      maxAreasPerChunk: 2,
      maxCharsPerChunk: 1000,
    });
    assert.strictEqual(chunks.length, 3);
    assert.deepStrictEqual(
      chunks.map((c) => c.address),
      ["A1,B1", "C1,D1", "E1"]
    );
    assert.deepStrictEqual(chunks.map((c) => c.areaCount), [2, 2, 1]);
  });

  it("splits into new chunks when the char cap fires", () => {
    const addresses = ["A1", "B1", "C1", "D1"];
    // "A1,B1" is 5 chars; adding ",C1" would make 8. Cap at 6 → "A1,B1" then "C1,D1".
    const chunks = packAddressesIntoChunks(addresses, {
      maxAreasPerChunk: 100,
      maxCharsPerChunk: 6,
    });
    assert.deepStrictEqual(
      chunks.map((c) => c.address),
      ["A1,B1", "C1,D1"]
    );
  });

  it("never produces an empty chunk", () => {
    const addresses = ["A1", "B1", "C1"];
    const chunks = packAddressesIntoChunks(addresses, {
      maxAreasPerChunk: 1,
      maxCharsPerChunk: 1,
    });
    for (const chunk of chunks) {
      assert.ok(chunk.areaCount > 0);
    }
  });
});

describe("buildBoldRangeAreasDiagnostics", () => {
  it("returns zero counts for an empty span set", () => {
    const summary = buildBoldRangeAreasDiagnostics([], 0, 0);
    assert.strictEqual(summary.areaCountTotal, 0);
    assert.strictEqual(summary.chunkCount, 0);
    assert.strictEqual(summary.commandCountEstimate, 0);
    assert.strictEqual(summary.maxAddressLength, 0);
  });

  it("packs many row-runs into one chunk under the defaults", () => {
    const spansByRow = Array.from({ length: 5 }, () => [{ start: 0, end: 2 }]);
    const summary = buildBoldRangeAreasDiagnostics(spansByRow, 0, 0);

    assert.strictEqual(summary.areaCountTotal, 5);
    assert.strictEqual(summary.chunkCount, 1);
    // One chunk → one getRanges + one .format.font.bold setter.
    assert.strictEqual(summary.commandCountEstimate, 2);
    assert.ok(summary.maxAddressLength > 0);
  });

  it("splits into multiple chunks when areas exceed the cap", () => {
    const spansByRow = Array.from({ length: 250 }, () => [{ start: 0, end: 0 }]);
    const summary = buildBoldRangeAreasDiagnostics(spansByRow, 0, 0, {
      maxAreasPerChunk: 100,
      maxCharsPerChunk: 100000,
    });
    assert.strictEqual(summary.areaCountTotal, 250);
    assert.strictEqual(summary.chunkCount, 3);
    assert.strictEqual(summary.commandCountEstimate, 6);
  });

  it("dramatically reduces the projected command count vs row-run count", () => {
    // 100 rows × 20 row-runs each = 2000 row-runs, modelling a wide workbook
    // mask. The current writer issues one command per row-run.
    const spansByRow = Array.from({ length: 100 }, () =>
      Array.from({ length: 20 }, (_, k) => ({ start: k * 2, end: k * 2 + 1 }))
    );
    const summary = buildBoldRangeAreasDiagnostics(spansByRow, 0, 0);

    const rowRunCount = 100 * 20;
    assert.strictEqual(summary.areaCountTotal, rowRunCount);
    // At default 100 areas/chunk, this packs into ~20 chunks → ~40 ops.
    assert.ok(summary.commandCountEstimate < rowRunCount / 10);
  });
});

describe("buildFillRangeAreasDiagnosticsByColor", () => {
  it("returns an empty summary when no spans carry colors", () => {
    const summary = buildFillRangeAreasDiagnosticsByColor([], 0, 0);
    assert.strictEqual(summary.colorCount, 0);
    assert.strictEqual(summary.areaCountTotal, 0);
    assert.strictEqual(summary.commandCountEstimate, 0);
  });

  it("groups spans by color and packs each color independently", () => {
    const spansByRow = [
      [{ start: 0, end: 1, color: "#GREEN" }],
      [{ start: 0, end: 1, color: "#ORANGE" }],
      [{ start: 2, end: 3, color: "#GREEN" }],
    ];
    const summary = buildFillRangeAreasDiagnosticsByColor(spansByRow, 0, 0);

    assert.strictEqual(summary.colorCount, 2);
    assert.strictEqual(summary.areaCountTotal, 3);

    const green = summary.perColor.find((entry) => entry.color === "#GREEN");
    const orange = summary.perColor.find((entry) => entry.color === "#ORANGE");
    assert.strictEqual(green.areaCount, 2);
    assert.strictEqual(orange.areaCount, 1);
    // Two colors with one chunk each → 4 queued ops total.
    assert.strictEqual(summary.commandCountEstimate, 4);
  });

  it("skips spans without a color (no fill needed)", () => {
    const spansByRow = [
      [{ start: 0, end: 1, color: "" }],
      [{ start: 0, end: 1, color: "#GREEN" }],
    ];
    const summary = buildFillRangeAreasDiagnosticsByColor(spansByRow, 0, 0);
    assert.strictEqual(summary.areaCountTotal, 1);
    assert.strictEqual(summary.colorCount, 1);
  });
});
