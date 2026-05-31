import { describe, it } from "node:test";
import assert from "node:assert/strict";

import {
  getSignificanceMarkerAlphabet,
  generateSignificanceLabels,
  isSignificanceMarkerLabel,
  getSignificanceMarkerCapacity,
  computeRequiredSignificanceLabelCount,
  detectSignificanceMarkerOverflow,
  getSignificanceLabelForColumnIndex,
  buildBannerLocalSignificanceLabelMap,
  applyComparisonResultsToFullCellResultMatrix,
  createEmptyCellResultMatrix,
  removeSignificanceMarkersFromText,
} from "../../src/core/significance.js";

// Cyrillic symbols that must never appear in generated markers (issue #312).
const EXCLUDED_CYRILLIC = [
  "а", "А", "В", "с", "С", "е", "Е", "К", "М", "Н",
  "о", "О", "р", "Р", "у", "х", "Х",
];

const CYRILLIC_RANGE = /[Ѐ-ӿ]/;

describe("getSignificanceMarkerAlphabet — Latin-only by default", () => {
  it("contains no Cyrillic characters when Cyrillic markers are disabled", () => {
    const alphabet = getSignificanceMarkerAlphabet({ useCyrillicMarkers: false });
    assert.ok(alphabet.every((label) => !CYRILLIC_RANGE.test(label)));
  });

  it("defaults to Latin-only when no options are passed", () => {
    const alphabet = getSignificanceMarkerAlphabet();
    assert.ok(alphabet.every((label) => !CYRILLIC_RANGE.test(label)));
  });

  it("excludes Latin t and T (reserved for Total markers)", () => {
    const alphabet = getSignificanceMarkerAlphabet({ useCyrillicMarkers: false });
    assert.ok(!alphabet.includes("t"));
    assert.ok(!alphabet.includes("T"));
  });

  it("starts with the historical Latin sequence a, b, c, ...", () => {
    const alphabet = getSignificanceMarkerAlphabet();
    assert.deepStrictEqual(alphabet.slice(0, 5), ["a", "b", "c", "d", "e"]);
  });
});

describe("getSignificanceMarkerAlphabet — Cyrillic exclusions", () => {
  const alphabet = getSignificanceMarkerAlphabet({ useCyrillicMarkers: true });

  it("excludes every visually confusable Cyrillic symbol", () => {
    for (const symbol of EXCLUDED_CYRILLIC) {
      assert.ok(!alphabet.includes(symbol), `expected ${symbol} to be excluded`);
    }
  });

  it("keeps uppercase Cyrillic У (distinct from Latin Y)", () => {
    assert.ok(alphabet.includes("У"));
  });

  it("excludes lowercase Cyrillic у (confusable with Latin y)", () => {
    assert.ok(!alphabet.includes("у"));
  });

  it("still excludes Cyrillic т and Т (existing Total exclusion)", () => {
    assert.ok(!alphabet.includes("т"));
    assert.ok(!alphabet.includes("Т"));
  });

  it("appends Cyrillic markers after the full Latin alphabet", () => {
    const latinCount = getSignificanceMarkerAlphabet({ useCyrillicMarkers: false }).length;
    assert.deepStrictEqual(
      alphabet.slice(0, latinCount),
      getSignificanceMarkerAlphabet({ useCyrillicMarkers: false })
    );
    assert.ok(CYRILLIC_RANGE.test(alphabet[latinCount]));
  });
});

describe("generateSignificanceLabels — under the single-character limit", () => {
  it("matches the previous Latin sequence for the first columns", () => {
    const labels = generateSignificanceLabels({ minimumCount: 6 });
    assert.deepStrictEqual(labels.slice(0, 6), ["a", "b", "c", "d", "e", "f"]);
  });

  it("does not extend to multi-character markers when allowance is off", () => {
    const labels = generateSignificanceLabels({ minimumCount: 1000 });
    // Latin-only single-character alphabet length.
    assert.strictEqual(labels.length, getSignificanceMarkerAlphabet().length);
    assert.ok(labels.every((label) => label.length === 1));
  });
});

describe("generateSignificanceLabels — multi-character overflow", () => {
  it("continues as aa, ab, ac after the single-character alphabet", () => {
    const singleCount = getSignificanceMarkerAlphabet().length;
    const labels = generateSignificanceLabels({
      allowMultiCharacterMarkers: true,
      minimumCount: singleCount + 3,
    });

    assert.strictEqual(labels.length, singleCount + 3);
    assert.deepStrictEqual(labels.slice(singleCount, singleCount + 3), ["aa", "ab", "ac"]);
  });

  it("only extends past the single-character alphabet when needed", () => {
    const labels = generateSignificanceLabels({
      allowMultiCharacterMarkers: true,
      minimumCount: 3,
    });
    assert.ok(labels.every((label) => label.length === 1));
  });
});

describe("isSignificanceMarkerLabel", () => {
  it("recognises single Latin markers", () => {
    assert.ok(isSignificanceMarkerLabel("a"));
    assert.ok(isSignificanceMarkerLabel("Z"));
  });

  it("recognises legacy single Cyrillic markers for cleanup", () => {
    // Even now-excluded look-alikes must still be removable on re-runs.
    assert.ok(isSignificanceMarkerLabel("а"));
    assert.ok(isSignificanceMarkerLabel("В"));
  });

  it("recognises multi-character overflow markers", () => {
    assert.ok(isSignificanceMarkerLabel("aa"));
    assert.ok(isSignificanceMarkerLabel("ac"));
  });

  it("rejects ordinary banner text", () => {
    assert.ok(!isSignificanceMarkerLabel("quarter"));
    assert.ok(!isSignificanceMarkerLabel("NE"));
    assert.ok(!isSignificanceMarkerLabel(""));
  });
});

describe("marker capacity and overflow detection", () => {
  it("capacity grows when Cyrillic markers are enabled", () => {
    const latinCapacity = getSignificanceMarkerCapacity({ useCyrillicMarkers: false });
    const cyrillicCapacity = getSignificanceMarkerCapacity({ useCyrillicMarkers: true });
    assert.ok(cyrillicCapacity > latinCapacity);
  });

  it("computes required label count without a Total column", () => {
    assert.strictEqual(computeRequiredSignificanceLabelCount(5, {}), 5);
  });

  it("excludes the Total column from the required count", () => {
    assert.strictEqual(
      computeRequiredSignificanceLabelCount(5, { firstColumnIsTotal: true }),
      4
    );
  });

  it("does not flag overflow for a normal table under the limit", () => {
    assert.ok(!detectSignificanceMarkerOverflow(10, {}));
  });

  it("flags overflow when columns exceed the Latin alphabet", () => {
    const capacity = getSignificanceMarkerCapacity({ useCyrillicMarkers: false });
    assert.ok(detectSignificanceMarkerOverflow(capacity + 1, {}));
  });

  it("uses the widest banner group for the required count", () => {
    const bannerStructure = {
      groups: [
        { groupKey: "g1", columnIndexes: [0, 1, 2] },
        { groupKey: "g2", columnIndexes: [3, 4, 5, 6] },
      ],
      totalColumnIndexes: [],
    };
    assert.strictEqual(
      computeRequiredSignificanceLabelCount(7, { respectBannerStructure: true }, bannerStructure),
      4
    );
  });
});

describe("getSignificanceLabelForColumnIndex — overflow labels", () => {
  it("returns multi-character labels past the alphabet when allowed", () => {
    const capacity = getSignificanceMarkerCapacity({ useCyrillicMarkers: false });
    const label = getSignificanceLabelForColumnIndex(capacity, {
      allowMultiCharacterMarkers: true,
    });
    assert.strictEqual(label, "aa");
  });

  it("returns empty string past the alphabet when overflow is not allowed", () => {
    const capacity = getSignificanceMarkerCapacity({ useCyrillicMarkers: false });
    assert.strictEqual(getSignificanceLabelForColumnIndex(capacity, {}), "");
  });
});

describe("buildBannerLocalSignificanceLabelMap — multi-character labels", () => {
  it("assigns multi-character labels to a wide banner group when allowed", () => {
    const capacity = getSignificanceMarkerCapacity({ useCyrillicMarkers: false });
    const columnIndexes = [];
    for (let i = 0; i <= capacity; i++) {
      columnIndexes.push(i);
    }

    const bannerStructure = {
      groups: [{ groupKey: "g1", columnIndexes }],
      totalColumnIndexes: [],
    };

    const labelMap = buildBannerLocalSignificanceLabelMap(bannerStructure, {
      allowMultiCharacterMarkers: true,
    });

    assert.strictEqual(labelMap.get(0), "a");
    assert.strictEqual(labelMap.get(capacity), "aa");
  });
});

// ─── Data-cell marker append / removal ────────────────────────────────────────

function buildSignificantComparison(firstColumnIndex, secondColumnIndex, direction) {
  return {
    firstColumnIndex,
    secondColumnIndex,
    comparisonType: "segment",
    result: { isSignificant: true, direction },
  };
}

describe("data-cell marker append — multi-character separation", () => {
  it("separates multiple multi-character markers with spaces", () => {
    // A wide table where column index `capacity+2` is higher than three lower
    // overflow columns, so its cell collects three multi-character markers.
    const capacity = getSignificanceMarkerCapacity({ useCyrillicMarkers: false });
    const target = capacity + 2; // label for this column is irrelevant; it is the higher one
    const lowerColumns = [capacity, capacity + 1, capacity + 3];

    const matrix = createEmptyCellResultMatrix(1, capacity + 5);
    const rowComparisons = lowerColumns.map((lower) =>
      buildSignificantComparison(target, lower, "first_higher")
    );

    applyComparisonResultsToFullCellResultMatrix(
      { comparisonRows: [{ valueRowIndex: 0, rowComparisons }] },
      matrix,
      { allowMultiCharacterMarkers: true }
    );

    const markers = matrix[0][target].markers;
    // Three multi-character markers, each separated by a single space.
    assert.ok(/^[a-z]{2}( [a-z]{2}){2}$/.test(markers), `got "${markers}"`);
  });

  it("keeps single-character markers concatenated under the limit", () => {
    const matrix = createEmptyCellResultMatrix(1, 4);
    const rowComparisons = [
      buildSignificantComparison(3, 1, "first_higher"), // appends label for col 1 = "b"
      buildSignificantComparison(3, 2, "first_higher"), // appends label for col 2 = "c"
    ];

    applyComparisonResultsToFullCellResultMatrix(
      { comparisonRows: [{ valueRowIndex: 0, rowComparisons }] },
      matrix,
      {}
    );

    assert.strictEqual(matrix[0][3].markers, "bc");
  });
});

describe("removeSignificanceMarkersFromText — multi-character markers", () => {
  it("removes a single trailing marker", () => {
    assert.strictEqual(removeSignificanceMarkersFromText("42% b"), "42%");
  });

  it("removes concatenated single-character markers", () => {
    assert.strictEqual(removeSignificanceMarkersFromText("42% bcd"), "42%");
  });

  it("removes space-separated multi-character markers", () => {
    assert.strictEqual(removeSignificanceMarkersFromText("42% aa ab ac"), "42%");
  });

  it("removes a Total marker followed by multi-character markers", () => {
    assert.strictEqual(removeSignificanceMarkersFromText("42% T aa ab"), "42%");
  });

  it("leaves a clean numeric value untouched", () => {
    assert.strictEqual(removeSignificanceMarkersFromText("42%"), "42%");
  });

  it("is idempotent across repeated removals", () => {
    const once = removeSignificanceMarkersFromText("21% aa ab");
    const twice = removeSignificanceMarkersFromText(once);
    assert.strictEqual(once, twice);
  });
});
