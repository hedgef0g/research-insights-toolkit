import { describe, it } from "node:test";
import assert from "node:assert/strict";

import {
  buildBannerDetectionDebugSummary,
  detectBannerStructure,
} from "../../src/core/banner-detector.js";
import {
  buildBannerLocalSignificanceLabelMap,
  buildColumnComparisonPairs,
} from "../../src/core/significance.js";

function getGroupLevelRowOffset(result) {
  return result.columnDescriptors[0].source.groupLevelRowOffset;
}

describe("detectBannerStructure - group level selection", () => {
  it("prefers a real-like semantic group row over a mixed technical wave value row", () => {
    const bannerContext = {
      selectedColumnCount: 14,
      lowerBannerRow: [
        "",
        "",
        "2025Q4",
        "2026Q1",
        "2025Q4",
        "2026Q1",
        "2025Q4",
        "2026Q1",
        "2025Q4",
        "2026Q1",
        "2025Q4",
        "2026Q1",
        "2025Q4",
        "2026Q1",
      ],
      upperScanRows: [
        [
          "2025Q4",
          "2026Q1",
          "Wave (quarter)",
          "",
          "Wave (quarter)",
          "",
          "Wave (quarter)",
          "",
          "Wave (quarter)",
          "",
          "Wave (quarter)",
          "",
          "Wave (quarter)",
          "",
        ],
        ["", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        [
          "Total",
          "",
          "",
          "",
          "Category usage",
          "",
          "",
          "",
          "Gender",
          "",
          "Age",
          "",
          "Geo",
          "",
        ],
      ],
    };

    const debugSummary = buildBannerDetectionDebugSummary(bannerContext);
    const level2Candidate = debugSummary.candidateRows.find((row) => row.bottomUpLevel === 2);

    assert.strictEqual(level2Candidate.isTechnicalWaveDescriptor, true);

    const result = detectBannerStructure(bannerContext, { autoDetectWaveBanners: true });

    assert.strictEqual(getGroupLevelRowOffset(result), -3);
    assert.ok(result.groups.length < 10);
    assert.deepStrictEqual(
      result.groups.map((group) => ({
        label: group.label,
        columnIndexes: group.columnIndexes,
      })),
      [
        { label: "Total", columnIndexes: [0, 1, 2, 3] },
        { label: "Category usage", columnIndexes: [4, 5, 6, 7] },
        { label: "Gender", columnIndexes: [8, 9] },
        { label: "Age", columnIndexes: [10, 11] },
        { label: "Geo", columnIndexes: [12, 13] },
      ]
    );

    const labelMap = buildBannerLocalSignificanceLabelMap(result, {
      autoDetectWaveBanners: true,
    });

    assert.deepStrictEqual(
      Array.from({ length: 6 }, (_, columnIndex) => labelMap.get(columnIndex) || ""),
      ["a", "b", "c", "d", "a", "b"]
    );
  });

  it("prefers a higher semantic group row over a technical wave descriptor row", () => {
    const result = detectBannerStructure(
      {
        selectedColumnCount: 6,
        lowerBannerRow: ["2025Q4", "2026Q1", "2025Q4", "2026Q1", "2025Q4", "2026Q1"],
        upperScanRows: [
          ["Волна (квартал)", "", "Волна (квартал)", "", "Волна (квартал)", ""],
          ["Всего", "", "Пользование категорией", "", "", ""],
        ],
      },
      { autoDetectWaveBanners: true }
    );

    assert.strictEqual(getGroupLevelRowOffset(result), -2);
    assert.strictEqual(result.globalTotalColumnIndex, null);
    assert.deepStrictEqual(result.totalColumnIndexes, []);

    assert.deepStrictEqual(
      result.groups.map((group) => ({
        label: group.label,
        columnIndexes: group.columnIndexes,
      })),
      [
        { label: "Всего", columnIndexes: [0, 1] },
        { label: "Пользование категорией", columnIndexes: [2, 3, 4, 5] },
      ]
    );

    const labelMap = buildBannerLocalSignificanceLabelMap(result, {
      autoDetectWaveBanners: true,
    });

    assert.deepStrictEqual(
      Array.from({ length: 6 }, (_, columnIndex) => labelMap.get(columnIndex) || ""),
      ["a", "b", "a", "b", "c", "d"]
    );
  });

  it("keeps a simple two-level semantic banner unchanged", () => {
    const result = detectBannerStructure({
      selectedColumnCount: 2,
      lowerBannerRow: ["Male", "Female"],
      upperScanRows: [["Gender", ""]],
    });

    assert.strictEqual(getGroupLevelRowOffset(result), -1);
    assert.deepStrictEqual(
      result.groups.map((group) => ({
        label: group.label,
        columnIndexes: group.columnIndexes,
      })),
      [{ label: "Gender", columnIndexes: [0, 1] }]
    );
  });

  it("keeps a pure wave-only banner as the group level when no semantic row exists", () => {
    const result = detectBannerStructure(
      {
        selectedColumnCount: 2,
        lowerBannerRow: ["2025Q4", "2026Q1"],
        upperScanRows: [["Wave (quarter)", ""]],
      },
      { autoDetectWaveBanners: true }
    );

    assert.strictEqual(getGroupLevelRowOffset(result), -1);
    assert.deepStrictEqual(
      result.groups.map((group) => ({
        label: group.label,
        columnIndexes: group.columnIndexes,
        recommendedComparisonMode: group.recommendedComparisonMode,
      })),
      [
        {
          label: "Wave (quarter)",
          columnIndexes: [0, 1],
          recommendedComparisonMode: "previousColumn",
        },
      ]
    );
  });
});

describe("detectBannerStructure - sparse lower banner totals", () => {
  it("marks columns with empty lower label as local Total when nearest upper cell is total-like", () => {
    const bannerContext = {
      selectedColumnCount: 6,
      lowerBannerRow: ["", "Male", "Female", "", "18-24", "25-34"],
      upperScanRows: [
        ["Всего", "M-prefix", "F-prefix", "Total", "Young", "Old"],
        ["", "Gender", "", "Age", "", ""],
      ],
    };

    const result = detectBannerStructure(bannerContext);

    const totalDescriptors = result.columnDescriptors.filter(
      (descriptor) => descriptor.isTotal
    );

    assert.deepStrictEqual(
      totalDescriptors.map((descriptor) => descriptor.columnIndex),
      [0, 3]
    );

    for (const columnIndex of [0, 3]) {
      const descriptor = result.columnDescriptors[columnIndex];
      assert.strictEqual(descriptor.isTotal, true);
      assert.strictEqual(descriptor.isLocalTotal, true);
      assert.strictEqual(descriptor.totalType, "local");
    }

    assert.ok(result.totalColumnIndexes.includes(0));
    assert.ok(result.totalColumnIndexes.includes(3));
    assert.strictEqual(result.globalTotalColumnIndex, null);

    const labelMap = buildBannerLocalSignificanceLabelMap(result);

    assert.deepStrictEqual(
      Array.from({ length: 6 }, (_, columnIndex) => labelMap.get(columnIndex) || ""),
      ["", "a", "b", "", "a", "b"]
    );

    assert.deepStrictEqual(
      result.groups.map((group) => ({
        label: group.label,
        columnIndexes: group.columnIndexes,
        localTotalColumnIndexes: group.localTotalColumnIndexes,
      })),
      [
        { label: "Gender", columnIndexes: [0, 1, 2], localTotalColumnIndexes: [0] },
        { label: "Age", columnIndexes: [3, 4, 5], localTotalColumnIndexes: [3] },
      ]
    );
  });

  it("produces local Total comparison pairs so non-total siblings compare against the sparse Total", () => {
    const bannerContext = {
      selectedColumnCount: 6,
      lowerBannerRow: ["", "Male", "Female", "", "18-24", "25-34"],
      upperScanRows: [
        ["Всего", "M-prefix", "F-prefix", "Total", "Young", "Old"],
        ["", "Gender", "", "Age", "", ""],
      ],
    };

    const result = detectBannerStructure(bannerContext);
    const pairs = buildColumnComparisonPairs(
      6,
      { respectBannerStructure: true },
      new Set(),
      result
    );

    const totalPairs = pairs
      .filter((pair) => pair.comparisonType === "bannerTotal")
      .map((pair) => [pair.firstColumnIndex, pair.secondColumnIndex]);

    assert.deepStrictEqual(totalPairs, [
      [0, 1],
      [0, 2],
      [3, 4],
      [3, 5],
    ]);

    const segmentPairs = pairs
      .filter((pair) => pair.comparisonType === "bannerGroup")
      .map((pair) => [pair.firstColumnIndex, pair.secondColumnIndex]);

    assert.deepStrictEqual(segmentPairs, [
      [1, 2],
      [4, 5],
    ]);
  });

  it("skips over an empty upper row and honors the next visible upper cell", () => {
    const bannerContext = {
      selectedColumnCount: 4,
      lowerBannerRow: ["", "Male", "", "Female"],
      upperScanRows: [
        ["", "", "", ""],
        ["Total", "M-prefix", "Всего", "F-prefix"],
      ],
    };

    const result = detectBannerStructure(bannerContext);

    assert.strictEqual(result.columnDescriptors[0].isTotal, true);
    assert.strictEqual(result.columnDescriptors[0].isLocalTotal, true);
    assert.strictEqual(result.columnDescriptors[2].isTotal, true);
    assert.strictEqual(result.columnDescriptors[2].isLocalTotal, true);
    assert.strictEqual(result.columnDescriptors[1].isTotal, false);
    assert.strictEqual(result.columnDescriptors[3].isTotal, false);
  });

  it("does not mark a column as total when the nearest upper cell is a non-total label", () => {
    const bannerContext = {
      selectedColumnCount: 4,
      lowerBannerRow: ["", "Male", "Female", "25-34"],
      upperScanRows: [
        ["Brand", "M-prefix", "F-prefix", "Age"],
        ["", "Gender", "", "Age"],
      ],
    };

    const result = detectBannerStructure(bannerContext);

    for (const descriptor of result.columnDescriptors) {
      assert.strictEqual(descriptor.isTotal, false);
      assert.strictEqual(descriptor.isLocalTotal, false);
    }

    assert.deepStrictEqual(result.totalColumnIndexes, []);
  });

  it("does not change behavior when the lower banner row already carries the Total label", () => {
    const bannerContext = {
      selectedColumnCount: 3,
      lowerBannerRow: ["Total", "Male", "Female"],
      upperScanRows: [["", "Gender", ""]],
    };

    const result = detectBannerStructure(bannerContext);

    assert.strictEqual(result.columnDescriptors[0].isTotal, true);
    assert.strictEqual(result.columnDescriptors[1].isTotal, false);
    assert.strictEqual(result.columnDescriptors[2].isTotal, false);

    const labelMap = buildBannerLocalSignificanceLabelMap(result);

    assert.deepStrictEqual(
      Array.from({ length: 3 }, (_, columnIndex) => labelMap.get(columnIndex) || ""),
      ["", "a", "b"]
    );
  });
});

describe("detectBannerStructure - global Total promotion", () => {
  it("promotes a stand-alone first Total column to global when no parent group covers it", () => {
    const bannerContext = {
      selectedColumnCount: 4,
      lowerBannerRow: ["Всего", "Daily", "Weekly", "Monthly"],
      upperScanRows: [
        ["", "Пользование категорией", "Пользование категорией", "Пользование категорией"],
      ],
    };

    const result = detectBannerStructure(bannerContext);

    assert.strictEqual(result.globalTotalColumnIndex, 0);
    assert.strictEqual(result.columnDescriptors[0].isGlobalTotal, true);
    assert.strictEqual(result.columnDescriptors[0].isLocalTotal, false);
    assert.strictEqual(result.columnDescriptors[0].totalType, "global");

    const labelMap = buildBannerLocalSignificanceLabelMap(result);

    assert.deepStrictEqual(
      Array.from({ length: 4 }, (_, columnIndex) => labelMap.get(columnIndex) || ""),
      ["", "a", "b", "c"]
    );

    const pairs = buildColumnComparisonPairs(
      4,
      { respectBannerStructure: true },
      new Set(),
      result
    );

    const totalPairs = pairs
      .filter((pair) => pair.comparisonType === "bannerTotal")
      .map((pair) => [pair.firstColumnIndex, pair.secondColumnIndex, pair.totalReferenceType]);

    assert.deepStrictEqual(totalPairs, [
      [0, 1, "global"],
      [0, 2, "global"],
      [0, 3, "global"],
    ]);
  });

  it("keeps repeated Total columns inside named groups as local Totals (no global promotion)", () => {
    const bannerContext = {
      selectedColumnCount: 6,
      lowerBannerRow: ["Всего", "Male", "Female", "Всего", "18-24", "25-34"],
      upperScanRows: [
        ["Gender", "Gender", "Gender", "Age", "Age", "Age"],
      ],
    };

    const result = detectBannerStructure(bannerContext);

    assert.strictEqual(result.globalTotalColumnIndex, null);
    assert.deepStrictEqual(result.totalColumnIndexes, [0, 3]);

    for (const columnIndex of [0, 3]) {
      const descriptor = result.columnDescriptors[columnIndex];
      assert.strictEqual(descriptor.isTotal, true);
      assert.strictEqual(descriptor.isLocalTotal, true);
      assert.strictEqual(descriptor.isGlobalTotal, false);
    }

    const pairs = buildColumnComparisonPairs(
      6,
      { respectBannerStructure: true },
      new Set(),
      result
    );

    const totalPairs = pairs
      .filter((pair) => pair.comparisonType === "bannerTotal")
      .map((pair) => [pair.firstColumnIndex, pair.secondColumnIndex, pair.totalReferenceType]);

    assert.deepStrictEqual(totalPairs, [
      [0, 1, "local"],
      [0, 2, "local"],
      [3, 4, "local"],
      [3, 5, "local"],
    ]);
  });

  it("keeps sparse upper-level Totals attached to adjacent named groups as local (no global promotion)", () => {
    const bannerContext = {
      selectedColumnCount: 6,
      lowerBannerRow: ["", "Male", "Female", "", "18-24", "25-34"],
      upperScanRows: [
        ["Всего", "M-prefix", "F-prefix", "Total", "Young", "Old"],
        ["", "Gender", "", "Age", "", ""],
      ],
    };

    const result = detectBannerStructure(bannerContext);

    assert.strictEqual(result.globalTotalColumnIndex, null);
    assert.strictEqual(result.columnDescriptors[0].isLocalTotal, true);
    assert.strictEqual(result.columnDescriptors[0].isGlobalTotal, false);
    assert.strictEqual(result.columnDescriptors[3].isLocalTotal, true);
    assert.strictEqual(result.columnDescriptors[3].isGlobalTotal, false);
  });

  it("does not promote when the first Total column sits inside a multi-column parent group span", () => {
    const bannerContext = {
      selectedColumnCount: 3,
      lowerBannerRow: ["Total", "Male", "Female"],
      upperScanRows: [["Total", "Total", "Total"]],
    };

    const result = detectBannerStructure(bannerContext);

    assert.strictEqual(result.globalTotalColumnIndex, null);
    assert.strictEqual(result.columnDescriptors[0].isLocalTotal, true);
  });
});
