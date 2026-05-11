import { describe, it } from "node:test";
import assert from "node:assert/strict";

import {
  buildBannerDetectionDebugSummary,
  detectBannerStructure,
} from "../../src/core/banner-detector.js";
import { buildBannerLocalSignificanceLabelMap } from "../../src/core/significance.js";

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
