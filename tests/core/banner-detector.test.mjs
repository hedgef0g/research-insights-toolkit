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
        recommendedComparisonMode: group.recommendedComparisonMode,
      })),
      [
        {
          label: "Total",
          columnIndexes: [0, 1, 2, 3],
          recommendedComparisonMode: "previousColumn",
        },
        {
          label: "Category usage",
          columnIndexes: [4, 5, 6, 7],
          recommendedComparisonMode: "previousColumn",
        },
        {
          label: "Gender",
          columnIndexes: [8, 9],
          recommendedComparisonMode: "previousColumn",
        },
        {
          label: "Age",
          columnIndexes: [10, 11],
          recommendedComparisonMode: "previousColumn",
        },
        {
          label: "Geo",
          columnIndexes: [12, 13],
          recommendedComparisonMode: "previousColumn",
        },
      ]
    );

    const labelMap = buildBannerLocalSignificanceLabelMap(result, {
      autoDetectWaveBanners: true,
    });

    assert.deepStrictEqual(
      Array.from({ length: 14 }, (_, columnIndex) => labelMap.get(columnIndex) || ""),
      ["", "", "", "", "", "", "", "", "", "", "", "", "", ""]
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
        recommendedComparisonMode: group.recommendedComparisonMode,
      })),
      [
        {
          label: "Всего",
          columnIndexes: [0, 1],
          recommendedComparisonMode: "previousColumn",
        },
        {
          label: "Пользование категорией",
          columnIndexes: [2, 3, 4, 5],
          recommendedComparisonMode: "previousColumn",
        },
      ]
    );

    const labelMap = buildBannerLocalSignificanceLabelMap(result, {
      autoDetectWaveBanners: true,
    });

    assert.deepStrictEqual(
      Array.from({ length: 6 }, (_, columnIndex) => labelMap.get(columnIndex) || ""),
      ["", "", "", "", "", ""]
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

  it("promotes every category group with a nested Волна (квартал) dimension to wave-aware", () => {
    const bannerContext = {
      selectedColumnCount: 12,
      lowerBannerRow: [
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
          "Волна (квартал)",
          "Волна (квартал)",
          "Волна (квартал)",
          "Волна (квартал)",
          "Волна (квартал)",
          "Волна (квартал)",
          "Волна (квартал)",
          "Волна (квартал)",
          "Волна (квартал)",
          "Волна (квартал)",
          "Волна (квартал)",
          "Волна (квартал)",
        ],
        [
          "Всего",
          "",
          "Всё покупаю сам",
          "",
          "Большую часть",
          "",
          "Половину",
          "",
          "Меньшую часть",
          "",
          "Почти не участвую",
          "",
        ],
      ],
    };

    const result = detectBannerStructure(bannerContext, { autoDetectWaveBanners: true });

    assert.strictEqual(result.globalTotalColumnIndex, null);
    assert.deepStrictEqual(result.totalColumnIndexes, []);

    assert.deepStrictEqual(
      result.groups.map((group) => ({
        label: group.label,
        columnIndexes: group.columnIndexes,
        recommendedComparisonMode: group.recommendedComparisonMode,
        semanticType: group.semanticType,
      })),
      [
        {
          label: "Всего",
          columnIndexes: [0, 1],
          recommendedComparisonMode: "previousColumn",
          semanticType: "wave",
        },
        {
          label: "Всё покупаю сам",
          columnIndexes: [2, 3],
          recommendedComparisonMode: "previousColumn",
          semanticType: "wave",
        },
        {
          label: "Большую часть",
          columnIndexes: [4, 5],
          recommendedComparisonMode: "previousColumn",
          semanticType: "wave",
        },
        {
          label: "Половину",
          columnIndexes: [6, 7],
          recommendedComparisonMode: "previousColumn",
          semanticType: "wave",
        },
        {
          label: "Меньшую часть",
          columnIndexes: [8, 9],
          recommendedComparisonMode: "previousColumn",
          semanticType: "wave",
        },
        {
          label: "Почти не участвую",
          columnIndexes: [10, 11],
          recommendedComparisonMode: "previousColumn",
          semanticType: "wave",
        },
      ]
    );

    const labelMap = buildBannerLocalSignificanceLabelMap(result, {
      autoDetectWaveBanners: true,
    });

    assert.deepStrictEqual(
      Array.from({ length: 12 }, (_, columnIndex) => labelMap.get(columnIndex) || ""),
      ["", "", "", "", "", "", "", "", "", "", "", ""]
    );

    const pairs = buildColumnComparisonPairs(
      12,
      { respectBannerStructure: true, autoDetectWaveBanners: true },
      new Set(),
      result
    );

    const previousColumnPairs = pairs
      .filter((pair) => pair.comparisonType === "previousColumn")
      .map((pair) => [pair.firstColumnIndex, pair.secondColumnIndex]);

    assert.deepStrictEqual(previousColumnPairs, [
      [0, 1],
      [2, 3],
      [4, 5],
      [6, 7],
      [8, 9],
      [10, 11],
    ]);

    const crossGroupPairs = pairs.filter(
      (pair) =>
        pair.comparisonType !== "previousColumn" &&
        pair.comparisonType !== "bannerTotal"
    );

    assert.deepStrictEqual(crossGroupPairs, []);
  });

  it("keeps nested-wave behavior off when autoDetectWaveBanners is disabled", () => {
    const bannerContext = {
      selectedColumnCount: 4,
      lowerBannerRow: ["2025Q4", "2026Q1", "2025Q4", "2026Q1"],
      upperScanRows: [
        ["Волна (квартал)", "Волна (квартал)", "Волна (квартал)", "Волна (квартал)"],
        ["Всего", "", "Всё покупаю сам", ""],
      ],
    };

    const result = detectBannerStructure(bannerContext, { autoDetectWaveBanners: false });

    for (const group of result.groups) {
      assert.notStrictEqual(group.recommendedComparisonMode, "previousColumn");
      assert.notStrictEqual(group.semanticType, "wave");
    }
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

  it("groups first wave-value pair together when label column was stripped and no explicit parent label exists", () => {
    // Reproduces the partial-banner-selection bug: after stripping one label
    // column, the bannerContext upper row starts with consecutive single-column
    // wave-value spans ("2025Q4", "2026Q1") with blank lower-banner cells,
    // while the remaining columns are covered by explicit "Волна (квартал)"
    // parent spans.  Without the merge fix, "2025Q4" and "2026Q1" each become
    // a one-column group and both receive local label "a".  After the fix they
    // form one two-column group and receive "a" and "b".
    const bannerContext = {
      selectedColumnCount: 12,
      lowerBannerRow: [
        "", "", "2025Q4", "2026Q1", "2025Q4", "2026Q1",
        "2025Q4", "2026Q1", "2025Q4", "2026Q1", "2025Q4", "2026Q1",
      ],
      upperScanRows: [
        [
          "2025Q4", "2026Q1",
          "Волна (квартал)", "", "Волна (квартал)", "",
          "Волна (квартал)", "", "Волна (квартал)", "",
          "Волна (квартал)", "",
        ],
      ],
    };

    const result = detectBannerStructure(bannerContext, { autoDetectWaveBanners: false });

    // Columns 0 and 1 must belong to the SAME group — they are the first wave
    // pair and are compared with each other.
    const groupForCol0 = result.groups.find((g) => g.columnIndexes.includes(0));
    const groupForCol1 = result.groups.find((g) => g.columnIndexes.includes(1));

    assert.ok(groupForCol0, "column 0 must be in a group");
    assert.ok(groupForCol1, "column 1 must be in a group");
    assert.strictEqual(
      groupForCol0.groupKey,
      groupForCol1.groupKey,
      "columns 0 and 1 must share the same group key"
    );
    assert.deepStrictEqual(
      groupForCol0.columnIndexes,
      [0, 1],
      "the first group must cover exactly columns 0 and 1"
    );

    // labelMap must assign distinct labels to cols 0 and 1.
    const labelMap = buildBannerLocalSignificanceLabelMap(result, { autoDetectWaveBanners: false });

    assert.strictEqual(labelMap.get(0), "a", "first data column must get label a");
    assert.strictEqual(labelMap.get(1), "b", "second data column must get label b");

    // Columns 2-3 must also form their own group (not collapsed into the first group).
    const groupForCol2 = result.groups.find((g) => g.columnIndexes.includes(2));
    const groupForCol3 = result.groups.find((g) => g.columnIndexes.includes(3));

    assert.ok(groupForCol2, "column 2 must be in a group");
    assert.strictEqual(
      groupForCol2.groupKey,
      groupForCol3.groupKey,
      "columns 2 and 3 must share the same group key"
    );
    assert.notStrictEqual(
      groupForCol0.groupKey,
      groupForCol2.groupKey,
      "the first group and the second group must be distinct"
    );

    assert.strictEqual(labelMap.get(2), "a", "first column of second group must get label a");
    assert.strictEqual(labelMap.get(3), "b", "second column of second group must get label b");
  });

  it("keeps real three-level subgroup splits separate when an earlier mixed-depth segment uses vertically merged wave labels", () => {
    const bannerContext = {
      selectedColumnCount: 6,
      lowerBannerRow: ["", "", "w6 (ноябрь 23)", "w7", "w6 (ноябрь 23)", "w7"],
      upperScanRows: [
        ["w6 (ноябрь 23)", "w7", "Мужской", "", "Женский", ""],
        ["Пол", "", "Пол", "", "", ""],
      ],
    };

    const result = detectBannerStructure(bannerContext, { autoDetectWaveBanners: false });

    assert.deepStrictEqual(
      result.groups.map((group) => ({
        label: group.label,
        bannerPath: group.bannerPath,
        columnIndexes: group.columnIndexes,
      })),
      [
        {
          label: "Пол",
          bannerPath: ["Пол"],
          columnIndexes: [0, 1],
        },
        {
          label: "Мужской",
          bannerPath: ["Пол", "Мужской"],
          columnIndexes: [2, 3],
        },
        {
          label: "Женский",
          bannerPath: ["Пол", "Женский"],
          columnIndexes: [4, 5],
        },
      ]
    );

    const labelMap = buildBannerLocalSignificanceLabelMap(result, {
      autoDetectWaveBanners: false,
    });

    assert.deepStrictEqual(
      Array.from({ length: 6 }, (_, columnIndex) => labelMap.get(columnIndex) || ""),
      ["a", "b", "a", "b", "a", "b"]
    );

    const pairs = buildColumnComparisonPairs(
      6,
      { respectBannerStructure: true, autoDetectWaveBanners: false },
      new Set(),
      result
    );

    const bannerGroupPairs = pairs
      .filter((pair) => pair.comparisonType === "bannerGroup")
      .map((pair) => [pair.firstColumnIndex, pair.secondColumnIndex, pair.groupLabel]);

    assert.deepStrictEqual(bannerGroupPairs, [
      [0, 1, "Пол"],
      [2, 3, "Мужской"],
      [4, 5, "Женский"],
    ]);
  });
});

describe("detectBannerStructure - compact w-number wave value labels", () => {
  it("compact w-number labels in upper row with blank lower cells form one merged group with distinct sibling markers when autoDetectWaveBanners is false", () => {
    // w6 / w7 in the upper row with blank lower cells must be recognised as
    // technical wave-value labels so that mergeAdjacentWaveValueSpans merges
    // them into a single two-column group instead of leaving two orphan
    // one-column groups that both receive local label "a".
    const bannerContext = {
      selectedColumnCount: 6,
      lowerBannerRow: ["", "", "w1", "w6", "w1", "w6"],
      upperScanRows: [
        ["w6", "w7", "Мужской", "", "Женский", ""],
      ],
    };

    const result = detectBannerStructure(bannerContext, { autoDetectWaveBanners: false });

    const groupForCol0 = result.groups.find((g) => g.columnIndexes.includes(0));
    const groupForCol1 = result.groups.find((g) => g.columnIndexes.includes(1));

    assert.ok(groupForCol0, "column 0 must be in a group");
    assert.ok(groupForCol1, "column 1 must be in a group");
    assert.strictEqual(
      groupForCol0.groupKey,
      groupForCol1.groupKey,
      "columns 0 and 1 must share the same group key"
    );
    assert.deepStrictEqual(groupForCol0.columnIndexes, [0, 1]);

    const groupForCol2 = result.groups.find((g) => g.columnIndexes.includes(2));
    const groupForCol3 = result.groups.find((g) => g.columnIndexes.includes(3));

    assert.ok(groupForCol2, "column 2 must be in a group");
    assert.strictEqual(
      groupForCol2.groupKey,
      groupForCol3.groupKey,
      "columns 2 and 3 must share the same group key"
    );
    assert.notStrictEqual(
      groupForCol0.groupKey,
      groupForCol2.groupKey,
      "wave group and sibling group must be distinct"
    );

    const labelMap = buildBannerLocalSignificanceLabelMap(result, { autoDetectWaveBanners: false });

    assert.strictEqual(labelMap.get(0), "a", "first column of wave group must get label a");
    assert.strictEqual(labelMap.get(1), "b", "second column of wave group must get label b");
    assert.strictEqual(labelMap.get(2), "a", "first column of sibling group must get label a");
    assert.strictEqual(labelMap.get(3), "b", "second column of sibling group must get label b");

    for (const group of result.groups) {
      assert.notStrictEqual(group.recommendedComparisonMode, "previousColumn");
    }
  });

  it("compact w-number labels in lower row promote previousColumn mode when autoDetectWaveBanners is true but not when false", () => {
    const bannerContext = {
      selectedColumnCount: 4,
      lowerBannerRow: ["w6", "w7", "w6", "w7"],
      upperScanRows: [
        ["Group A", "", "Group B", ""],
      ],
    };

    const resultTrue = detectBannerStructure(bannerContext, { autoDetectWaveBanners: true });

    assert.deepStrictEqual(
      resultTrue.groups.map((group) => ({
        label: group.label,
        columnIndexes: group.columnIndexes,
        recommendedComparisonMode: group.recommendedComparisonMode,
      })),
      [
        { label: "Group A", columnIndexes: [0, 1], recommendedComparisonMode: "previousColumn" },
        { label: "Group B", columnIndexes: [2, 3], recommendedComparisonMode: "previousColumn" },
      ]
    );

    const resultFalse = detectBannerStructure(bannerContext, { autoDetectWaveBanners: false });

    for (const group of resultFalse.groups) {
      assert.notStrictEqual(
        group.recommendedComparisonMode,
        "previousColumn",
        "previousColumn mode must not be set when autoDetectWaveBanners is false"
      );
    }
  });

  it("compact lower-row w18/w19 labels under semantic parent groups promote wave mode only when autoDetectWaveBanners is true", () => {
    // Regression test for realistic compact wave labels like "w18 (ноябрь 24)" in
    // the lowerBannerRow under named semantic groups.  Covers both the true and
    // false autoDetectWaveBanners paths.
    const bannerContext = {
      selectedColumnCount: 6,
      lowerBannerRow: [
        "w18 (ноябрь 24)", "w19 (май 25)",
        "w18 (ноябрь 24)", "w19 (май 25)",
        "w18 (ноябрь 24)", "w19 (май 25)",
      ],
      upperScanRows: [
        ["Всего", "", "Мужской", "", "Женский", ""],
      ],
    };

    const resultTrue = detectBannerStructure(bannerContext, { autoDetectWaveBanners: true });

    assert.deepStrictEqual(
      resultTrue.groups.map((g) => ({
        label: g.label,
        columnIndexes: g.columnIndexes,
        recommendedComparisonMode: g.recommendedComparisonMode,
      })),
      [
        { label: "Всего",   columnIndexes: [0, 1], recommendedComparisonMode: "previousColumn" },
        { label: "Мужской", columnIndexes: [2, 3], recommendedComparisonMode: "previousColumn" },
        { label: "Женский", columnIndexes: [4, 5], recommendedComparisonMode: "previousColumn" },
      ]
    );

    const resultFalse = detectBannerStructure(bannerContext, { autoDetectWaveBanners: false });

    // sibling groups must remain separate and marker labels must restart
    assert.strictEqual(resultFalse.groups.length, 3, "three sibling groups expected");

    const groupKeys = resultFalse.groups.map((g) => g.groupKey);
    assert.strictEqual(new Set(groupKeys).size, 3, "all three groups must have distinct keys");

    const labelMap = buildBannerLocalSignificanceLabelMap(resultFalse, { autoDetectWaveBanners: false });

    for (let base = 0; base < 6; base += 2) {
      assert.strictEqual(labelMap.get(base),     "a", `column ${base} must get label a`);
      assert.strictEqual(labelMap.get(base + 1), "b", `column ${base + 1} must get label b`);
    }

    for (const group of resultFalse.groups) {
      assert.notStrictEqual(group.recommendedComparisonMode, "previousColumn");
    }
  });

  it("compact w-number labels in upper scan row (vertically merged-like) promote wave groups only when autoDetectWaveBanners is true", () => {
    // Reproduces the Excel smoke failure: Office.js returns blank text for the
    // non-top cell of a vertically-merged banner cell, so the wave-value label
    // (e.g. "w18 (ноябрь 24)") appears in upperScanRows rather than
    // lowerBannerRow.  detectGroupKeysWithNestedWaveDimension must therefore
    // also scan upper scan rows for wave value labels.
    const bannerContext = {
      selectedColumnCount: 4,
      lowerBannerRow: ["", "", "", ""],
      upperScanRows: [
        ["w18 (ноябрь 24)", "w19 (май 25)", "w18 (ноябрь 24)", "w19 (май 25)"],
        ["Всего", "", "Мужской", ""],
      ],
    };

    const resultTrue = detectBannerStructure(bannerContext, { autoDetectWaveBanners: true });

    assert.deepStrictEqual(
      resultTrue.groups.map((g) => ({
        label: g.label,
        columnIndexes: g.columnIndexes,
        recommendedComparisonMode: g.recommendedComparisonMode,
      })),
      [
        { label: "Всего",   columnIndexes: [0, 1], recommendedComparisonMode: "previousColumn" },
        { label: "Мужской", columnIndexes: [2, 3], recommendedComparisonMode: "previousColumn" },
      ]
    );

    const resultFalse = detectBannerStructure(bannerContext, { autoDetectWaveBanners: false });

    // sibling groups stay separate with restarting labels even without wave mode
    assert.strictEqual(resultFalse.groups.length, 2, "two sibling groups expected");

    const g0 = resultFalse.groups.find((g) => g.columnIndexes.includes(0));
    const g2 = resultFalse.groups.find((g) => g.columnIndexes.includes(2));
    assert.notStrictEqual(g0.groupKey, g2.groupKey, "groups must be distinct");

    const labelMap = buildBannerLocalSignificanceLabelMap(resultFalse, { autoDetectWaveBanners: false });

    assert.strictEqual(labelMap.get(0), "a");
    assert.strictEqual(labelMap.get(1), "b");
    assert.strictEqual(labelMap.get(2), "a");
    assert.strictEqual(labelMap.get(3), "b");

    for (const group of resultFalse.groups) {
      assert.notStrictEqual(group.recommendedComparisonMode, "previousColumn");
    }
  });

  it("w-label variants with space / hyphen / underscore separator all promote wave groups when autoDetectWaveBanners is true", () => {
    // Root-cause regression for the real-workbook smoke failure:
    // Excel cells may render the wave number with a space ("w 18 (ноябрь 24)"),
    // a hyphen ("w-18"), or an underscore ("w_18").
    // normalizeLookupText does NOT strip hyphens or underscores, so the old
    // /^w\d+$/ first-token test failed for those variants.
    // The new regex /^w[\s\-_]?\d+(?!\w)/ on the full normalized label fixes all.
    const separator_variants = [
      ["w 18 (ноябрь 24)", "w 19 (май 25)"],   // space — the most common real-world format
      ["w-18 (ноябрь 24)", "w-19 (май 25)"],    // hyphen
      ["w_18 (ноябрь 24)", "w_19 (май 25)"],    // underscore
      ["W 18 (ноябрь 24)", "W 19 (май 25)"],    // uppercase + space
    ];

    for (const [label1, label2] of separator_variants) {
      const bannerContext = {
        selectedColumnCount: 4,
        lowerBannerRow: [label1, label2, label1, label2],
        upperScanRows: [["Всего", "", "Мужской", ""]],
      };

      const result = detectBannerStructure(bannerContext, { autoDetectWaveBanners: true });

      for (const group of result.groups) {
        assert.strictEqual(
          group.recommendedComparisonMode,
          "previousColumn",
          `variant "${label1}" / "${label2}" should promote previousColumn but got ${group.recommendedComparisonMode}`
        );
      }

      // autoDetectWaveBanners: false must never enable previousColumn
      const resultFalse = detectBannerStructure(bannerContext, { autoDetectWaveBanners: false });

      for (const group of resultFalse.groups) {
        assert.notStrictEqual(
          group.recommendedComparisonMode,
          "previousColumn",
          `variant "${label1}" must not trigger previousColumn when autoDetectWaveBanners is false`
        );
      }
    }
  });

  it("ordinary w-words (woman, work) are not recognised as wave value labels", () => {
    // Ensures the broadened regex does not create false positives for ordinary
    // words that start with the letter w but are not wave-number labels.
    // Note: "wave awareness" is intentionally omitted here because it does
    // trigger the existing WAVE_GROUP_LABEL_KEYWORDS descriptor path ("wave" is
    // in that list); that is expected and correct behaviour, not a false positive.
    const nonWaveWords = ["woman", "work"];

    for (const word of nonWaveWords) {
      const bannerContext = {
        selectedColumnCount: 4,
        lowerBannerRow: [word, word, word, word],
        upperScanRows: [["Group A", "", "Group B", ""]],
      };

      const result = detectBannerStructure(bannerContext, { autoDetectWaveBanners: true });

      for (const group of result.groups) {
        assert.notStrictEqual(
          group.recommendedComparisonMode,
          "previousColumn",
          `"${word}" must not trigger wave detection`
        );
      }
    }
  });

  it("compact w-number labels as the only upper scan row still promote wave groups when autoDetectWaveBanners is true", () => {
    // Root-cause regression for the groupLevelRowIndex skip bug:
    // When wave-value labels appear in the only upper scan row, the detector
    // elects that row as the group level (it is the only candidate, even though
    // it is technical).  The old hasWaveValueLabelsInIntermediateUpperRow skipped
    // the group-level row and therefore returned false, preventing wave promotion.
    // After removing the skip, the function finds the wave values and promotes.
    const bannerContext = {
      selectedColumnCount: 4,
      lowerBannerRow: ["", "", "", ""],
      upperScanRows: [
        ["w18 (ноябрь 24)", "w19 (май 25)", "w18 (ноябрь 24)", "w19 (май 25)"],
        // Only one upper row — after mergeAdjacentWaveValueSpans it becomes a
        // four-column technical span elected as the group level.
      ],
    };

    const resultTrue = detectBannerStructure(bannerContext, { autoDetectWaveBanners: true });

    for (const group of resultTrue.groups) {
      assert.strictEqual(
        group.recommendedComparisonMode,
        "previousColumn",
        "all groups in a pure wave-value banner should be wave-aware"
      );
    }

    const resultFalse = detectBannerStructure(bannerContext, { autoDetectWaveBanners: false });

    for (const group of resultFalse.groups) {
      assert.notStrictEqual(
        group.recommendedComparisonMode,
        "previousColumn",
        "previousColumn must not be enabled when autoDetectWaveBanners is false"
      );
    }
  });

  it("quarter value labels (2025Q4 / 2026Q1) in upper scan row promote wave groups only when autoDetectWaveBanners is true", () => {
    // Same scenario as the compact-w test above, but with existing quarter-style
    // labels so there is a regression test for the vertically-merged layout for
    // both label families in the same fix.
    const bannerContext = {
      selectedColumnCount: 4,
      lowerBannerRow: ["", "", "", ""],
      upperScanRows: [
        ["2025Q4", "2026Q1", "2025Q4", "2026Q1"],
        ["Всего", "", "Мужской", ""],
      ],
    };

    const resultTrue = detectBannerStructure(bannerContext, { autoDetectWaveBanners: true });

    assert.deepStrictEqual(
      resultTrue.groups.map((g) => ({
        label: g.label,
        columnIndexes: g.columnIndexes,
        recommendedComparisonMode: g.recommendedComparisonMode,
      })),
      [
        { label: "Всего",   columnIndexes: [0, 1], recommendedComparisonMode: "previousColumn" },
        { label: "Мужской", columnIndexes: [2, 3], recommendedComparisonMode: "previousColumn" },
      ]
    );

    const resultFalse = detectBannerStructure(bannerContext, { autoDetectWaveBanners: false });

    assert.strictEqual(resultFalse.groups.length, 2, "two sibling groups expected");

    const g0 = resultFalse.groups.find((g) => g.columnIndexes.includes(0));
    const g2 = resultFalse.groups.find((g) => g.columnIndexes.includes(2));
    assert.notStrictEqual(g0.groupKey, g2.groupKey, "groups must be distinct");

    const labelMap = buildBannerLocalSignificanceLabelMap(resultFalse, { autoDetectWaveBanners: false });

    assert.strictEqual(labelMap.get(0), "a");
    assert.strictEqual(labelMap.get(1), "b");
    assert.strictEqual(labelMap.get(2), "a");
    assert.strictEqual(labelMap.get(3), "b");

    for (const group of resultFalse.groups) {
      assert.notStrictEqual(group.recommendedComparisonMode, "previousColumn");
    }
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
