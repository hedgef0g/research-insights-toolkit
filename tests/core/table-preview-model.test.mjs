import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { buildTablePreviewModel } from "../../src/core/table-preview-model.js";
import { normalizeSelectedRange } from "../../src/core/range-normalizer.js";

function makeSimplePreviewModel(labels, options = {}) {
  const values = [
    [0.21, 0.45, 0.33],
    [0.15, 0.28, 0.57],
    [100, 200, 150],
  ];
  const leftLabelValues = labels.map((label) => [label]);
  return buildTablePreviewModel({ values, leftLabelValues, ...options });
}

function makeCustomPreviewModel(values, labels, options = {}) {
  const leftLabelValues = labels.map((label) => [label]);
  return buildTablePreviewModel({ values, leftLabelValues, ...options });
}

function buildNormalizedPreviewModel(rawText, cleanedValues = rawText, settings = {}) {
  const normalized = normalizeSelectedRange(cleanedValues, rawText);
  assert.strictEqual(normalized.normalizationApplied, true, "normalization must succeed first");

  const model = buildTablePreviewModel({
    values: normalized.valuesForCalculation,
    leftLabelValues: normalized.leftLabelValues,
    bannerContext: normalized.bannerContext,
    settings,
  });

  return { normalized, model };
}

function warningCodes(model) {
  return model.warnings.map((warning) => warning.code);
}

function findBlock(model, metricType) {
  return model.calculationBlocks.find((block) => block.metricType === metricType);
}

describe("buildTablePreviewModel - normalized full-table regression coverage", () => {
  it("builds a stable preview model for normalized full-table proportions", () => {
    const rawText = [
      ["Usage table", "", "", ""],
      ["", "Total", "Male", "Female"],
      ["", "2025Q4", "2025Q4", "2025Q4"],
      ["Agree", "44%", "41%", "39%"],
      ["Disagree", "56%", "59%", "61%"],
      ["BASE", "5605", "1320", "3083"],
    ];
    const cleanedValues = rawText.map((row, rowIndex) =>
      rowIndex >= 3 ? ["", ...row.slice(1)] : [...row]
    );

    const { normalized, model } = buildNormalizedPreviewModel(rawText, cleanedValues);
    const proportionBlock = findBlock(model, "proportion");

    assert.deepStrictEqual(normalized.titleRows, [0]);
    assert.deepStrictEqual(normalized.bannerRows, [1, 2]);
    assert.deepStrictEqual(normalized.leftLabelValues, [["Agree"], ["Disagree"], ["BASE"]]);

    assert.deepStrictEqual(model.rowDiagnostics.map((row) => row.primaryLabel), [
      "Agree",
      "Disagree",
      "BASE",
    ]);
    assert.deepStrictEqual(model.summary, {
      rowCount: 3,
      columnCount: 3,
      detectedMetricRows: 1,
      detectedBlocks: 1,
      baseRows: 1,
      hasNps: false,
      hasMeans: false,
      hasBanner: false,
      hasGlobalTotal: false,
      hasWaveGroups: false,
    });
    assert.deepStrictEqual(proportionBlock, {
      metricType: "proportion",
      valueRowIndexes: [0, 1],
      valueRowIndex: null,
      baseRowIndex: 2,
      baseSubtype: null,
      baseSelection: {
        selectedBaseRowIndex: 2,
        selectedBaseSubtype: null,
        selectedBaseLabel: "BASE",
        isWeightedFallback: false,
      },
      promotersRowIndex: null,
      detractorsRowIndex: null,
      neutralRowIndex: null,
      sdRowIndex: null,
      varianceRowIndex: null,
      notes: [],
    });
    assert.ok(
      !warningCodes(model).includes("MISSING_ROW_LABEL_WITH_DATA"),
      `unexpected warning set: ${JSON.stringify(model.warnings)}`
    );
  });

  it("accepts normalizer-style multi-level banner context and detects groups", () => {
    const model = makeSimplePreviewModel(["Agree", "Disagree", "BASE"], {
      bannerContext: {
        scanRows: [
          ["Gender", "", ""],
          ["Total", "Male", "Female"],
        ],
        columnCount: 3,
      },
      settings: { respectBannerStructure: true },
    });

    assert.strictEqual(model.bannerStructure.isDetected, true);
    assert.strictEqual(model.summary.hasBanner, true);
    assert.strictEqual(model.bannerStructure.mode, "twoLevel");
    assert.deepStrictEqual(model.bannerStructure.totalColumnIndexes, [0]);
    assert.deepStrictEqual(
      model.bannerStructure.groups.map((group) => ({
        label: group.label,
        columnIndexes: group.columnIndexes,
        localTotalColumnIndexes: group.localTotalColumnIndexes,
      })),
      [{ label: "Gender", columnIndexes: [0, 1, 2], localTotalColumnIndexes: [0] }]
    );
  });

  it("keeps local Total groups aligned with banner-detector behavior", () => {
    const model = makeSimplePreviewModel(["Agree", "Disagree", "BASE"], {
      bannerContext: {
        scanRows: [
          ["", "Gender", "", "Age", "", ""],
          ["Total", "M-prefix", "F-prefix", "Total", "Young", "Old"],
          ["", "Male", "Female", "", "18-24", "25-34"],
        ],
        columnCount: 6,
      },
      settings: { respectBannerStructure: true },
    });

    assert.strictEqual(model.bannerStructure.globalTotalColumnIndex, null);
    assert.deepStrictEqual(model.bannerStructure.totalColumnIndexes, [0, 3]);
    assert.deepStrictEqual(
      model.bannerStructure.groups.map((group) => ({
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

  it("promotes a stand-alone first Total column to global Total in preview", () => {
    const model = makeSimplePreviewModel(["Agree", "Disagree", "BASE"], {
      bannerContext: {
        scanRows: [
          ["", "Category usage", "Category usage", "Category usage"],
          ["Total", "Daily", "Weekly", "Monthly"],
        ],
        columnCount: 4,
      },
      settings: { respectBannerStructure: true },
    });

    assert.strictEqual(model.bannerStructure.globalTotalColumnIndex, 0);
    assert.strictEqual(model.summary.hasBanner, true);
    assert.strictEqual(model.summary.hasGlobalTotal, true);
    assert.deepStrictEqual(model.bannerStructure.totalColumnIndexes, [0]);
  });

  it("reports nested wave-aware groups from normalizer banner context", () => {
    const rawText = [
      ["Wave table", "", "", "", "", "", "", "", "", "", "", "", ""],
      ["", "", "", "Category usage", "", "", "", "", "", "", "", "", ""],
      ["", "Wave (quarter)", "", "Buys all myself", "", "Most of it", "", "Half", "", "Less", "", "Rarely", ""],
      ["", "2025Q4", "2026Q1", "Wave (quarter)", "", "Wave (quarter)", "", "Wave (quarter)", "", "Wave (quarter)", "", "Wave (quarter)", ""],
      ["", "", "", "2025Q4", "2026Q1", "2025Q4", "2026Q1", "2025Q4", "2026Q1", "2025Q4", "2026Q1", "2025Q4", "2026Q1"],
      ["Male", "50%", "41%", "39%", "41%", "44%", "42%", "48%", "43%", "51%", "49%", "46%", "45%"],
      ["Female", "50%", "59%", "61%", "59%", "56%", "58%", "52%", "57%", "49%", "51%", "54%", "55%"],
      ["BASE", "5605", "1320", "3083", "1045", "2200", "900", "1800", "760", "1500", "700", "1300", "620"],
    ];
    const cleanedValues = rawText.map((row, rowIndex) =>
      rowIndex >= 5 ? ["", ...row.slice(1)] : [...row]
    );

    const { model } = buildNormalizedPreviewModel(rawText, cleanedValues, {
      respectBannerStructure: true,
      autoDetectWaveBanners: true,
    });

    assert.strictEqual(model.summary.hasBanner, true);
    assert.strictEqual(model.summary.hasWaveGroups, true);
    assert.strictEqual(model.bannerStructure.recommendedComparisonMode, "mixed");
    assert.deepStrictEqual(
      model.bannerStructure.groups.map((group) => ({
        columnIndexes: group.columnIndexes,
        semanticType: group.semanticType,
        recommendedComparisonMode: group.recommendedComparisonMode,
      })),
      [
        {
          columnIndexes: [0, 1],
          semanticType: "wave",
          recommendedComparisonMode: "previousColumn",
        },
        {
          columnIndexes: [2, 3],
          semanticType: "wave",
          recommendedComparisonMode: "previousColumn",
        },
        {
          columnIndexes: [4, 5],
          semanticType: "wave",
          recommendedComparisonMode: "previousColumn",
        },
        {
          columnIndexes: [6, 7],
          semanticType: "wave",
          recommendedComparisonMode: "previousColumn",
        },
        {
          columnIndexes: [8, 9],
          semanticType: "wave",
          recommendedComparisonMode: "previousColumn",
        },
        {
          columnIndexes: [10, 11],
          semanticType: "wave",
          recommendedComparisonMode: "previousColumn",
        },
      ]
    );
  });

  it("preserves extended NPS labels and builds the expected NPS block", () => {
    const rawText = [
      ["NPS full table", "", ""],
      ["", "Total", "Segment A"],
      ["", "2025Q4", "2026Q1"],
      ["1", "1%", "2%"],
      ["2", "2%", "3%"],
      ["3", "3%", "4%"],
      ["4", "4%", "5%"],
      ["5", "5%", "6%"],
      ["6", "6%", "7%"],
      ["7", "7%", "8%"],
      ["8", "8%", "9%"],
      ["9", "9%", "10%"],
      ["10", "10%", "11%"],
      ["Bottom-3", "6%", "9%"],
      ["Mid-4", "22%", "26%"],
      ["Top-3", "72%", "65%"],
      ["Detractors", "6%", "9%"],
      ["Neutral", "22%", "26%"],
      ["Promoters", "72%", "65%"],
      ["NPS", "66%", "56%"],
      ["BASE", "1000", "800"],
    ];
    const cleanedValues = rawText.map((row, rowIndex) =>
      rowIndex >= 3 ? ["", ...row.slice(1)] : [...row]
    );

    const { normalized, model } = buildNormalizedPreviewModel(rawText, cleanedValues);
    const npsBlock = findBlock(model, "npsStructure");

    assert.deepStrictEqual(
      normalized.leftLabelValues.slice(0, 10).map((row) => row[0]),
      ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"]
    );
    assert.deepStrictEqual(
      model.rowDiagnostics.slice(0, 10).map((row) => row.primaryLabel),
      ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"]
    );
    assert.deepStrictEqual(npsBlock, {
      metricType: "npsStructure",
      valueRowIndexes: [16],
      valueRowIndex: 16,
      baseRowIndex: 17,
      baseSubtype: null,
      baseSelection: {
        selectedBaseRowIndex: 17,
        selectedBaseSubtype: null,
        selectedBaseLabel: "BASE",
        isWeightedFallback: false,
      },
      promotersRowIndex: 15,
      detractorsRowIndex: 13,
      neutralRowIndex: null,
      sdRowIndex: null,
      varianceRowIndex: null,
      notes: [],
    });
    assert.strictEqual(model.rowDiagnostics[17].rowType, "base");
    assert.ok(
      !warningCodes(model).includes("MISSING_ROW_LABEL_WITH_DATA"),
      `unexpected warning set: ${JSON.stringify(model.warnings)}`
    );
  });

  it("treats mean, variance, and BASE as service rows in normalized full-table preview", () => {
    const rawText = [
      ["Age table", "", ""],
      ["", "Total", ""],
      ["", "Wave (quarter)", ""],
      ["", "2025Q4", "2026Q1"],
      ["", "(a)", "(a)"],
      ["18-24", "19%", "15%"],
      ["25-34", "37%", "32%"],
      ["35-44", "28%", "28%"],
      ["45+", "17%", "24%"],
      ["mean", "29.4", "31.5"],
      ["variance", "103.6", "132.6"],
      ["BASE", "5605", "1320"],
      ["All respondents", "", ""],
    ];
    const cleanedValues = rawText.map((row, rowIndex) =>
      rowIndex >= 5 && rowIndex <= 11 ? ["", ...row.slice(1)] : [...row]
    );

    const { model } = buildNormalizedPreviewModel(rawText, cleanedValues);
    const meanBlock = findBlock(model, "mean");

    assert.deepStrictEqual(
      model.rowDiagnostics.slice(4).map((row) => ({
        label: row.primaryLabel,
        rowType: row.rowType,
      })),
      [
        { label: "mean", rowType: "mean" },
        { label: "variance", rowType: "variance" },
        { label: "BASE", rowType: "base" },
      ]
    );
    assert.deepStrictEqual(meanBlock, {
      metricType: "mean",
      valueRowIndexes: [4],
      valueRowIndex: 4,
      baseRowIndex: 6,
      baseSubtype: null,
      baseSelection: {
        selectedBaseRowIndex: 6,
        selectedBaseSubtype: null,
        selectedBaseLabel: "BASE",
        isWeightedFallback: false,
      },
      promotersRowIndex: null,
      detractorsRowIndex: null,
      neutralRowIndex: null,
      sdRowIndex: null,
      varianceRowIndex: 5,
      notes: [],
    });
    assert.strictEqual(model.summary.hasMeans, true);
    assert.ok(
      !warningCodes(model).includes("MISSING_ROW_LABEL_WITH_DATA"),
      `unexpected warning set: ${JSON.stringify(model.warnings)}`
    );
  });
});

describe("buildTablePreviewModel - explicit base requirement", () => {
  it("does not report a complete proportion block when base is missing", () => {
    const model = makeCustomPreviewModel(
      [
        [0.21, 0.45, 0.33],
        [0.15, 0.28, 0.57],
      ],
      ["Agree", "Disagree"]
    );

    assert.deepStrictEqual(model.calculationBlocks, []);
    assert.strictEqual(model.summary.detectedBlocks, 0);
  });

  it("still reports a proportion block when an explicit base is present", () => {
    const model = makeSimplePreviewModel(["Agree", "Disagree", "BASE"]);
    const proportionBlock = findBlock(model, "proportion");

    assert.ok(proportionBlock, "expected a proportion block");
    assert.deepStrictEqual(proportionBlock.valueRowIndexes, [0, 1]);
    assert.strictEqual(proportionBlock.baseRowIndex, 2);
  });

  it("does not report a mean block when variance is present but base is missing", () => {
    const model = makeCustomPreviewModel(
      [
        [29.4, 31.5],
        [103.6, 132.6],
      ],
      ["Mean", "Variance"]
    );

    assert.deepStrictEqual(model.calculationBlocks, []);
    assert.strictEqual(model.summary.detectedBlocks, 0);
  });

  it("still reports a mean block when variance and explicit base are present", () => {
    const model = makeCustomPreviewModel(
      [
        [29.4, 31.5],
        [103.6, 132.6],
        [5605, 1320],
      ],
      ["Mean", "Variance", "BASE"]
    );
    const meanBlock = findBlock(model, "mean");

    assert.ok(meanBlock, "expected a mean block");
    assert.strictEqual(meanBlock.baseRowIndex, 2);
    assert.strictEqual(meanBlock.varianceRowIndex, 1);
  });

  it("does not report an extended NPS block when base is missing", () => {
    const model = makeCustomPreviewModel(
      [
        [0.01, 0.02],
        [0.02, 0.03],
        [0.03, 0.04],
        [0.04, 0.05],
        [0.05, 0.06],
        [0.06, 0.07],
        [0.07, 0.08],
        [0.08, 0.09],
        [0.09, 0.1],
        [0.1, 0.11],
        [0.06, 0.09],
        [0.22, 0.26],
        [0.72, 0.65],
        [0.06, 0.09],
        [0.22, 0.26],
        [0.72, 0.65],
        [0.66, 0.56],
      ],
      [
        "1",
        "2",
        "3",
        "4",
        "5",
        "6",
        "7",
        "8",
        "9",
        "10",
        "Bottom-3",
        "Mid-4",
        "Top-3",
        "Detractors",
        "Neutral",
        "Promoters",
        "NPS",
      ]
    );

    assert.deepStrictEqual(model.calculationBlocks, []);
    assert.strictEqual(model.summary.detectedBlocks, 0);
  });

  it("still reports an extended NPS block when an explicit base is present", () => {
    const model = makeCustomPreviewModel(
      [
        [0.01, 0.02],
        [0.02, 0.03],
        [0.03, 0.04],
        [0.04, 0.05],
        [0.05, 0.06],
        [0.06, 0.07],
        [0.07, 0.08],
        [0.08, 0.09],
        [0.09, 0.1],
        [0.1, 0.11],
        [0.06, 0.09],
        [0.22, 0.26],
        [0.72, 0.65],
        [0.06, 0.09],
        [0.22, 0.26],
        [0.72, 0.65],
        [0.66, 0.56],
        [1000, 800],
      ],
      [
        "1",
        "2",
        "3",
        "4",
        "5",
        "6",
        "7",
        "8",
        "9",
        "10",
        "Bottom-3",
        "Mid-4",
        "Top-3",
        "Detractors",
        "Neutral",
        "Promoters",
        "NPS",
        "BASE",
      ]
    );
    const npsBlock = findBlock(model, "npsStructure");

    assert.ok(npsBlock, "expected an NPS block");
    assert.strictEqual(npsBlock.baseRowIndex, 17);
  });
});

describe("buildTablePreviewModel - header-only block suppression", () => {
  it("does not report blocks for banner-only descriptor selections", () => {
    const model = makeCustomPreviewModel(
      [
        ["Wave (quarter)", "Wave (quarter)"],
        ["2025Q4", "2026Q1"],
      ],
      ["Wave (quarter)", "2025Q4"]
    );

    assert.deepStrictEqual(model.calculationBlocks, []);
    assert.strictEqual(model.summary.detectedBlocks, 0);
  });

  it("does not report blocks for quarter-label-only selections with year-quarter values", () => {
    const model = makeCustomPreviewModel(
      [
        ["2025Q4", "2026Q1"],
        ["2026Q2", "2026Q3"],
      ],
      ["2025Q4", "2026Q2"]
    );

    assert.deepStrictEqual(model.calculationBlocks, []);
    assert.strictEqual(model.summary.detectedBlocks, 0);
  });

  it("still reports blocks for real banner-heavy tables with numeric body rows", () => {
    const rawText = [
      ["Usage table", "", "", ""],
      ["", "Wave (quarter)", "Wave (quarter)", "Wave (quarter)"],
      ["", "2025Q4", "2026Q1", "2026Q2"],
      ["Agree", "44%", "41%", "39%"],
      ["Disagree", "56%", "59%", "61%"],
      ["BASE", "5605", "1320", "3083"],
    ];
    const cleanedValues = rawText.map((row, rowIndex) =>
      rowIndex >= 3 ? ["", ...row.slice(1)] : [...row]
    );

    const { model } = buildNormalizedPreviewModel(rawText, cleanedValues);
    const proportionBlock = findBlock(model, "proportion");

    assert.ok(proportionBlock, "expected a proportion block for numeric body rows");
    assert.deepStrictEqual(proportionBlock.valueRowIndexes, [0, 1]);
    assert.strictEqual(proportionBlock.baseRowIndex, 2);
  });
});

describe("buildTablePreviewModel - label warning heuristics", () => {
  it('"q12" emits SUSPICIOUS_CODE_LIKE_LABEL', () => {
    const model = makeSimplePreviewModel(["q12", "Agreement", "Base"]);
    assert.ok(
      warningCodes(model).includes("SUSPICIOUS_CODE_LIKE_LABEL"),
      `expected SUSPICIOUS_CODE_LIKE_LABEL; got: ${JSON.stringify(warningCodes(model))}`
    );
  });

  it('"Concept Test" does not emit SUSPICIOUS_PLACEHOLDER_LABEL', () => {
    const model = makeSimplePreviewModel(["Concept Test", "Agreement", "Base"]);
    assert.ok(
      !warningCodes(model).includes("SUSPICIOUS_PLACEHOLDER_LABEL"),
      `expected no SUSPICIOUS_PLACEHOLDER_LABEL; got: ${JSON.stringify(warningCodes(model))}`
    );
  });

  it('"test row" emits SUSPICIOUS_PLACEHOLDER_LABEL', () => {
    const model = makeSimplePreviewModel(["test row", "Agreement", "Base"]);
    assert.ok(
      warningCodes(model).includes("SUSPICIOUS_PLACEHOLDER_LABEL"),
      `expected SUSPICIOUS_PLACEHOLDER_LABEL; got: ${JSON.stringify(warningCodes(model))}`
    );
  });
});
