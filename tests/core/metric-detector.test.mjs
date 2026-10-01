import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  buildCalculationBlocks,
  classifyMetricLabel,
  detectMetricRowsFromLeftLabels,
  extractRowLabelFromLeftCells,
} from "../../src/core/metric-detector.js";
import { SIGNIFICANCE_FOOTNOTE_MARKER } from "../../src/core/significance-footnote.js";
import { BACKLINK_MARKER } from "../../src/core/generated-rows.js";

function detectBlocks(values, labels) {
  const leftLabelValues = labels.map((label) => [label]);
  const detectionResult = detectMetricRowsFromLeftLabels(values, leftLabelValues);
  return buildCalculationBlocks(detectionResult);
}

function detectBlocksWithPreference(values, labels, preferredBase) {
  const leftLabelValues = labels.map((label) => [label]);
  const detectionResult = detectMetricRowsFromLeftLabels(values, leftLabelValues);
  return buildCalculationBlocks(detectionResult, { preferredBase });
}

// ─── Suffix base-subtype label detection ──────────────────────────────────────

describe("classifyMetricLabel — suffix base-subtype patterns", () => {
  it('"BASE weighted" is classified as weighted base', () => {
    const r = classifyMetricLabel("BASE weighted");
    assert.strictEqual(r.rowType, "base");
    assert.strictEqual(r.baseSubtype, "weighted");
  });

  it('"Base weighted" is classified as weighted base', () => {
    const r = classifyMetricLabel("Base weighted");
    assert.strictEqual(r.rowType, "base");
    assert.strictEqual(r.baseSubtype, "weighted");
  });

  it('"Base unweighted" is classified as unweighted base', () => {
    const r = classifyMetricLabel("Base unweighted");
    assert.strictEqual(r.rowType, "base");
    assert.strictEqual(r.baseSubtype, "unweighted");
  });

  it('"Base effective" is classified as effective base', () => {
    const r = classifyMetricLabel("Base effective");
    assert.strictEqual(r.rowType, "base");
    assert.strictEqual(r.baseSubtype, "effective");
  });

  it('"Base" (plain) remains plain base — no baseSubtype set', () => {
    const r = classifyMetricLabel("Base");
    assert.strictEqual(r.rowType, "base");
    assert.strictEqual(r.baseSubtype, undefined);
  });

  it('"BASE" (all-caps plain) remains plain base — no baseSubtype set', () => {
    const r = classifyMetricLabel("BASE");
    assert.strictEqual(r.rowType, "base");
    assert.strictEqual(r.baseSubtype, undefined);
  });
});

// ─── Split two-column label extraction ───────────────────────────────────────

describe("extractRowLabelFromLeftCells — split two-column base-subtype labels", () => {
  it('["Base", "weighted"] → concatenated label "Base weighted"', () => {
    assert.strictEqual(extractRowLabelFromLeftCells(["Base", "weighted"]), "Base weighted");
  });

  it('["Base", "unweighted"] → concatenated label "Base unweighted"', () => {
    assert.strictEqual(extractRowLabelFromLeftCells(["Base", "unweighted"]), "Base unweighted");
  });

  it('["Base", "effective"] → concatenated label "Base effective"', () => {
    assert.strictEqual(extractRowLabelFromLeftCells(["Base", "effective"]), "Base effective");
  });

  it('single-cell ["Base"] still returns "Base" unchanged', () => {
    assert.strictEqual(extractRowLabelFromLeftCells(["Base"]), "Base");
  });

  it('numeric cell is skipped — ["Base", 100] → "Base"', () => {
    assert.strictEqual(extractRowLabelFromLeftCells(["Base", 100]), "Base");
  });

  it('empty cell is skipped — ["", "Agree"] → "Agree"', () => {
    assert.strictEqual(extractRowLabelFromLeftCells(["", "Agree"]), "Agree");
  });
});

// ─── Split two-column labels — end-to-end block detection ────────────────────

describe("split two-column base-subtype labels — end-to-end block detection", () => {
  function detectSplitBlocks(rowLabels, values) {
    // rowLabels is an array of [col1, col2] or [col1] arrays
    const detectionResult = detectMetricRowsFromLeftLabels(values, rowLabels);
    return buildCalculationBlocks(detectionResult);
  }

  it('[Base][weighted] split label → proportion block with weighted baseSubtype', () => {
    const blocks = detectSplitBlocks(
      [["Agree"], ["Disagree"], ["Base", "weighted"]],
      [[0.4, 0.6], [0.6, 0.4], [100, 200]]
    );
    assert.strictEqual(blocks.length, 1);
    assert.strictEqual(blocks[0].metricType, "proportion");
    assert.strictEqual(blocks[0].baseSubtype, "weighted");
  });

  it('[Base][unweighted] split label → proportion block with unweighted baseSubtype', () => {
    const blocks = detectSplitBlocks(
      [["Agree"], ["Disagree"], ["Base", "unweighted"]],
      [[0.4, 0.6], [0.6, 0.4], [100, 200]]
    );
    assert.strictEqual(blocks.length, 1);
    assert.strictEqual(blocks[0].metricType, "proportion");
    assert.strictEqual(blocks[0].baseSubtype, "unweighted");
  });

  it('[Base][effective] split label → proportion block with effective baseSubtype', () => {
    const blocks = detectSplitBlocks(
      [["Agree"], ["Disagree"], ["Base", "effective"]],
      [[0.4, 0.6], [0.6, 0.4], [100, 200]]
    );
    assert.strictEqual(blocks.length, 1);
    assert.strictEqual(blocks[0].metricType, "proportion");
    assert.strictEqual(blocks[0].baseSubtype, "effective");
  });
});

// ─── Explicit base requirement ────────────────────────────────────────────────

describe("buildCalculationBlocks - explicit base requirement", () => {
  it("does not invent a proportion base from the last selected row", () => {
    const blocks = detectBlocks(
      [
        [0.21, 0.45, 0.33],
        [0.15, 0.28, 0.57],
      ],
      ["Agree", "Disagree"]
    );

    assert.deepStrictEqual(blocks, []);
  });

  it("still detects a proportion block when an explicit base row is present", () => {
    const blocks = detectBlocks(
      [
        [0.21, 0.45, 0.33],
        [0.15, 0.28, 0.57],
        [100, 200, 150],
      ],
      ["Agree", "Disagree", "BASE"]
    );

    assert.deepStrictEqual(blocks, [
      {
        metricType: "proportion",
        valueRowIndexes: [0, 1],
        baseRowIndex: 2,
      },
    ]);
  });

  it("does not invent a base for mean plus variance when no base row is selected", () => {
    const blocks = detectBlocks(
      [
        [29.4, 31.5],
        [103.6, 132.6],
      ],
      ["Mean", "Variance"]
    );

    assert.deepStrictEqual(blocks, []);
  });

  it("does not invent a base for extended NPS when no base row is selected", () => {
    const blocks = detectBlocks(
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

    assert.deepStrictEqual(blocks, []);
  });
});

// ─── preferredBase option ────────────────────────────────────────────────────
//
// Table used across most cases:
//   Agree / Disagree / Weighted Base / Base / Unweighted Base / Effective Base

const ALL_BASE_TYPES_VALUES = [
  [0.4, 0.6], [0.6, 0.4],
  [200, 300], [180, 280], [170, 260], [160, 250],
];
const ALL_BASE_TYPES_LABELS = [
  "Agree", "Disagree",
  "Base weighted", "Base", "Base unweighted", "Base effective",
];

describe("buildCalculationBlocks — preferredBase option", () => {
  it("auto: selects effective base (highest priority) when all types present", () => {
    const blocks = detectBlocksWithPreference(ALL_BASE_TYPES_VALUES, ALL_BASE_TYPES_LABELS, "auto");
    assert.strictEqual(blocks.length, 1);
    assert.strictEqual(blocks[0].baseSubtype, "effective");
    assert.strictEqual(blocks[0].baseRowIndex, 5);
  });

  it("no option (omitted): auto priority unchanged — selects effective base", () => {
    const blocks = detectBlocks(ALL_BASE_TYPES_VALUES, ALL_BASE_TYPES_LABELS);
    assert.strictEqual(blocks.length, 1);
    assert.strictEqual(blocks[0].baseSubtype, "effective");
    assert.strictEqual(blocks[0].baseRowIndex, 5);
  });

  it("prefer effective: selects effective base", () => {
    const blocks = detectBlocksWithPreference(ALL_BASE_TYPES_VALUES, ALL_BASE_TYPES_LABELS, "effective");
    assert.strictEqual(blocks.length, 1);
    assert.strictEqual(blocks[0].baseSubtype, "effective");
    assert.strictEqual(blocks[0].baseRowIndex, 5);
  });

  it("prefer unweighted: selects unweighted base", () => {
    const blocks = detectBlocksWithPreference(ALL_BASE_TYPES_VALUES, ALL_BASE_TYPES_LABELS, "unweighted");
    assert.strictEqual(blocks.length, 1);
    assert.strictEqual(blocks[0].baseSubtype, "unweighted");
    assert.strictEqual(blocks[0].baseRowIndex, 4);
  });

  it("prefer plain: selects plain Base (no baseSubtype)", () => {
    const blocks = detectBlocksWithPreference(ALL_BASE_TYPES_VALUES, ALL_BASE_TYPES_LABELS, "plain");
    assert.strictEqual(blocks.length, 1);
    assert.strictEqual(blocks[0].baseSubtype, undefined);
    assert.strictEqual(blocks[0].baseRowIndex, 3);
  });

  it("prefer weighted: selects weighted base", () => {
    const blocks = detectBlocksWithPreference(ALL_BASE_TYPES_VALUES, ALL_BASE_TYPES_LABELS, "weighted");
    assert.strictEqual(blocks.length, 1);
    assert.strictEqual(blocks[0].baseSubtype, "weighted");
    assert.strictEqual(blocks[0].baseRowIndex, 2);
  });

  it("prefer effective but only weighted and unweighted available: falls back to auto (unweighted)", () => {
    const values = [[0.4, 0.6], [0.6, 0.4], [200, 300], [180, 280]];
    const labels = ["Agree", "Disagree", "Base weighted", "Base unweighted"];
    const blocks = detectBlocksWithPreference(values, labels, "effective");
    assert.strictEqual(blocks.length, 1);
    // Auto priority among weighted and unweighted picks unweighted (priority 1 < 3).
    assert.strictEqual(blocks[0].baseSubtype, "unweighted");
  });

  it("prefer plain but only weighted and effective available: falls back to auto (effective)", () => {
    const values = [[0.4, 0.6], [200, 300], [160, 250]];
    const labels = ["Agree", "Base weighted", "Base effective"];
    const blocks = detectBlocksWithPreference(values, labels, "plain");
    assert.strictEqual(blocks.length, 1);
    assert.strictEqual(blocks[0].baseSubtype, "effective");
  });

  it("prefer weighted but only plain Base available: falls back to auto (plain)", () => {
    const values = [[0.4, 0.6], [0.6, 0.4], [100, 200]];
    const labels = ["Agree", "Disagree", "Base"];
    const blocks = detectBlocksWithPreference(values, labels, "weighted");
    assert.strictEqual(blocks.length, 1);
    assert.strictEqual(blocks[0].baseSubtype, undefined);
  });
});

// ─── Silent above-block Base fallback (issue #310) ───────────────────────────

describe("buildCalculationBlocks — silent above-block Base fallback", () => {
  it("1. proportion block with Base below: behavior unchanged", () => {
    const blocks = detectBlocks(
      [[0.4, 0.6], [0.6, 0.4], [100, 200]],
      ["Agree", "Disagree", "BASE"]
    );
    assert.deepStrictEqual(blocks, [
      { metricType: "proportion", valueRowIndexes: [0, 1], baseRowIndex: 2 },
    ]);
  });

  it("2. proportion block with Base above and none below: uses the above Base", () => {
    const blocks = detectBlocks(
      [[100, 200], [0.4, 0.6], [0.6, 0.4]],
      ["BASE", "Agree", "Disagree"]
    );
    assert.deepStrictEqual(blocks, [
      { metricType: "proportion", valueRowIndexes: [1, 2], baseRowIndex: 0 },
    ]);
  });

  it("3. mean + SD block with Base above and none below: uses the above Base", () => {
    const blocks = detectBlocks(
      [[100, 200], [29.4, 31.5], [5.1, 6.2]],
      ["BASE", "Mean", "SD"]
    );
    assert.deepStrictEqual(blocks, [
      {
        metricType: "mean",
        valueRowIndex: 1,
        spreadRowIndex: 2,
        spreadType: "standardDeviation",
        baseRowIndex: 0,
      },
    ]);
  });

  it("4. mean + Variance block with Base above and none below: uses the above Base", () => {
    const blocks = detectBlocks(
      [[100, 200], [29.4, 31.5], [103.6, 132.6]],
      ["BASE", "Mean", "Variance"]
    );
    assert.deepStrictEqual(blocks, [
      {
        metricType: "mean",
        valueRowIndex: 1,
        spreadRowIndex: 2,
        spreadType: "variance",
        baseRowIndex: 0,
      },
    ]);
  });

  it("5. Base both above and below: the below Base wins", () => {
    const blocks = detectBlocks(
      [[100, 200], [0.4, 0.6], [0.6, 0.4], [110, 210]],
      ["BASE", "Agree", "Disagree", "BASE"]
    );
    assert.deepStrictEqual(blocks, [
      { metricType: "proportion", valueRowIndexes: [1, 2], baseRowIndex: 3 },
    ]);
  });

  it("6. no Base below or above: existing skip behavior remains", () => {
    const blocks = detectBlocks(
      [[0.4, 0.6], [0.6, 0.4]],
      ["Agree", "Disagree"]
    );
    assert.deepStrictEqual(blocks, []);
  });

  it("7. above Base separated by a blank row: not used", () => {
    const blocks = detectBlocks(
      [[100, 200], [null, null], [0.4, 0.6], [0.6, 0.4]],
      ["BASE", "", "Agree", "Disagree"]
    );
    assert.deepStrictEqual(blocks, []);
  });

  it("8. two adjacent tables: lower table does not steal the upper table's Base", () => {
    const blocks = detectBlocks(
      [[0.4, 0.6], [0.6, 0.4], [100, 200], [0.3, 0.7], [0.7, 0.3]],
      ["Agree", "Disagree", "BASE", "Agree", "Disagree"]
    );
    // Only the upper table forms a block (Base below it, consumed). The lower
    // table finds no Base below and must not steal the upper table's Base above.
    assert.deepStrictEqual(blocks, [
      { metricType: "proportion", valueRowIndexes: [0, 1], baseRowIndex: 2 },
    ]);
  });

  it("9. generated footnote row between block and above Base: Base not crossed", () => {
    const footnoteLabel = `${SIGNIFICANCE_FOOTNOTE_MARKER}Уровень значимости: 95%.`;
    const blocks = detectBlocks(
      [[100, 200], [0, 0], [0.4, 0.6], [0.6, 0.4]],
      ["BASE", footnoteLabel, "Agree", "Disagree"]
    );
    assert.deepStrictEqual(blocks, []);
  });

  it("10. generated backlink row between block and above Base: Base not crossed", () => {
    const blocks = detectBlocks(
      [[100, 200], [0, 0], [0.4, 0.6], [0.6, 0.4]],
      ["BASE", BACKLINK_MARKER, "Agree", "Disagree"]
    );
    assert.deepStrictEqual(blocks, []);
  });

  // ── Shared above-Base reuse within one continuous table segment ────────────

  it("shared Base above several proportion rows: one block uses the above Base", () => {
    const blocks = detectBlocks(
      [[100, 200], [0.4, 0.6], [0.6, 0.4], [0.2, 0.3]],
      ["BASE", "Agree", "Disagree", "Other %"]
    );
    assert.deepStrictEqual(blocks, [
      { metricType: "proportion", valueRowIndexes: [1, 2, 3], baseRowIndex: 0 },
    ]);
  });

  it("shared Base above proportions + mean block: both blocks use the same above Base", () => {
    const blocks = detectBlocks(
      [[100, 200], [0.4, 0.6], [0.6, 0.4], [29.4, 31.5], [5.1, 6.2]],
      ["BASE", "Agree", "Disagree", "Mean", "SD"]
    );
    const proportionBlock = blocks.find((b) => b.metricType === "proportion");
    const meanBlock = blocks.find((b) => b.metricType === "mean");
    assert.ok(proportionBlock, "expected a proportion block");
    assert.ok(meanBlock, "expected a mean block");
    assert.deepStrictEqual(proportionBlock.valueRowIndexes, [1, 2]);
    assert.strictEqual(proportionBlock.baseRowIndex, 0);
    assert.strictEqual(meanBlock.valueRowIndex, 3);
    assert.strictEqual(meanBlock.spreadRowIndex, 4);
    assert.strictEqual(meanBlock.spreadType, "standardDeviation");
    assert.strictEqual(meanBlock.baseRowIndex, 0);
  });

  it("shared Base above two mean + SD blocks: both mean blocks use the same above Base", () => {
    const blocks = detectBlocks(
      [[100, 200], [29.4, 31.5], [5.1, 6.2], [42.0, 40.1], [7.3, 6.9]],
      ["BASE", "Mean A", "SD", "Mean B", "SD"]
    );
    const meanBlocks = blocks.filter((b) => b.metricType === "mean");
    assert.strictEqual(meanBlocks.length, 2);
    assert.strictEqual(meanBlocks[0].valueRowIndex, 1);
    assert.strictEqual(meanBlocks[0].baseRowIndex, 0);
    assert.strictEqual(meanBlocks[1].valueRowIndex, 3);
    assert.strictEqual(meanBlocks[1].baseRowIndex, 0);
  });

  it("uses nearest above Base run with priority when several Bases sit above", () => {
    const blocks = detectBlocks(
      [[200, 300], [160, 250], [0.4, 0.6], [0.6, 0.4]],
      ["Base weighted", "Base effective", "Agree", "Disagree"]
    );
    // Both Bases form a consecutive run above the block; auto priority prefers
    // the effective Base over the weighted one.
    assert.strictEqual(blocks.length, 1);
    assert.strictEqual(blocks[0].metricType, "proportion");
    assert.deepStrictEqual(blocks[0].valueRowIndexes, [2, 3]);
    assert.strictEqual(blocks[0].baseRowIndex, 1);
    assert.strictEqual(blocks[0].baseSubtype, "effective");
  });
});
