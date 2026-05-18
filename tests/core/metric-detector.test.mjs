import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  buildCalculationBlocks,
  classifyMetricLabel,
  detectMetricRowsFromLeftLabels,
  extractRowLabelFromLeftCells,
} from "../../src/core/metric-detector.js";

function detectBlocks(values, labels) {
  const leftLabelValues = labels.map((label) => [label]);
  const detectionResult = detectMetricRowsFromLeftLabels(values, leftLabelValues);
  return buildCalculationBlocks(detectionResult);
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
