import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  buildCalculationBlocks,
  detectMetricRowsFromLeftLabels,
} from "../../src/core/metric-detector.js";

function detectBlocks(values, labels) {
  const leftLabelValues = labels.map((label) => [label]);
  const detectionResult = detectMetricRowsFromLeftLabels(values, leftLabelValues);
  return buildCalculationBlocks(detectionResult);
}

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
