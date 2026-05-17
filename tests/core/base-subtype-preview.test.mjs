import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { buildTablePreviewModel } from "../../src/core/table-preview-model.js";

function makeModel(labels, values) {
  const leftLabelValues = labels.map((label) => [label]);
  return buildTablePreviewModel({ values, leftLabelValues });
}

function findBlock(model, metricType) {
  return model.calculationBlocks.find((b) => b.metricType === metricType);
}

function warningCodes(model) {
  return model.warnings.map((w) => w.code);
}

describe("base subtype selection — preview model metadata", () => {
  it("Weighted Base only → selected as fallback, WEIGHTED_BASE_FALLBACK warning present", () => {
    const model = makeModel(
      ["Agree", "Disagree", "Weighted Base"],
      [
        [0.4, 0.6],
        [0.6, 0.4],
        [100, 200],
      ]
    );
    const block = findBlock(model, "proportion");

    assert.ok(block, "expected a proportion block");
    assert.strictEqual(block.baseSubtype, "weighted", "base subtype must be 'weighted'");
    assert.ok(block.baseSelection, "baseSelection must be populated");
    assert.strictEqual(block.baseSelection.isWeightedFallback, true);
    assert.ok(
      warningCodes(model).includes("WEIGHTED_BASE_FALLBACK"),
      `expected WEIGHTED_BASE_FALLBACK warning; got: ${JSON.stringify(warningCodes(model))}`
    );
  });

  it("Unweighted Base + Weighted Base → Unweighted selected, no WEIGHTED_BASE_FALLBACK", () => {
    const model = makeModel(
      ["Agree", "Disagree", "Unweighted Base", "Weighted Base"],
      [
        [0.4, 0.6],
        [0.6, 0.4],
        [100, 200],
        [120, 210],
      ]
    );
    const block = findBlock(model, "proportion");

    assert.ok(block, "expected a proportion block");
    assert.strictEqual(block.baseSubtype, "unweighted", "Unweighted Base must win over Weighted");
    assert.strictEqual(block.baseSelection.isWeightedFallback, false);
    assert.ok(
      !warningCodes(model).includes("WEIGHTED_BASE_FALLBACK"),
      `unexpected WEIGHTED_BASE_FALLBACK warning; got: ${JSON.stringify(warningCodes(model))}`
    );
  });

  it("Effective Base + Unweighted Base + Weighted Base → Effective selected", () => {
    const model = makeModel(
      ["Agree", "Disagree", "Effective Base", "Unweighted Base", "Weighted Base"],
      [
        [0.4, 0.6],
        [0.6, 0.4],
        [80, 160],
        [100, 200],
        [120, 210],
      ]
    );
    const block = findBlock(model, "proportion");

    assert.ok(block, "expected a proportion block");
    assert.strictEqual(block.baseSubtype, "effective", "Effective Base must have highest priority");
    assert.strictEqual(block.baseSelection.isWeightedFallback, false);
    assert.ok(
      !warningCodes(model).includes("WEIGHTED_BASE_FALLBACK"),
      `unexpected WEIGHTED_BASE_FALLBACK warning; got: ${JSON.stringify(warningCodes(model))}`
    );
  });

  it("plain Base + Weighted Base → plain Base selected, no WEIGHTED_BASE_FALLBACK", () => {
    const model = makeModel(
      ["Agree", "Disagree", "BASE", "Weighted Base"],
      [
        [0.4, 0.6],
        [0.6, 0.4],
        [100, 200],
        [120, 210],
      ]
    );
    const block = findBlock(model, "proportion");

    assert.ok(block, "expected a proportion block");
    assert.strictEqual(block.baseSubtype, null, "plain Base has no subtype");
    assert.strictEqual(block.baseSelection.isWeightedFallback, false);
    assert.ok(
      !warningCodes(model).includes("WEIGHTED_BASE_FALLBACK"),
      `unexpected WEIGHTED_BASE_FALLBACK warning; got: ${JSON.stringify(warningCodes(model))}`
    );
  });

  it("NPS with Effective Base + Weighted Base → Effective selected, no warning", () => {
    // NPS-first format: NPS / Promoters / Detractors / BASE
    const model = makeModel(
      ["NPS", "Promoters", "Detractors", "Effective Base", "Weighted Base"],
      [
        [0.66, 0.56],
        [0.72, 0.65],
        [0.06, 0.09],
        [1000, 800],
        [1200, 900],
      ]
    );
    const npsBlock = findBlock(model, "npsStructure");

    assert.ok(npsBlock, "expected an npsStructure block");
    assert.strictEqual(npsBlock.baseSubtype, "effective", "Effective Base must win");
    assert.strictEqual(npsBlock.baseSelection.isWeightedFallback, false);
    assert.ok(
      !warningCodes(model).includes("WEIGHTED_BASE_FALLBACK"),
      `unexpected WEIGHTED_BASE_FALLBACK warning; got: ${JSON.stringify(warningCodes(model))}`
    );
  });

  it("ordinary BASE table → baseSubtype null, baseSelection populated, no warning", () => {
    const model = makeModel(
      ["Agree", "Disagree", "BASE"],
      [
        [0.4, 0.6],
        [0.6, 0.4],
        [100, 200],
      ]
    );
    const block = findBlock(model, "proportion");

    assert.ok(block, "expected a proportion block");
    assert.strictEqual(block.baseSubtype, null, "plain BASE has no subtype");
    assert.ok(block.baseSelection, "baseSelection must be populated");
    assert.strictEqual(block.baseSelection.selectedBaseRowIndex, 2);
    assert.strictEqual(block.baseSelection.selectedBaseSubtype, null);
    assert.strictEqual(block.baseSelection.selectedBaseLabel, "BASE");
    assert.strictEqual(block.baseSelection.isWeightedFallback, false);
    assert.ok(
      !warningCodes(model).includes("WEIGHTED_BASE_FALLBACK"),
      `unexpected WEIGHTED_BASE_FALLBACK warning; got: ${JSON.stringify(warningCodes(model))}`
    );
    // rowSubtype on row diagnostics is also null for plain base
    assert.strictEqual(model.rowDiagnostics[2].rowSubtype, null);
  });
});
