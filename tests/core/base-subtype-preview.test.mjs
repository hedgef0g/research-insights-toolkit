import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { buildTablePreviewModel } from "../../src/core/table-preview-model.js";

function makeModel(labels, values, settings = {}) {
  const leftLabelValues = labels.map((label) => [label]);
  return buildTablePreviewModel({ values, leftLabelValues, settings });
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

  it('"BASE weighted" suffix label → weighted fallback warning', () => {
    const model = makeModel(
      ["Agree", "Disagree", "BASE weighted"],
      [[0.4, 0.6], [0.6, 0.4], [100, 200]]
    );
    const block = findBlock(model, "proportion");
    assert.ok(block, "expected a proportion block");
    assert.strictEqual(block.baseSubtype, "weighted");
    assert.ok(
      warningCodes(model).includes("WEIGHTED_BASE_FALLBACK"),
      `expected WEIGHTED_BASE_FALLBACK; got: ${JSON.stringify(warningCodes(model))}`
    );
  });

  it('"Base unweighted" suffix label → unweighted selected, no warning', () => {
    const model = makeModel(
      ["Agree", "Disagree", "Base unweighted"],
      [[0.4, 0.6], [0.6, 0.4], [100, 200]]
    );
    const block = findBlock(model, "proportion");
    assert.ok(block, "expected a proportion block");
    assert.strictEqual(block.baseSubtype, "unweighted");
    assert.ok(
      !warningCodes(model).includes("WEIGHTED_BASE_FALLBACK"),
      `unexpected WEIGHTED_BASE_FALLBACK; got: ${JSON.stringify(warningCodes(model))}`
    );
  });

  it('"Base effective" suffix label → effective selected, no warning', () => {
    const model = makeModel(
      ["Agree", "Disagree", "Base effective"],
      [[0.4, 0.6], [0.6, 0.4], [100, 200]]
    );
    const block = findBlock(model, "proportion");
    assert.ok(block, "expected a proportion block");
    assert.strictEqual(block.baseSubtype, "effective");
    assert.ok(
      !warningCodes(model).includes("WEIGHTED_BASE_FALLBACK"),
      `unexpected WEIGHTED_BASE_FALLBACK; got: ${JSON.stringify(warningCodes(model))}`
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

// ─── Vertically merged / sparse Base labels ───────────────────────────────────

describe("vertically merged / sparse Base labels", () => {
  it("basic merged Base: labeled row detected as base, empty rows below fall outside the block", () => {
    // In Excel a merged Base cell spanning rows 3-5 exposes the label only
    // on the first physical row; subsequent rows arrive as empty strings with
    // no meaningful data.  Current behavior: the labeled row is detected as
    // the base; the empty/null rows below form orphaned empty rows that are
    // not included in any calculation block.
    const model = makeModel(
      ["Agree", "Disagree", "Base", "", ""],
      [
        [0.4, 0.6],
        [0.6, 0.4],
        [100, 200],
        [null, null],
        [null, null],
      ]
    );
    const block = findBlock(model, "proportion");
    assert.ok(block, "proportion block must be detected");
    assert.deepStrictEqual(block.valueRowIndexes, [0, 1]);
    assert.strictEqual(block.baseRowIndex, 2, "base must be the labeled row");
  });

  it("split-data merged Base: base data spread across multiple rows — NOT SUPPORTED (regression documentation)", () => {
    // Edge case: if the data values for a merged Base cell are spread across
    // its physical rows (different columns on different rows), only the data
    // in the labeled row (index 2) is used as the base.  Subsequent rows
    // (indexes 3-4) are misclassified as empty proportion rows and their
    // data is ignored.
    //
    // This shape is NOT supported.  Documented here so any future fix has a
    // baseline.  Do not change these assertions without a dedicated issue.
    const model = makeModel(
      ["Agree", "Disagree", "Base", "", ""],
      [
        [0.4, null, null],
        [null, 0.6, null],
        [100, null, null],
        [null, 200, null],
        [null, null, 150],
      ]
    );
    const block = findBlock(model, "proportion");
    // The preview model may or may not surface a block here (depends on
    // numeric-evidence checks for the base row).  What must NOT happen is
    // that rows 3 or 4 are selected as the base row.
    if (block !== undefined) {
      assert.strictEqual(block.baseRowIndex, 2, "if a block is detected, base must be the labeled row");
    }
  });
});

// ─── checkSelectedBaseValidity ────────────────────────────────────────────────

describe("checkSelectedBaseValidity — selected base row quality issues", () => {
  // Convenience: proportion table with a configurable base row and settings.
  function baseModel(baseRow, settings = {}) {
    return makeModel(
      ["Agree", "Disagree", "BASE"],
      [[0.4, 0.6], [0.6, 0.4], baseRow],
      settings
    );
  }

  // ── Valid base — no new issues ─────────────────────────────────────────────

  it("valid ordinary BASE row → no base-validity issues", () => {
    const model = baseModel([100, 200]);
    const codes = warningCodes(model);
    assert.ok(!codes.includes("BASE_NO_VALID_VALUES"), "unexpected BASE_NO_VALID_VALUES");
    assert.ok(!codes.includes("BASE_BLANK_VALUES"), "unexpected BASE_BLANK_VALUES");
    assert.ok(!codes.includes("BASE_NON_NUMERIC_VALUES"), "unexpected BASE_NON_NUMERIC_VALUES");
    assert.ok(!codes.includes("BASE_NON_POSITIVE_VALUES"), "unexpected BASE_NON_POSITIVE_VALUES");
    assert.ok(!codes.includes("BASE_BELOW_THRESHOLD"), "unexpected BASE_BELOW_THRESHOLD");
  });

  // ── Partial invalids — warning, not critical ───────────────────────────────

  it("one blank base cell → BASE_BLANK_VALUES warning (not critical)", () => {
    const model = baseModel([100, ""]);
    const codes = warningCodes(model);
    assert.ok(codes.includes("BASE_BLANK_VALUES"),
      `expected BASE_BLANK_VALUES; got: ${JSON.stringify(codes)}`);
    assert.ok(!codes.includes("BASE_NO_VALID_VALUES"), "must not be critical while some valid remain");
    const issue = model.dataQualityIssues.find((i) => i.code === "BASE_BLANK_VALUES");
    assert.strictEqual(issue.severity, "warning");
    assert.strictEqual(issue.evidence.blankCount, 1);
  });

  it("null base cell → BASE_BLANK_VALUES warning", () => {
    const model = baseModel([100, null]);
    assert.ok(warningCodes(model).includes("BASE_BLANK_VALUES"),
      "null cell should count as blank");
  });

  it("non-numeric text base cell → BASE_NON_NUMERIC_VALUES warning (not critical)", () => {
    const model = baseModel([100, "n/a"]);
    const codes = warningCodes(model);
    assert.ok(codes.includes("BASE_NON_NUMERIC_VALUES"),
      `expected BASE_NON_NUMERIC_VALUES; got: ${JSON.stringify(codes)}`);
    assert.ok(!codes.includes("BASE_NO_VALID_VALUES"), "must not be critical while some valid remain");
    const issue = model.dataQualityIssues.find((i) => i.code === "BASE_NON_NUMERIC_VALUES");
    assert.strictEqual(issue.severity, "warning");
  });

  it("zero base cell → BASE_NON_POSITIVE_VALUES warning (not critical)", () => {
    const model = baseModel([0, 200]);
    const codes = warningCodes(model);
    assert.ok(codes.includes("BASE_NON_POSITIVE_VALUES"),
      `expected BASE_NON_POSITIVE_VALUES; got: ${JSON.stringify(codes)}`);
    assert.ok(!codes.includes("BASE_NO_VALID_VALUES"), "must not be critical while some valid remain");
    const issue = model.dataQualityIssues.find((i) => i.code === "BASE_NON_POSITIVE_VALUES");
    assert.strictEqual(issue.severity, "warning");
  });

  it("negative base cell → BASE_NON_POSITIVE_VALUES warning (not critical)", () => {
    const model = baseModel([-5, 200]);
    const codes = warningCodes(model);
    assert.ok(codes.includes("BASE_NON_POSITIVE_VALUES"),
      `expected BASE_NON_POSITIVE_VALUES; got: ${JSON.stringify(codes)}`);
    assert.ok(!codes.includes("BASE_NO_VALID_VALUES"), "must not be critical while some valid remain");
  });

  // ── All invalid — critical ─────────────────────────────────────────────────

  it("all zero/negative → BASE_NO_VALID_VALUES critical", () => {
    const model = baseModel([0, -10]);
    const codes = warningCodes(model);
    assert.ok(codes.includes("BASE_NO_VALID_VALUES"),
      `expected BASE_NO_VALID_VALUES; got: ${JSON.stringify(codes)}`);
    const issue = model.dataQualityIssues.find((i) => i.code === "BASE_NO_VALID_VALUES");
    assert.strictEqual(issue.severity, "critical");
    assert.ok(model.qualitySummary.hasBlockingIssues, "qualitySummary must reflect critical");
    assert.ok(!codes.includes("BASE_NON_POSITIVE_VALUES"), "partial code must not also fire");
  });

  it("mix of zero and blank → BASE_NO_VALID_VALUES critical", () => {
    const model = baseModel([0, ""]);
    const codes = warningCodes(model);
    assert.ok(codes.includes("BASE_NO_VALID_VALUES"),
      `expected BASE_NO_VALID_VALUES; got: ${JSON.stringify(codes)}`);
    assert.strictEqual(
      model.dataQualityIssues.find((i) => i.code === "BASE_NO_VALID_VALUES").severity,
      "critical"
    );
  });

  // ── Regression: block filtered but check still fires ──────────────────────
  // blockHasPreviewEvidence drops a block when rowHasNumericEvidence returns
  // false for the base row (all-blank or all-non-numeric strings fail
  // isPreviewNumericCellValue).  checkSelectedBaseValidity must still fire
  // because it runs against rawBlocks, not the filtered calculationBlocks.

  it("all-blank base row → BASE_NO_VALID_VALUES critical even when block is filtered (regression)", () => {
    const model = makeModel(
      ["Agree", "Disagree", "BASE"],
      [[0.4, 0.6], [0.6, 0.4], ["", ""]]
    );
    assert.strictEqual(model.calculationBlocks.length, 0,
      "block must be absent — confirms the regression scenario");
    const codes = warningCodes(model);
    assert.ok(codes.includes("BASE_NO_VALID_VALUES"),
      `expected BASE_NO_VALID_VALUES via rawBlocks path; got: ${JSON.stringify(codes)}`);
    assert.strictEqual(
      model.dataQualityIssues.find((i) => i.code === "BASE_NO_VALID_VALUES").severity,
      "critical"
    );
  });

  it("all-non-numeric base row → BASE_NO_VALID_VALUES critical even when block is filtered (regression)", () => {
    const model = makeModel(
      ["Agree", "Disagree", "BASE"],
      [[0.4, 0.6], [0.6, 0.4], ["n/a", "n/a"]]
    );
    assert.strictEqual(model.calculationBlocks.length, 0,
      "block must be absent — confirms the regression scenario");
    const codes = warningCodes(model);
    assert.ok(codes.includes("BASE_NO_VALID_VALUES"),
      `expected BASE_NO_VALID_VALUES for all-non-numeric base; got: ${JSON.stringify(codes)}`);
  });

  // ── Small-base threshold ──────────────────────────────────────────────────

  it("base below smallBaseThreshold → BASE_BELOW_THRESHOLD warning", () => {
    const model = baseModel([25, 200], { smallBaseThreshold: 30 });
    const codes = warningCodes(model);
    assert.ok(codes.includes("BASE_BELOW_THRESHOLD"),
      `expected BASE_BELOW_THRESHOLD; got: ${JSON.stringify(codes)}`);
    assert.ok(!codes.includes("BASE_NO_VALID_VALUES"),
      "below-threshold values are positive numeric, not critical");
    const issue = model.dataQualityIssues.find((i) => i.code === "BASE_BELOW_THRESHOLD");
    assert.strictEqual(issue.severity, "warning");
    assert.strictEqual(issue.evidence.belowThresholdCount, 1);
    assert.strictEqual(issue.evidence.threshold, 30);
  });

  it("no BASE_BELOW_THRESHOLD when smallBaseThreshold is absent from settings", () => {
    const model = baseModel([25, 200]);
    assert.ok(!warningCodes(model).includes("BASE_BELOW_THRESHOLD"),
      "threshold check must be opt-in via settings");
  });

  // ── Deduplication ─────────────────────────────────────────────────────────

  it("shared base row across multiple blocks is checked only once (no duplicate issues)", () => {
    // NPS structure and proportion blocks often share the same base row.
    const model = makeModel(
      ["NPS", "Promoters", "Detractors", "BASE"],
      [[0.66, 0.56], [0.72, 0.65], [0.06, 0.09], [100, 0]]
    );
    const baseIssues = model.dataQualityIssues.filter(
      (i) => i.code === "BASE_NON_POSITIVE_VALUES" && i.rowIndex === 3
    );
    assert.strictEqual(baseIssues.length, 1, "same base row must not produce duplicate issues");
  });

  // ── Weighted Base fallback co-existence ───────────────────────────────────

  it("Weighted Base fallback warning still fires alongside base-validity issues", () => {
    const model = makeModel(
      ["Agree", "Disagree", "Weighted Base"],
      [[0.4, 0.6], [0.6, 0.4], [100, 0]]
    );
    const codes = warningCodes(model);
    assert.ok(codes.includes("WEIGHTED_BASE_FALLBACK"),
      `expected WEIGHTED_BASE_FALLBACK; got: ${JSON.stringify(codes)}`);
    assert.ok(codes.includes("BASE_NON_POSITIVE_VALUES"),
      `expected BASE_NON_POSITIVE_VALUES alongside WEIGHTED_BASE_FALLBACK; got: ${JSON.stringify(codes)}`);
  });
});
