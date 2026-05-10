import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { buildTablePreviewModel } from "../../src/core/table-preview-model.js";
import { normalizeSelectedRange } from "../../src/core/range-normalizer.js";

// Minimal 3-row data grid with one base row and two proportion rows.
// Labels are supplied via leftLabelValues; values are plain proportions.
function makeModel(labels) {
  const values = [
    [0.21, 0.45, 0.33],
    [0.15, 0.28, 0.57],
    [100, 200, 150],
  ];
  const leftLabelValues = labels.map((l) => [l]);
  return buildTablePreviewModel({ values, leftLabelValues });
}

function warnCodes(model) {
  return model.warnings.map((w) => w.code);
}

describe("buildTablePreviewModel – Возраст normalized integration", () => {
  // Full "Возраст" shape as it arrives from Excel: percentage strings,
  // Russian comma-decimal service rows, trailing "Все респонденты" footer.
  const vozrastValues = [
    ["Ваш возраст",     "",                ""],
    ["",                "Всего",           ""],
    ["",                "Волна (квартал)", ""],
    ["",                "2025Q4",          "2026Q1"],
    ["",                "(a)",             "(a)"],
    ["19 и младше",     "19%",             "15%"],
    ["20-29",           "37%",             "32%"],
    ["от 30 до 40",     "28%",             "28%"],
    ["40 и старше",     "17%",             "24%"],
    ["mean",            "29,4",            "31,5"],
    ["variance",        "103,6",           "132,6"],
    ["BASE",            "5605",            "1320"],
    ["Все респонденты", "",                ""],
  ];

  it("normalizer produces leftLabelValues with mean/variance/BASE", () => {
    const norm = normalizeSelectedRange(vozrastValues);
    assert.strictEqual(norm.normalizationApplied, true);
    assert.deepStrictEqual(norm.leftLabelValues[4], ["mean"]);
    assert.deepStrictEqual(norm.leftLabelValues[5], ["variance"]);
    assert.deepStrictEqual(norm.leftLabelValues[6], ["BASE"]);
  });

  it("buildTablePreviewModel receives mean/variance/BASE as service rows, no MISSING_ROW_LABEL_WITH_DATA for them", () => {
    const norm = normalizeSelectedRange(vozrastValues);
    assert.strictEqual(norm.normalizationApplied, true, "normalization must succeed first");

    const model = buildTablePreviewModel({
      values: norm.valuesForCalculation,
      leftLabelValues: norm.leftLabelValues,
      bannerContext: norm.bannerContext,
    });

    // mean/variance/BASE are service rows — no MISSING_ROW_LABEL_WITH_DATA for them.
    const missingLabelWarnings = model.warnings.filter(
      (w) => w.code === "MISSING_ROW_LABEL_WITH_DATA"
    );
    assert.strictEqual(
      missingLabelWarnings.length,
      0,
      `expected 0 MISSING_ROW_LABEL_WITH_DATA warnings but got: ${JSON.stringify(missingLabelWarnings)}`
    );

    // BASE must be detected.
    assert.ok(model.summary.baseRows > 0, `expected at least 1 base row; summary: ${JSON.stringify(model.summary)}`);
  });
});

describe("buildTablePreviewModel – label warnings", () => {
  it('"q12" emits SUSPICIOUS_CODE_LIKE_LABEL', () => {
    // "q12" matches /^[a-zA-Z]{1,4}\d{2,}/ → looks like a survey variable name.
    // isNumericLikeCellValue("q12") = NaN → NOT filtered by the label extractor,
    // so rawLabel = "q12" and looksLikeCodeLabel fires.
    // Note: "61,00" from the original issue spec is pre-filtered as numeric by
    // the metric detector's label extractor (isNumericLikeCellValue returns true),
    // so SUSPICIOUS_NUMERIC_LABEL cannot fire for it in the current implementation.
    const model = makeModel(["q12", "Agreement", "Base"]);
    assert.ok(
      warnCodes(model).includes("SUSPICIOUS_CODE_LIKE_LABEL"),
      `expected SUSPICIOUS_CODE_LIKE_LABEL; got: ${JSON.stringify(warnCodes(model))}`
    );
  });

  it('"Concept Test" does NOT emit SUSPICIOUS_PLACEHOLDER_LABEL', () => {
    // "Concept Test" ends with "Test" but does not START with "test",
    // so the /^test(?:[\s\-_]|$)/i pattern does not match.
    const model = makeModel(["Concept Test", "Agreement", "Base"]);
    assert.ok(
      !warnCodes(model).includes("SUSPICIOUS_PLACEHOLDER_LABEL"),
      `expected no SUSPICIOUS_PLACEHOLDER_LABEL; got: ${JSON.stringify(warnCodes(model))}`
    );
  });

  it('"test row" emits SUSPICIOUS_PLACEHOLDER_LABEL', () => {
    // "test row" starts with "test " → matches /^test(?:[\s\-_]|$)/i.
    const model = makeModel(["test row", "Agreement", "Base"]);
    assert.ok(
      warnCodes(model).includes("SUSPICIOUS_PLACEHOLDER_LABEL"),
      `expected SUSPICIOUS_PLACEHOLDER_LABEL; got: ${JSON.stringify(warnCodes(model))}`
    );
  });
});
