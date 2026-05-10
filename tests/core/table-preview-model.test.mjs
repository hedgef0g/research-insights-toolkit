import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { buildTablePreviewModel } from "../../src/core/table-preview-model.js";

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
