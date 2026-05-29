import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  SIGNIFICANCE_FOOTNOTE_MARKER,
  isGeneratedSignificanceFootnoteRow,
  collectStatisticTypeLabels,
  buildSignificanceFootnoteVisibleText,
  buildSignificanceFootnoteCellValue,
} from "../../src/core/significance-footnote.js";

describe("significance footnote helpers", () => {
  it("marker is the fixed invisible prefix \\u2063\\u2063\\u2060\\u2063\\u2060", () => {
    assert.strictEqual(
      [...SIGNIFICANCE_FOOTNOTE_MARKER].map((c) => c.codePointAt(0).toString(16)).join(","),
      "2063,2063,2060,2063,2060"
    );
  });

  it("detects generated footnote rows via startsWith", () => {
    assert.strictEqual(isGeneratedSignificanceFootnoteRow(SIGNIFICANCE_FOOTNOTE_MARKER + "any text"), true);
    assert.strictEqual(isGeneratedSignificanceFootnoteRow("Уровень значимости: 95%"), false);
    assert.strictEqual(isGeneratedSignificanceFootnoteRow(""), false);
    assert.strictEqual(isGeneratedSignificanceFootnoteRow(null), false);
    assert.strictEqual(isGeneratedSignificanceFootnoteRow(undefined), false);
    assert.strictEqual(isGeneratedSignificanceFootnoteRow(42), false);
  });

  it("collects statistic labels from actual blocks, de-duplicated and ordered", () => {
    assert.deepStrictEqual(
      collectStatisticTypeLabels([{ metricType: "proportion" }]),
      ["Z-критерий для долей"]
    );
    assert.deepStrictEqual(
      collectStatisticTypeLabels([
        { metricType: "mean" },
        { metricType: "proportion" },
        { metricType: "proportion" },
      ]),
      ["Z-критерий для долей", "t-тест для средних"]
    );
    assert.deepStrictEqual(collectStatisticTypeLabels([]), []);
    assert.deepStrictEqual(collectStatisticTypeLabels(null), []);
  });

  it("builds two-sided proportion footnote text", () => {
    assert.strictEqual(
      buildSignificanceFootnoteVisibleText({
        confidenceLevel: "95",
        oneTailedTest: false,
        statisticLabels: ["Z-критерий для долей"],
      }),
      "Уровень значимости: 95%; тест: двусторонний; статистика: Z-критерий для долей."
    );
  });

  it("builds mixed-statistic, one-sided footnote text", () => {
    assert.strictEqual(
      buildSignificanceFootnoteVisibleText({
        confidenceLevel: "90",
        oneTailedTest: true,
        statisticLabels: ["Z-критерий для долей", "t-тест для средних"],
      }),
      "Уровень значимости: 90%; тест: односторонний; статистика: Z-критерий для долей, t-тест для средних."
    );
  });

  it("cell value prefixes the visible text with the invisible marker", () => {
    const value = buildSignificanceFootnoteCellValue({
      confidenceLevel: "95",
      oneTailedTest: false,
      statisticLabels: ["Z-критерий для долей"],
    });
    assert.ok(value.startsWith(SIGNIFICANCE_FOOTNOTE_MARKER));
    assert.ok(isGeneratedSignificanceFootnoteRow(value));
    assert.strictEqual(
      value.slice(SIGNIFICANCE_FOOTNOTE_MARKER.length),
      "Уровень значимости: 95%; тест: двусторонний; статистика: Z-критерий для долей."
    );
  });
});
