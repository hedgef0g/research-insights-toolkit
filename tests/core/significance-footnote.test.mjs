import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  SIGNIFICANCE_FOOTNOTE_MARKER,
  isGeneratedSignificanceFootnoteRow,
  collectStatisticTypeLabels,
  buildSignificanceFootnoteVisibleText,
  buildSignificanceFootnoteCellValue,
  buildProcessedRangeFootnoteSuffix,
  buildProcessedBannerGroupsFootnoteSuffix,
  appendFootnoteScopeDetail,
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

  it("base footnote text is unchanged when no scope detail is passed", () => {
    // Test 4 (issue #308): existing basic text must remain byte-for-byte identical.
    const withoutDetail = buildSignificanceFootnoteVisibleText({
      confidenceLevel: "95",
      oneTailedTest: false,
      statisticLabels: ["Z-критерий для долей"],
    });
    assert.strictEqual(
      withoutDetail,
      "Уровень значимости: 95%; тест: двусторонний; статистика: Z-критерий для долей."
    );
    // Undefined / empty / whitespace scope details are no-ops.
    for (const empty of [undefined, "", "   "]) {
      assert.strictEqual(
        buildSignificanceFootnoteVisibleText({
          confidenceLevel: "95",
          oneTailedTest: false,
          statisticLabels: ["Z-критерий для долей"],
          scopeDetail: empty,
        }),
        withoutDetail
      );
    }
  });

  it("marker prefix is unchanged when a scope detail is appended", () => {
    // Test 5 (issue #308): the marker prefix behavior must not change.
    const value = buildSignificanceFootnoteCellValue({
      confidenceLevel: "95",
      oneTailedTest: false,
      statisticLabels: ["Z-критерий для долей"],
      scopeDetail: " Обработано: B12:F34.",
    });
    assert.ok(value.startsWith(SIGNIFICANCE_FOOTNOTE_MARKER));
    assert.ok(isGeneratedSignificanceFootnoteRow(value));
    assert.strictEqual(
      value.slice(SIGNIFICANCE_FOOTNOTE_MARKER.length),
      "Уровень значимости: 95%; тест: двусторонний; статистика: Z-критерий для долей. Обработано: B12:F34."
    );
  });
});

describe("processed-scope footnote suffixes", () => {
  it("appends a processed range suffix, stripping the sheet prefix", () => {
    // Test 1 (issue #308): append processed range.
    assert.strictEqual(buildProcessedRangeFootnoteSuffix("B12:F34"), " Обработано: B12:F34.");
    assert.strictEqual(buildProcessedRangeFootnoteSuffix("Sheet1!B12:F34"), " Обработано: B12:F34.");
    assert.strictEqual(
      buildProcessedRangeFootnoteSuffix("'My Sheet'!B12:F34"),
      " Обработано: B12:F34."
    );
    assert.strictEqual(
      buildSignificanceFootnoteVisibleText({
        confidenceLevel: "95",
        oneTailedTest: false,
        statisticLabels: ["Z-критерий для долей"],
        scopeDetail: buildProcessedRangeFootnoteSuffix("Sheet1!B12:F34"),
      }),
      "Уровень значимости: 95%; тест: двусторонний; статистика: Z-критерий для долей. Обработано: B12:F34."
    );
  });

  it("returns no suffix for an empty / non-string range", () => {
    assert.strictEqual(buildProcessedRangeFootnoteSuffix(""), "");
    assert.strictEqual(buildProcessedRangeFootnoteSuffix("   "), "");
    assert.strictEqual(buildProcessedRangeFootnoteSuffix(null), "");
    assert.strictEqual(buildProcessedRangeFootnoteSuffix(undefined), "");
    assert.strictEqual(buildProcessedRangeFootnoteSuffix(42), "");
  });

  it("appends a processed groups suffix, de-duplicated and ordered", () => {
    // Test 2 (issue #308): append processed groups.
    assert.strictEqual(
      buildProcessedBannerGroupsFootnoteSuffix(["Мужской", "Женский", "Total"]),
      " Обработаны группы: Мужской, Женский, Total."
    );
    assert.strictEqual(
      buildProcessedBannerGroupsFootnoteSuffix(["Мужской", " Мужской ", "Женский", ""]),
      " Обработаны группы: Мужской, Женский."
    );
  });

  it("uses the banner-path wording for hierarchical labels", () => {
    assert.strictEqual(
      buildProcessedBannerGroupsFootnoteSuffix(["Пол / Мужской", "Пол / Женский"]),
      " Обработаны группы баннера: Пол / Мужской; Пол / Женский."
    );
  });

  it("falls back (returns '') when groups are empty or unhelpful", () => {
    // Test 3 (issue #308): fallback when groups are empty/unhelpful.
    assert.strictEqual(buildProcessedBannerGroupsFootnoteSuffix([]), "");
    assert.strictEqual(buildProcessedBannerGroupsFootnoteSuffix(null), "");
    assert.strictEqual(buildProcessedBannerGroupsFootnoteSuffix(["", "  ", null]), "");
    // A single distinct label is not informative enough.
    assert.strictEqual(buildProcessedBannerGroupsFootnoteSuffix(["Total"]), "");
    assert.strictEqual(buildProcessedBannerGroupsFootnoteSuffix(["Total", "Total"]), "");
    // Too many distinct labels → too noisy.
    const many = Array.from({ length: 9 }, (_, i) => `Группа ${i + 1}`);
    assert.strictEqual(buildProcessedBannerGroupsFootnoteSuffix(many), "");
    // Too long even within the count limit → too noisy.
    const long = ["А".repeat(120), "Б".repeat(120)];
    assert.strictEqual(buildProcessedBannerGroupsFootnoteSuffix(long), "");
  });

  it("appendFootnoteScopeDetail leaves base text unchanged for blank suffixes", () => {
    assert.strictEqual(appendFootnoteScopeDetail("base.", undefined), "base.");
    assert.strictEqual(appendFootnoteScopeDetail("base.", ""), "base.");
    assert.strictEqual(appendFootnoteScopeDetail("base.", "   "), "base.");
    assert.strictEqual(appendFootnoteScopeDetail("base.", " X."), "base. X.");
  });
});
