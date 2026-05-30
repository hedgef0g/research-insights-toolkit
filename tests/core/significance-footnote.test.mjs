import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  SIGNIFICANCE_FOOTNOTE_MARKER,
  isGeneratedSignificanceFootnoteRow,
  collectStatisticTypeLabels,
  buildSignificanceFootnoteVisibleText,
  buildSignificanceFootnoteCellValue,
  buildProcessedRangeFootnoteSuffix,
  appendFootnoteScopeDetail,
  classifyFootnoteScanRow,
  resolveFootnotePlacement,
  resolveFootnoteRemovalRow,
  isFootnoteScanCellBlank,
  FOOTNOTE_SCAN_WINDOW_ROWS,
} from "../../src/core/significance-footnote.js";

const MARKER = SIGNIFICANCE_FOOTNOTE_MARKER;
const footnoteCell = (text = "Уровень значимости: 95%.") => MARKER + text;

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

  it("appendFootnoteScopeDetail leaves base text unchanged for blank suffixes", () => {
    assert.strictEqual(appendFootnoteScopeDetail("base.", undefined), "base.");
    assert.strictEqual(appendFootnoteScopeDetail("base.", ""), "base.");
    assert.strictEqual(appendFootnoteScopeDetail("base.", "   "), "base.");
    assert.strictEqual(appendFootnoteScopeDetail("base.", " X."), "base. X.");
  });
});

describe("footnote placement (idempotent insert/replace)", () => {
  it("classifies scanned rows by content and data-region population", () => {
    // dataColStartOffset = 1: index 0 is a label/margin column, index >= 1 is data.
    assert.strictEqual(classifyFootnoteScanRow([null, "", "  "], 1), "blank");
    assert.strictEqual(classifyFootnoteScanRow([footnoteCell(), "", ""], 1), "marker");
    // Marker is authoritative even when offset into the row (left blank-gap case).
    assert.strictEqual(classifyFootnoteScanRow(["", "", footnoteCell()], 1), "marker");
    // Populated only in the label/margin region → ordinary user note.
    assert.strictEqual(classifyFootnoteScanRow(["Все респонденты", "", ""], 1), "note");
    // Populated in the data region → start of another table.
    assert.strictEqual(classifyFootnoteScanRow(["Возраст", "18-24", "25-34"], 1), "table");
  });

  it("isFootnoteScanCellBlank treats nullish / whitespace as blank", () => {
    assert.strictEqual(isFootnoteScanCellBlank(null), true);
    assert.strictEqual(isFootnoteScanCellBlank(undefined), true);
    assert.strictEqual(isFootnoteScanCellBlank("   "), true);
    assert.strictEqual(isFootnoteScanCellBlank(0), false);
    assert.strictEqual(isFootnoteScanCellBlank("x"), false);
  });

  it("updates in place when a generated footnote is immediately below the table", () => {
    // Regression 1: existing generated footnote is updated, not duplicated.
    const scan = [[footnoteCell(), "", ""], [null, null, null]];
    assert.deepStrictEqual(resolveFootnotePlacement(scan, 10, 1), {
      mode: "update",
      rowIndex: 10,
    });
  });

  it("finds the generated footnote even with a label/data blank gap (left-offset marker)", () => {
    // Regression 2: marker sits in a left column because of a blank gap column;
    // the windowed scan still detects it and updates in place.
    const scan = [["", "", footnoteCell()], [null, null, null]];
    assert.deepStrictEqual(resolveFootnotePlacement(scan, 20, 2), {
      mode: "update",
      rowIndex: 20,
    });
  });

  it("inserts below an ordinary user note row, preserving it", () => {
    // Regression 3: note immediately below → footnote inserted on the next row.
    const scan = [["Все респонденты", "", ""], [null, null, null]];
    assert.deepStrictEqual(resolveFootnotePlacement(scan, 30, 1), {
      mode: "insert",
      rowIndex: 31,
    });
  });

  it("updates a generated footnote that sits below an ordinary user note row", () => {
    // Regression 4: note + existing generated footnote below it → update in place.
    const scan = [["Все респонденты", "", ""], [footnoteCell(), "", ""], [null, null, null]];
    assert.deepStrictEqual(resolveFootnotePlacement(scan, 40, 1), {
      mode: "update",
      rowIndex: 41,
    });
  });

  it("does not attach to the next table below the current one", () => {
    // Regression 7: next table starts right below → insert at the first row below
    // this table (pushing the next table down), never updating the next table.
    const scan = [["Возраст", "18-24", "25-34"], ["BASE", "100", "120"]];
    assert.deepStrictEqual(resolveFootnotePlacement(scan, 50, 1), {
      mode: "insert",
      rowIndex: 50,
    });
  });

  it("stops at a blank separation rather than adopting a distant marker", () => {
    // A marker beyond a blank row belongs to another table; do not adopt it.
    const scan = [[null, null, null], [footnoteCell(), "", ""]];
    assert.deepStrictEqual(resolveFootnotePlacement(scan, 60, 1), {
      mode: "insert",
      rowIndex: 60,
    });
  });

  it("scan window constant is small and bounded", () => {
    assert.ok(FOOTNOTE_SCAN_WINDOW_ROWS >= 2 && FOOTNOTE_SCAN_WINDOW_ROWS <= 6);
  });
});

describe("footnote removal (Clear, marker-based bounded scan)", () => {
  it("removes a generated footnote immediately below the table", () => {
    // 1. marker immediately below → remove that row.
    const scan = [[footnoteCell(), "", ""], [null, null, null]];
    assert.strictEqual(resolveFootnoteRemovalRow(scan, 10, 1), 10);
  });

  it("removes a generated footnote below a user note, preserving the note", () => {
    // 2. note row then marker → remove the marker row (note is only traversed).
    const scan = [["Все респонденты", "", ""], [footnoteCell(), "", ""], [null, null, null]];
    assert.strictEqual(resolveFootnoteRemovalRow(scan, 20, 1), 21);
  });

  it("removes nothing when a user note is followed by a blank row", () => {
    // 3. note then blank → no generated footnote → remove nothing.
    const scan = [["Все респонденты", "", ""], [null, null, null]];
    assert.strictEqual(resolveFootnoteRemovalRow(scan, 30, 1), null);
  });

  it("does not adopt a marker beyond a blank separation", () => {
    // 4. blank then marker → marker belongs to another table → remove nothing.
    const scan = [[null, null, null], [footnoteCell(), "", ""]];
    assert.strictEqual(resolveFootnoteRemovalRow(scan, 40, 1), null);
  });

  it("stops at a next-table boundary before any marker", () => {
    // 5. next-table row before marker → remove nothing.
    const scan = [["Возраст", "18-24", "25-34"], [footnoteCell(), "", ""]];
    assert.strictEqual(resolveFootnoteRemovalRow(scan, 50, 1), null);
  });

  it("finds a left-offset marker (label/data blank gap)", () => {
    // 6. marker offset into the row because of a blank gap column.
    const scan = [["", "", footnoteCell()], [null, null, null]];
    assert.strictEqual(resolveFootnoteRemovalRow(scan, 60, 2), 60);
  });

  it("returns null for empty / non-array input", () => {
    assert.strictEqual(resolveFootnoteRemovalRow([], 70, 1), null);
    assert.strictEqual(resolveFootnoteRemovalRow(null, 70, 1), null);
  });
});
