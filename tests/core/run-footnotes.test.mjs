import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { buildSignificanceFootnoteJob } from "../../src/taskpane/run-footnotes.js";
import {
  SIGNIFICANCE_FOOTNOTE_MARKER,
  isGeneratedSignificanceFootnoteRow,
} from "../../src/core/significance-footnote.js";

const baseSettings = {
  addTableFootnote: true,
  labelsOnLeftSide: false,
  confidenceLevel: "95",
  oneTailedTest: false,
};

function buildJob(overrides = {}) {
  return buildSignificanceFootnoteJob({
    sheetName: "Sheet1",
    dataStartRowIndex: 10,
    dataStartColIndex: 4,
    dataRowCount: 3,
    dataColCount: 4,
    leftLabelValues: [["Label"], ["Base"], ["Value"]],
    adjacentLabelColumnCount: 1,
    calculationBlocks: [{ metricType: "proportion" }],
    calculationSettings: baseSettings,
    ...overrides,
  });
}

describe("run footnote job builder", () => {
  it("returns null when the table footnote setting is disabled", () => {
    assert.strictEqual(
      buildJob({
        calculationSettings: {
          ...baseSettings,
          addTableFootnote: false,
        },
      }),
      null
    );
  });

  it("returns null when labels-on-left mode is enabled", () => {
    assert.strictEqual(
      buildJob({
        calculationSettings: {
          ...baseSettings,
          labelsOnLeftSide: true,
        },
      }),
      null
    );
  });

  it("builds the expected job shape for valid data geometry", () => {
    const job = buildJob();

    assert.deepStrictEqual(Object.keys(job), [
      "sheetName",
      "tableBottomRowIndex",
      "tableLeftColIndex",
      "tableRightColIndex",
      "dataStartColIndex",
      "footnoteCellValue",
    ]);
    assert.strictEqual(job.sheetName, "Sheet1");
    assert.strictEqual(job.tableBottomRowIndex, 12);
    assert.strictEqual(job.tableLeftColIndex, 3);
    assert.strictEqual(job.tableRightColIndex, 7);
    assert.strictEqual(job.dataStartColIndex, 4);
    assert.ok(isGeneratedSignificanceFootnoteRow(job.footnoteCellValue));
    assert.strictEqual(
      job.footnoteCellValue.slice(SIGNIFICANCE_FOOTNOTE_MARKER.length),
      "Уровень значимости: 95%; тест: двусторонний; статистика: Z-критерий для долей."
    );
  });

  it("includes the manual processed range suffix when supplied", () => {
    const job = buildJob({
      processedScopeSuffix: " Обработано: B12:F34.",
    });

    assert.strictEqual(
      job.footnoteCellValue.slice(SIGNIFICANCE_FOOTNOTE_MARKER.length),
      "Уровень значимости: 95%; тест: двусторонний; статистика: Z-критерий для долей. Обработано: B12:F34."
    );
  });
});
