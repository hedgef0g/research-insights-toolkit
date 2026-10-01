import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { SIGNIFICANCE_FOOTNOTE_MARKER } from "../../src/core/significance-footnote.js";
import { resolveClearTargetBodyRange } from "../../src/taskpane/selected-range-interpreter.js";

function toDisplayText(values) {
  return values.map((row) =>
    row.map((cell) => (cell === null || cell === undefined ? "" : String(cell)))
  );
}

function sliceMatrix(matrix, rowOffset, colOffset, rowCount, colCount) {
  return matrix
    .slice(rowOffset, rowOffset + rowCount)
    .map((row) => row.slice(colOffset, colOffset + colCount));
}

function resolveCurrentClearTargetBody(selectedValues, selectedText = toDisplayText(selectedValues)) {
  const target = resolveClearTargetBodyRange({
    values: selectedValues,
    text: selectedText,
  });

  if (target.state === "blocked") {
    return target;
  }

  return {
    ...target,
    targetValues: sliceMatrix(
      target.cleanedValues,
      target.rowOffset,
      target.colOffset,
      target.rowCount,
      target.colCount
    ),
  };
}

describe("current Clear target body resolution", () => {
  it("keeps a strict numeric selected range as the clear target", () => {
    const values = [
      [0.44, 0.41],
      [0.56, 0.59],
      [5605, 1320],
    ];

    const target = resolveCurrentClearTargetBody(values);

    assert.strictEqual(target.state, "passThrough");
    assert.deepStrictEqual(
      {
        rowOffset: target.rowOffset,
        colOffset: target.colOffset,
        rowCount: target.rowCount,
        colCount: target.colCount,
      },
      { rowOffset: 0, colOffset: 0, rowCount: 3, colCount: 2 }
    );
    assert.deepStrictEqual(target.targetValues, [
      ["0.44", "0.41"],
      ["0.56", "0.59"],
      ["5605", "1320"],
    ]);
  });

  it("normalizes a sloppy full-table selection to the data body below banner/header rows", () => {
    const values = [
      ["Ваш пол:", "", ""],
      ["", "Всего", ""],
      ["", "2025Q4", "2026Q1"],
      ["Мужской", 0.44, 0.41],
      ["Женский", 0.56, 0.59],
      ["BASE", 5605, 1320],
    ];

    const target = resolveCurrentClearTargetBody(values);

    assert.strictEqual(target.state, "normalized");
    assert.deepStrictEqual(
      {
        rowOffset: target.rowOffset,
        colOffset: target.colOffset,
        rowCount: target.rowCount,
        colCount: target.colCount,
      },
      { rowOffset: 3, colOffset: 1, rowCount: 3, colCount: 2 }
    );
    assert.deepStrictEqual(target.targetValues, [
      ["0.44", "0.41"],
      ["0.56", "0.59"],
      ["5605", "1320"],
    ]);
  });

  it("excludes adjacent row-label columns from the clear target", () => {
    const values = [
      ["Agree", 0.44, 0.41],
      ["Disagree", 0.56, 0.59],
      ["BASE", 5605, 1320],
    ];

    const target = resolveCurrentClearTargetBody(values);

    assert.strictEqual(target.state, "normalized");
    assert.deepStrictEqual(
      {
        rowOffset: target.rowOffset,
        colOffset: target.colOffset,
        rowCount: target.rowCount,
        colCount: target.colCount,
      },
      { rowOffset: 0, colOffset: 1, rowCount: 3, colCount: 2 }
    );
    assert.deepStrictEqual(target.targetValues, [
      ["0.44", "0.41"],
      ["0.56", "0.59"],
      ["5605", "1320"],
    ]);
  });

  it("excludes a leading empty structural column where the current pass-through clear path does", () => {
    const values = [
      ["", 4.2, 4.4],
      ["", 1.1, 1.2],
      ["", 500, 600],
    ];

    const target = resolveCurrentClearTargetBody(values);

    assert.strictEqual(target.state, "passThrough");
    assert.deepStrictEqual(
      {
        rowOffset: target.rowOffset,
        colOffset: target.colOffset,
        rowCount: target.rowCount,
        colCount: target.colCount,
      },
      { rowOffset: 0, colOffset: 1, rowCount: 3, colCount: 2 }
    );
    assert.deepStrictEqual(target.targetValues, [
      ["4.2", "4.4"],
      ["1.1", "1.2"],
      ["500", "600"],
    ]);
  });

  it("preserves far-left labels mode limitation: clear resolution has no settings branch", () => {
    // In labels-on-left-side mode, labels live outside the selected data range.
    // Current Clear target resolution does not accept calculation settings, so a
    // strict numeric selection remains the target and far-left labels are not
    // part of the body-resolution decision.
    const values = [
      [4.2, 4.4],
      [1.1, 1.2],
      [500, 600],
    ];

    const target = resolveCurrentClearTargetBody(values);

    assert.strictEqual(target.state, "passThrough");
    assert.deepStrictEqual(
      {
        rowOffset: target.rowOffset,
        colOffset: target.colOffset,
        rowCount: target.rowCount,
        colCount: target.colCount,
      },
      { rowOffset: 0, colOffset: 0, rowCount: 3, colCount: 2 }
    );
  });

  it("does not include a generated RIT significance footnote row as data body", () => {
    const values = [
      ["Ваш пол:", "", ""],
      ["", "2025Q4", "2026Q1"],
      ["Мужской", 0.44, 0.41],
      ["Женский", 0.56, 0.59],
      ["BASE", 5605, 1320],
      [`${SIGNIFICANCE_FOOTNOTE_MARKER}Уровень значимости: 95%.`, "", ""],
    ];

    const target = resolveCurrentClearTargetBody(values);

    assert.strictEqual(target.state, "normalized");
    assert.deepStrictEqual(
      {
        rowOffset: target.rowOffset,
        colOffset: target.colOffset,
        rowCount: target.rowCount,
        colCount: target.colCount,
      },
      { rowOffset: 2, colOffset: 1, rowCount: 3, colCount: 2 }
    );
    assert.deepStrictEqual(target.targetValues, [
      ["0.44", "0.41"],
      ["0.56", "0.59"],
      ["5605", "1320"],
    ]);
  });

  it("does not include an ordinary user note below the table as data body", () => {
    const values = [
      ["Ваш пол:", "", ""],
      ["", "2025Q4", "2026Q1"],
      ["Мужской", 0.44, 0.41],
      ["Женский", 0.56, 0.59],
      ["BASE", 5605, 1320],
      ["Все респонденты", "", ""],
    ];

    const target = resolveCurrentClearTargetBody(values);

    assert.strictEqual(target.state, "normalized");
    assert.deepStrictEqual(
      {
        rowOffset: target.rowOffset,
        colOffset: target.colOffset,
        rowCount: target.rowCount,
        colCount: target.colCount,
      },
      { rowOffset: 2, colOffset: 1, rowCount: 3, colCount: 2 }
    );
    assert.deepStrictEqual(target.targetValues, [
      ["0.44", "0.41"],
      ["0.56", "0.59"],
      ["5605", "1320"],
    ]);
  });

  it("recognizes marker-stripped numeric cells as the current clear target body", () => {
    const values = [
      ["Agree", "44% a", "41% b"],
      ["Disagree", "56% b", "59% a"],
      ["BASE", 5605, 1320],
    ];

    const target = resolveCurrentClearTargetBody(values);

    assert.strictEqual(target.state, "normalized");
    assert.deepStrictEqual(
      {
        rowOffset: target.rowOffset,
        colOffset: target.colOffset,
        rowCount: target.rowCount,
        colCount: target.colCount,
      },
      { rowOffset: 0, colOffset: 1, rowCount: 3, colCount: 2 }
    );
    assert.deepStrictEqual(target.targetValues, [
      ["44%", "41%"],
      ["56%", "59%"],
      ["5605", "1320"],
    ]);
  });

  it("blocks a selected range spanning two close tables instead of bleeding into the neighbor", () => {
    const values = [
      ["Таблица 1", "", ""],
      ["", "Всего", "Кат A"],
      ["Вариант A", 0.4, 0.3],
      ["Вариант B", 0.6, 0.7],
      ["BASE", 1000, 500],
      ["Таблица 2", "", ""],
      ["", "Всего", "Кат B"],
      ["Вариант X", 0.5, 0.4],
      ["Вариант Y", 0.5, 0.6],
      ["BASE", 800, 400],
    ];

    const target = resolveCurrentClearTargetBody(values);

    assert.strictEqual(target.state, "blocked");
    assert.ok(
      target.blockingReasons.includes("BODY_APPEARS_MULTI_TABLE") ||
        target.blockingReasons.includes("HEADER_AREA_TOO_LARGE"),
      `expected multi-table blocking reason, got ${JSON.stringify(target.blockingReasons)}`
    );
  });

  it("preserves minimal labeled-table limitation: two-row labels are not stripped in pass-through", () => {
    // Current behavior: this tiny labeled shape is too small to normalize, then
    // marker stripping removes pure-text labels before detectEmbeddedLabelColumns
    // can see them. Clear therefore targets the whole selection. This documents
    // the limitation for the extraction PR rather than changing behavior here.
    const values = [
      ["Agree", 0.5],
      ["BASE", 100],
    ];

    const target = resolveCurrentClearTargetBody(values);

    assert.strictEqual(target.state, "passThrough");
    assert.deepStrictEqual(
      {
        rowOffset: target.rowOffset,
        colOffset: target.colOffset,
        rowCount: target.rowCount,
        colCount: target.colCount,
      },
      { rowOffset: 0, colOffset: 0, rowCount: 2, colCount: 2 }
    );
    assert.deepStrictEqual(target.targetValues, [
      ["", "0.5"],
      ["", "100"],
    ]);
  });
});
