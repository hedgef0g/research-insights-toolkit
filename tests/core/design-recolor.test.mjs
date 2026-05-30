import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  countAdjacentLabelColumns,
  resolveAdjacentLabelColumnCount,
  buildDesignRecolorJob,
} from "../../src/core/design-recolor.js";

const ON = { recolorBannerAndLabels: true, labelsOnLeftSide: false, bannerLabelFillColor: "#FFF2CC" };

describe("countAdjacentLabelColumns", () => {
  it("returns 0 for empty / non-array input", () => {
    assert.strictEqual(countAdjacentLabelColumns(null), 0);
    assert.strictEqual(countAdjacentLabelColumns([]), 0);
    assert.strictEqual(countAdjacentLabelColumns([[]]), 0);
  });

  it("counts a single adjacent label column", () => {
    const labels = [["Agree"], ["Disagree"], ["Base"]];
    assert.strictEqual(countAdjacentLabelColumns(labels), 1);
  });

  it("counts two adjacent label columns when both have content", () => {
    const labels = [
      ["Top 2 box", "Agree"],
      ["", "Disagree"],
      ["Base", "n"],
    ];
    assert.strictEqual(countAdjacentLabelColumns(labels), 2);
  });

  it("returns 1 when only the right/data-adjacent column has content", () => {
    // Leftmost column is blank (excluded); data-adjacent column carries labels.
    const labels = [
      ["", "Agree"],
      ["", "Disagree"],
      ["", "Base"],
    ];
    assert.strictEqual(countAdjacentLabelColumns(labels), 1);
  });

  it("returns 2 when only the left column has content and the right is blank", () => {
    // Two-column label area with the right (data-adjacent) cell blank — e.g.
    // labels stored only in the left cell. Not a gap: spans both columns.
    const labels = [
      ["Agree", ""],
      ["Disagree", ""],
      ["Base", ""],
    ];
    assert.strictEqual(countAdjacentLabelColumns(labels), 2);
  });

  it("returns 2 for merged-like labels stored only in the left column", () => {
    const labels = [
      ["Mean", ""],
      ["Standard deviation", ""],
      ["Base", ""],
    ];
    assert.strictEqual(countAdjacentLabelColumns(labels), 2);
  });

  it("returns 0 when both columns are blank", () => {
    const labels = [
      ["", ""],
      ["", ""],
    ];
    assert.strictEqual(countAdjacentLabelColumns(labels), 0);
  });

  it("treats numeric 0 and unit indicators as content", () => {
    const labels = [
      ["Mean", 0],
      ["SD", 0],
    ];
    assert.strictEqual(countAdjacentLabelColumns(labels), 2);
  });
});

describe("resolveAdjacentLabelColumnCount", () => {
  it("normalized: uses the data column offset as the label-area width", () => {
    // Mean + SD/Variance + Base with a merged/left-stored two-column label area:
    // normalizer label col (1) + stripped blank data-adjacent label col (1) →
    // effective data column offset 2. Label text lives only in the left column.
    const count = resolveAdjacentLabelColumnCount({
      state: "normalized",
      labelsOnLeftSide: false,
      dataColumnOffset: 2,
      leftLabelValues: [["Mean"], ["Standard deviation"], ["Base"]],
    });
    assert.strictEqual(count, 2);
  });

  it("normalized: Mean + Variance + Base merged labels also resolve to 2", () => {
    const count = resolveAdjacentLabelColumnCount({
      state: "normalized",
      labelsOnLeftSide: false,
      dataColumnOffset: 2,
      leftLabelValues: [["Mean"], ["Variance"], ["Base"]],
    });
    assert.strictEqual(count, 2);
  });

  it("normalized: a true single-column label table resolves to 1", () => {
    const count = resolveAdjacentLabelColumnCount({
      state: "normalized",
      labelsOnLeftSide: false,
      dataColumnOffset: 1,
      leftLabelValues: [["Agree"], ["Disagree"], ["Base"]],
    });
    assert.strictEqual(count, 1);
  });

  it("normalized: external labels (offset 0) fall back to scanning the matrix", () => {
    const count = resolveAdjacentLabelColumnCount({
      state: "normalized",
      labelsOnLeftSide: false,
      dataColumnOffset: 0,
      leftLabelValues: [["", "Agree"], ["", "Disagree"], ["", "Base"]],
    });
    assert.strictEqual(count, 1);
  });

  it("pass-through: scans the adjacent label matrix (right-blank two-column → 2)", () => {
    const count = resolveAdjacentLabelColumnCount({
      state: "passThrough",
      labelsOnLeftSide: false,
      leadingEmptyColumns: 0,
      leftLabelValues: [["Mean", ""], ["SD", ""], ["Base", ""]],
    });
    assert.strictEqual(count, 2);
  });

  it("pass-through: a real blank helper gap (leadingEmptyColumns > 0) resolves to 0", () => {
    const count = resolveAdjacentLabelColumnCount({
      state: "passThrough",
      labelsOnLeftSide: false,
      leadingEmptyColumns: 1,
      leftLabelValues: [["Mean", ""], ["SD", ""], ["Base", ""]],
    });
    assert.strictEqual(count, 0);
  });

  it("labelsOnLeftSide forces 0 in any state", () => {
    assert.strictEqual(
      resolveAdjacentLabelColumnCount({
        state: "normalized",
        labelsOnLeftSide: true,
        dataColumnOffset: 2,
        leftLabelValues: [["Mean"], ["SD"]],
      }),
      0
    );
    assert.strictEqual(
      resolveAdjacentLabelColumnCount({
        state: "passThrough",
        labelsOnLeftSide: true,
        leftLabelValues: [["Mean", "x"]],
      }),
      0
    );
  });

  it("blocked state resolves to 0", () => {
    assert.strictEqual(
      resolveAdjacentLabelColumnCount({ state: "blocked", labelsOnLeftSide: false }),
      0
    );
  });
});

describe("buildDesignRecolorJob", () => {
  const baseGeometry = {
    sheetName: "Sheet1",
    dataStartRowIndex: 5,
    dataStartColIndex: 3,
    dataRowCount: 4,
    dataColCount: 6,
    adjacentLabelColumnCount: 1,
    bannerRowCount: 1,
  };

  it("returns null when the feature is off", () => {
    const job = buildDesignRecolorJob({
      ...baseGeometry,
      calculationSettings: { recolorBannerAndLabels: false, bannerLabelFillColor: "#FFF2CC" },
    });
    assert.strictEqual(job, null);
  });

  it("returns null when labelsOnLeftSide is enabled (forced off)", () => {
    const job = buildDesignRecolorJob({
      ...baseGeometry,
      calculationSettings: { ...ON, labelsOnLeftSide: true },
    });
    assert.strictEqual(job, null);
  });

  it("returns null when the color is missing/empty", () => {
    const job = buildDesignRecolorJob({
      ...baseGeometry,
      calculationSettings: { recolorBannerAndLabels: true, labelsOnLeftSide: false, bannerLabelFillColor: "" },
    });
    assert.strictEqual(job, null);
  });

  it("returns null for unusable geometry", () => {
    assert.strictEqual(
      buildDesignRecolorJob({ ...baseGeometry, dataRowCount: 0, calculationSettings: ON }),
      null
    );
    assert.strictEqual(
      buildDesignRecolorJob({ ...baseGeometry, dataStartRowIndex: NaN, calculationSettings: ON }),
      null
    );
  });

  it("recolors exactly 1 adjacent label column", () => {
    const job = buildDesignRecolorJob({
      ...baseGeometry,
      adjacentLabelColumnCount: 1,
      bannerRowCount: 0,
      calculationSettings: ON,
    });
    assert.deepStrictEqual(job.rects, [
      { rowIndex: 5, columnIndex: 2, rowCount: 4, columnCount: 1 },
    ]);
    assert.strictEqual(job.color, "#FFF2CC");
    assert.strictEqual(job.sheetName, "Sheet1");
  });

  it("recolors exactly 2 adjacent label columns", () => {
    const job = buildDesignRecolorJob({
      ...baseGeometry,
      adjacentLabelColumnCount: 2,
      bannerRowCount: 0,
      calculationSettings: ON,
    });
    assert.deepStrictEqual(job.rects, [
      { rowIndex: 5, columnIndex: 1, rowCount: 4, columnCount: 2 },
    ]);
  });

  it("skips label recolor when there are no adjacent label columns", () => {
    const job = buildDesignRecolorJob({
      ...baseGeometry,
      adjacentLabelColumnCount: 0,
      bannerRowCount: 1,
      calculationSettings: ON,
    });
    // Only the banner rect remains.
    assert.deepStrictEqual(job.rects, [
      { rowIndex: 4, columnIndex: 3, rowCount: 1, columnCount: 6 },
    ]);
  });

  it("produces both banner and label rects (L-shape, corner untouched)", () => {
    const job = buildDesignRecolorJob({
      ...baseGeometry,
      adjacentLabelColumnCount: 2,
      bannerRowCount: 2,
      calculationSettings: ON,
    });
    assert.deepStrictEqual(job.rects, [
      // banner: 2 rows above data body, data columns only
      { rowIndex: 3, columnIndex: 3, rowCount: 2, columnCount: 6 },
      // labels: 2 columns left of data body, data rows only
      { rowIndex: 5, columnIndex: 1, rowCount: 4, columnCount: 2 },
    ]);
  });

  it("skips the banner rect when it would run above row 0", () => {
    const job = buildDesignRecolorJob({
      ...baseGeometry,
      dataStartRowIndex: 0,
      bannerRowCount: 1,
      adjacentLabelColumnCount: 1,
      calculationSettings: ON,
    });
    // No banner (would be row -1); only labels.
    assert.deepStrictEqual(job.rects, [
      { rowIndex: 0, columnIndex: 2, rowCount: 4, columnCount: 1 },
    ]);
  });

  it("caps the label span at the left sheet edge", () => {
    const job = buildDesignRecolorJob({
      ...baseGeometry,
      dataStartColIndex: 1,
      adjacentLabelColumnCount: 2,
      bannerRowCount: 0,
      calculationSettings: ON,
    });
    assert.deepStrictEqual(job.rects, [
      { rowIndex: 5, columnIndex: 0, rowCount: 4, columnCount: 1 },
    ]);
  });

  it("returns null when neither region produces a rect", () => {
    const job = buildDesignRecolorJob({
      ...baseGeometry,
      dataStartRowIndex: 0,
      bannerRowCount: 1,
      adjacentLabelColumnCount: 0,
      calculationSettings: ON,
    });
    assert.strictEqual(job, null);
  });

  it("regression: Mean + SD + Base with labels only in the left of a 2-column label area recolors 2 columns", () => {
    // Label text lives only in the left label column; the right (data-adjacent)
    // label column is blank (merged / left-stored labels). Calculation reads the
    // labels fine, and recolor should cover the full 2-column label area.
    const leftLabelValues = [
      ["Mean", ""],
      ["Standard deviation", ""],
      ["Base", ""],
    ];
    const adjacentLabelColumnCount = countAdjacentLabelColumns(leftLabelValues);
    assert.strictEqual(adjacentLabelColumnCount, 2);

    const job = buildDesignRecolorJob({
      sheetName: "Sheet1",
      dataStartRowIndex: 5,
      dataStartColIndex: 3,
      dataRowCount: 3,
      dataColCount: 6,
      adjacentLabelColumnCount,
      bannerRowCount: 0,
      calculationSettings: ON,
    });
    assert.deepStrictEqual(job.rects, [
      { rowIndex: 5, columnIndex: 1, rowCount: 3, columnCount: 2 },
    ]);
  });

  it("regression: normalized Mean + SD + Base (blank data-adjacent label col) recolors 2 columns", () => {
    // Interpreter path: normalizer detects 1 text label column, the blank
    // data-adjacent label column is stripped from the body (effective data column
    // offset 2). The shared resolver yields 2; recolor covers both label columns.
    const adjacentLabelColumnCount = resolveAdjacentLabelColumnCount({
      state: "normalized",
      labelsOnLeftSide: false,
      dataColumnOffset: 2,
      leftLabelValues: [["Mean"], ["Standard deviation"], ["Base"]],
    });
    assert.strictEqual(adjacentLabelColumnCount, 2);

    const job = buildDesignRecolorJob({
      sheetName: "Sheet1",
      dataStartRowIndex: 4,
      dataStartColIndex: 3, // data body starts at col 3 → label cols 1 and 2
      dataRowCount: 3,
      dataColCount: 5,
      adjacentLabelColumnCount,
      bannerRowCount: 0,
      calculationSettings: ON,
    });
    assert.deepStrictEqual(job.rects, [
      { rowIndex: 4, columnIndex: 1, rowCount: 3, columnCount: 2 },
    ]);
  });
});
