import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { normalizeSelectedRange, hasEmptyDataRowGap } from "../../src/core/range-normalizer.js";

function makeRunCleanedValues(rawText, bodyRowIndexes) {
  return rawText.map((row, rowIndex) =>
    bodyRowIndexes.includes(rowIndex) ? ["", ...row.slice(1)] : [...row]
  );
}

describe("normalizeSelectedRange", () => {
  it("numeric-only selection passes through unchanged", () => {
    const values = [
      [10, 20, 30],
      [40, 50, 60],
      [70, 80, 90],
    ];
    const result = normalizeSelectedRange(values);
    assert.strictEqual(result.normalizationNeeded, false);
    assert.strictEqual(result.normalizationApplied, false);
    assert.deepStrictEqual(result.blockingReasons, []);
  });

  it("sparse title row + numeric body is normalized", () => {
    // Row 0: single text cell (sparse → title-like). Rows 1-3: all numeric.
    const values = [
      ["Survey Results", null, null, null],
      [10, 20, 30, 40],
      [50, 60, 70, 80],
      [90, 10, 20, 30],
    ];
    const result = normalizeSelectedRange(values);
    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, true);
    assert.deepStrictEqual(result.titleRows, [0]);
    assert.strictEqual(result.dataRowOffset, 1);
    assert.deepStrictEqual(result.labelColumns, []);
    assert.deepStrictEqual(result.blockingReasons, []);
  });

  it("label column + numeric data columns is normalized", () => {
    // Col 0: text labels. Cols 1-3: numeric data.
    const values = [
      ["Alpha", 10, 20, 30],
      ["Beta", 40, 50, 60],
      ["Gamma", 70, 80, 90],
    ];
    const result = normalizeSelectedRange(values);
    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, true);
    assert.deepStrictEqual(result.labelColumns, [0]);
    assert.strictEqual(result.dataColOffset, 1);
    assert.deepStrictEqual(result.titleRows, []);
    assert.deepStrictEqual(result.blockingReasons, []);
  });

  it("sparse merged-like banner rows with title are normalized correctly", () => {
    // Simulates a full-table selection where Excel merged cells produce sparse
    // banner rows: text appears only in one column per row because Office.js
    // returns the merge value only in the top-left cell.
    //
    // Row 0:   sparse title ("Ваш возраст" in col 0, rest empty)
    // Rows 1-3: sparse banner (col 0 is empty; text only in col 1 or cols 1-2)
    // Rows 4-11: body with label column (col 0) + numeric data (cols 1-2)
    const values = [
      ["Ваш возраст", "", ""],
      ["", "Всего", ""],
      ["", "Волна", ""],
      ["", "2025Q4", "2026Q1"],
      ["19 и младше", 0.19, 0.15],
      ["20-29", 0.37, 0.32],
      ["от 30 до 40", 0.28, 0.28],
      ["40 и старше", 0.17, 0.24],
      ["mean", 29.4, 31.5],
      ["variance", 103.6, 132.6],
      ["BASE", 5605, 1320],
      ["Все респонденты", 6925, 2640],
    ];
    const result = normalizeSelectedRange(values);

    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, true);
    assert.deepStrictEqual(result.titleRows, [0], "title row should be row 0");
    assert.deepStrictEqual(
      result.bannerRows,
      [1, 2, 3],
      "sparse banner rows 1-3 should be detected"
    );
    assert.deepStrictEqual(result.labelColumns, [0], "col 0 should be identified as label column");
    assert.strictEqual(result.dataRowOffset, 4, "body should start at row 4");
    assert.strictEqual(result.dataColOffset, 1, "data columns should start at col 1");
    assert.deepStrictEqual(
      result.valuesForCalculation[0],
      [0.19, 0.15],
      "first valuesForCalculation row should be the first body data row"
    );
    assert.deepStrictEqual(result.blockingReasons, []);
  });

  it("full table with 2025Q4-style banner labels and percent strings normalizes correctly", () => {
    // Simulates the full "Ваш пол" shape as it actually arrives from Excel:
    //   - row 0: sparse title
    //   - rows 1-4: sparse banner rows whose labels include "2025Q4", "(a)" etc.
    //     These must be classified as TEXT (not numeric) so they are treated as
    //     banner rows, not body rows.
    //   - rows 5-7: data rows where values are percentage strings ("44%") or
    //     numbers (5605).  "44%" must be classified as NUMERIC.
    //   - row 8: trailing footer with empty data columns → trimmed away.
    const values = [
      ["Ваш пол:", "", "", ""],
      ["", "Всего", "", "Пользование категорией"],
      ["", "Волна (квартал)", "Всё покупаю сам(а)", "Большую часть"],
      ["", "2025Q4", "2026Q1", "2025Q4"],
      ["", "(a)", "(a)", "(a)"],
      ["Мужской", "44%", "41%", "39%"],
      ["Женский", "56%", "59%", "61%"],
      ["BASE", 5605, 1320, 3083],
      ["Все респонденты", "", "", ""],
    ];
    const result = normalizeSelectedRange(values);

    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, true, "must normalize, not block");
    assert.deepStrictEqual(result.titleRows, [0], "row 0 is the sparse title");
    assert.deepStrictEqual(result.bannerRows, [1, 2, 3, 4], "rows 1-4 are banner rows");
    assert.deepStrictEqual(result.labelColumns, [0], "col 0 is the label column");
    assert.strictEqual(result.dataRowOffset, 5, "body starts at row 5");
    assert.strictEqual(result.valuesForCalculation.length, 3, "footer row excluded");
    assert.deepStrictEqual(
      result.valuesForCalculation[0],
      ["44%", "41%", "39%"],
      "first data row is Мужской with percent strings"
    );
    assert.deepStrictEqual(result.blockingReasons, []);
  });

  it("real sparse 4-row banner + proportion body normalizes from Run cleaned values", () => {
    const rawText = [
      ["Ваш пол:", "", "", "", "", "", "", "", "", "", "", "", ""],
      ["", "", "", "Пользование кат", "", "", "", "", "", "", "", "", ""],
      [
        "",
        "Волна (квартал)",
        "",
        "Всё покупаю сам(а)",
        "",
        "Большую часть",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
      ],
      [
        "",
        "2025Q4",
        "2026Q1",
        "Волна (квартал)",
        "",
        "Волна (квартал)",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
      ],
      [
        "",
        "",
        "",
        "2025Q4",
        "2026Q1",
        "2025Q4",
        "2026Q1",
        "2025Q4",
        "2026Q1",
        "2025Q4",
        "2026Q1",
        "2025Q4",
        "2026Q1",
      ],
      [
        "Мужской",
        "0.5",
        "0.4136",
        "0.39",
        "0.41",
        "0.44",
        "0.42",
        "0.48",
        "0.43",
        "0.51",
        "0.49",
        "0.46",
        "0.45",
      ],
      [
        "Женский",
        "0.5",
        "0.5863",
        "0.61",
        "0.59",
        "0.56",
        "0.58",
        "0.52",
        "0.57",
        "0.49",
        "0.51",
        "0.54",
        "0.55",
      ],
      [
        "BASE",
        "5605",
        "1320",
        "3083",
        "1045",
        "2200",
        "900",
        "1800",
        "760",
        "1500",
        "700",
        "1300",
        "620",
      ],
    ];
    const cleanedValues = makeRunCleanedValues(rawText, [5, 6, 7]);

    const result = normalizeSelectedRange(cleanedValues, rawText);

    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, true);
    assert.deepStrictEqual(result.titleRows, [0]);
    assert.deepStrictEqual(result.bannerRows, [1, 2, 3, 4]);
    assert.deepStrictEqual(result.labelColumns, [0]);
    assert.strictEqual(result.dataRowOffset, 5);
    assert.strictEqual(result.dataColOffset, 1);
    assert.strictEqual(result.valuesForCalculation.length, 3);
    assert.strictEqual(result.valuesForCalculation[0].length, 12);
    assert.deepStrictEqual(result.valuesForCalculation[0].slice(0, 3), ["0.5", "0.4136", "0.39"]);
    assert.deepStrictEqual(result.leftLabelValues, [["Мужской"], ["Женский"], ["BASE"]]);
    assert.strictEqual(result.bannerContext.scanRows.length, 4);
    assert.strictEqual(result.bannerContext.columnCount, 12);
    assert.deepStrictEqual(result.bannerContext.scanRows[0].slice(0, 4), [
      "",
      "",
      "Пользование кат",
      "",
    ]);
  });

  it("real sparse 4-row banner + marked proportion body normalizes on repeat Run", () => {
    const rawText = [
      ["NPS-free proportion table", "", "", "", "", "", "", "", "", "", "", "", ""],
      ["", "", "", "Category usage", "", "", "", "", "", "", "", "", ""],
      ["", "Wave (quarter)", "", "Buys all myself", "", "Most of it", "", "", "", "", "", "", ""],
      ["", "2025Q4", "2026Q1", "Wave (quarter)", "", "Wave (quarter)", "", "", "", "", "", "", ""],
      [
        "",
        "",
        "",
        "2025Q4",
        "2026Q1",
        "2025Q4",
        "2026Q1",
        "2025Q4",
        "2026Q1",
        "2025Q4",
        "2026Q1",
        "2025Q4",
        "2026Q1",
      ],
      [
        "Male",
        "50% b",
        "41% a",
        "39% a b",
        "41%",
        "44% b",
        "42% a",
        "48% b",
        "43% a",
        "51% b",
        "49% a",
        "46% b",
        "45% a",
      ],
      [
        "Female",
        "50% a",
        "59% b",
        "61% b",
        "59% b",
        "56% a",
        "58% b",
        "52% a",
        "57% b",
        "49% a",
        "51% b",
        "54% a",
        "55% b",
      ],
      [
        "BASE",
        "5605",
        "1320",
        "3083",
        "1045",
        "2200",
        "900",
        "1800",
        "760",
        "1500",
        "700",
        "1300",
        "620",
      ],
    ];
    const cleanedValues = [
      ...rawText.slice(0, 5).map((row) => [...row]),
      ["", "50%", "41%", "39%", "41%", "44%", "42%", "48%", "43%", "51%", "49%", "46%", "45%"],
      ["", "50%", "59%", "61%", "59%", "56%", "58%", "52%", "57%", "49%", "51%", "54%", "55%"],
      [
        "",
        "5605",
        "1320",
        "3083",
        "1045",
        "2200",
        "900",
        "1800",
        "760",
        "1500",
        "700",
        "1300",
        "620",
      ],
    ];

    const result = normalizeSelectedRange(cleanedValues, rawText);

    assert.strictEqual(result.normalizationApplied, true);
    assert.deepStrictEqual(result.labelColumns, [0]);
    assert.strictEqual(result.dataRowOffset, 5);
    assert.strictEqual(result.dataColOffset, 1);
    assert.deepStrictEqual(result.valuesForCalculation[0].slice(0, 3), ["50%", "41%", "39%"]);
    assert.notStrictEqual(result.valuesForCalculation[0][0], "Male");
    assert.deepStrictEqual(result.leftLabelValues, [["Male"], ["Female"], ["BASE"]]);
    assert.strictEqual(result.bannerContext.scanRows.length, 4);
    assert.deepStrictEqual(result.bannerContext.scanRows[0].slice(0, 4), [
      "",
      "",
      "Category usage",
      "",
    ]);
    assert.deepStrictEqual(result.bannerContext.scanRows[3].slice(0, 4), [
      "",
      "",
      "2025Q4",
      "2026Q1",
    ]);
  });

  it("real sparse 4-row banner + means body preserves service labels from Run text", () => {
    const rawText = [
      ["Ваш возраст", "", "", "", "", "", "", "", "", "", "", "", ""],
      ["", "", "", "Пользование кат", "", "", "", "", "", "", "", "", ""],
      [
        "",
        "Волна (квартал)",
        "",
        "Всё покупаю сам(а)",
        "",
        "Большую часть",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
      ],
      [
        "",
        "2025Q4",
        "2026Q1",
        "Волна (квартал)",
        "",
        "Волна (квартал)",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
      ],
      [
        "",
        "",
        "",
        "2025Q4",
        "2026Q1",
        "2025Q4",
        "2026Q1",
        "2025Q4",
        "2026Q1",
        "2025Q4",
        "2026Q1",
        "2025Q4",
        "2026Q1",
      ],
      [
        "Среднее",
        "29.4",
        "31.5",
        "30.1",
        "32.2",
        "28.9",
        "30.7",
        "33.0",
        "34.1",
        "27.5",
        "28.0",
        "35.2",
        "36.4",
      ],
      [
        "variance",
        "103.6",
        "132.6",
        "120.1",
        "140.4",
        "99.2",
        "118.8",
        "150.0",
        "160.1",
        "90.5",
        "95.2",
        "170.4",
        "180.3",
      ],
      [
        "BASE",
        "5605",
        "1320",
        "3083",
        "1045",
        "2200",
        "900",
        "1800",
        "760",
        "1500",
        "700",
        "1300",
        "620",
      ],
    ];
    const cleanedValues = makeRunCleanedValues(rawText, [5, 6, 7]);

    const result = normalizeSelectedRange(cleanedValues, rawText);

    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, true);
    assert.deepStrictEqual(result.titleRows, [0]);
    assert.deepStrictEqual(result.bannerRows, [1, 2, 3, 4]);
    assert.deepStrictEqual(result.labelColumns, [0]);
    assert.strictEqual(result.dataRowOffset, 5);
    assert.strictEqual(result.dataColOffset, 1);
    assert.strictEqual(result.valuesForCalculation.length, 3);
    assert.deepStrictEqual(result.leftLabelValues, [["Среднее"], ["variance"], ["BASE"]]);
    assert.strictEqual(result.bannerContext.scanRows.length, 4);
  });

  it("real-like multi-table broad selection with Run text still blocks", () => {
    const rawText = [
      ["Таблица 1", "", "", ""],
      ["", "Всего", "Кат A", ""],
      ["", "2025Q4", "2025Q4", ""],
      ["Вариант A", "0.4", "0.3", ""],
      ["Вариант B", "0.6", "0.7", ""],
      ["BASE", "1000", "500", ""],
      ["Таблица 2", "", "", ""],
      ["", "Всего", "Кат B", ""],
      ["", "2026Q1", "2026Q1", ""],
      ["Вариант X", "0.5", "0.4", ""],
      ["Вариант Y", "0.5", "0.6", ""],
      ["BASE", "800", "400", ""],
    ];
    const cleanedValues = makeRunCleanedValues(rawText, [3, 4, 5, 9, 10, 11]);

    const result = normalizeSelectedRange(cleanedValues, rawText);

    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, false);
    assert.ok(
      result.blockingReasons.includes("BODY_APPEARS_MULTI_TABLE") ||
        result.blockingReasons.includes("HEADER_AREA_TOO_LARGE"),
      `expected multi-table blocking reason, got: ${JSON.stringify(result.blockingReasons)}`
    );
  });

  it("two-table broad selection blocks with BODY_APPEARS_MULTI_TABLE", () => {
    // Two complete research tables stacked vertically.  The second table's sparse
    // title row (content only in col 0, data cols empty) falls inside what the
    // normalizer computes as the body, creating an all-empty-data-col gap that
    // must trigger BODY_APPEARS_MULTI_TABLE.
    const values = [
      // Table 1
      ["Таблица 1", "", "", ""],
      ["", "Всего", "Кат A", ""],
      ["", "2025Q4", "2025Q4", ""],
      ["Вариант A", 0.4, 0.3, ""],
      ["Вариант B", 0.6, 0.7, ""],
      ["BASE", 1000, 500, ""],
      // Table 2 — sparse title provides all-empty data cols → gap
      ["Таблица 2", "", "", ""],
      ["", "Всего", "Кат B", ""],
      ["", "2026Q1", "2026Q1", ""],
      ["Вариант X", 0.5, 0.4, ""],
      ["Вариант Y", 0.5, 0.6, ""],
      ["BASE", 800, 400, ""],
    ];
    const result = normalizeSelectedRange(values);

    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, false);
    assert.ok(
      result.blockingReasons.includes("BODY_APPEARS_MULTI_TABLE") ||
        result.blockingReasons.includes("HEADER_AREA_TOO_LARGE"),
      `expected BODY_APPEARS_MULTI_TABLE or HEADER_AREA_TOO_LARGE, got: ${JSON.stringify(result.blockingReasons)}`
    );
  });

  it("leftLabelValues uses rawText when values have significance markers stripped", () => {
    // Root-cause regression test.
    //
    // taskpane.js calls removeSignificanceMarkersFromMatrix(selectedRange.values)
    // before passing to normalizeSelectedRange. The marker-strip regex is
    // /\s*[allLetters]+$/ — for cells whose entire content is letters, the whole
    // value is erased: "mean" → "", "variance" → "", "BASE" → "".
    //
    // The fix: buildNormalizedModel slices leftLabelValues from rawText
    // (Office.js selectedRange.text, never marker-stripped) rather than from the
    // stripped values grid.
    //
    // This test simulates the stripping by passing "" in the label column for the
    // three service rows while passing the original values as rawText.
    const rawText = [
      ["Ваш возраст", "", ""],
      ["", "Всего", ""],
      ["", "Волна", ""],
      ["", "2025Q4", "2026Q1"],
      ["", "(a)", "(a)"],
      ["19 и младше", "19%", "15%"],
      ["20-29", "37%", "32%"],
      ["от 30 до 40", "28%", "28%"],
      ["40 и старше", "17%", "24%"],
      ["mean", "29,4", "31,5"],
      ["variance", "103,6", "132,6"],
      ["BASE", "5605", "1320"],
      ["Все респонденты", "", ""],
    ];
    // Simulate significance marker removal: pure-letter label cells → "".
    const strippedValues = rawText.map((row, r) =>
      [9, 10, 11].includes(r)
        ? ["", row[1], row[2]] // "mean"/"variance"/"BASE" erased
        : row
    );

    const result = normalizeSelectedRange(strippedValues, rawText);

    assert.strictEqual(result.normalizationApplied, true);
    // Labels must come from rawText, not from strippedValues.
    assert.deepStrictEqual(result.leftLabelValues[4], ["mean"], "mean label preserved from text");
    assert.deepStrictEqual(
      result.leftLabelValues[5],
      ["variance"],
      "variance label preserved from text"
    );
    assert.deepStrictEqual(result.leftLabelValues[6], ["BASE"], "BASE label preserved from text");
  });

  it("Возраст full-table shape with service rows normalizes correctly", () => {
    // Exact shape from Excel smoke test, values as strings (Russian comma-decimal,
    // percent strings, quarter labels). Verifies leftLabelValues contains the
    // service-row labels ("mean", "variance", "BASE") so buildTablePreviewModel
    // can classify them correctly.
    const values = [
      ["Ваш возраст", "", ""],
      ["", "Всего", ""],
      ["", "Волна (квартал)", ""],
      ["", "2025Q4", "2026Q1"],
      ["", "(a)", "(a)"],
      ["19 и младше", "19%", "15%"],
      ["20-29", "37%", "32%"],
      ["от 30 до 40", "28%", "28%"],
      ["40 и старше", "17%", "24%"],
      ["mean", "29,4", "31,5"],
      ["variance", "103,6", "132,6"],
      ["BASE", "5605", "1320"],
      ["Все респонденты", "", ""],
    ];
    const result = normalizeSelectedRange(values);

    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, true, "must normalize successfully");
    assert.deepStrictEqual(result.titleRows, [0]);
    assert.deepStrictEqual(result.bannerRows, [1, 2, 3, 4]);
    assert.deepStrictEqual(result.labelColumns, [0]);
    assert.strictEqual(result.dataRowOffset, 5);
    assert.strictEqual(result.valuesForCalculation.length, 7, "trailing footer excluded");
    assert.strictEqual(result.leftLabelValues.length, 7, "leftLabelValues must have 7 rows");
    assert.deepStrictEqual(result.leftLabelValues[0], ["19 и младше"], "first body label");
    assert.deepStrictEqual(result.leftLabelValues[4], ["mean"], "mean row label");
    assert.deepStrictEqual(result.leftLabelValues[5], ["variance"], "variance row label");
    assert.deepStrictEqual(result.leftLabelValues[6], ["BASE"], "BASE row label");
    assert.deepStrictEqual(result.blockingReasons, []);
  });

  it("extended NPS full-table shape treats numeric scale rows as left labels", () => {
    const rawText = [
      ["NPS full table", "", ""],
      ["", "Total", "Segment A"],
      ["", "2025Q4", "2026Q1"],
      ["1", "1%", "2%"],
      ["2", "2%", "3%"],
      ["3", "3%", "4%"],
      ["4", "4%", "5%"],
      ["5", "5%", "6%"],
      ["6", "6%", "7%"],
      ["7", "7%", "8%"],
      ["8", "8%", "9%"],
      ["9", "9%", "10%"],
      ["10", "10%", "11%"],
      ["Bottom-3", "6%", "9%"],
      ["Mid-4", "22%", "26%"],
      ["Top-3", "72%", "65%"],
      ["Detractors", "6%", "9%"],
      ["Neutral", "22%", "26%"],
      ["Promoters", "72%", "65%"],
      ["NPS", "66%", "56%"],
      ["BASE", "1000", "800"],
    ];

    const cleanedValues = rawText.map((row, rowIndex) =>
      rowIndex >= 3 ? ["", ...row.slice(1)] : [...row]
    );

    const result = normalizeSelectedRange(cleanedValues, rawText);

    assert.strictEqual(result.normalizationApplied, true);
    assert.deepStrictEqual(result.labelColumns, [0]);
    assert.strictEqual(result.dataRowOffset, 3);
    assert.strictEqual(result.dataColOffset, 1);
    assert.strictEqual(result.valuesForCalculation[0][0], "1%");
    assert.notStrictEqual(result.valuesForCalculation[0][0], "1");
    assert.deepStrictEqual(result.leftLabelValues[13], ["Detractors"]);
    assert.deepStrictEqual(result.leftLabelValues[14], ["Neutral"]);
    assert.deepStrictEqual(result.leftLabelValues[15], ["Promoters"]);
    assert.deepStrictEqual(result.leftLabelValues[16], ["NPS"]);
    assert.deepStrictEqual(result.leftLabelValues[17], ["BASE"]);
  });

  it("trailing label-only footer row is excluded from body, normalization succeeds", () => {
    // Simulates "Ваш пол" shape: sparse banner rows at top, two data rows,
    // then "Все респонденты" which has a label in col[0] but empty data columns.
    // The footer row must be trimmed so it does not trigger BODY_APPEARS_MULTI_TABLE.
    const values = [
      ["Ваш пол:", "", ""],
      ["", "Всего", ""],
      ["", "2025Q4", "2026Q1"],
      ["Мужской", 0.44, 0.41],
      ["Женский", 0.56, 0.59],
      ["BASE", 5605, 1320],
      ["Все респонденты", "", ""],
    ];
    const result = normalizeSelectedRange(values);
    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, true);
    assert.deepStrictEqual(
      result.blockingReasons,
      [],
      "footer row must not trigger BODY_APPEARS_MULTI_TABLE"
    );
    assert.strictEqual(
      result.valuesForCalculation.length,
      3,
      "footer row must be excluded from valuesForCalculation"
    );
    assert.deepStrictEqual(result.valuesForCalculation[0], [0.44, 0.41], "first data row");
    assert.deepStrictEqual(result.valuesForCalculation[2], [5605, 1320], "BASE row included");
  });

  it("multi-table selection with mid-body empty row blocks normalization with BODY_APPEARS_MULTI_TABLE", () => {
    // Two groups of data rows separated by a fully empty row.
    // The empty row lies between actual numeric rows so bodyDataEndRow does NOT
    // trim it away — the gap check must fire and block normalization.
    const values = [
      ["Группа", "A", "B", "C"],
      ["Label1", 10, 20, 30],
      ["Label2", 40, 50, 60],
      [null, null, null, null],
      ["Label3", 70, 80, 90],
      ["Label4", 10, 20, 30],
    ];
    const result = normalizeSelectedRange(values);
    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, false);
    assert.ok(
      result.blockingReasons.includes("BODY_APPEARS_MULTI_TABLE"),
      `expected BODY_APPEARS_MULTI_TABLE in blockingReasons, got: ${JSON.stringify(result.blockingReasons)}`
    );
  });

  it("empty row inside body blocks normalization with BODY_APPEARS_MULTI_TABLE", () => {
    // Row 0: wide banner header. Rows 1-2 and 4-5: data. Row 3: empty gap.
    // The gap triggers the multi-table blocking reason.
    const values = [
      ["Group", "A", "B", "C"],
      ["Label1", 10, 20, 30],
      ["Label2", 40, 50, 60],
      [null, null, null, null],
      ["Label3", 70, 80, 90],
      ["Label4", 10, 20, 30],
    ];
    const result = normalizeSelectedRange(values);
    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, false);
    assert.ok(
      result.blockingReasons.includes("BODY_APPEARS_MULTI_TABLE"),
      `expected BODY_APPEARS_MULTI_TABLE in blockingReasons, got: ${JSON.stringify(result.blockingReasons)}`
    );
  });

  // ── Issue #118: embedded % unit column handling ───────────────────────────────

  it("[label | 0%-unit-col | data] strips both label and unit column (standard table)", () => {
    // Simulates a selection where the '%' unit column stores value=0 with a
    // standard "0%" number format so Excel's .text is "0%".
    // isNumericCell("0%") = true, so without the fix the normalizer would only
    // strip col[0] (the label) and leave col[1] ("0%") as the first data column.
    const rawText = [
      ["Metric A", "0%", "21%", "35%", "42%"],
      ["Metric B", "0%", "15%", "22%", "33%"],
      ["Base",     "0%", "100", "200", "150"],
    ];
    // cleanedValues mirrors how Run strips markers: body rows have col[0] zeroed
    // out because the label cell was a pure-text label stripped by marker removal.
    // The "0%" column stores numeric 0 in .values — after removeSignificanceMarkersFromMatrix
    // numeric 0 becomes string "0", so we simulate that here.
    const cleanedValues = [
      ["Metric A", "0", "0.21", "0.35", "0.42"],
      ["Metric B", "0", "0.15", "0.22", "0.33"],
      ["Base",     "0", "100",  "200",  "150"],
    ];

    const result = normalizeSelectedRange(cleanedValues, rawText);

    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, true, "should normalize, not block");
    assert.deepStrictEqual(result.labelColumns, [0, 1], "both label and unit columns should be detected");
    assert.strictEqual(result.dataColOffset, 2, "data should start at col 2");
    assert.deepStrictEqual(
      result.valuesForCalculation[0],
      ["0.21", "0.35", "0.42"],
      "valuesForCalculation must not include the unit column"
    );
    assert.deepStrictEqual(result.blockingReasons, []);
  });

  it("[label | 0%-unit-col | data] leftLabelValues includes both label columns, valuesForCalculation starts at col 2", () => {
    const rawText = [
      ["Alpha", "0%", "44%", "41%"],
      ["Beta",  "0%", "56%", "59%"],
      ["BASE",  "0%", "500", "600"],
    ];
    const cleanedValues = [
      ["Alpha", "0", "0.44", "0.41"],
      ["Beta",  "0", "0.56", "0.59"],
      ["BASE",  "0", "500",  "600"],
    ];

    const result = normalizeSelectedRange(cleanedValues, rawText);

    assert.strictEqual(result.normalizationApplied, true);
    assert.strictEqual(result.dataColOffset, 2);
    // leftLabelValues covers both detected label columns (col 0 = label, col 1 = unit).
    // The row label itself is the first element in each row.
    assert.deepStrictEqual(result.leftLabelValues, [["Alpha", "0%"], ["Beta", "0%"], ["BASE", "0%"]]);
    assert.deepStrictEqual(result.valuesForCalculation[0], ["0.44", "0.41"]);
  });

  it("[NPS-scale-label | 0%-unit-col | data] strips both columns for NPS table", () => {
    // Extended NPS table where col[0] is the scale label column (mostly numeric)
    // and col[1] is a '%' unit column with "0%" in every row.
    // detectLabelColumns uses the NPS early-return path, so col[1] was previously
    // never checked.
    const rawText = [
      ["0",          "0%", "2%", "3%", "4%"],
      ["1",          "0%", "3%", "4%", "5%"],
      ["2",          "0%", "5%", "7%", "8%"],
      ["3",          "0%", "4%", "3%", "2%"],
      ["4",          "0%", "6%", "5%", "4%"],
      ["5",          "0%", "8%", "9%", "10%"],
      ["6",          "0%", "7%", "8%", "9%"],
      ["7",          "0%", "9%", "10%", "11%"],
      ["8",          "0%", "12%", "13%", "14%"],
      ["9",          "0%", "8%",  "7%",  "6%"],
      ["10",         "0%", "36%", "33%", "27%"],
      ["Detractors", "0%", "12%", "14%", "10%"],
      ["Neutral",    "0%", "14%", "16%", "19%"],
      ["Promoters",  "0%", "74%", "70%", "71%"],
      ["NPS",        "0%", "62%", "56%", "61%"],
      ["Base",       "0%", "500", "600", "400"],
    ];
    const cleanedValues = rawText.map((row) => [row[0], "0", ...row.slice(2)]);

    const result = normalizeSelectedRange(cleanedValues, rawText);

    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, true, "NPS + unit column must normalize");
    assert.strictEqual(result.dataColOffset, 2, "data must start at col 2, skipping NPS label and unit col");
    assert.deepStrictEqual(result.labelColumns, [0, 1]);
    assert.deepStrictEqual(result.valuesForCalculation[0], ["2%", "3%", "4%"]);
    assert.deepStrictEqual(result.blockingReasons, []);
  });

  it("[NPS-scale-label | '%'-unit-col | data] strips both columns when unit column shows '%'", () => {
    // Real Excel smoke case: col[1] contains the literal text "%" in every row
    // (cells are text, not value=0 formatted as %).  isTextOnlyCell("%") = true,
    // but the NPS early-return paths skip the col1Frac check and went straight to
    // isUniformZeroPercentColumn which only accepted "0%".  isUniformUnitColumn now
    // also accepts the bare "%" string.
    const rawText = [
      ["1 - Совершенно неудовлет", "%", "2%",  "3%",  "4%"],
      ["2",          "%", "3%",  "4%",  "5%"],
      ["3",          "%", "5%",  "7%",  "8%"],
      ["4",          "%", "4%",  "3%",  "2%"],
      ["5",          "%", "6%",  "5%",  "4%"],
      ["6",          "%", "8%",  "9%",  "10%"],
      ["7",          "%", "7%",  "8%",  "9%"],
      ["8",          "%", "9%",  "10%", "11%"],
      ["9",          "%", "8%",  "7%",  "6%"],
      ["10",         "%", "36%", "33%", "27%"],
      ["Detractors", "%", "12%", "14%", "10%"],
      ["Neutral",    "%", "14%", "16%", "19%"],
      ["Promoters",  "%", "74%", "70%", "71%"],
      ["NPS",        "%", "62%", "56%", "61%"],
      ["Base",       "%", "500", "600", "400"],
    ];
    const cleanedValues = rawText.map((row) => [row[0], "%", ...row.slice(2)]);

    const result = normalizeSelectedRange(cleanedValues, rawText);

    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, true, "NPS + '%' unit column must normalize");
    assert.strictEqual(result.dataColOffset, 2, "data must start at col 2");
    assert.deepStrictEqual(result.labelColumns, [0, 1]);
    assert.deepStrictEqual(result.valuesForCalculation[0], ["2%", "3%", "4%"]);
    assert.deepStrictEqual(result.blockingReasons, []);
  });

  // ── Issue #132: 2-column partial banner + data subset selections ─────────────

  it("2-col banner+data subset normalizes: banner rows separated from data body", () => {
    // Simulates selecting cols B-C of "Ваш пол" where col A (row labels) is outside
    // the selection.  The selection contains 1 banner row at top and data rows below.
    // With GATE_MIN_COLS=2 the gate passes and banner is properly separated.
    const values = [
      ["2025Q4", "2026Q1"],   // banner row: wide text
      [0.44,     0.41],       // Мужской
      [0.56,     0.59],       // Женский
      [5605,     1320],       // BASE
    ];
    const result = normalizeSelectedRange(values);

    assert.strictEqual(result.normalizationNeeded, true, "2-col banner+data must trigger normalization");
    assert.strictEqual(result.normalizationApplied, true, "must normalize, not block");
    assert.deepStrictEqual(result.bannerRows, [0], "row 0 is the banner row");
    assert.strictEqual(result.dataRowOffset, 1, "data body starts at row 1");
    assert.strictEqual(result.valuesForCalculation.length, 3, "three data rows");
    assert.deepStrictEqual(result.valuesForCalculation[0], [0.44, 0.41], "first data row");
    assert.deepStrictEqual(result.valuesForCalculation[2], [5605, 1320], "BASE row included");
    assert.deepStrictEqual(result.blockingReasons, []);
  });

  it("2-col banner+data subset with rawText preserves banner labels in bannerContext", () => {
    // 2-column selection of two data columns (no label column in selection).
    // Two banner rows at top carry column headers.  rawText carries the original
    // text; cleanedValues has the numeric body after significance marker removal.
    // bannerContext.scanRows must be sourced from rawText, not cleanedValues.
    const rawText = [
      ["Волна (квартал)", "Всего"],   // wide banner row 1
      ["2025Q4", "2026Q1"],           // wide banner row 2
      ["44%",  "41%"],                // data row (Мужской)
      ["56%",  "59%"],                // data row (Женский)
      ["5605", "1320"],               // BASE
    ];
    const cleanedValues = [
      ["Волна (квартал)", "Всего"],   // unchanged (no markers)
      ["2025Q4",          "2026Q1"],  // unchanged
      [0.44, 0.41],                   // cleaned numeric
      [0.56, 0.59],
      [5605, 1320],
    ];

    const result = normalizeSelectedRange(cleanedValues, rawText);

    assert.strictEqual(result.normalizationApplied, true, "must normalize");
    assert.deepStrictEqual(result.bannerRows, [0, 1], "both banner rows detected");
    assert.strictEqual(result.dataRowOffset, 2, "data starts after 2 banner rows");
    assert.strictEqual(result.valuesForCalculation.length, 3, "three data rows");
    assert.deepStrictEqual(result.bannerContext.scanRows[0], ["Волна (квартал)", "Всего"],
      "first banner row from rawText");
    assert.deepStrictEqual(result.bannerContext.scanRows[1], ["2025Q4", "2026Q1"],
      "second banner row from rawText");
    assert.strictEqual(result.bannerContext.columnCount, 2);
    assert.deepStrictEqual(result.blockingReasons, []);
  });

  it("2-col purely numeric selection still passes through (no normalization)", () => {
    // A 2-col selection of pure numbers must not trigger normalization — there is
    // nothing structural to decompose.
    const values = [
      [0.44, 0.41],
      [0.56, 0.59],
      [5605, 1320],
    ];
    const result = normalizeSelectedRange(values);

    assert.strictEqual(result.normalizationNeeded, false, "pure numeric 2-col must pass through");
    assert.strictEqual(result.normalizationApplied, false);
    assert.deepStrictEqual(result.blockingReasons, []);
  });

  it("2-col banner-only selection (no data rows) blocks with BODY_TOO_SHORT", () => {
    // Selecting only the banner header rows without any data body must be blocked —
    // there is nothing to calculate.
    const values = [
      ["Всего",  "Кат A"],  // banner
      ["2025Q4", "2025Q4"], // banner
    ];
    const result = normalizeSelectedRange(values);

    // Either not needed (too few rows for the gate) or blocked — must not produce
    // a successful normalization with valuesForCalculation.
    const noDataToCalculate =
      result.normalizationNeeded === false ||
      (result.normalizationApplied === false && result.blockingReasons.length > 0);
    assert.ok(noDataToCalculate, `expected pass-through or blocked, got: ${JSON.stringify(result)}`);
    assert.ok(
      !result.normalizationApplied || result.valuesForCalculation.length === 0,
      "must not produce a non-empty valuesForCalculation for a banner-only selection"
    );
  });

  it("[label | 0%-unit-col | data] with mixed 0% values does not strip non-uniform column", () => {
    // If col[1] has different values (not all uniform), it must NOT be treated as a
    // unit column even if some cells are "0%".  This protects real data columns.
    const rawText = [
      ["Metric A", "0%",  "21%", "35%"],
      ["Metric B", "10%", "15%", "22%"],
      ["Base",     "0%",  "100", "200"],
    ];
    const cleanedValues = [
      ["Metric A", "0",    "0.21", "0.35"],
      ["Metric B", "0.10", "0.15", "0.22"],
      ["Base",     "0",    "100",  "200"],
    ];

    const result = normalizeSelectedRange(cleanedValues, rawText);

    assert.strictEqual(result.normalizationApplied, true);
    // col[1] is not uniform ("0%" and "10%" differ), so only col[0] is stripped.
    assert.strictEqual(result.dataColOffset, 1, "non-uniform col[1] must not be treated as unit column");
    assert.deepStrictEqual(result.labelColumns, [0]);
  });
});

describe("hasEmptyDataRowGap", () => {
  it("returns false for a clean numeric grid with no empty rows", () => {
    const values = [
      [10, 20, 30],
      [40, 50, 60],
      [100, 200, 300],
    ];
    assert.strictEqual(hasEmptyDataRowGap(values), false);
  });

  it("returns true when one row is all-null", () => {
    const values = [
      [10, 20, 30],
      [null, null, null],
      [40, 50, 60],
    ];
    assert.strictEqual(hasEmptyDataRowGap(values), true);
  });

  it("returns true when one row is all-empty-string", () => {
    const values = [
      [10, 20, 30],
      ["", "", ""],
      [40, 50, 60],
    ];
    assert.strictEqual(hasEmptyDataRowGap(values), true);
  });

  it("returns true when one row is all-undefined", () => {
    const values = [
      [10, 20, 30],
      [undefined, undefined, undefined],
      [40, 50, 60],
    ];
    assert.strictEqual(hasEmptyDataRowGap(values), true);
  });

  it("returns true when gap row is mixed null and empty-string", () => {
    const values = [
      [10, 20],
      [null, ""],
      [40, 50],
    ];
    assert.strictEqual(hasEmptyDataRowGap(values), true);
  });

  it("returns false when a row has at least one non-blank cell", () => {
    const values = [
      [10, 20, 30],
      [0, null, null],
      [40, 50, 60],
    ];
    assert.strictEqual(hasEmptyDataRowGap(values), false, "numeric 0 is non-blank");
  });

  it("returns false for a single-row grid", () => {
    assert.strictEqual(hasEmptyDataRowGap([[10, 20, 30]]), false);
  });

  it("returns false for an empty array", () => {
    assert.strictEqual(hasEmptyDataRowGap([]), false);
  });

  it("returns false for a non-array input", () => {
    assert.strictEqual(hasEmptyDataRowGap(null), false);
    assert.strictEqual(hasEmptyDataRowGap(undefined), false);
  });

  it("returns true for two-table pass-through scenario: numeric tables separated by an empty row", () => {
    // Mirrors the smoke scenario: two purely numeric tables that both pass
    // isNormalizationNeeded=false, joined by an empty row.
    // normalizeSelectedRange's validateBody would catch this for normalized ranges;
    // hasEmptyDataRowGap catches it for pass-through valuesForCalculation.
    const values = [
      [30, 40, 50],   // table A row 1
      [100, 200, 300], // table A base
      ["", "", ""],    // empty separator
      [20, 30, 10],   // table B row 1
      [80, 150, 180], // table B base
    ];
    assert.strictEqual(hasEmptyDataRowGap(values), true);
  });
});
