import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { normalizeSelectedRange } from "../../src/core/range-normalizer.js";

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
      ["Ваш возраст",     "",       ""      ],
      ["",                "Всего",  ""      ],
      ["",                "Волна",  ""      ],
      ["",                "2025Q4", "2026Q1"],
      ["19 и младше",     0.19,     0.15    ],
      ["20-29",           0.37,     0.32    ],
      ["от 30 до 40",     0.28,     0.28    ],
      ["40 и старше",     0.17,     0.24    ],
      ["mean",            29.4,     31.5    ],
      ["variance",        103.6,    132.6   ],
      ["BASE",            5605,     1320    ],
      ["Все респонденты", 6925,     2640    ],
    ];
    const result = normalizeSelectedRange(values);

    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, true);
    assert.deepStrictEqual(result.titleRows, [0], "title row should be row 0");
    assert.deepStrictEqual(result.bannerRows, [1, 2, 3], "sparse banner rows 1-3 should be detected");
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
      ["Ваш пол:",          "",                    "",                   ""             ],
      ["",                  "Всего",               "",                   "Пользование категорией"],
      ["",                  "Волна (квартал)",     "Всё покупаю сам(а)", "Большую часть"],
      ["",                  "2025Q4",              "2026Q1",             "2025Q4"       ],
      ["",                  "(a)",                 "(a)",                "(a)"          ],
      ["Мужской",           "44%",                 "41%",                "39%"          ],
      ["Женский",           "56%",                 "59%",                "61%"          ],
      ["BASE",              5605,                  1320,                 3083           ],
      ["Все респонденты",   "",                    "",                   ""             ],
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
      ["", "Волна (квартал)", "", "Всё покупаю сам(а)", "", "Большую часть", "", "", "", "", "", "", ""],
      ["", "2025Q4", "2026Q1", "Волна (квартал)", "", "Волна (квартал)", "", "", "", "", "", "", ""],
      ["", "", "", "2025Q4", "2026Q1", "2025Q4", "2026Q1", "2025Q4", "2026Q1", "2025Q4", "2026Q1", "2025Q4", "2026Q1"],
      ["Мужской", "0.5", "0.4136", "0.39", "0.41", "0.44", "0.42", "0.48", "0.43", "0.51", "0.49", "0.46", "0.45"],
      ["Женский", "0.5", "0.5863", "0.61", "0.59", "0.56", "0.58", "0.52", "0.57", "0.49", "0.51", "0.54", "0.55"],
      ["BASE", "5605", "1320", "3083", "1045", "2200", "900", "1800", "760", "1500", "700", "1300", "620"],
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
    assert.deepStrictEqual(result.bannerContext.scanRows[0].slice(0, 4), ["", "", "Пользование кат", ""]);
  });

  it("real sparse 4-row banner + means body preserves service labels from Run text", () => {
    const rawText = [
      ["Ваш возраст", "", "", "", "", "", "", "", "", "", "", "", ""],
      ["", "", "", "Пользование кат", "", "", "", "", "", "", "", "", ""],
      ["", "Волна (квартал)", "", "Всё покупаю сам(а)", "", "Большую часть", "", "", "", "", "", "", ""],
      ["", "2025Q4", "2026Q1", "Волна (квартал)", "", "Волна (квартал)", "", "", "", "", "", "", ""],
      ["", "", "", "2025Q4", "2026Q1", "2025Q4", "2026Q1", "2025Q4", "2026Q1", "2025Q4", "2026Q1", "2025Q4", "2026Q1"],
      ["Среднее", "29.4", "31.5", "30.1", "32.2", "28.9", "30.7", "33.0", "34.1", "27.5", "28.0", "35.2", "36.4"],
      ["variance", "103.6", "132.6", "120.1", "140.4", "99.2", "118.8", "150.0", "160.1", "90.5", "95.2", "170.4", "180.3"],
      ["BASE", "5605", "1320", "3083", "1045", "2200", "900", "1800", "760", "1500", "700", "1300", "620"],
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
      ["Таблица 1",  "",       "",       ""],
      ["",           "Всего",  "Кат A",  ""],
      ["",           "2025Q4", "2025Q4", ""],
      ["Вариант A",  0.4,      0.3,      ""],
      ["Вариант B",  0.6,      0.7,      ""],
      ["BASE",       1000,     500,      ""],
      // Table 2 — sparse title provides all-empty data cols → gap
      ["Таблица 2",  "",       "",       ""],
      ["",           "Всего",  "Кат B",  ""],
      ["",           "2026Q1", "2026Q1", ""],
      ["Вариант X",  0.5,      0.4,      ""],
      ["Вариант Y",  0.5,      0.6,      ""],
      ["BASE",       800,      400,      ""],
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
      ["Ваш возраст",     "",       ""      ],
      ["",                "Всего",  ""      ],
      ["",                "Волна",  ""      ],
      ["",                "2025Q4", "2026Q1"],
      ["",                "(a)",    "(a)"   ],
      ["19 и младше",     "19%",    "15%"   ],
      ["20-29",           "37%",    "32%"   ],
      ["от 30 до 40",     "28%",    "28%"   ],
      ["40 и старше",     "17%",    "24%"   ],
      ["mean",            "29,4",   "31,5"  ],
      ["variance",        "103,6",  "132,6" ],
      ["BASE",            "5605",   "1320"  ],
      ["Все респонденты", "",       ""      ],
    ];
    // Simulate significance marker removal: pure-letter label cells → "".
    const strippedValues = rawText.map((row, r) =>
      [9, 10, 11].includes(r)
        ? ["", row[1], row[2]]   // "mean"/"variance"/"BASE" erased
        : row
    );

    const result = normalizeSelectedRange(strippedValues, rawText);

    assert.strictEqual(result.normalizationApplied, true);
    // Labels must come from rawText, not from strippedValues.
    assert.deepStrictEqual(result.leftLabelValues[4], ["mean"],     "mean label preserved from text");
    assert.deepStrictEqual(result.leftLabelValues[5], ["variance"], "variance label preserved from text");
    assert.deepStrictEqual(result.leftLabelValues[6], ["BASE"],     "BASE label preserved from text");
  });

  it("Возраст full-table shape with service rows normalizes correctly", () => {
    // Exact shape from Excel smoke test, values as strings (Russian comma-decimal,
    // percent strings, quarter labels). Verifies leftLabelValues contains the
    // service-row labels ("mean", "variance", "BASE") so buildTablePreviewModel
    // can classify them correctly.
    const values = [
      ["Ваш возраст",     "",                   ""      ],
      ["",                "Всего",              ""      ],
      ["",                "Волна (квартал)",    ""      ],
      ["",                "2025Q4",             "2026Q1"],
      ["",                "(a)",                "(a)"   ],
      ["19 и младше",     "19%",                "15%"   ],
      ["20-29",           "37%",                "32%"   ],
      ["от 30 до 40",     "28%",                "28%"   ],
      ["40 и старше",     "17%",                "24%"   ],
      ["mean",            "29,4",               "31,5"  ],
      ["variance",        "103,6",              "132,6" ],
      ["BASE",            "5605",               "1320"  ],
      ["Все респонденты", "",                   ""      ],
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
    assert.deepStrictEqual(result.leftLabelValues[4], ["mean"],     "mean row label");
    assert.deepStrictEqual(result.leftLabelValues[5], ["variance"], "variance row label");
    assert.deepStrictEqual(result.leftLabelValues[6], ["BASE"],     "BASE row label");
    assert.deepStrictEqual(result.blockingReasons, []);
  });

  it("trailing label-only footer row is excluded from body, normalization succeeds", () => {
    // Simulates "Ваш пол" shape: sparse banner rows at top, two data rows,
    // then "Все респонденты" which has a label in col[0] but empty data columns.
    // The footer row must be trimmed so it does not trigger BODY_APPEARS_MULTI_TABLE.
    const values = [
      ["Ваш пол:",         "",       ""      ],
      ["",                 "Всего",  ""      ],
      ["",                 "2025Q4", "2026Q1"],
      ["Мужской",          0.44,     0.41    ],
      ["Женский",          0.56,     0.59    ],
      ["BASE",             5605,     1320    ],
      ["Все респонденты",  "",       ""      ],
    ];
    const result = normalizeSelectedRange(values);
    assert.strictEqual(result.normalizationNeeded, true);
    assert.strictEqual(result.normalizationApplied, true);
    assert.deepStrictEqual(result.blockingReasons, [], "footer row must not trigger BODY_APPEARS_MULTI_TABLE");
    assert.strictEqual(result.valuesForCalculation.length, 3, "footer row must be excluded from valuesForCalculation");
    assert.deepStrictEqual(result.valuesForCalculation[0], [0.44, 0.41], "first data row");
    assert.deepStrictEqual(result.valuesForCalculation[2], [5605, 1320], "BASE row included");
  });

  it("multi-table selection with mid-body empty row blocks normalization with BODY_APPEARS_MULTI_TABLE", () => {
    // Two groups of data rows separated by a fully empty row.
    // The empty row lies between actual numeric rows so bodyDataEndRow does NOT
    // trim it away — the gap check must fire and block normalization.
    const values = [
      ["Группа",  "A",  "B",  "C" ],
      ["Label1",  10,   20,   30  ],
      ["Label2",  40,   50,   60  ],
      [null,      null, null, null],
      ["Label3",  70,   80,   90  ],
      ["Label4",  10,   20,   30  ],
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
});
