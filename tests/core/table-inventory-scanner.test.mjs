import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { scanWorksheetForTables } from "../../src/core/table-inventory-scanner.js";

const SHEET = "Sheet1";
const OFFSET = { usedRangeRowOffset: 0, usedRangeColOffset: 0, sheetName: SHEET };

describe("scanWorksheetForTables", () => {
  it("two tables separated by an empty row return 2 inventory items", () => {
    const values = [
      ["Label1", 10, 20, 30],
      ["Label2", 40, 50, 60],
      ["Base", 100, 100, 100],
      [null, null, null, null],  // empty separator
      ["Label3", 70, 80, 90],
      ["Label4", 10, 20, 30],
      ["Base", 200, 200, 200],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.strictEqual(items.length, 2, `expected 2 items, got ${items.length}`);
    assert.strictEqual(items[0].sheetName, SHEET);
    assert.strictEqual(items[1].sheetName, SHEET);
  });

  it("first-row merged-like title is detected and surfaced on the item", () => {
    // Row 0: single text cell (sparse, no numeric) → title-like.
    // Rows 1-3: numeric data band.
    const values = [
      ["My Survey", null, null, null],
      ["Label1", 10, 20, 30],
      ["Label2", 40, 50, 60],
      ["Base", 100, 100, 100],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    const item = items[0];
    assert.strictEqual(item.title, "My Survey");
    assert.strictEqual(item.titleSource, "firstRowOfBand");
  });

  it("column A=label, column B=Total keeps Total as data (not a second label column)", () => {
    // Col 0 is text labels. Col 1 header is "Total" (text in row 0 only).
    // Remaining rows of col 1 are numeric → col 1 text fraction < threshold
    // → classified as data, not a second label column.
    const values = [
      ["Category", "Total", "Male", "Female"],
      ["Label1", 100, 60, 40],
      ["Label2", 200, 120, 80],
      ["Base", 300, 180, 120],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    const item = items[0];
    // Confident split means col 0 was correctly identified as the only label column.
    assert.strictEqual(
      item.labelSplitConfidence,
      "confident",
      `expected confident split; got ${item.labelSplitConfidence}`
    );
    // Total band width is 4 columns (cols 0-3).
    assert.strictEqual(item.columnCount, 4);
  });

  it("quarter-like labels alone are not treated as numeric evidence", () => {
    const values = [
      ["Wave", "2025Q4", "2026Q1"],
      ["Segment", "2025Q4", "2026Q1"],
      ["Note", "Q4 2025", "Q1 2026"],
    ];

    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.strictEqual(items.length, 0, `expected 0 items, got ${items.length}`);
  });

  it("banner-heavy tables with quarter-like headers still produce one inventory item", () => {
    const values = [
      ["Usage table", null, null, null],
      ["", "Total", "Male", "Female"],
      ["", "2025Q4", "2025Q4", "2025Q4"],
      ["Agree", "44%", "41%", "39%"],
      ["Disagree", "56%", "59%", "61%"],
      ["BASE", "5605", "1320", "3083"],
    ];

    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.strictEqual(items.length, 1, `expected 1 item, got ${items.length}`);

    const item = items[0];
    assert.strictEqual(item.title, "Usage table");
    assert.strictEqual(item.titleSource, "firstRowOfBand");
    assert.strictEqual(item.labelSplitConfidence, "confident");
    assert.strictEqual(item.columnCount, 4);
    assert.strictEqual(item.isLikelyTable, true);
  });

  // ─── New hardening tests for issue #110 ──────────────────────────────────

  it("item does not expose canRunSignificance — scanner is a candidate finder only", () => {
    const values = [
      ["Label1", 10, 20, 30],
      ["Label2", 40, 50, 60],
      ["Base", 100, 100, 100],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    assert.strictEqual(
      items[0].canRunSignificance,
      undefined,
      "canRunSignificance must not exist on inventory items"
    );
    assert.ok(
      ["available", "uncertain", "rejected"].includes(items[0].candidateStatus),
      `candidateStatus must be available/uncertain/rejected; got ${items[0].candidateStatus}`
    );
  });

  it("clean table with base produces candidateStatus=available", () => {
    const values = [
      ["Label1", 10, 20, 30],
      ["Label2", 40, 50, 60],
      ["Base", 100, 100, 100],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1);
    assert.strictEqual(items[0].candidateStatus, "available");
  });

  it("table with no explicit Base row is not presented as Run-ready — no canRunSignificance field", () => {
    const values = [
      ["Label1", 10, 20, 30],
      ["Label2", 40, 50, 60],
      ["Label3", 70, 80, 90],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    // Without a Base row the scanner may produce 0 or 1 items depending on metric detection.
    // If an item exists, it must not carry canRunSignificance (the key safety property).
    if (items.length > 0) {
      assert.strictEqual(items[0].canRunSignificance, undefined, "canRunSignificance must not exist");
      // "rejected" is acceptable — no Base means no complete metric blocks.
      assert.ok(
        ["available", "uncertain", "rejected"].includes(items[0].candidateStatus),
        `candidateStatus must be a recognised value; got ${items[0].candidateStatus}`
      );
    }
    // If items.length === 0, the scanner correctly found nothing to over-promise on.
  });

  it("candidate with preview warnings is surfaced as uncertain", () => {
    // All-100 rows trigger a quality warning in the preview model.
    const values = [
      ["Label1", 100, 100, 100],
      ["Label2", 100, 100, 100],
      ["Base", 1000, 1000, 1000],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    const item = items[0];
    // Warnings must be reflected in the candidate status.
    if (item.warningsCount > 0) {
      assert.strictEqual(
        item.candidateStatus,
        "uncertain",
        "item with preview warnings must be uncertain, not available"
      );
    }
  });

  it("side-by-side tables in one row band appear as one candidate and expose no canRunSignificance", () => {
    // Known limitation: scanner cannot split side-by-side tables within one band.
    // This test documents the limitation and ensures no over-promising.
    const values = [
      ["", "Total", "Male", "", "Total", "Female"],
      ["Metric A", 100, 60, "Metric B", 200, 120],
      ["Metric A2", 150, 90, "Metric B2", 250, 150],
      ["Base", 1000, 600, "Base", 2000, 1200],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    // Entire band is treated as one candidate (known limitation).
    assert.strictEqual(items.length, 1, "side-by-side tables are reported as one band (known limitation)");
    assert.strictEqual(items[0].canRunSignificance, undefined, "canRunSignificance must not exist");
    assert.ok(
      ["available", "uncertain", "rejected"].includes(items[0].candidateStatus),
      "candidateStatus must be a recognised value"
    );
  });

  it("non-empty commentary row between two tables merges them — no split, no false confidence", () => {
    // Tables separated by a non-empty row are merged into one band (known limitation).
    // The merged candidate must not report canRunSignificance.
    const values = [
      ["Label1", 10, 20, 30],
      ["Label2", 40, 50, 60],
      ["Base", 100, 100, 100],
      ["Note: preliminary data", null, null, null],  // non-empty, prevents band split
      ["Label3", 70, 80, 90],
      ["Label4", 10, 20, 30],
      ["Base", 200, 200, 200],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    // Should be 1 merged band, not 2.
    assert.strictEqual(items.length, 1, "non-empty note row prevents band split — expect 1 merged candidate");
    assert.strictEqual(items[0].canRunSignificance, undefined, "canRunSignificance must not exist");
  });

  it("title inferred two rows above band via empty separator gets medium confidence", () => {
    // Row 0: section title. Row 1: empty separator. Rows 2-4: table band.
    // twoRowsAbove inference should be medium confidence, not high.
    const values = [
      ["Section Title", null, null],
      [null, null, null],
      ["Label1", 10, 20],
      ["Label2", 30, 40],
      ["Base", 100, 100],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    const item = items[0];
    if (item.titleSource === "twoRowsAbove") {
      assert.strictEqual(
        item.titleConfidence,
        "medium",
        "twoRowsAbove title confidence must be medium (text may belong to a preceding section)"
      );
    }
  });

  it("item exposes candidateNotes not reasonsIfNotRunnable", () => {
    const values = [
      ["Label1", 10, 20],
      ["Label2", 30, 40],
      ["Base", 100, 100],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1);
    assert.ok(Array.isArray(items[0].candidateNotes), "candidateNotes must be an array");
    assert.strictEqual(items[0].reasonsIfNotRunnable, undefined, "reasonsIfNotRunnable must not exist");
  });

  it("item exposes labelColCount and qualityIssueCodes for diagnostics", () => {
    // labelColCount must match the number of left label columns detected.
    // qualityIssueCodes must be an array of {code, severity} objects — no raw issue objects.
    const values = [
      ["Label1", 10, 20, 30],
      ["Label2", 40, 50, 60],
      ["Base", 100, 100, 100],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    const item = items[0];
    assert.strictEqual(typeof item.labelColCount, "number", "labelColCount must be a number");
    assert.ok(item.labelColCount >= 0, "labelColCount must be non-negative");
    assert.ok(Array.isArray(item.qualityIssueCodes), "qualityIssueCodes must be an array");
    for (const entry of item.qualityIssueCodes) {
      assert.ok(typeof entry.code === "string", "each qualityIssueCodes entry must have a string code");
      assert.ok(
        entry.severity === "warning" || entry.severity === "critical",
        `each qualityIssueCodes entry must have severity warning or critical; got ${entry.severity}`
      );
    }
  });

  it("rating/NPS scale labels in the first column keep the label split confident", () => {
    const values = [
      ["Score", "Total", "Male", "Female"],
      ["1 - very unlikely", 5, 4, 6],
      ["2", 6, 5, 7],
      ["3", 7, 6, 8],
      ["4", 8, 7, 9],
      ["5", 9, 8, 10],
      ["6", 10, 9, 11],
      ["7", 11, 10, 12],
      ["8", 12, 11, 13],
      ["9", 13, 12, 14],
      ["10 - very likely", 14, 13, 15],
      ["Bottom-3", 18, 17, 19],
      ["Mid-4", 32, 30, 34],
      ["Top-3", 50, 53, 47],
      ["Detractors", 20, 19, 18],
      ["Neutral", 30, 29, 31],
      ["Promoters", 50, 52, 51],
      ["NPS", 30, 33, 33],
      ["BASE", 1000, 500, 500],
    ];

    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");

    const item = items[0];
    assert.strictEqual(item.labelSplitConfidence, "confident");
    assert.ok(
      !item.candidateNotes.includes("Граница лейблов/данных не определена"),
      "rating/NPS label column should not produce an uncertain label/data boundary note"
    );
  });

  it("pure numeric first columns still stay uncertain when they look like data, not labels", () => {
    const values = [
      [1, 10, 20, 30],
      [2, 40, 50, 60],
      [3, 70, 80, 90],
      [4, 100, 110, 120],
    ];

    const items = scanWorksheetForTables({ values, ...OFFSET });

    if (items.length > 0) {
      assert.strictEqual(
        items[0].labelSplitConfidence,
        "uncertain",
        "pure numeric first columns must not become confidently label-like"
      );
    }
  });

  // ─── Issue #153: two-column row label tables ──────────────────────────────

  it("non-NPS table with pure ordinal col0 + text answer col1 is available, not uncertain", () => {
    // Pattern: [scale code] [answer label] [data...]
    // col0 is integers 1-5 (ordinal scale codes, not text), col1 is text answer labels.
    const values = [
      [1, "Strongly agree", 40, 35, 42],
      [2, "Agree", 30, 32, 28],
      [3, "Neither agree nor disagree", 15, 17, 13],
      [4, "Disagree", 10, 12, 10],
      [5, "Strongly disagree", 5, 4, 7],
      ["Base", "", 100, 100, 100],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    const item = items[0];
    // "twoColumn" when col0 fails the text test but is ordinal;
    // "confident" when isLikelyRatingScaleLabelColumn already accepts the mixed col0.
    assert.ok(
      item.labelSplitConfidence === "twoColumn" || item.labelSplitConfidence === "confident",
      `expected twoColumn or confident split; got ${item.labelSplitConfidence}`
    );
    assert.strictEqual(
      item.candidateStatus,
      "available",
      `valid two-column label table must not be uncertain; got ${item.candidateStatus}`
    );
  });

  it("non-NPS table with ordinal string col0 (\"1\"..\"5\") + text col1 is available", () => {
    // Same pattern but col0 values are strings "1".."5" rather than integers.
    const values = [
      ["1", "Very satisfied", 44, 41, 45],
      ["2", "Satisfied", 30, 32, 28],
      ["3", "Neither", 15, 17, 13],
      ["4", "Dissatisfied", 8, 7, 10],
      ["5", "Very dissatisfied", 3, 3, 4],
      ["Base", "", 500, 450, 550],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    const item = items[0];
    assert.ok(
      item.labelSplitConfidence === "twoColumn" || item.labelSplitConfidence === "confident",
      `expected twoColumn or confident split; got ${item.labelSplitConfidence}`
    );
    assert.strictEqual(
      item.candidateStatus,
      "available",
      `two-column label table must not be uncertain; got ${item.candidateStatus}`
    );
  });

  it("pure ordinal col0 with no text cell at all triggers twoColumn path when col1 is text", () => {
    // col0 has only integers 1-5 (no "Base" or any text cell).
    // isLikelyRatingScaleLabelColumn requires textOnlyCount >= 1, so it returns false here.
    // The new isMostlyOrdinalLabelColumn + col1 text check should set twoColumn.
    // A Base row is provided so the preview model can build a complete metric block.
    // col0 of the Base row is null so col0 stays purely ordinal (no text cells).
    const values = [
      [1, "Strongly agree", 40, 35, 42],
      [2, "Agree", 30, 32, 28],
      [3, "Neither", 15, 17, 13],
      [4, "Disagree", 10, 12, 10],
      [5, "Strongly disagree", 5, 4, 7],
      [null, "BASE", 100, 100, 100],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    const item = items[0];
    assert.strictEqual(
      item.labelSplitConfidence,
      "twoColumn",
      `pure-ordinal col0 with text col1 should yield twoColumn; got ${item.labelSplitConfidence}`
    );
    assert.strictEqual(
      item.candidateStatus,
      "available",
      `twoColumn split must not make the candidate uncertain; got ${item.candidateStatus}`
    );
    assert.ok(
      item.candidateNotes.some((n) => n.includes("Двухколоночные метки")),
      "candidateNotes must explain two-column labels for twoColumn split"
    );
  });

  it("ordinal col0 + numeric col1 stays uncertain (col1 data, not a second label)", () => {
    // col0 is ordinal 1-5 but col1 is numeric data — should not be classified as
    // a two-column label table.
    const values = [
      [1, 40, 35, 42],
      [2, 30, 32, 28],
      [3, 15, 17, 13],
      [4, 10, 12, 10],
      [5, 5, 4, 7],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    if (items.length > 0) {
      assert.notStrictEqual(
        items[0].labelSplitConfidence,
        "twoColumn",
        "ordinal col0 with numeric col1 must not be classified as two-column label"
      );
    }
  });

  it("ordinal col0 + single Base text cell yields confident split (isLikelyRatingScaleLabelColumn fix)", () => {
    // col0 has ordinal values 1-5 + one text cell "Base".
    // isLikelyRatingScaleLabelColumn should now accept textOnlyCount=1 when all cells are label-like.
    const values = [
      ["1", "Very likely", 50, 48, 52],
      ["2", "Likely", 25, 26, 24],
      ["3", "Unlikely", 15, 16, 14],
      ["4", "Very unlikely", 10, 10, 10],
      ["Base", "", 1000, 950, 1050],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    const item = items[0];
    // With only 4 ordinal values in col0, isLikelyRatingScaleLabelColumn needs
    // ordinalScaleCount >= 3 — should pass (4 ordinal + 1 text, all label-like).
    assert.strictEqual(
      item.candidateStatus,
      "available",
      `table with two-column labels and Base row must not be uncertain; got ${item.candidateStatus}`
    );
  });

  it("mean/variance/base table with empty gutter column is available, not uncertain", () => {
    // Smoke-failing case from PR #154 review: metric labels in col0, col1 is an empty
    // visual gutter, col2+ is numeric data.  The gutter must not land in the data matrix
    // and cause quality warnings that flip the candidate to uncertain.
    const values = [
      ["Сколько (примерно) потратили?", null, null, null, null],  // title row
      ["Среднее", null, 5000, 4500, 5500],
      ["variance", null, 2500, 2300, 2700],
      ["BASE", null, 500, 450, 550],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    const item = items[0];
    assert.strictEqual(
      item.candidateStatus,
      "available",
      `mean/variance/base with gutter column must not be uncertain; got ${item.candidateStatus}`
    );
    assert.strictEqual(item.isLikelyTable, true);
  });

  it("mean/variance/base with non-ordinal code col0 + text metric col1 is available", () => {
    // Variant: col0 carries a non-ordinal code (e.g. year/wave number > 10) that fails
    // the text test, col1 has the metric labels.  The generalised twoColumn path must
    // still classify this as available.
    const values = [
      [2025, "Среднее", 5000, 4500, 5500],
      [2025, "variance", 2500, 2300, 2700],
      [2025, "BASE", 500, 450, 550],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    const item = items[0];
    assert.strictEqual(
      item.labelSplitConfidence,
      "twoColumn",
      `non-ordinal code col0 + text metric col1 must yield twoColumn; got ${item.labelSplitConfidence}`
    );
    assert.strictEqual(
      item.candidateStatus,
      "available",
      `non-ordinal code col0 + text metric col1 must not be uncertain; got ${item.candidateStatus}`
    );
  });

  it("metric table with text+sample banner rows in col1 and null gutter in body rows is available", () => {
    // Real-world structure: title row, two banner rows (col0=null, col1=segment labels/sample sizes),
    // then body rows with col0=metric label, col1=null gutter, col2+=data.
    // col1 banner rows carry text values ("Total", "(n=500)") but col1 is null for all body rows.
    // isGutterColumnForBodyRows must recognise col1 as a gutter (ignoring banner rows).
    const values = [
      ["Сколько (примерно) потратили?", null, null, null, null],
      [null, "Total", "Male", "Female", "Other"],
      [null, "(n=500)", "(n=250)", "(n=250)", "(n=100)"],
      ["Среднее", null, 5000, 4500, 4800],
      ["variance", null, 2500, 2300, 2400],
      ["BASE", null, 500, 450, 100],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected item");
    const item = items[0];
    assert.strictEqual(item.candidateStatus, "available",
      `${item.candidateStatus} split=${item.labelSplitConfidence} warn=${item.warningsCount} notes=${JSON.stringify(item.candidateNotes)}`);
  });

  it("metric table with numeric year in banner col1 and null gutter in body rows is available", () => {
    // col1 banner row carries a numeric year (2025) which is not a text value.
    // isGutterColumnForBodyRows must ignore the banner row (col0=null) and only
    // test body rows where col0 has content — those have col1=null → gutter detected.
    const values = [
      ["Сколько (примерно) потратили?", null, null, null, null],
      [null, 2025, "Male", "Female", "Other"],
      ["Среднее", null, 5000, 4500, 4800],
      ["variance", null, 2500, 2300, 2700],
      ["BASE", null, 500, 450, 100],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected item");
    const item = items[0];
    assert.strictEqual(item.candidateStatus, "available",
      `${item.candidateStatus} split=${item.labelSplitConfidence} warn=${item.warningsCount} notes=${JSON.stringify(item.candidateNotes)}`);
  });

  it("metric table with percent values in banner col1 and null gutter in body rows is available", () => {
    // col1 banner row carries percent strings ("44%", "41%", "39%") — numeric-like but text.
    // isGutterColumnForBodyRows must ignore that row and detect col1 as a gutter via body rows only.
    const values = [
      ["Сколько (примерно) потратили?", null, null, null],
      [null, "44%", "41%", "39%"],
      ["Среднее", null, 5000, 4500],
      ["variance", null, 2500, 2300],
      ["BASE", null, 500, 450],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected item");
    const item = items[0];
    assert.strictEqual(item.candidateStatus, "available",
      `${item.candidateStatus} split=${item.labelSplitConfidence} warn=${item.warningsCount} notes=${JSON.stringify(item.candidateNotes)}`);
  });

  it("metric table with two banner rows (text then numeric) in col1 and null gutter in body is available", () => {
    // Two banner rows: first has text ("Total"/"Male"/...), second has numeric sample sizes (1000/500/...).
    // Both have col0=null so isGutterColumnForBodyRows skips them.
    // Body rows (col0 non-empty) have col1=null → gutter correctly detected → no BASE_BLANK_VALUES warning.
    const values = [
      ["Сколько (примерно) потратили?", null, null, null, null],
      [null, "Total", "Male", "Female", "Other"],
      [null, 1000, 500, 450, 50],
      ["Среднее", null, 5000, 4500, 4800],
      ["variance", null, 2500, 2300, 2400],
      ["BASE", null, 500, 450, 50],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected item");
    const item = items[0];
    assert.strictEqual(item.candidateStatus, "available",
      `${item.candidateStatus} split=${item.labelSplitConfidence} warn=${item.warningsCount} notes=${JSON.stringify(item.candidateNotes)}`);
  });

  it("banner header rows wider than body data rows do not cause spurious BASE_BLANK_VALUES", () => {
    // banner row spans 4 data columns (C-F), but mean/variance/base only have
    // values in 3 data columns (C-E). Col F is present only in the banner header.
    // Trailing empty data column must be trimmed before quality checks so that
    // the base row does not get a spurious BASE_BLANK_VALUES warning.
    const values = [
      //  A               B      C         D           E           F (banner-only)
      ["Title",       null,  null,     null,       null,       null],   // title row
      ["Question?",   null,  null,     null,       null,       null],   // question title (unknownText)
      [null,          null,  "Total",  "Group A",  "Group B",  "Extra"],// banner — F has header
      [null,          null,  "Sub 1",  "Sub 2",    null,       null],   // banner row 2
      ["Mean",        null,  50,       45,         55,         null],   // F is empty
      ["Variance",    null,  100,      80,         120,        null],   // F is empty
      ["Base",        null,  500,      200,        300,        null],   // F is empty → BASE_BLANK_VALUES
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    const item = items[0];
    const hasBbv = item.qualityIssueCodes.some((e) => e.code === "BASE_BLANK_VALUES");
    assert.strictEqual(hasBbv, false,
      `should not produce BASE_BLANK_VALUES for a banner-extension trailing column; got codes=${JSON.stringify(item.qualityIssueCodes)}`);
    assert.strictEqual(
      item.candidateStatus,
      "available",
      `expected available; got ${item.candidateStatus} split=${item.labelSplitConfidence} warn=${item.warningsCount} codes=${JSON.stringify(item.qualityIssueCodes)}`
    );
  });

  it("extended NPS inventory behavior is preserved with two-column layout", () => {
    // Extended NPS: col0 has scale values 0-10 + text bucket rows + NPS + Base.
    // This previously worked; confirm it still does after the #153 changes.
    const values = [
      ["Score", "Total", "Male", "Female"],
      ["1 - very unlikely", 5, 4, 6],
      ["2", 6, 5, 7],
      ["3", 7, 6, 8],
      ["4", 8, 7, 9],
      ["5", 9, 8, 10],
      ["6", 10, 9, 11],
      ["7", 11, 10, 12],
      ["8", 12, 11, 13],
      ["9", 13, 12, 14],
      ["10 - very likely", 14, 13, 15],
      ["Bottom-3", 18, 17, 19],
      ["Top-3", 50, 53, 47],
      ["Detractors", 20, 19, 18],
      ["Promoters", 50, 52, 51],
      ["NPS", 30, 33, 33],
      ["BASE", 1000, 500, 500],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    const item = items[0];
    assert.strictEqual(item.labelSplitConfidence, "confident");
    assert.strictEqual(item.candidateStatus, "available");
    assert.ok(
      !item.candidateNotes.includes("Граница лейблов/данных не определена"),
      "extended NPS must not produce uncertain label/data boundary note"
    );
  });

  it("item exposes detectedMetricRows, detectedBaseRows, detectedBlocks, hasNps, hasMeans for a proportion table", () => {
    const values = [
      ["Label1", 10, 20, 30],
      ["Label2", 40, 50, 60],
      ["Base", 100, 100, 100],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    const item = items[0];
    assert.ok(typeof item.detectedMetricRows === "number", "detectedMetricRows should be a number");
    assert.ok(typeof item.detectedBaseRows === "number", "detectedBaseRows should be a number");
    assert.ok(typeof item.detectedBlocks === "number", "detectedBlocks should be a number");
    assert.strictEqual(typeof item.hasNps, "boolean", "hasNps should be a boolean");
    assert.strictEqual(typeof item.hasMeans, "boolean", "hasMeans should be a boolean");
    assert.ok(item.detectedMetricRows > 0, "proportion table should have metric rows");
    assert.ok(item.detectedBaseRows > 0, "proportion table should have at least one base row");
    assert.strictEqual(item.hasNps, false, "plain proportion table should not have NPS");
    assert.strictEqual(item.hasMeans, false, "plain proportion table should not have means");
  });

  it("item exposes hasNps=true for an NPS table", () => {
    const values = [
      ["Promoters", 50, 52, 51],
      ["Detractors", 20, 19, 18],
      ["NPS", 30, 33, 33],
      ["Base", 1000, 500, 500],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.ok(items.length >= 1, "expected at least one item");
    const item = items[0];
    assert.strictEqual(item.hasNps, true, "NPS table should have hasNps=true");
  });
});

// ─── preferredBase threading through inventory scanner ───────────────────────
//
// Observable: WEIGHTED_BASE_FALLBACK appears in qualityIssueCodes when weighted
// base is selected (either by preference or auto fallback when only weighted is
// available). It is absent when a non-weighted base is chosen.
//
// Table: effective + weighted bases available.

describe("scanWorksheetForTables — preferredBase setting threading", () => {
  const values = [
    ["Agree",          0.4, 0.6],
    ["Disagree",       0.6, 0.4],
    ["Base weighted",  200, 300],
    ["Base effective", 160, 250],
  ];

  it("no settings (omitted): auto picks effective — no WEIGHTED_BASE_FALLBACK", () => {
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.strictEqual(items.length, 1);
    assert.ok(
      !items[0].qualityIssueCodes.some((e) => e.code === "WEIGHTED_BASE_FALLBACK"),
      "auto should pick effective, not weighted"
    );
  });

  it("settings.preferredBase=auto: same as omitted — no WEIGHTED_BASE_FALLBACK", () => {
    const items = scanWorksheetForTables({ values, ...OFFSET, settings: { preferredBase: "auto" } });
    assert.strictEqual(items.length, 1);
    assert.ok(
      !items[0].qualityIssueCodes.some((e) => e.code === "WEIGHTED_BASE_FALLBACK"),
      "auto should pick effective, not weighted"
    );
  });

  it("settings.preferredBase=weighted: weighted base selected — WEIGHTED_BASE_FALLBACK present", () => {
    const items = scanWorksheetForTables({ values, ...OFFSET, settings: { preferredBase: "weighted" } });
    assert.strictEqual(items.length, 1);
    assert.ok(
      items[0].qualityIssueCodes.some((e) => e.code === "WEIGHTED_BASE_FALLBACK"),
      "WEIGHTED_BASE_FALLBACK should be present when weighted base is preferred"
    );
  });

  it("preferredBase=weighted: WEIGHTED_BASE_FALLBACK must not make candidateStatus uncertain", () => {
    // A table that is otherwise clean (valid base values, confident label split)
    // must remain 'available' even when WEIGHTED_BASE_FALLBACK is emitted.
    const items = scanWorksheetForTables({ values, ...OFFSET, settings: { preferredBase: "weighted" } });
    assert.strictEqual(items.length, 1);
    assert.strictEqual(
      items[0].candidateStatus,
      "available",
      `WEIGHTED_BASE_FALLBACK is advisory and must not downgrade candidateStatus; got ${items[0].candidateStatus}`
    );
  });

  it("real availability-affecting warning (SUSPICIOUS_ALL_100) still makes candidate uncertain", () => {
    // All-100 rows trigger SUSPICIOUS_ALL_100 — a non-advisory warning that should
    // still downgrade candidateStatus to uncertain.
    const allHundredValues = [
      ["Label1",       100, 100, 100],
      ["Label2",       100, 100, 100],
      ["Base weighted", 200, 300, 400],
      ["Base effective",160, 250, 350],
    ];
    const items = scanWorksheetForTables({ values: allHundredValues, ...OFFSET, settings: { preferredBase: "weighted" } });
    assert.ok(items.length >= 1, "expected at least one item");
    // warningsCount must reflect all warnings including advisory ones for display.
    assert.ok(items[0].warningsCount > 0, "warningsCount must include advisory warnings for display");
    assert.strictEqual(
      items[0].candidateStatus,
      "uncertain",
      "SUSPICIOUS_ALL_100 must still make candidateStatus uncertain"
    );
  });
});

// ─── Advisory vs availability-affecting warning codes ────────────────────────

describe("scanWorksheetForTables — advisory issue codes do not affect candidateStatus", () => {
  it("BASE_BELOW_THRESHOLD: stays in qualityIssueCodes but candidateStatus remains available", () => {
    // Base values all < 50 → BASE_BELOW_THRESHOLD fires when threshold = 50.
    // It must be visible in qualityIssueCodes but must not downgrade status.
    const values = [
      ["Agree",    0.4, 0.6, 0.5],
      ["Disagree", 0.6, 0.4, 0.5],
      ["Base",      30,  35,  40],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET, settings: { smallBaseThreshold: 50 } });
    assert.strictEqual(items.length, 1);
    assert.ok(
      items[0].qualityIssueCodes.some((e) => e.code === "BASE_BELOW_THRESHOLD"),
      "BASE_BELOW_THRESHOLD must be present in qualityIssueCodes"
    );
    assert.ok(items[0].warningsCount > 0, "warningsCount must count advisory warnings for display");
    assert.strictEqual(
      items[0].candidateStatus,
      "available",
      `BASE_BELOW_THRESHOLD is advisory and must not downgrade candidateStatus; got ${items[0].candidateStatus}`
    );
  });

  it("BASE_BLANK_VALUES: is NOT advisory and still makes candidateStatus uncertain", () => {
    // A base row with blank cells → BASE_BLANK_VALUES is a real data-quality warning
    // that genuinely signals the table may not calculate significance correctly.
    const values = [
      ["Agree",    0.4,  0.6, 0.5],
      ["Disagree", 0.6,  0.4, 0.5],
      ["Base",     100, null, 200],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.strictEqual(items.length, 1);
    assert.ok(
      items[0].qualityIssueCodes.some((e) => e.code === "BASE_BLANK_VALUES"),
      "BASE_BLANK_VALUES must be present in qualityIssueCodes"
    );
    assert.strictEqual(
      items[0].candidateStatus,
      "uncertain",
      "BASE_BLANK_VALUES must still downgrade candidateStatus to uncertain"
    );
  });
});

// ─── selectedBaseSubtypeLabel field ──────────────────────────────────────────

describe("scanWorksheetForTables — selectedBaseSubtypeLabel field", () => {
  it("plain 'Base' row → selectedBaseSubtypeLabel is 'Base'", () => {
    const values = [
      ["Agree",    0.4, 0.6],
      ["Disagree", 0.6, 0.4],
      ["Base",     100, 200],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.strictEqual(items.length, 1);
    assert.strictEqual(items[0].selectedBaseSubtypeLabel, "Base");
  });

  it("'Base effective' row → selectedBaseSubtypeLabel is 'Effective Base'", () => {
    const values = [
      ["Agree",          0.4, 0.6],
      ["Disagree",       0.6, 0.4],
      ["Base effective", 100, 200],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.strictEqual(items.length, 1);
    assert.strictEqual(items[0].selectedBaseSubtypeLabel, "Effective Base");
  });

  it("'Base unweighted' row → selectedBaseSubtypeLabel is 'Unweighted Base'", () => {
    const values = [
      ["Agree",           0.4, 0.6],
      ["Disagree",        0.6, 0.4],
      ["Base unweighted", 100, 200],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.strictEqual(items.length, 1);
    assert.strictEqual(items[0].selectedBaseSubtypeLabel, "Unweighted Base");
  });

  it("'Base weighted' row → selectedBaseSubtypeLabel is 'Weighted Base'", () => {
    const values = [
      ["Agree",         0.4, 0.6],
      ["Disagree",      0.6, 0.4],
      ["Base weighted", 100, 200],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    assert.strictEqual(items.length, 1);
    assert.strictEqual(items[0].selectedBaseSubtypeLabel, "Weighted Base");
  });

  it("preferredBase=effective with effective+weighted available → 'Effective Base'", () => {
    const values = [
      ["Agree",          0.4, 0.6],
      ["Disagree",       0.6, 0.4],
      ["Base weighted",  200, 300],
      ["Base effective", 160, 250],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET, settings: { preferredBase: "effective" } });
    assert.strictEqual(items.length, 1);
    assert.strictEqual(items[0].selectedBaseSubtypeLabel, "Effective Base");
  });

  it("preferredBase=weighted with effective+weighted available → 'Weighted Base'", () => {
    const values = [
      ["Agree",          0.4, 0.6],
      ["Disagree",       0.6, 0.4],
      ["Base weighted",  200, 300],
      ["Base effective", 160, 250],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET, settings: { preferredBase: "weighted" } });
    assert.strictEqual(items.length, 1);
    assert.strictEqual(items[0].selectedBaseSubtypeLabel, "Weighted Base");
  });

  it("no calculation blocks (proportion rows only, no base) → selectedBaseSubtypeLabel is ''", () => {
    // Without a base row no blocks can be formed so the label must be empty.
    const values = [
      ["Agree",    0.4, 0.6],
      ["Disagree", 0.6, 0.4],
    ];
    const items = scanWorksheetForTables({ values, ...OFFSET });
    // No base row — scanner may or may not produce a candidate, but if it does
    // selectedBaseSubtypeLabel must be empty string.
    for (const item of items) {
      assert.strictEqual(
        item.selectedBaseSubtypeLabel,
        "",
        "no-base table must have empty selectedBaseSubtypeLabel"
      );
    }
  });
});
