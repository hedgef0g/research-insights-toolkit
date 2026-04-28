/**
 * Default z-threshold for a two-tailed 95% significance test.
 *
 * TODO:
 * In future versions, make confidence level configurable:
 * 90%  -> 1.645
 * 95%  -> 1.960
 * 99%  -> 2.576
 */
export const DEFAULT_Z_THRESHOLD_95 = 1.96;

/**
 * Converts a user-entered percentage/share into a decimal proportion.
 *
 * PURPOSE:
 * Users may enter values either as:
 * - 0.42 meaning 42%
 * - 42 meaning 42%
 *
 * INPUT:
 * rawValue - any value from a spreadsheet cell.
 *
 * OUTPUT:
 * Number between 0 and 1, or null if the value cannot be used.
 */
export function normalizeShare(rawValue) {
  if (rawValue === null || rawValue === undefined || rawValue === "") {
    return null;
  }

  const textValue = String(rawValue).trim(); // Spreadsheet value as text.

  const isPercentText = textValue.endsWith("%"); // Example: "42%".

  const cleanedTextValue = isPercentText
    ? textValue.replace("%", "").trim()
    : textValue;

  const numericValue = Number(cleanedTextValue.replace(",", "."));

  if (Number.isNaN(numericValue)) {
    return null;
  }

  if (isPercentText) {
    return numericValue / 100;
  }

  return numericValue > 1 ? numericValue / 100 : numericValue;
}

/**
 * Calculates statistical significance between two proportions.
 *
 * PURPOSE:
 * Compare two percentage values using a pooled z-test for proportions.
 *
 * INPUT:
 * firstRawValue  - first percentage/share value.
 * firstRawBase   - base size for the first value.
 * secondRawValue - second percentage/share value.
 * secondRawBase  - base size for the second value.
 *
 * OUTPUT:
 * Object with calculation result, or null if input is invalid.
 *
 * MVP LIMITATIONS:
 * - Assumes values are independent proportions.
 * - Assumes two-tailed 95% test.
 * - Does not yet support means, NPS, weighted bases, or multiple columns.
 */
export function calculateProportionSignificance(
  firstRawValue,
  firstRawBase,
  secondRawValue,
  secondRawBase
) {
  const firstProportion = normalizeShare(firstRawValue); // First value converted to 0–1 proportion.
  const secondProportion = normalizeShare(secondRawValue); // Second value converted to 0–1 proportion.

  const firstBase = Number(firstRawBase); // Number of respondents behind the first value.
  const secondBase = Number(secondRawBase); // Number of respondents behind the second value.

  if (firstProportion === null || secondProportion === null) {
    return null;
  }

  if (firstBase <= 0 || secondBase <= 0) {
    return null;
  }

  if (
    firstProportion < 0 ||
    firstProportion > 1 ||
    secondProportion < 0 ||
    secondProportion > 1
  ) {
    return null;
  }

  // Pooled proportion is used because the test checks whether
  // both values may come from the same underlying population proportion.
  const pooledProportion =
    (firstProportion * firstBase + secondProportion * secondBase) /
    (firstBase + secondBase);

  // Standard error for the difference between two proportions.
  const standardError = Math.sqrt(
    pooledProportion *
      (1 - pooledProportion) *
      (1 / firstBase + 1 / secondBase)
  );

  if (standardError === 0) {
    return null;
  }

  // z-score shows how many standard errors separate the two proportions.
  const zScore = (firstProportion - secondProportion) / standardError;

  // Absolute z-score is used for a two-tailed test.
  const absoluteZScore = Math.abs(zScore);

  // Difference is significant if absolute z-score reaches 95% threshold.
  const isSignificant = absoluteZScore >= DEFAULT_Z_THRESHOLD_95;

  // Direction will later help us decide where to place visual markers.
  const direction =
    zScore > 0 ? "first_higher" : zScore < 0 ? "second_higher" : "equal";

  return {
    firstProportion,
    secondProportion,
    firstBase,
    secondBase,
    pooledProportion,
    standardError,
    zScore,
    absoluteZScore,
    isSignificant,
    direction,
    confidenceLevel: 0.95,
    zThreshold: DEFAULT_Z_THRESHOLD_95,
  };
}

/**
 * Compares all column pairs inside one row of proportions.
 *
 * PURPOSE:
 * For one row of values, compare every column with every other column.
 * Example: columns A, B, C produce comparisons:
 * A vs B, A vs C, B vs C.
 *
 * INPUT:
 * valueRow - array of values from one spreadsheet row.
 * baseRow  - array of bases, usually from the bottom row of selected range.
 *
 * OUTPUT:
 * Array of comparison result objects.
 */
export function compareAllProportionsInRow(valueRow, baseRow) {
  const rowComparisons = []; // Stores all pairwise comparisons for this row.

  for (let firstColumnIndex = 0; firstColumnIndex < valueRow.length; firstColumnIndex++) {
    for (
      let secondColumnIndex = firstColumnIndex + 1;
      secondColumnIndex < valueRow.length;
      secondColumnIndex++
    ) {
      const firstValue = valueRow[firstColumnIndex]; // Value from the first compared column.
      const secondValue = valueRow[secondColumnIndex]; // Value from the second compared column.

      const firstBase = baseRow[firstColumnIndex]; // Base for the first compared column.
      const secondBase = baseRow[secondColumnIndex]; // Base for the second compared column.

      const significanceResult = calculateProportionSignificance(
        firstValue,
        firstBase,
        secondValue,
        secondBase
      );

      rowComparisons.push({
        firstColumnIndex,
        secondColumnIndex,
        firstValue,
        secondValue,
        firstBase,
        secondBase,
        result: significanceResult,
      });
    }
  }

  return rowComparisons;
}

/**
 * Compares all rows of selected table data using the last row as bases.
 *
 * PURPOSE:
 * This is the main parser/calculation function for MVP v0.2.
 * It assumes:
 * - selected range contains multiple columns;
 * - last row contains bases;
 * - all rows above the last row contain values;
 * - values are proportions/percentages;
 * - each value row is compared column-by-column, all pairs.
 *
 * INPUT:
 * selectedValues - 2D array from Excel or Google Sheets.
 *
 * OUTPUT:
 * Object with base row and comparison results for every value row.
 *
 * MVP LIMITATIONS:
 * - Bases must be in the last selected row.
 * - Does not yet auto-detect weighted bases.
 * - Does not yet use text labels.
 * - Does not yet write significance markers back to the table.
 */
export function compareAllRowsUsingBottomBases(selectedValues) {
  if (!selectedValues || selectedValues.length < 2) {
    return null;
  }

  const totalRows = selectedValues.length; // Number of selected spreadsheet rows.
  const baseRowIndex = totalRows - 1; // Last row is treated as the base row.
  const baseRow = selectedValues[baseRowIndex]; // Bases for all columns.

  const comparisonRows = []; // Stores results for each value row.

  for (let valueRowIndex = 0; valueRowIndex < baseRowIndex; valueRowIndex++) {
    const valueRow = selectedValues[valueRowIndex]; // Current row of percentages/shares.

    const rowComparisons = compareAllProportionsInRow(valueRow, baseRow);

    comparisonRows.push({
      valueRowIndex,
      rowComparisons,
    });
  }

  return {
    baseRowIndex,
    baseRow,
    comparisonRows,
  };
}

/**
 * Generates column significance labels.
 *
 * PURPOSE:
 * Create readable column labels for significance notation:
 * a-z, A-Z, а-я, А-Я.
 *
 * IMPORTANT:
 * Excludes:
 * - Latin t / T
 * - Cyrillic т / Т
 *
 * OUTPUT:
 * Array of labels.
 */
export function generateSignificanceLabels() {
  const latinLowercaseLabels = "abcdefghijklmnopqrsuvwxyz".split(""); // Latin a-z without t.
  const latinUppercaseLabels = "ABCDEFGHIJKLMNOPQRSUVWXYZ".split(""); // Latin A-Z without T.

  const cyrillicLowercaseLabels = "абвгдежзийклмнопрсуфхцчшщъыьэюя".split(""); // Cyrillic а-я without т.
  const cyrillicUppercaseLabels = "АБВГДЕЖЗИЙКЛМНОПРСУФХЦЧШЩЪЫЬЭЮЯ".split(""); // Cyrillic А-Я without Т.

  return [
    ...latinLowercaseLabels,
    ...latinUppercaseLabels,
    ...cyrillicLowercaseLabels,
    ...cyrillicUppercaseLabels,
  ];
}

/**
 * Builds empty marker storage for each value cell.
 *
 * PURPOSE:
 * For every value row and column, prepare a place where significance letters
 * will be collected before writing them back to Excel.
 */
export function createEmptyMarkerMatrix(rowCount, columnCount) {
  const markerMatrix = []; // 2D array: row -> column -> marker letters.

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
    const markerRow = []; // Marker letters for one value row.

    for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      markerRow.push(""); // Empty string means no significance markers yet.
    }

    markerMatrix.push(markerRow);
  }

  return markerMatrix;
}

/**
 * Converts pairwise comparison results into cell markers.
 *
 * PURPOSE:
 * If one column is significantly higher than another column,
 * add the lower column's label to the higher value cell.
 *
 * EXAMPLE:
 * If column 1 is significantly higher than column 2,
 * column 1 receives marker "b".
 *
 * INPUT:
 * allResults - result of compareAllRowsUsingBottomBases().
 *
 * OUTPUT:
 * 2D marker matrix for value rows only.
 */
export function buildSignificanceMarkerMatrix(allResults) {
  const valueRowCount = allResults.baseRowIndex; // Number of rows above the base row.
  const columnCount = allResults.baseRow.length; // Number of selected columns.

  const significanceLabels = generateSignificanceLabels(); // Labels assigned to selected columns.
  const markerMatrix = createEmptyMarkerMatrix(valueRowCount, columnCount); // Output marker storage.

  for (const comparisonRow of allResults.comparisonRows) {
    const valueRowIndex = comparisonRow.valueRowIndex; // Row where markers will be applied.

    for (const comparison of comparisonRow.rowComparisons) {
      if (comparison.result === null) {
        continue;
      }

      if (!comparison.result.isSignificant) {
        continue;
      }

      const firstColumnIndex = comparison.firstColumnIndex; // First compared column.
      const secondColumnIndex = comparison.secondColumnIndex; // Second compared column.

      const firstColumnLabel = significanceLabels[firstColumnIndex]; // Label for first column.
      const secondColumnLabel = significanceLabels[secondColumnIndex]; // Label for second column.

      if (comparison.result.direction === "first_higher") {
        markerMatrix[valueRowIndex][firstColumnIndex] += secondColumnLabel;
      }

      if (comparison.result.direction === "second_higher") {
        markerMatrix[valueRowIndex][secondColumnIndex] += firstColumnLabel;
      }
    }
  }

  return markerMatrix;
}

/**
 * Removes significance marker letters from the end of a cell text.
 *
 * PURPOSE:
 * If a previous macro run changed "42bC" into a marked value,
 * this function restores visible value text back to "42".
 *
 * IMPORTANT:
 * This only removes marker characters from the END of the text.
 */
export function removeSignificanceMarkersFromText(rawText) {
  if (rawText === null || rawText === undefined) {
    return rawText;
  }

  const textValue = String(rawText); // Cell value converted to text.

  const markerCharacters =
    "abcdefghijklmnopqrsuvwxyz" +
    "ABCDEFGHIJKLMNOPQRSUVWXYZ" +
    "абвгдежзийклмнопрсуфхцчшщъыьэюя" +
    "АБВГДЕЖЗИЙКЛМНОПРСУФХЦЧШЩЪЫЬЭЮЯ";

  const markerSuffixPattern = new RegExp(`[${markerCharacters}]+$`);

  return textValue.replace(markerSuffixPattern, "");
}

/**
 * Removes significance markers from a 2D spreadsheet values array.
 *
 * PURPOSE:
 * Clean the selected range before recalculating significance.
 */
export function removeSignificanceMarkersFromMatrix(valuesMatrix) {
  return valuesMatrix.map((row) =>
    row.map((cellValue) => removeSignificanceMarkersFromText(cellValue))
  );
}