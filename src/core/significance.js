import {
  normalizeShare,
  normalizeVariance,
  normalizeNpsValue,
  normalizeNpsSpread,
} from "./normalizers";

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
  secondRawBase,
  calculationSettings = { confidenceLevel: "95" }
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

  if (firstProportion < 0 || firstProportion > 1 || secondProportion < 0 || secondProportion > 1) {
    return null;
  }

  // Pooled proportion is used because the test checks whether
  // both values may come from the same underlying population proportion.
  const pooledProportion =
    (firstProportion * firstBase + secondProportion * secondBase) / (firstBase + secondBase);

  // Standard error for the difference between two proportions.
  const standardError = Math.sqrt(
    pooledProportion * (1 - pooledProportion) * (1 / firstBase + 1 / secondBase)
  );

  if (standardError === 0) {
    return null;
  }

  // z-score shows how many standard errors separate the two proportions.
  const zScore = (firstProportion - secondProportion) / standardError;

  // Absolute z-score is used for a two-tailed test.
  const absoluteZScore = Math.abs(zScore);

  const confidenceLevel = calculationSettings.confidenceLevel;
  const zThreshold = getZThresholdForConfidence(confidenceLevel);

  // Difference is significant if absolute z-score reaches 95% threshold.
  const isSignificant = absoluteZScore >= zThreshold;

  // Direction will later help us decide where to place visual markers.
  const direction = zScore > 0 ? "first_higher" : zScore < 0 ? "second_higher" : "equal";

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
    confidenceLevel,
    zThreshold,
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
export function compareAllProportionsInRow(
  valueRow,
  baseRow,
  calculationSettings = { confidenceLevel: "95" }
) {
  const rowComparisons = [];

  const comparisonPairs = buildColumnComparisonPairs(
    valueRow.length,
    calculationSettings,
    calculationSettings.excludedColumnIndexes || new Set()
  );

  for (const comparisonPair of comparisonPairs) {
    const firstColumnIndex = comparisonPair.firstColumnIndex;
    const secondColumnIndex = comparisonPair.secondColumnIndex;

    const firstValue = valueRow[firstColumnIndex];
    const secondValue = valueRow[secondColumnIndex];

    const firstBase = baseRow[firstColumnIndex];
    const secondBase = baseRow[secondColumnIndex];

    const significanceResult = calculateProportionSignificance(
      firstValue,
      firstBase,
      secondValue,
      secondBase,
      calculationSettings
    );

    rowComparisons.push({
      firstColumnIndex,
      secondColumnIndex,
      comparisonType: comparisonPair.comparisonType,
      firstValue,
      secondValue,
      firstBase,
      secondBase,
      result: significanceResult,
    });
  }

  return rowComparisons;
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
 * Builds column comparison pairs based on current comparison settings.
 *
 * MODES:
 * - Default: all columns are compared pairwise.
 * - firstColumnIsTotal:
 *   Total is compared with every segment, and segments are also compared pairwise.
 * - firstColumnIsTotal + compareOnlyWithTotal:
 *   only Total-vs-segment comparisons are performed.
 * - firstColumnIsTotal + excludeTotalFromComparisons:
 *   only segment-vs-segment comparisons are performed.
 *
 * excludedColumnIndexes:
 * Columns excluded before calculation, for example because of small base.
 */
export function buildColumnComparisonPairs(
  columnCount,
  calculationSettings = {},
  excludedColumnIndexes = new Set()
) {
  const pairs = [];

  if (columnCount < 2) {
    return pairs;
  }

  const firstColumnIsTotal = calculationSettings.firstColumnIsTotal;
  const compareOnlyWithTotal = calculationSettings.compareOnlyWithTotal;
  const excludeTotalFromComparisons = calculationSettings.excludeTotalFromComparisons;
  const compareWithPreviousColumn = calculationSettings.compareWithPreviousColumn;

  const isExcluded = (columnIndex) => excludedColumnIndexes.has(columnIndex);

  if (compareWithPreviousColumn) {
    const startColumnIndex = firstColumnIsTotal && excludeTotalFromComparisons ? 2 : 1;

    for (let columnIndex = startColumnIndex; columnIndex < columnCount; columnIndex++) {
      const previousColumnIndex = columnIndex - 1;

      if (isExcluded(previousColumnIndex) || isExcluded(columnIndex)) {
        continue;
      }

      pairs.push({
        firstColumnIndex: previousColumnIndex,
        secondColumnIndex: columnIndex,
        comparisonType: "previousColumn",
      });
    }

    return pairs;
  }

  if (firstColumnIsTotal) {
    if (!excludeTotalFromComparisons && !isExcluded(0)) {
      for (let columnIndex = 1; columnIndex < columnCount; columnIndex++) {
        if (isExcluded(columnIndex)) {
          continue;
        }

        pairs.push({
          firstColumnIndex: 0,
          secondColumnIndex: columnIndex,
          comparisonType: "total",
        });
      }
    }

    if (compareOnlyWithTotal) {
      return pairs;
    }

    for (let firstColumnIndex = 1; firstColumnIndex < columnCount; firstColumnIndex++) {
      if (isExcluded(firstColumnIndex)) {
        continue;
      }

      for (
        let secondColumnIndex = firstColumnIndex + 1;
        secondColumnIndex < columnCount;
        secondColumnIndex++
      ) {
        if (isExcluded(secondColumnIndex)) {
          continue;
        }

        pairs.push({
          firstColumnIndex,
          secondColumnIndex,
          comparisonType: "segment",
        });
      }
    }

    return pairs;
  }

  for (let firstColumnIndex = 0; firstColumnIndex < columnCount; firstColumnIndex++) {
    if (isExcluded(firstColumnIndex)) {
      continue;
    }

    for (
      let secondColumnIndex = firstColumnIndex + 1;
      secondColumnIndex < columnCount;
      secondColumnIndex++
    ) {
      if (isExcluded(secondColumnIndex)) {
        continue;
      }

      pairs.push({
        firstColumnIndex,
        secondColumnIndex,
        comparisonType: "segment",
      });
    }
  }

  return pairs;
}

/**
 * Returns visible significance label for a selected column.
 *
 * In firstColumnIsTotal mode:
 * - column 0 has no label;
 * - column 1 gets "a";
 * - column 2 gets "b";
 * - etc.
 */
export function getSignificanceLabelForColumnIndex(columnIndex, calculationSettings = {}) {
  const labels = generateSignificanceLabels();

  if (calculationSettings.firstColumnIsTotal) {
    if (columnIndex === 0) {
      return "";
    }

    return labels[columnIndex - 1] || "";
  }

  return labels[columnIndex] || "";
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

export const CELL_FILL_REASONS = {
  NONE: "none",
  SIGNIFICANT: "significant",
  LOWER_THAN_TOTAL: "lowerThanTotal",
  SMALL_BASE: "smallBase",
};

export const CELL_FILL_PRIORITIES = {
  [CELL_FILL_REASONS.NONE]: 0,
  [CELL_FILL_REASONS.SIGNIFICANT]: 10,
  [CELL_FILL_REASONS.LOWER_THAN_TOTAL]: 20,
  [CELL_FILL_REASONS.SMALL_BASE]: 30,
};

/**
 * Builds empty cell result storage for each selected cell.
 *
 * PURPOSE:
 * Stores both visible significance markers and formatting reasons.
 * This allows fill priority logic:
 * small base > lower than Total > normal significance > none.
 */
export function createEmptyCellResultMatrix(rowCount, columnCount) {
  const cellResultMatrix = [];

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
    const cellResultRow = [];

    for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      cellResultRow.push({
        markers: "",
        fillReason: CELL_FILL_REASONS.NONE,
        fillPriority: CELL_FILL_PRIORITIES[CELL_FILL_REASONS.NONE],
        hasPositiveTotalComparison: false,
        previousColumnArrow: "",
        previousColumnArrowDirection: "",
      });
    }

    cellResultMatrix.push(cellResultRow);
  }

  return cellResultMatrix;
}

/**
 * Applies fill reason only if it has higher priority than current fill.
 */
function applyFillReasonToCellResult(cellResult, fillReason) {
  const nextPriority = CELL_FILL_PRIORITIES[fillReason];

  if (nextPriority === undefined) {
    return;
  }

  if (nextPriority > cellResult.fillPriority) {
    cellResult.fillReason = fillReason;
    cellResult.fillPriority = nextPriority;
  }
}

/**
 * Appends ordinary segment marker and applies normal significance fill.
 */
function appendSignificanceMarkerToCellResult(cellResult, marker) {
  if (!marker) {
    return;
  }

  cellResult.markers += marker;
  applyFillReasonToCellResult(cellResult, CELL_FILL_REASONS.SIGNIFICANT);
}

/**
 * Prepends Total marker and applies corresponding fill.
 *
 * T = segment is significantly higher than Total.
 * t = segment is significantly lower than Total.
 */
function prependTotalMarkerToCellResult(cellResult, totalMarker) {
  const markerText = cellResult.markers || "";
  const markerTextWithoutOldTotalMarker = markerText.replace(/[tT]/g, "");

  cellResult.markers = `${totalMarker}${markerTextWithoutOldTotalMarker}`;

  if (totalMarker === "T") {
    cellResult.hasPositiveTotalComparison = true;
    applyFillReasonToCellResult(cellResult, CELL_FILL_REASONS.SIGNIFICANT);
    return;
  }

  cellResult.hasPositiveTotalComparison = false;
  applyFillReasonToCellResult(cellResult, CELL_FILL_REASONS.LOWER_THAN_TOTAL);
}

/**
 * Converts pairwise comparison results into cell markers.
 *
 * PURPOSE:
 * If one column is significantly higher than another column,
 * add the lower column's label to the higher value cell.
 *
 * INPUT:
 * allResults - calculation result object with comparisonRows.
 * markerRowCount - how many rows should receive markers.
 *
 * WHY markerRowCount EXISTS:
 * For proportions, all rows above the base row may receive markers.
 * For means and NPS, only the first row should receive markers.
 */
export function buildSignificanceMarkerMatrix(allResults, markerRowCount = null) {
  const valueRowCount = markerRowCount === null ? allResults.baseRowIndex : markerRowCount;

  const columnCount = allResults.baseRow.length; // Number of selected columns.

  const significanceLabels = generateSignificanceLabels(); // Labels assigned to selected columns.
  const markerMatrix = createEmptyMarkerMatrix(valueRowCount, columnCount); // Output marker storage.

  for (const comparisonRow of allResults.comparisonRows) {
    const valueRowIndex = comparisonRow.valueRowIndex; // Row where markers should be applied.

    // Skip rows that are not intended to receive markers.
    if (valueRowIndex >= valueRowCount) {
      continue;
    }

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
    "abcdefghijklmnopqrstuvwxyz" +
    "ABCDEFGHIJKLMNOPQRSTUVWXYZ" +
    "абвгдежзийклмнопрсуфхцчшщъыьэюя" +
    "АБВГДЕЖЗИЙКЛМНОПРСУФХЦЧШЩЪЫЬЭЮЯ";

  const markerSuffixPattern = new RegExp(`\\s*[${markerCharacters}↑↓]+$`);

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

/**
 * Default confidence level for two-tailed significance tests.
 */
export const DEFAULT_CONFIDENCE_LEVEL = 0.95;

/**
 * Calculates statistical significance between two means using Welch's t-test.
 *
 * PURPOSE:
 * Compare two means where bases and either standard deviations or variances are known.
 *
 * WHY WELCH:
 * Welch's t-test is safer than Student's pooled t-test because it does not assume
 * equal variances between groups.
 *
 * INPUT:
 * firstRawMean      - first mean value.
 * firstRawSpread    - first SD or variance.
 * firstRawBase      - first base size.
 * secondRawMean     - second mean value.
 * secondRawSpread   - second SD or variance.
 * secondRawBase     - second base size.
 * spreadType        - "standardDeviation" or "variance".
 *
 * OUTPUT:
 * Object with calculation result, or null if input is invalid.
 */
export function calculateMeanSignificance(
  firstRawMean,
  firstRawSpread,
  firstRawBase,
  secondRawMean,
  secondRawSpread,
  secondRawBase,
  spreadType,
  calculationSettings = { confidenceLevel: "95" }
) {
  const firstMean = Number(String(firstRawMean).replace(",", ".")); // First mean value.
  const secondMean = Number(String(secondRawMean).replace(",", ".")); // Second mean value.

  const firstBase = Number(firstRawBase); // Base size for first mean.
  const secondBase = Number(secondRawBase); // Base size for second mean.

  const firstVariance = normalizeVariance(firstRawSpread, spreadType); // First variance.
  const secondVariance = normalizeVariance(secondRawSpread, spreadType); // Second variance.

  if (Number.isNaN(firstMean) || Number.isNaN(secondMean)) {
    return null;
  }

  if (firstBase <= 1 || secondBase <= 1) {
    return null;
  }

  if (firstVariance === null || secondVariance === null) {
    return null;
  }

  const firstStandardErrorPart = firstVariance / firstBase; // Variance contribution from first group.
  const secondStandardErrorPart = secondVariance / secondBase; // Variance contribution from second group.

  const standardError = Math.sqrt(firstStandardErrorPart + secondStandardErrorPart);

  if (standardError === 0) {
    return null;
  }

  const tScore = (firstMean - secondMean) / standardError; // Welch t-score.

  const degreesOfFreedomNumerator = (firstStandardErrorPart + secondStandardErrorPart) ** 2;

  const degreesOfFreedomDenominator =
    firstStandardErrorPart ** 2 / (firstBase - 1) + secondStandardErrorPart ** 2 / (secondBase - 1);

  const degreesOfFreedom = degreesOfFreedomNumerator / degreesOfFreedomDenominator;

  const confidenceLevel = calculationSettings.confidenceLevel;
  const tThreshold = getTThresholdForConfidence(confidenceLevel, degreesOfFreedom);
  const absoluteTScore = Math.abs(tScore); // Two-tailed comparison uses absolute t.

  const isSignificant = absoluteTScore >= tThreshold;

  const direction = tScore > 0 ? "first_higher" : tScore < 0 ? "second_higher" : "equal";

  return {
    firstMean,
    secondMean,
    firstVariance,
    secondVariance,
    firstBase,
    secondBase,
    standardError,
    tScore,
    absoluteTScore,
    degreesOfFreedom,
    tThreshold,
    isSignificant,
    direction,
    confidenceLevel,
    spreadType,
  };
}

/**
 * Compares all column pairs inside one row of means.
 *
 * PURPOSE:
 * For one row of means, compare every column with every other column.
 */
export function compareAllMeansInRow(
  meanRow,
  spreadRow,
  baseRow,
  spreadType,
  calculationSettings = { confidenceLevel: "95" }
) {
  const rowComparisons = [];

  const comparisonPairs = buildColumnComparisonPairs(
    meanRow.length,
    calculationSettings,
    calculationSettings.excludedColumnIndexes || new Set()
  );

  for (const comparisonPair of comparisonPairs) {
    const firstColumnIndex = comparisonPair.firstColumnIndex;
    const secondColumnIndex = comparisonPair.secondColumnIndex;

    const significanceResult = calculateMeanSignificance(
      meanRow[firstColumnIndex],
      spreadRow[firstColumnIndex],
      baseRow[firstColumnIndex],
      meanRow[secondColumnIndex],
      spreadRow[secondColumnIndex],
      baseRow[secondColumnIndex],
      spreadType,
      calculationSettings
    );

    rowComparisons.push({
      firstColumnIndex,
      secondColumnIndex,
      comparisonType: comparisonPair.comparisonType,
      firstValue: meanRow[firstColumnIndex],
      secondValue: meanRow[secondColumnIndex],
      firstSpread: spreadRow[firstColumnIndex],
      secondSpread: spreadRow[secondColumnIndex],
      firstBase: baseRow[firstColumnIndex],
      secondBase: baseRow[secondColumnIndex],
      result: significanceResult,
    });
  }

  return rowComparisons;
}

/**
 * Calculates significance between two NPS values using promoter/detractor structure.
 *
 * IMPORTANT:
 * We do NOT trust the displayed NPS row for calculation.
 * Instead, NPS is recalculated from:
 * NPS = promoters share - detractors share
 *
 * NPS is treated as mean of:
 * promoter = +1
 * passive = 0
 * detractor = -1
 *
 * Variance = P(promoter) + P(detractor) - NPS²
 */
export function calculateNpsSignificanceFromStructure(
  firstRawNps,
  firstRawPromoters,
  firstRawDetractors,
  firstRawBase,
  secondRawNps,
  secondRawPromoters,
  secondRawDetractors,
  secondRawBase,
  calculationSettings = { confidenceLevel: "95" }
) {
  const firstPromoters = normalizeShare(firstRawPromoters); // First promoter share.
  const secondPromoters = normalizeShare(secondRawPromoters); // Second promoter share.

  const firstDetractors = normalizeShare(firstRawDetractors); // First detractor share.
  const secondDetractors = normalizeShare(secondRawDetractors); // Second detractor share.

  const firstBase = Number(firstRawBase); // First base size.
  const secondBase = Number(secondRawBase); // Second base size.

  if (
    firstPromoters === null ||
    secondPromoters === null ||
    firstDetractors === null ||
    secondDetractors === null
  ) {
    return null;
  }

  if (firstBase <= 1 || secondBase <= 1) {
    return null;
  }

  if (
    firstPromoters < 0 ||
    firstPromoters > 1 ||
    secondPromoters < 0 ||
    secondPromoters > 1 ||
    firstDetractors < 0 ||
    firstDetractors > 1 ||
    secondDetractors < 0 ||
    secondDetractors > 1
  ) {
    return null;
  }

  // Promoters + detractors cannot logically exceed 100%.
  if (firstPromoters + firstDetractors > 1 || secondPromoters + secondDetractors > 1) {
    return null;
  }

  // Recalculate NPS from structure instead of trusting the visible NPS row.
  const firstNps = firstPromoters - firstDetractors;
  const secondNps = secondPromoters - secondDetractors;

  // Variance of NPS score where promoter = +1, passive = 0, detractor = -1.
  const firstVariance = firstPromoters + firstDetractors - firstNps * firstNps;

  const secondVariance = secondPromoters + secondDetractors - secondNps * secondNps;

  if (firstVariance < 0 || secondVariance < 0) {
    return null;
  }

  return calculateMeanSignificance(
    firstNps,
    firstVariance,
    firstBase,
    secondNps,
    secondVariance,
    secondBase,
    "variance",
    calculationSettings
  );
}

export function calculateNpsSignificanceFromSpread(
  firstRawNps,
  firstRawSpread,
  firstRawBase,
  secondRawNps,
  secondRawSpread,
  secondRawBase,
  spreadType,
  calculationSettings = { confidenceLevel: "95" }
) {
  const firstNps = normalizeNpsValue(firstRawNps); // First NPS on -1..1 scale.
  const secondNps = normalizeNpsValue(secondRawNps); // Second NPS on -1..1 scale.

  const firstSpread = normalizeNpsSpread(firstRawSpread, spreadType); // First SD/variance on -1..1 scale.
  const secondSpread = normalizeNpsSpread(secondRawSpread, spreadType); // Second SD/variance on -1..1 scale.

  if (firstNps === null || secondNps === null || firstSpread === null || secondSpread === null) {
    return null;
  }

  return calculateMeanSignificance(
    firstNps,
    firstSpread,
    firstRawBase,
    secondNps,
    secondSpread,
    secondRawBase,
    spreadType,
    calculationSettings
  );
}

/**
 * Compares selected proportion rows against a specific base row.
 *
 * PURPOSE:
 * Complex tables may contain proportions, means, and NPS in one range.
 * This function calculates only selected proportion value rows.
 */
export function compareProportionRowsUsingBaseRow(
  selectedValues,
  valueRowIndexes,
  baseRowIndex,
  calculationSettings = { confidenceLevel: "95" }
) {
  const baseRow = selectedValues[baseRowIndex]; // Base row for these proportion rows.
  const comparisonRows = []; // Results for each proportion row.

  for (const valueRowIndex of valueRowIndexes) {
    const valueRow = selectedValues[valueRowIndex]; // Current proportion row.

    const rowComparisons = compareAllProportionsInRow(valueRow, baseRow, calculationSettings);

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
 * Compares one mean row using explicitly provided spread and base rows.
 *
 * PURPOSE:
 * Supports complex tables where mean block can be located anywhere
 * inside selected range.
 */
export function compareMeanBlockByRowIndexes(
  selectedValues,
  valueRowIndex,
  spreadRowIndex,
  baseRowIndex,
  spreadType,
  calculationSettings = { confidenceLevel: "95" }
) {
  const meanRow = selectedValues[valueRowIndex]; // Mean values.
  const spreadRow = selectedValues[spreadRowIndex]; // SD or variance values.
  const baseRow = selectedValues[baseRowIndex]; // Bases.

  const rowComparisons = compareAllMeansInRow(
    meanRow,
    spreadRow,
    baseRow,
    spreadType,
    calculationSettings
  );

  return {
    baseRowIndex,
    baseRow,
    comparisonRows: [
      {
        valueRowIndex,
        rowComparisons,
      },
    ],
  };
}

/**
 * Compares one NPS structure block using explicitly provided rows.
 *
 * PURPOSE:
 * Supports complex tables where NPS structure block can be located anywhere.
 */
export function compareNpsStructureBlockByRowIndexes(
  selectedValues,
  valueRowIndex,
  promotersRowIndex,
  detractorsRowIndex,
  baseRowIndex,
  calculationSettings = { confidenceLevel: "95" }
) {
  const npsRow = selectedValues[valueRowIndex]; // Visible NPS row.
  const promotersRow = selectedValues[promotersRowIndex]; // Promoter shares.
  const detractorsRow = selectedValues[detractorsRowIndex]; // Detractor shares.
  const baseRow = selectedValues[baseRowIndex]; // Bases.

  const rowComparisons = []; // Pairwise NPS comparisons.

  const comparisonPairs = buildColumnComparisonPairs(
    npsRow.length,
    calculationSettings,
    calculationSettings.excludedColumnIndexes || new Set()
  );

  for (const comparisonPair of comparisonPairs) {
    const firstColumnIndex = comparisonPair.firstColumnIndex;
    const secondColumnIndex = comparisonPair.secondColumnIndex;

    const significanceResult = calculateNpsSignificanceFromStructure(
      npsRow[firstColumnIndex],
      promotersRow[firstColumnIndex],
      detractorsRow[firstColumnIndex],
      baseRow[firstColumnIndex],
      npsRow[secondColumnIndex],
      promotersRow[secondColumnIndex],
      detractorsRow[secondColumnIndex],
      baseRow[secondColumnIndex],
      calculationSettings
    );

    rowComparisons.push({
      firstColumnIndex,
      secondColumnIndex,
      comparisonType: comparisonPair.comparisonType,
      result: significanceResult,
    });
  }

  return {
    baseRowIndex,
    baseRow,
    comparisonRows: [
      {
        valueRowIndex,
        rowComparisons,
      },
    ],
  };
}

/**
 * Compares one NPS spread block using explicitly provided rows.
 *
 * PURPOSE:
 * Supports complex tables where NPS spread block can be located anywhere.
 */
export function compareNpsSpreadBlockByRowIndexes(
  selectedValues,
  valueRowIndex,
  spreadRowIndex,
  baseRowIndex,
  spreadType,
  calculationSettings = { confidenceLevel: "95" }
) {
  const npsRow = selectedValues[valueRowIndex]; // NPS values.
  const spreadRow = selectedValues[spreadRowIndex]; // SD or variance.
  const baseRow = selectedValues[baseRowIndex]; // Bases.

  const rowComparisons = []; // Pairwise NPS comparisons.

  const comparisonPairs = buildColumnComparisonPairs(
    npsRow.length,
    calculationSettings,
    calculationSettings.excludedColumnIndexes || new Set()
  );

  for (const comparisonPair of comparisonPairs) {
    const firstColumnIndex = comparisonPair.firstColumnIndex;
    const secondColumnIndex = comparisonPair.secondColumnIndex;

    const significanceResult = calculateNpsSignificanceFromSpread(
      npsRow[firstColumnIndex],
      spreadRow[firstColumnIndex],
      baseRow[firstColumnIndex],
      npsRow[secondColumnIndex],
      spreadRow[secondColumnIndex],
      baseRow[secondColumnIndex],
      spreadType,
      calculationSettings
    );

    rowComparisons.push({
      firstColumnIndex,
      secondColumnIndex,
      comparisonType: comparisonPair.comparisonType,
      result: significanceResult,
    });
  }

  return {
    baseRowIndex,
    baseRow,
    comparisonRows: [
      {
        valueRowIndex,
        rowComparisons,
      },
    ],
  };
}

/**
 * Adds significance results directly into full cell result matrix.
 *
 * PURPOSE:
 * Used for block-plan mode where different metric blocks
 * occupy arbitrary rows inside one selected range.
 *
 * TOTAL LOGIC:
 * If firstColumnIsTotal is enabled:
 * - Total column never receives markers.
 * - Total-vs-segment comparisons write:
 *   - "T" into segment cell if segment is significantly higher than Total.
 *   - "t" into segment cell if segment is significantly lower than Total.
 * - Segment-vs-segment comparisons use normal labels starting from the second selected column.
 */
export function applyComparisonResultsToFullCellResultMatrix(
  blockResults,
  fullCellResultMatrix,
  calculationSettings = {}
) {
  for (const comparisonRow of blockResults.comparisonRows) {
    const valueRowIndex = comparisonRow.valueRowIndex;

    for (const comparison of comparisonRow.rowComparisons) {
      if (!comparison.result || !comparison.result.isSignificant) {
        continue;
      }

      const firstColumnIndex = comparison.firstColumnIndex;
      const secondColumnIndex = comparison.secondColumnIndex;

      if (comparison.comparisonType === "previousColumn") {
        applyPreviousColumnArrowToCellResultMatrix(
          fullCellResultMatrix,
          valueRowIndex,
          comparison,
          calculationSettings
        );

        continue;
      }

      if (calculationSettings.firstColumnIsTotal && comparison.comparisonType === "total") {
        applyTotalComparisonMarkerToFullCellResultMatrix(
          fullCellResultMatrix,
          valueRowIndex,
          comparison
        );

        continue;
      }

      if (calculationSettings.firstColumnIsTotal) {
        if (firstColumnIndex === 0 || secondColumnIndex === 0) {
          continue;
        }
      }

      const firstLabel = getSignificanceLabelForColumnIndex(firstColumnIndex, calculationSettings);

      const secondLabel = getSignificanceLabelForColumnIndex(
        secondColumnIndex,
        calculationSettings
      );

      if (!firstLabel || !secondLabel) {
        continue;
      }

      if (comparison.result.direction === "first_higher") {
        appendSignificanceMarkerToCellResult(
          fullCellResultMatrix[valueRowIndex][firstColumnIndex],
          secondLabel
        );
      }

      if (comparison.result.direction === "second_higher") {
        appendSignificanceMarkerToCellResult(
          fullCellResultMatrix[valueRowIndex][secondColumnIndex],
          firstLabel
        );
      }
    }
  }
}

/**
 * Applies special Total comparison marker.
 *
 * RULES:
 * - Total column is column 0.
 * - Marker is always written into the segment column.
 * - "T" means segment is significantly higher than Total.
 * - "t" means segment is significantly lower than Total.
 * - Total marker has priority and must appear before segment markers.
 */
function applyTotalComparisonMarkerToFullCellResultMatrix(
  fullCellResultMatrix,
  valueRowIndex,
  comparison
) {
  const firstColumnIndex = comparison.firstColumnIndex;
  const secondColumnIndex = comparison.secondColumnIndex;

  const segmentColumnIndex = firstColumnIndex === 0 ? secondColumnIndex : firstColumnIndex;
  const segmentIsFirstColumn = segmentColumnIndex === firstColumnIndex;

  const segmentIsHigher =
    (segmentIsFirstColumn && comparison.result.direction === "first_higher") ||
    (!segmentIsFirstColumn && comparison.result.direction === "second_higher");

  const totalMarker = segmentIsHigher ? "T" : "t";

  prependTotalMarkerToCellResult(
    fullCellResultMatrix[valueRowIndex][segmentColumnIndex],
    totalMarker
  );
}

/**
 * Applies special Total comparison marker.
 *
 * RULES:
 * - Total column is column 0.
 * - Marker is always written into the segment column.
 * - "T" means segment is significantly higher than Total.
 * - "t" means segment is significantly lower than Total.
 * - Total marker has priority and must appear before segment markers.
 */
function applyTotalComparisonMarkerToFullMarkerMatrix(fullMarkerMatrix, valueRowIndex, comparison) {
  const firstColumnIndex = comparison.firstColumnIndex;
  const secondColumnIndex = comparison.secondColumnIndex;

  const segmentColumnIndex = firstColumnIndex === 0 ? secondColumnIndex : firstColumnIndex;
  const segmentIsFirstColumn = segmentColumnIndex === firstColumnIndex;

  const segmentIsHigher =
    (segmentIsFirstColumn && comparison.result.direction === "first_higher") ||
    (!segmentIsFirstColumn && comparison.result.direction === "second_higher");

  const totalMarker = segmentIsHigher ? "T" : "t";

  fullMarkerMatrix[valueRowIndex][segmentColumnIndex] = prependTotalMarker(
    fullMarkerMatrix[valueRowIndex][segmentColumnIndex],
    totalMarker
  );
}

/**
 * Puts Total marker before ordinary segment markers.
 *
 * Example:
 * - existing "ab" + "T" -> "Tab"
 * - existing "tbc" + "T" -> "Tbc"
 */
function prependTotalMarker(existingMarkers, totalMarker) {
  const markerText = existingMarkers || "";
  const markerTextWithoutOldTotalMarker = markerText.replace(/[tT]/g, "");

  return `${totalMarker}${markerTextWithoutOldTotalMarker}`;
}

/**
 * Clears markers from rows that are not allowed to receive markers.
 *
 * PURPOSE:
 * Defensive protection for complex tables.
 * Service rows may still receive fill formatting, for example small-base fill,
 * but they must not receive significance markers.
 */
export function keepMarkersOnlyInAllowedRows(cellResultMatrix, allowedMarkerRows) {
  for (let rowIndex = 0; rowIndex < cellResultMatrix.length; rowIndex++) {
    if (allowedMarkerRows.has(rowIndex)) {
      continue;
    }

    for (let columnIndex = 0; columnIndex < cellResultMatrix[rowIndex].length; columnIndex++) {
      cellResultMatrix[rowIndex][columnIndex].markers = "";
    }
  }

  return cellResultMatrix;
}

/**
 * Supported two-tailed z-thresholds by confidence level.
 *
 * IMPORTANT:
 * Keys are ordered from highest confidence to lowest confidence.
 */
export const Z_THRESHOLDS_BY_CONFIDENCE_LEVEL = {
  99: 2.576,
  95: 1.96,
  90: 1.645,
  80: 1.282,
  66.6: 0.967,
};

/**
 * Returns z-threshold for selected two-tailed confidence level.
 *
 * PURPOSE:
 * Do not silently fallback to another confidence level.
 */
export function getZThresholdForConfidence(confidenceLevel) {
  const confidenceKey = String(confidenceLevel);
  const threshold = Z_THRESHOLDS_BY_CONFIDENCE_LEVEL[confidenceKey];

  if (threshold === undefined) {
    throw new Error(`Unsupported confidence level: ${confidenceLevel}`);
  }

  return threshold;
}

/**
 * Approximate two-tailed t-thresholds by confidence level.
 *
 * PURPOSE:
 * Used for means and NPS spread calculations.
 */
export const T_THRESHOLDS_BY_CONFIDENCE_LEVEL = {
  99: [
    { df: 1, value: 63.657 },
    { df: 2, value: 9.925 },
    { df: 5, value: 4.032 },
    { df: 10, value: 3.169 },
    { df: 20, value: 2.845 },
    { df: 30, value: 2.75 },
    { df: 60, value: 2.66 },
  ],
  95: [
    { df: 1, value: 12.706 },
    { df: 2, value: 4.303 },
    { df: 5, value: 2.571 },
    { df: 10, value: 2.228 },
    { df: 20, value: 2.086 },
    { df: 30, value: 2.042 },
    { df: 60, value: 2.0 },
  ],
  90: [
    { df: 1, value: 6.314 },
    { df: 2, value: 2.92 },
    { df: 5, value: 2.015 },
    { df: 10, value: 1.812 },
    { df: 20, value: 1.725 },
    { df: 30, value: 1.697 },
    { df: 60, value: 1.671 },
  ],
  80: [
    { df: 1, value: 3.078 },
    { df: 2, value: 1.886 },
    { df: 5, value: 1.476 },
    { df: 10, value: 1.372 },
    { df: 20, value: 1.325 },
    { df: 30, value: 1.31 },
    { df: 60, value: 1.296 },
  ],
  66.6: [
    { df: 1, value: 1.376 },
    { df: 2, value: 0.816 },
    { df: 5, value: 0.727 },
    { df: 10, value: 0.7 },
    { df: 20, value: 0.687 },
    { df: 30, value: 0.683 },
    { df: 60, value: 0.677 },
  ],
};

/**
 * Returns approximate two-tailed t-threshold for selected confidence level.
 *
 * PURPOSE:
 * For means and NPS spread calculations.
 */
export function getTThresholdForConfidence(confidenceLevel, degreesOfFreedom) {
  const confidenceKey = String(confidenceLevel);
  const thresholdTable = T_THRESHOLDS_BY_CONFIDENCE_LEVEL[confidenceKey];

  if (!thresholdTable) {
    throw new Error(`Unsupported confidence level: ${confidenceLevel}`);
  }

  const zThreshold = getZThresholdForConfidence(confidenceKey);

  if (degreesOfFreedom >= 60) {
    return zThreshold;
  }

  for (const thresholdPoint of thresholdTable) {
    if (degreesOfFreedom <= thresholdPoint.df) {
      return thresholdPoint.value;
    }
  }

  return zThreshold;
}

/**
 * Applies small-base formatting for one calculation block
 * and returns columns that must be excluded from significance calculations.
 *
 * IMPORTANT:
 * Small-base logic runs before statistical tests.
 *
 * RULES:
 * - If base < threshold, the column is excluded from comparisons in this block.
 * - Small-base fill is applied to the whole column inside this block.
 * - Base row itself is also filled.
 * - If Total column has small base, calculation is stopped with an error.
 */
export function applySmallBaseRulesForCalculationBlock(
  selectedValues,
  calculationBlock,
  fullCellResultMatrix,
  calculationSettings = {}
) {
  const excludedColumnIndexes = new Set();

  if (!calculationSettings.excludeSmallBasesFromComparisons) {
    return {
      excludedColumnIndexes,
      errorMessage: "",
    };
  }

  const threshold = Number(calculationSettings.smallBaseThreshold);

  if (Number.isNaN(threshold) || threshold < 0) {
    return {
      excludedColumnIndexes,
      errorMessage: "Некорректный порог маленькой базы. Проверьте настройку “База <”.",
    };
  }

  const baseRow = selectedValues[calculationBlock.baseRowIndex];

  if (!baseRow) {
    return {
      excludedColumnIndexes,
      errorMessage: "",
    };
  }

  const blockRowIndexes = getRowIndexesCoveredByCalculationBlock(calculationBlock);

  for (let columnIndex = 0; columnIndex < baseRow.length; columnIndex++) {
    const baseValue = parseBaseValue(baseRow[columnIndex]);

    if (baseValue === null) {
      continue;
    }

    if (baseValue >= threshold) {
      continue;
    }

    if (calculationSettings.firstColumnIsTotal && columnIndex === 0) {
      return {
        excludedColumnIndexes,
        errorMessage:
          "В колонке Тотал обнаружена маленькая база. Проверьте данные: база Тотала не должна быть меньше баз сегментов. Расчёт остановлен.",
      };
    }

    excludedColumnIndexes.add(columnIndex);

    applySmallBaseFillToBlockColumn(fullCellResultMatrix, blockRowIndexes, columnIndex);
  }

  return {
    excludedColumnIndexes,
    errorMessage: "",
  };
}

/**
 * Applies small-base fill to every row of one calculation block
 * in the selected column.
 */
function applySmallBaseFillToBlockColumn(fullCellResultMatrix, blockRowIndexes, columnIndex) {
  for (const rowIndex of blockRowIndexes) {
    if (!fullCellResultMatrix[rowIndex]) {
      continue;
    }

    const cellResult = fullCellResultMatrix[rowIndex][columnIndex];

    if (!cellResult) {
      continue;
    }

    cellResult.markers = "";
    cellResult.hasPositiveTotalComparison = false;
    cellResult.previousColumnArrow = "";
    cellResult.previousColumnArrowDirection = "";

    applyFillReasonToCellResult(cellResult, CELL_FILL_REASONS.SMALL_BASE);
  }
}

/**
 * Parses base value safely.
 */
function parseBaseValue(rawValue) {
  if (rawValue === null || rawValue === undefined || rawValue === "") {
    return null;
  }

  const numericValue = Number(String(rawValue).trim().replace(",", "."));

  if (Number.isNaN(numericValue)) {
    return null;
  }

  return numericValue;
}

/**
 * Returns all selected-range row indexes covered by one calculation block.
 *
 * PURPOSE:
 * Small-base fill is applied within the current calculation block,
 * including the block's Base row.
 */
function getRowIndexesCoveredByCalculationBlock(calculationBlock) {
  if (calculationBlock.metricType === "proportion") {
    return [...calculationBlock.valueRowIndexes, calculationBlock.baseRowIndex];
  }

  if (calculationBlock.metricType === "mean") {
    return [
      calculationBlock.valueRowIndex,
      calculationBlock.spreadRowIndex,
      calculationBlock.baseRowIndex,
    ];
  }

  if (calculationBlock.metricType === "npsStructure") {
    return [
      calculationBlock.valueRowIndex,
      calculationBlock.promotersRowIndex,
      calculationBlock.detractorsRowIndex,
      calculationBlock.baseRowIndex,
    ];
  }

  if (calculationBlock.metricType === "npsSpread") {
    return [
      calculationBlock.valueRowIndex,
      calculationBlock.spreadRowIndex,
      calculationBlock.baseRowIndex,
    ];
  }

  return [];
}

/**
 * Applies previous-column arrow to the right/current column of a comparison pair.
 *
 * RULE:
 * - Previous column is firstColumnIndex.
 * - Current column is secondColumnIndex.
 * - Arrow is written only into current/right column.
 */
function applyPreviousColumnArrowToCellResultMatrix(
  fullCellResultMatrix,
  valueRowIndex,
  comparison,
  calculationSettings = {}
) {
  const currentColumnIndex = comparison.secondColumnIndex;
  const cellResult = fullCellResultMatrix[valueRowIndex][currentColumnIndex];

  if (!cellResult) {
    return;
  }

  if (comparison.result.direction === "second_higher") {
    cellResult.previousColumnArrow = "↑";
    cellResult.previousColumnArrowDirection = "up";

    if (calculationSettings.applyPreviousColumnFill) {
      applyFillReasonToCellResult(cellResult, CELL_FILL_REASONS.SIGNIFICANT);
    }

    return;
  }

  if (comparison.result.direction === "first_higher") {
    cellResult.previousColumnArrow = "↓";
    cellResult.previousColumnArrowDirection = "down";

    if (calculationSettings.applyPreviousColumnFill) {
      applyFillReasonToCellResult(cellResult, CELL_FILL_REASONS.LOWER_THAN_TOTAL);
    }
  }
}
