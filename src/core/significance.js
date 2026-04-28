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
  const numericValue = Number(rawValue); // Numeric version of the spreadsheet cell value.

  if (Number.isNaN(numericValue)) {
    return null;
  }

  // If value is greater than 1, assume it was entered as a percentage.
  // Example: 42 becomes 0.42.
  if (numericValue > 1) {
    return numericValue / 100;
  }

  // If value is already between 0 and 1, treat it as a decimal share.
  return numericValue;
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