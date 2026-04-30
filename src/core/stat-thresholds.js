import normalQuantile from "@stdlib/stats-base-dists-normal-quantile";
import tQuantile from "@stdlib/stats-base-dists-t-quantile";

/**
 * Converts UI confidence level string to numeric confidence.
 *
 * Examples:
 * "95" -> 0.95
 * "66.6" -> 0.666
 */
export function parseConfidenceLevel(confidenceLevel) {
  const numericConfidenceLevel = Number(confidenceLevel);

  if (Number.isNaN(numericConfidenceLevel)) {
    throw new Error(`Unsupported confidence level: ${confidenceLevel}`);
  }

  const confidence = numericConfidenceLevel / 100;

  if (confidence <= 0 || confidence >= 1) {
    throw new Error(`Unsupported confidence level: ${confidenceLevel}`);
  }

  return confidence;
}

/**
 * Returns upper-tail probability point for a two-tailed confidence interval.
 *
 * Example:
 * confidence = 0.95
 * alpha = 0.05
 * p = 0.975
 */
function getUpperProbability(confidenceLevel, options = {}) {
  const confidence = parseConfidenceLevel(confidenceLevel);
  const alpha = 1 - confidence;

  if (options.oneTailedTest) {
    return confidence;
  }

  return 1 - alpha / 2;
}

/**
 * Returns z-threshold for a two-tailed normal test.
 */
export function getZThresholdForConfidence(confidenceLevel, options = {}) {
  const p = getUpperProbability(confidenceLevel, options);
  const threshold = normalQuantile(p, 0, 1);

  if (!Number.isFinite(threshold)) {
    throw new Error(`Could not calculate z-threshold for confidence level: ${confidenceLevel}`);
  }

  return threshold;
}

/**
 * Returns t-threshold for a two-tailed Student's t-test.
 */
export function getTThresholdForConfidence(confidenceLevel, degreesOfFreedom, options = {}) {
  const df = Number(degreesOfFreedom);

  if (!Number.isFinite(df) || df <= 0) {
    throw new Error(`Invalid degrees of freedom: ${degreesOfFreedom}`);
  }

  const p = getUpperProbability(confidenceLevel, options);
  const threshold = tQuantile(p, df);

  if (!Number.isFinite(threshold)) {
    throw new Error(
      `Could not calculate t-threshold for confidence level ${confidenceLevel} and df ${degreesOfFreedom}`
    );
  }

  return threshold;
}
