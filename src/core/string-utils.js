/**
 * Shared text normalization utilities for lookup and display.
 *
 * PURPOSE:
 * Single source of truth for the two normalization operations used across
 * metric-detector and banner-detector.
 *
 * normalizeLookupText — for dictionary/keyword matching.
 * normalizeDisplayText — for preserving user-facing text as-is.
 */

/**
 * Normalizes a raw value for dictionary/keyword lookup.
 *
 * - Lowercases
 * - Replaces ё → е
 * - Collapses punctuation to spaces
 * - Collapses whitespace runs
 */
export function normalizeLookupText(rawValue) {
  if (rawValue === null || rawValue === undefined) {
    return "";
  }

  return String(rawValue)
    .toLowerCase()
    .replace(/ё/g, "е")
    .replace(/[.,:;()]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Normalizes a raw cell value for display — trims whitespace, preserves case.
 */
export function normalizeDisplayText(rawValue) {
  if (rawValue === null || rawValue === undefined) {
    return "";
  }

  return String(rawValue).trim();
}
