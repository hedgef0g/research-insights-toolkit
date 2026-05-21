/**
 * Shared constants and helpers for action + scope processing.
 *
 * Provides a single source of truth for the three supported action values
 * (run, clear, check) and three supported scope values (current_table,
 * current_sheet, whole_workbook), plus lightweight factories for the compact
 * result objects returned by each action pipeline.
 *
 * This module is Office.js-free and has no side effects.
 */

// ─── Scope constants ──────────────────────────────────────────────────────────

export const SCOPES = Object.freeze({
  CURRENT_TABLE: "current_table",
  CURRENT_SHEET: "current_sheet",
  WHOLE_WORKBOOK: "whole_workbook",
});

const VALID_SCOPES = new Set(Object.values(SCOPES));

// ─── Action constants ─────────────────────────────────────────────────────────

export const ACTIONS = Object.freeze({
  RUN: "run",
  CLEAR: "clear",
  CHECK: "check",
});

const VALID_ACTIONS = new Set(Object.values(ACTIONS));

// ─── Validators ───────────────────────────────────────────────────────────────

/**
 * Normalizes and validates an action value.
 * Returns the canonical lowercase action string, or null if unrecognized.
 */
export function normalizeAction(value) {
  if (value === null || value === undefined) return null;
  const lower = String(value).toLowerCase().trim();
  return VALID_ACTIONS.has(lower) ? lower : null;
}

/**
 * Normalizes and validates a scope value.
 * Returns the canonical lowercase scope string, or null if unrecognized.
 */
export function normalizeScope(value) {
  if (value === null || value === undefined) return null;
  const lower = String(value).toLowerCase().trim();
  return VALID_SCOPES.has(lower) ? lower : null;
}

// ─── Result factories ─────────────────────────────────────────────────────────

/**
 * Creates a result object for a successfully processed (Run) table.
 * @param {string} rangeAddress
 * @param {number} [blocksProcessed]
 * @param {string} [message]
 */
export function makeProcessedResult(rangeAddress, blocksProcessed, message) {
  return { status: "processed", rangeAddress, blocksProcessed: blocksProcessed ?? null, message: message ?? "" };
}

/**
 * Creates a result object for a successfully cleared table.
 * @param {string} [rangeAddress]
 * @param {string} [message]
 */
export function makeClearedResult(rangeAddress, message) {
  return { status: "cleared", rangeAddress: rangeAddress ?? null, message: message ?? "" };
}

/**
 * Creates a result object for a successfully checked table.
 * @param {string} [rangeAddress]
 * @param {string} [message]
 */
export function makeCheckedResult(rangeAddress, message) {
  return { status: "checked", rangeAddress: rangeAddress ?? null, message: message ?? "" };
}

/**
 * Creates a result object for a skipped table (no mutation performed).
 * @param {string} [rangeAddress]
 * @param {string} [message]
 */
export function makeSkippedResult(rangeAddress, message) {
  return { status: "skipped", rangeAddress: rangeAddress ?? null, message: message ?? "" };
}

/**
 * Creates a result object for a blocked table (unsafe selection, no mutation performed).
 * @param {string} [rangeAddress]
 * @param {string} [message]
 */
export function makeBlockedResult(rangeAddress, message) {
  return { status: "blocked", rangeAddress: rangeAddress ?? null, message: message ?? "" };
}

/**
 * Creates a result object for an error during processing.
 * @param {string} [rangeAddress]
 * @param {string} [message]
 */
export function makeErrorResult(rangeAddress, message) {
  return { status: "error", rangeAddress: rangeAddress ?? null, message: message ?? "" };
}

// ─── Reason codes ─────────────────────────────────────────────────────────────

/**
 * Centralized reason codes used across Run / Clear / Check summaries.
 * Values are stable string identifiers; display text lives in the caller.
 */
export const SKIP_REASONS = Object.freeze({
  TOO_LITTLE_DATA: "too_little_data",
  NO_CALCULATION_BLOCKS: "no_calculation_blocks",
  NO_DATA_AFTER_NORMALIZATION: "no_data_after_normalization",
  BLOCKED_SELECTION: "blocked_selection",
  EMPTY_RANGE: "empty_range",
  CONTENT_SHEET: "content_sheet",
  CANDIDATE_NOT_AVAILABLE: "candidate_not_available",
  OVER_CELL_LIMIT: "over_cell_limit",
});
