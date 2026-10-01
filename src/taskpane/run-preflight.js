/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Excel */

import {
  detectSignificanceMarkerOverflow,
  detectBatchMarkerOverflow,
} from "../core/significance";

import { detectBannerStructure } from "../core/banner-detector";

import { interpretSelectedRange } from "./selected-range-interpreter";

/**
 * Read-only marker-overflow check for a single range, mirroring the front half
 * of runSignificanceForRangeInContext (load -> interpret -> detect banner) but
 * WITHOUT writing anything. Returns true when the table would need more
 * single-character markers than the alphabet provides.
 *
 * Used by the banner-aware batch preflight so overflow is judged on the exact
 * group-local required count rather than raw table width.
 */
export async function computeRangeMarkerOverflowInContext(
  context,
  sheetName,
  rangeAddress,
  calculationSettings
) {
  const worksheet = context.workbook.worksheets.getItem(sheetName);
  const sourceRange = worksheet.getRange(rangeAddress);

  sourceRange.load(["values", "text", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
  await context.sync();

  const selectedValues = sourceRange.values;
  const selectedText = sourceRange.text;

  if (!selectedValues || selectedValues.length < 2 || selectedValues[0].length < 2) {
    return false;
  }

  const interpretation = await interpretSelectedRange(
    context,
    sourceRange,
    selectedValues,
    selectedText,
    calculationSettings
  );

  if (interpretation.state === "blocked") {
    return false;
  }

  const { valuesForCalculation, bannerContext } = interpretation;

  if (
    !valuesForCalculation ||
    valuesForCalculation.length < 2 ||
    !valuesForCalculation[0] ||
    valuesForCalculation[0].length < 2
  ) {
    return false;
  }

  let bannerStructure = null;
  if (calculationSettings.respectBannerStructure) {
    bannerStructure = detectBannerStructure(bannerContext, calculationSettings);
  }

  return detectSignificanceMarkerOverflow(
    valuesForCalculation[0].length,
    calculationSettings,
    bannerStructure
  );
}

/**
 * Exact, read-only batch overflow pass for banner-aware runs.
 *
 * Interprets and banner-detects each eligible table (no writes) and returns
 * true as soon as one would overflow on its group-local required count. Errors
 * during preflight are ignored here -- the per-table run reports them -- so a
 * single bad table never blocks the decision.
 */
export async function detectBatchMarkerOverflowExact(eligible, calculationSettings) {
  let overflow = false;

  await Excel.run(async (context) => {
    for (const candidate of eligible) {
      try {
        if (
          await computeRangeMarkerOverflowInContext(
            context,
            candidate.sheetName,
            candidate.rangeAddress,
            calculationSettings
          )
        ) {
          overflow = true;
          return;
        }
      } catch (_) {
        /* preflight read failure is non-fatal; per-table run handles it */
      }
    }
  });

  return overflow;
}

/**
 * Operation-level marker-overflow preflight for batch runs (current-sheet /
 * workbook). Decides BEFORE any table is written, so Stop leaves the whole
 * operation without partial results.
 *
 * Comparison-mode aware:
 * - previous-column mode uses arrow markers, never letter labels -> no dialog;
 * - non-banner modes -> conservative raw column-count gate;
 * - banner-aware mode -> raw width is not meaningful (Total/label columns and
 *   per-group widths), so when the cheap gate says overflow is *possible* an
 *   exact read-only pass confirms it on the group-local required count.
 *
 * Returns true when the user chose to stop (caller must abort before writing
 * any table). On Continue it sets allowMultiCharacterMarkers on the shared
 * calculationSettings so every table uses multi-character markers.
 *
 * The per-table preflight inside runSignificanceForRangeInContext stays as a
 * safety net; because the shared decider's choice is cached here, it never
 * re-prompts.
 */
export async function preflightBatchMarkerOverflow(
  eligible,
  itemMap,
  calculationSettings,
  decider
) {
  // Previous-column mode never uses letter labels -- capacity is irrelevant.
  if (calculationSettings.compareWithPreviousColumn) {
    return false;
  }

  const columnCounts = eligible.map((candidate) => {
    const item = itemMap.get(`${candidate.sheetName}!${candidate.rangeAddress}`);
    return item ? item.columnCount : undefined;
  });

  // Cheap gate first: if no table is even wider than the alphabet, nothing can
  // overflow (the real required count is always <= raw column count).
  if (!detectBatchMarkerOverflow(columnCounts, calculationSettings)) {
    return false;
  }

  // A table is wide enough that overflow is possible. In banner-aware mode raw
  // width over-counts (Total columns, per-group widths), so confirm with an
  // exact read-only pass before prompting. Non-banner modes use the raw gate.
  if (calculationSettings.respectBannerStructure) {
    const exactOverflow = await detectBatchMarkerOverflowExact(eligible, calculationSettings);
    if (!exactOverflow) {
      return false;
    }
  }

  if ((await decider.resolve()) === "stop") {
    return true;
  }

  calculationSettings.allowMultiCharacterMarkers = true;
  return false;
}
