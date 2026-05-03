/**
 * Table preview data model for Research Insights Toolkit.
 *
 * PURPOSE:
 * Provides a read-only snapshot of how RIT interprets the selected table —
 * row types, calculation blocks, banner structure, and data quality issues —
 * without calculating significance and without writing anything to Excel.
 *
 * INTENDED USE:
 * Future "Проверить таблицу" / "Check table" feature.
 * Safe to call from taskpane.js but must not be wired into it in this PR.
 *
 * DOES NOT:
 * - Write to Excel or call excel-writer.js
 * - Calculate statistical significance or call significance.js
 * - Scan worksheets or workbooks automatically
 * - Depend on Office.js
 * - Implement weighted base logic
 * - Implement multi-column row label calculation (only represents the data)
 */

import {
  detectMetricRowsFromLeftLabels,
  buildCalculationBlocks,
  normalizeLabelText,
} from "./metric-detector";

import { detectBannerStructure } from "./banner-detector";

import { normalizeShare, normalizeNpsValue } from "./normalizers";

// NPS mismatch display tolerance.
// Differences below this threshold are assumed to be rounding artefacts.
const NPS_MISMATCH_ROUNDING_TOLERANCE = 0.02;
// Differences above this threshold are elevated from warning to critical.
const NPS_MISMATCH_CRITICAL_THRESHOLD = 0.05;

// ─── Main export ───────────────────────────────────────────────────────────

/**
 * Builds a read-only preview model of how RIT interprets the selected table.
 *
 * @param {object} input
 * @param {Array}  input.values           - 2D array [row][col] of cleaned cell values.
 * @param {Array}  input.leftLabelValues  - 2D array of left-side label cells (up to LABEL_SCAN_COLUMNS_LEFT columns).
 * @param {Array}  [input.numberFormats]  - 2D array of Excel number format strings (reserved for future use).
 * @param {object} [input.bannerContext]  - Banner detection context (upper scan rows, column count, etc.).
 * @param {object} [input.settings]       - Calculation settings object.
 * @returns {object} Preview model — plain JSON-compatible object.
 */
export function buildTablePreviewModel(input) {
  const { values, leftLabelValues, bannerContext, settings } = input || {};

  const safeValues = Array.isArray(values) ? values : [];
  const safeLeft = Array.isArray(leftLabelValues) ? leftLabelValues : [];
  const safeSettings = settings || {};

  // Detect row types.
  const detectionResult = detectMetricRowsFromLeftLabels(safeValues, safeLeft);

  // Banner structure (only populated when respectBannerStructure is enabled).
  const bannerStructure = buildBannerPreview(bannerContext, safeSettings);

  // Enrich outputs for the preview layer.
  const rawRowDiagnostics = detectionResult?.rowDiagnostics || [];
  const rawBlocks = buildCalculationBlocks(detectionResult);
  const rowDiagnostics = buildPreviewRowDiagnostics(rawRowDiagnostics, safeLeft);
  const calculationBlocks = buildPreviewBlocks(rawBlocks, rawRowDiagnostics);

  // Data quality analysis.
  const dataQualityIssues = [
    ...checkNumericLikeLabels(rowDiagnostics),
    ...checkBaseConsistency(safeValues, calculationBlocks, bannerStructure),
    ...checkNpsMismatch(safeValues, calculationBlocks),
  ];

  const qualitySummary = buildQualitySummary(dataQualityIssues);
  const summary = buildSummary(safeValues, rowDiagnostics, calculationBlocks, bannerStructure);

  // Flat warnings array — a convenience alias for UI compatibility.
  const warnings = dataQualityIssues.map((issue) => ({
    code: issue.code,
    severity: issue.severity,
    text: issue.message,
    rowIndex: issue.rowIndex ?? null,
    columnIndex: issue.columnIndex ?? null,
  }));

  return {
    rowDiagnostics,
    calculationBlocks,
    bannerStructure,
    dataQualityIssues,
    qualitySummary,
    summary,
    warnings,
  };
}

// ─── Banner ────────────────────────────────────────────────────────────────

function buildBannerPreview(bannerContext, settings) {
  const isEnabled = !!(settings && settings.respectBannerStructure);

  if (!isEnabled || !bannerContext) {
    return {
      isEnabled,
      isDetected: false,
      mode: null,
      groups: [],
      totalColumnIndexes: [],
      globalTotalColumnIndex: null,
      hasWaveGroups: false,
      recommendedComparisonMode: null,
      messages: [],
    };
  }

  const detected = detectBannerStructure(bannerContext, settings);

  return {
    isEnabled,
    isDetected: detected.isDetected || false,
    mode: detected.mode || null,
    groups: detected.groups || [],
    totalColumnIndexes: detected.totalColumnIndexes || [],
    globalTotalColumnIndex: detected.globalTotalColumnIndex ?? null,
    hasWaveGroups: detected.hasWaveGroups || false,
    recommendedComparisonMode: detected.recommendedComparisonMode || null,
    messages: detected.messages || [],
  };
}

// ─── Row diagnostics ───────────────────────────────────────────────────────

function buildPreviewRowDiagnostics(rawDiagnostics, leftLabelValues) {
  return (rawDiagnostics || []).map((diag) => {
    const leftRow = leftLabelValues[diag.rowIndex] || [];
    const labelParts = extractLabelPartsFromRow(leftRow);

    // Rightmost part is closest to the data column and treated as the primary label.
    const primaryLabel = labelParts.length > 0 ? labelParts[labelParts.length - 1] : "";
    const secondaryLabel = labelParts.length > 1 ? labelParts[labelParts.length - 2] : null;
    // Future: combinedLabel may join hierarchical parts (e.g. "Gender / Male").
    const combinedLabel = labelParts.length > 1 ? labelParts.join(" / ") : primaryLabel;

    return {
      rowIndex: diag.rowIndex,

      // Raw and normalized label values.
      label: diag.rawLabel,
      labelParts,
      combinedLabel,
      primaryLabel,
      secondaryLabel,
      normalizedLabel: diag.normalizedLabel,
      normalizedLabelParts: labelParts.map(normalizeLabelText),

      // Classification.
      rowType: diag.rowType,
      // rowSubtype is reserved for future weighted/unweighted/effective base variants.
      rowSubtype: null,
      confidence: inferRowTypeConfidence(diag.rowType),
      notes: [],
    };
  });
}

/**
 * Collects non-empty values from the left label cells of one row.
 * Left-to-right order is preserved. Numeric-looking labels are preserved
 * because they can be valid NPS scale rows, age groups, or wave labels.
 */
function extractLabelPartsFromRow(leftRow) {
  const parts = [];

  for (const cellValue of leftRow) {
    if (cellValue === null || cellValue === undefined || cellValue === "") continue;
    const text = String(cellValue).trim();
    if (!text) continue;
    parts.push(text);
  }

  return parts;
}

function inferRowTypeConfidence(rowType) {
  if (rowType === "empty" || rowType === "unknownText") return "low";
  return "high";
}

// ─── Calculation blocks ────────────────────────────────────────────────────

function buildPreviewBlocks(rawBlocks, rowDiagnostics) {
  const blocks = (rawBlocks || []).map(normalizeBlockShape);
  enrichNpsBlocksWithNeutralRow(blocks);
  return blocks;
}

/**
 * Normalises the varied shapes emitted by buildCalculationBlocks into a
 * single consistent shape for the preview layer.
 *
 * Fields that do not apply to a given metricType are set to null.
 */
function normalizeBlockShape(block) {
  const base = {
    metricType: block.metricType,
    valueRowIndexes: null,
    valueRowIndex: null,
    baseRowIndex: block.baseRowIndex ?? null,
    promotersRowIndex: null,
    detractorsRowIndex: null,
    neutralRowIndex: null,
    sdRowIndex: null,
    varianceRowIndex: null,
    notes: [],
  };

  switch (block.metricType) {
    case "proportion":
      return {
        ...base,
        valueRowIndexes: block.valueRowIndexes || [],
      };

    case "mean":
      return {
        ...base,
        valueRowIndex: block.valueRowIndex ?? null,
        valueRowIndexes: block.valueRowIndex != null ? [block.valueRowIndex] : [],
        sdRowIndex:
          block.spreadType === "standardDeviation" ? (block.spreadRowIndex ?? null) : null,
        varianceRowIndex: block.spreadType === "variance" ? (block.spreadRowIndex ?? null) : null,
      };

    case "npsStructure":
      return {
        ...base,
        valueRowIndex: block.valueRowIndex ?? null,
        valueRowIndexes: block.valueRowIndex != null ? [block.valueRowIndex] : [],
        promotersRowIndex: block.promotersRowIndex ?? null,
        detractorsRowIndex: block.detractorsRowIndex ?? null,
        // neutralRowIndex is filled in by enrichNpsBlocksWithNeutralRow.
      };

    case "npsSpread":
      return {
        ...base,
        valueRowIndex: block.valueRowIndex ?? null,
        valueRowIndexes: block.valueRowIndex != null ? [block.valueRowIndex] : [],
        sdRowIndex:
          block.spreadType === "standardDeviation" ? (block.spreadRowIndex ?? null) : null,
        varianceRowIndex: block.spreadType === "variance" ? (block.spreadRowIndex ?? null) : null,
      };

    default:
      return base;
  }
}

/**
 * Fills in neutralRowIndex for npsStructure blocks where a Neutral/Passive row
 * exists between Promoters and Detractors.
 *
 * Strategy: find the proportion block that shares the same baseRowIndex and
 * contains both promoters and detractors in its valueRowIndexes. Any remaining
 * index in that proportion block is the neutral row.
 */
function enrichNpsBlocksWithNeutralRow(previewBlocks) {
  for (const block of previewBlocks) {
    if (
      block.metricType !== "npsStructure" ||
      block.promotersRowIndex === null ||
      block.detractorsRowIndex === null
    ) {
      continue;
    }

    const proportionBlock = previewBlocks.find(
      (b) =>
        b.metricType === "proportion" &&
        b.baseRowIndex === block.baseRowIndex &&
        Array.isArray(b.valueRowIndexes) &&
        b.valueRowIndexes.includes(block.promotersRowIndex) &&
        b.valueRowIndexes.includes(block.detractorsRowIndex)
    );

    if (!proportionBlock) continue;

    const neutralCandidates = proportionBlock.valueRowIndexes.filter(
      (idx) => idx !== block.promotersRowIndex && idx !== block.detractorsRowIndex
    );

    if (neutralCandidates.length === 1) {
      block.neutralRowIndex = neutralCandidates[0];
    }
  }
}

// ─── Data quality checks ───────────────────────────────────────────────────

/**
 * Flags row labels that look purely numeric or like numeric ranges and were not
 * classified as a known metric type.
 *
 * Note: numeric labels can be valid in NPS 1–10 scales, age groups (18–24),
 * or wave numbers. This check produces warnings, not errors.
 */
function checkNumericLikeLabels(rowDiagnostics) {
  const issues = [];

  for (const row of rowDiagnostics) {
    if (!row.label) continue;
    if (row.rowType !== "unknownText" && row.rowType !== "empty") continue;
    if (!looksNumericOrRange(row.label)) continue;

    issues.push({
      code: "NUMERIC_LIKE_LABEL",
      severity: "warning",
      message: `Row ${row.rowIndex + 1}: label "${row.label}" looks numeric but was not matched to a known metric type. May be valid for NPS scales, age groups, or wave numbers.`,
      rowIndex: row.rowIndex,
      columnIndex: null,
      relatedRowIndexes: [],
      relatedColumnIndexes: [],
      evidence: { label: row.label, rowType: row.rowType },
    });
  }

  return issues;
}

/** Returns true for purely numeric labels and simple numeric ranges (e.g. "1", "18–24"). */
function looksNumericOrRange(label) {
  const s = String(label).trim();
  if (/^\d+$/.test(s)) return true;
  if (/^\d+[.,]\d+$/.test(s)) return true;
  if (/^\d+[\-–—]\d+$/.test(s)) return true;
  return false;
}

/**
 * Checks base-size consistency across columns using banner structure.
 *
 * Implements issues 2–5 from the spec:
 * 2. Global Total base < any other column base  → critical
 * 3. Local Total base < any group member base   → critical
 * 4. Local Total base < sum of group members    → warning
 * 5. Global Total base < sum of any group's members → warning
 *
 * Only runs when banner detection is enabled and detected.
 */
function checkBaseConsistency(values, calculationBlocks, bannerStructure) {
  const issues = [];

  if (!bannerStructure || !bannerStructure.isEnabled || !bannerStructure.isDetected) {
    return issues;
  }

  const { groups, globalTotalColumnIndex } = bannerStructure;
  const checkedBaseRows = new Set();

  for (const block of calculationBlocks) {
    if (block.baseRowIndex === null) continue;
    if (checkedBaseRows.has(block.baseRowIndex)) continue;
    checkedBaseRows.add(block.baseRowIndex);

    const baseValues = extractRowNumbers(values, block.baseRowIndex);
    if (baseValues.length === 0) continue;

    const rowLabel = `Row ${block.baseRowIndex + 1}`;

    // Check 2: Global Total base < any other column base.
    if (globalTotalColumnIndex !== null && globalTotalColumnIndex < baseValues.length) {
      const globalTotalBase = baseValues[globalTotalColumnIndex];

      if (globalTotalBase !== null) {
        const violatingColumns = baseValues
          .map((v, idx) => ({ v, idx }))
          .filter(
            ({ v, idx }) => idx !== globalTotalColumnIndex && v !== null && v > globalTotalBase
          )
          .map(({ idx }) => idx);

        if (violatingColumns.length > 0) {
          issues.push({
            code: "GLOBAL_TOTAL_BASE_TOO_SMALL",
            severity: "critical",
            message: `${rowLabel}: Global Total base (${globalTotalBase}) is smaller than ${violatingColumns.length} other column(s).`,
            rowIndex: block.baseRowIndex,
            columnIndex: globalTotalColumnIndex,
            relatedRowIndexes: [],
            relatedColumnIndexes: violatingColumns,
            evidence: { globalTotalBase, violatingColumns },
          });
        }
      }
    }

    // Checks 3 & 4: Local Total consistency within each group.
    for (const group of groups || []) {
      const localTotalIndexes = group.localTotalColumnIndexes || [];
      if (localTotalIndexes.length === 0) continue;

      const memberIndexes = (group.columnIndexes || []).filter(
        (idx) => !localTotalIndexes.includes(idx)
      );
      if (memberIndexes.length === 0) continue;

      const memberBases = memberIndexes
        .map((idx) => (idx < baseValues.length ? baseValues[idx] : null))
        .filter((v) => v !== null);

      for (const localTotalIdx of localTotalIndexes) {
        if (localTotalIdx >= baseValues.length) continue;
        const localTotalBase = baseValues[localTotalIdx];
        if (localTotalBase === null) continue;

        // Check 3: Local Total < any single member.
        const violatingMembers = memberIndexes.filter((idx) => {
          const mb = idx < baseValues.length ? baseValues[idx] : null;
          return mb !== null && mb > localTotalBase;
        });

        if (violatingMembers.length > 0) {
          issues.push({
            code: "LOCAL_TOTAL_BASE_TOO_SMALL",
            severity: "critical",
            message: `${rowLabel}: Local Total base (${localTotalBase}) in group "${group.label}" is smaller than ${violatingMembers.length} member column(s).`,
            rowIndex: block.baseRowIndex,
            columnIndex: localTotalIdx,
            relatedRowIndexes: [],
            relatedColumnIndexes: violatingMembers,
            evidence: { localTotalBase, groupLabel: group.label, violatingMembers },
          });
        }

        // Check 4: Local Total < sum of group members.
        if (memberBases.length > 0) {
          const memberSum = memberBases.reduce((a, b) => a + b, 0);
          if (memberSum > localTotalBase) {
            issues.push({
              code: "LOCAL_TOTAL_BASE_LESS_THAN_SUM",
              severity: "warning",
              message: `${rowLabel}: Local Total base (${localTotalBase}) in group "${group.label}" is less than the sum of member bases (${memberSum}). May be valid with overlapping groups.`,
              rowIndex: block.baseRowIndex,
              columnIndex: localTotalIdx,
              relatedRowIndexes: [],
              relatedColumnIndexes: memberIndexes,
              evidence: { localTotalBase, memberSum, groupLabel: group.label },
            });
          }
        }
      }
    }

    // Check 5: Global Total < sum of any group's non-total members.
    if (globalTotalColumnIndex !== null && globalTotalColumnIndex < baseValues.length) {
      const globalTotalBase = baseValues[globalTotalColumnIndex];

      if (globalTotalBase !== null) {
        for (const group of groups || []) {
          const localTotalIndexes = group.localTotalColumnIndexes || [];
          const memberIndexes = (group.columnIndexes || []).filter(
            (idx) => !localTotalIndexes.includes(idx) && idx !== globalTotalColumnIndex
          );
          if (memberIndexes.length === 0) continue;

          const memberBases = memberIndexes
            .map((idx) => (idx < baseValues.length ? baseValues[idx] : null))
            .filter((v) => v !== null);

          if (memberBases.length === 0) continue;

          const memberSum = memberBases.reduce((a, b) => a + b, 0);
          if (memberSum > globalTotalBase) {
            issues.push({
              code: "GLOBAL_TOTAL_BASE_LESS_THAN_GROUP_SUM",
              severity: "warning",
              message: `${rowLabel}: Global Total base (${globalTotalBase}) is less than the sum of member bases (${memberSum}) in group "${group.label}". May be valid with overlapping groups.`,
              rowIndex: block.baseRowIndex,
              columnIndex: globalTotalColumnIndex,
              relatedRowIndexes: [],
              relatedColumnIndexes: memberIndexes,
              evidence: { globalTotalBase, memberSum, groupLabel: group.label },
            });
          }
        }
      }
    }
  }

  return issues;
}

/**
 * Checks whether the displayed NPS value is consistent with Promoters − Detractors.
 *
 * Uses tolerances to account for rounding in displayed values:
 * - difference < 2pp → rounding noise, no issue
 * - 2–5pp → warning
 * - > 5pp → critical
 *
 * Only fires for npsStructure blocks that have all three rows present.
 */
function checkNpsMismatch(values, calculationBlocks) {
  const issues = [];

  for (const block of calculationBlocks) {
    if (block.metricType !== "npsStructure") continue;
    if (
      block.valueRowIndex === null ||
      block.promotersRowIndex === null ||
      block.detractorsRowIndex === null
    ) {
      continue;
    }

    const npsRow = values[block.valueRowIndex] || [];
    const promotersRow = values[block.promotersRowIndex] || [];
    const detractorsRow = values[block.detractorsRowIndex] || [];
    const columnCount = npsRow.length;

    for (let colIdx = 0; colIdx < columnCount; colIdx++) {
      const displayedNps = normalizeNpsValue(npsRow[colIdx]);
      const promotersValue = normalizeShare(promotersRow[colIdx]);
      const detractorsValue = normalizeShare(detractorsRow[colIdx]);

      if (displayedNps === null || promotersValue === null || detractorsValue === null) continue;

      const recalculatedNps = promotersValue - detractorsValue;
      const difference = Math.abs(displayedNps - recalculatedNps);

      if (difference <= NPS_MISMATCH_ROUNDING_TOLERANCE) continue;

      const severity = difference > NPS_MISMATCH_CRITICAL_THRESHOLD ? "critical" : "warning";

      issues.push({
        code: "NPS_MISMATCH",
        severity,
        message:
          `Row ${block.valueRowIndex + 1}, col ${colIdx + 1}: displayed NPS ` +
          `(${(displayedNps * 100).toFixed(1)}%) does not match ` +
          `Promoters − Detractors (${(recalculatedNps * 100).toFixed(1)}%). ` +
          `Difference: ${(difference * 100).toFixed(1)}%.`,
        rowIndex: block.valueRowIndex,
        columnIndex: colIdx,
        relatedRowIndexes: [block.promotersRowIndex, block.detractorsRowIndex],
        relatedColumnIndexes: [],
        evidence: {
          displayedNps,
          recalculatedNps,
          promotersValue,
          detractorsValue,
          difference,
          tolerance: NPS_MISMATCH_ROUNDING_TOLERANCE,
        },
      });
    }
  }

  return issues;
}

/** Reads a row from values as an array of numbers (null for non-numeric cells). */
function extractRowNumbers(values, rowIndex) {
  if (!values || rowIndex == null || !values[rowIndex]) return [];
  return values[rowIndex].map((v) => {
    if (v === null || v === undefined || v === "") return null;
    const n = Number(String(v).trim().replace(",", "."));
    return Number.isNaN(n) ? null : n;
  });
}

// ─── Quality summary ───────────────────────────────────────────────────────

function buildQualitySummary(issues) {
  const criticalCount = issues.filter((i) => i.severity === "critical").length;
  const warningCount = issues.filter((i) => i.severity === "warning").length;
  const infoCount = issues.filter((i) => i.severity === "info").length;

  return {
    criticalCount,
    warningCount,
    infoCount,
    hasBlockingIssues: criticalCount > 0,
  };
}

// ─── Summary ───────────────────────────────────────────────────────────────

function buildSummary(values, rowDiagnostics, calculationBlocks, bannerStructure) {
  const rowCount = values.length;
  const columnCount = rowCount > 0 ? (values[0] || []).length : 0;

  const detectedMetricRows = rowDiagnostics.filter(
    (r) => r.rowType !== "empty" && r.rowType !== "unknownText"
  ).length;

  const baseRows = rowDiagnostics.filter((r) => r.rowType === "base").length;

  const hasNps = calculationBlocks.some(
    (b) => b.metricType === "npsStructure" || b.metricType === "npsSpread"
  );
  const hasMeans = calculationBlocks.some((b) => b.metricType === "mean");

  const hasBanner = !!(bannerStructure && bannerStructure.isDetected);
  const hasGlobalTotal = !!(bannerStructure && bannerStructure.globalTotalColumnIndex !== null);
  const hasWaveGroups = !!(bannerStructure && bannerStructure.hasWaveGroups);

  return {
    rowCount,
    columnCount,
    detectedMetricRows,
    detectedBlocks: calculationBlocks.length,
    baseRows,
    hasNps,
    hasMeans,
    hasBanner,
    hasGlobalTotal,
    hasWaveGroups,
  };
}
