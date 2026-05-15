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

// All-100 row detection thresholds.
const ALL_100_MIN_COLUMNS = 2;
const ALL_100_MIN_FRACTION = 0.8;

// Row types recognized as metric/service rows — exempt from the all-100 check.
const METRIC_SERVICE_ROW_TYPES = new Set([
  "base",
  "nps",
  "npsScore",
  "mean",
  "sd",
  "standardDeviation",
  "variance",
]);

// Spreadsheet formula errors and programming error strings that must not appear as
// client-facing row labels.
const ERROR_LABEL_PATTERNS = [
  /^#(N\/A|VALUE!|REF!|DIV\/0!|NUM!|NAME\?|NULL!|GETTING_DATA)$/i,
  /\bnull\b/i,
  /\bnan\b/i,
  /\bundefined\b/i,
  /\[object\s+object\]/i,
];

// Placeholder and test-row keywords (English and Russian).
//
// "test" and "тест" are restricted to start-of-string to avoid false positives
// on legitimate research concepts such as "Concept Test" or "Product Test".
// "тестовая" is matched only when followed by "строка" (with space/dash/underscore)
// to avoid flagging valid labels like "Тестовая концепция" or "Тестовая упаковка".
// Cyrillic patterns use start-of-string anchors instead of \b because JavaScript
// \b is ASCII-only and does not form word boundaries around Cyrillic characters.
const PLACEHOLDER_LABEL_PATTERNS = [
  /\btodo\b/i,
  /\btbd\b/i,
  /^test(?:[\s\-_]|$)/i,
  /\bdummy\b/i,
  /\bplaceholder\b/i,
  /\btemp\b/i,
  /\bdelete\b/i,
  /\bremove\b/i,
  /\bignore\b/i,
  /\bxxx\b/i,
  /\basdf\b/i,
  /\bqwerty\b/i,
  /lorem\s+ipsum/i,
  /^тест(?:[\s\-_]|$)/i,
  /^тестовая[\s\-_]строка/i,
  /^удалить(?:[\s\-_]|$)/i,
  /^не\s+использовать(?:[\s\-_]|$)/i,
  /^заглушка(?:[\s\-_]|$)/i,
  /^временно(?:[\s\-_]|$)/i,
  /^черновик(?:[\s\-_]|$)/i,
];

const PREVIEW_NUMERIC_CELL_RE = /^[+-]?(\d+([.,]\d*)?|\d*[.,]\d+)%?$/;
const PREVIEW_NUMERIC_WITH_MARKER_SUFFIX_RE =
  /^([+-]?(\d+([.,]\d*)?|\d*[.,]\d+)%?)(\s+[\p{L}↑↓]+)+$/u;

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
  const calculationBlocks = buildPreviewBlocks(rawBlocks, rawRowDiagnostics, safeValues);

  // Data quality analysis.
  const dataQualityIssues = [
    ...checkNumericLikeLabels(rowDiagnostics),
    ...checkSuspiciousNumericLabels(rowDiagnostics),
    ...checkSuspiciousAll100Rows(safeValues, rowDiagnostics),
    ...checkMissingRowLabelWithData(safeValues, rowDiagnostics),
    ...checkSuspiciousErrorLabels(rowDiagnostics),
    ...checkSuspiciousPlaceholderLabels(rowDiagnostics),
    ...checkSuspiciousCodeLikeLabels(rowDiagnostics),
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

  const detected = detectBannerStructure(adaptBannerContextForDetection(bannerContext), settings);

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

// Accept both normalizer-style banner context ({ scanRows, columnCount })
// and detector-style banner context ({ selectedColumnCount, lowerBannerRow, upperScanRows }).
function adaptBannerContextForDetection(bannerContext) {
  if (!bannerContext) {
    return null;
  }

  if (bannerContext.selectedColumnCount !== undefined) {
    return bannerContext;
  }

  const scanRows = Array.isArray(bannerContext.scanRows) ? bannerContext.scanRows : [];
  const selectedColumnCount = bannerContext.columnCount || 0;

  if (!selectedColumnCount || scanRows.length === 0) {
    return {
      selectedColumnCount,
      lowerBannerRow: [],
      upperScanRows: [],
      messages: bannerContext.messages || [],
    };
  }

  return {
    selectedColumnCount,
    lowerBannerRow: scanRows[scanRows.length - 1],
    upperScanRows: scanRows.slice(0, -1).reverse(),
    messages: bannerContext.messages || [],
  };
}

// ─── Row diagnostics ───────────────────────────────────────────────────────

function buildPreviewRowDiagnostics(rawDiagnostics, leftLabelValues) {
  return (rawDiagnostics || []).map((diag) => {
    const leftRow = leftLabelValues[diag.rowIndex] || [];
    const labelParts = extractLabelPartsFromRow(leftRow);

    // Rightmost meaningful part (non-symbol-only) is closest to the data column
    // and treated as the primary label. Symbol-only trailing parts like "%" are
    // skipped so a unit column never shadows the real descriptive label.
    // Falls back to the actual rightmost part when all parts are symbol-only.
    const primaryLabel =
      labelParts.length > 0
        ? ([...labelParts].reverse().find((p) => !isEmptyOrSymbolOnlyLabel(p)) ||
            labelParts[labelParts.length - 1])
        : "";
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

function buildPreviewBlocks(rawBlocks, rowDiagnostics, values) {
  const blocks = (rawBlocks || [])
    .map(normalizeBlockShape)
    .filter((block) => blockHasPreviewEvidence(block, rowDiagnostics, values));
  enrichNpsBlocksWithNeutralRow(blocks);
  return blocks;
}

function blockHasPreviewEvidence(block, rowDiagnostics, values) {
  const valueRowIndexes = Array.isArray(block.valueRowIndexes) ? block.valueRowIndexes : [];

  switch (block.metricType) {
    case "proportion":
      if (valueRowIndexes.length === 0) return false;
      if (!valueRowIndexes.some((rowIndex) => rowHasNumericEvidence(values, rowIndex))) return false;
      if (!rowHasNumericEvidence(values, block.baseRowIndex)) return false;
      if (isFallbackOnlyUnknownBlock(block, rowDiagnostics) && !hasNonHeaderLikeValueLabel(block, rowDiagnostics)) {
        return false;
      }
      return true;

    case "mean":
      return (
        rowHasNumericEvidence(values, block.valueRowIndex) &&
        rowHasNumericEvidence(values, block.baseRowIndex)
      );

    case "npsStructure":
      return (
        rowHasNumericEvidence(values, block.valueRowIndex) &&
        rowHasNumericEvidence(values, block.promotersRowIndex) &&
        rowHasNumericEvidence(values, block.detractorsRowIndex) &&
        rowHasNumericEvidence(values, block.baseRowIndex)
      );

    case "npsSpread":
      return (
        rowHasNumericEvidence(values, block.valueRowIndex) &&
        (rowHasNumericEvidence(values, block.sdRowIndex) ||
          rowHasNumericEvidence(values, block.varianceRowIndex)) &&
        rowHasNumericEvidence(values, block.baseRowIndex)
      );

    default:
      return false;
  }
}

function rowHasNumericEvidence(values, rowIndex) {
  if (rowIndex === null || rowIndex === undefined) return false;
  const row = Array.isArray(values) ? values[rowIndex] : null;
  if (!Array.isArray(row)) return false;
  return row.some((cell) => isPreviewNumericCellValue(cell));
}

function isFallbackOnlyUnknownBlock(block, rowDiagnostics) {
  if (block.metricType !== "proportion") return false;

  const rowIndexes = [...(block.valueRowIndexes || []), block.baseRowIndex].filter(
    (rowIndex) => rowIndex !== null && rowIndex !== undefined
  );

  if (rowIndexes.length < 2) return false;

  return rowIndexes.every((rowIndex) => {
    const rowType = rowDiagnostics[rowIndex]?.rowType;
    return rowType === "unknownText" || rowType === "empty";
  });
}

function hasNonHeaderLikeValueLabel(block, rowDiagnostics) {
  return (block.valueRowIndexes || []).some((rowIndex) => {
    const row = rowDiagnostics[rowIndex];
    const label = row?.primaryLabel || row?.label || "";
    return !!label && !looksLikePreviewHeaderLabel(label);
  });
}

function looksLikePreviewHeaderLabel(label) {
  const text = String(label || "").trim();
  if (!text) return true;

  const normalized = normalizeLabelText(text);
  if (!normalized) return true;

  if (
    normalized.includes("wave") ||
    normalized.includes("quarter") ||
    normalized.includes("period") ||
    normalized.includes("volna") ||
    normalized.includes("kvartal") ||
    normalized.includes("period")
  ) {
    return true;
  }

  return /^(\d{4}q[1-4]|q[1-4]\d{4})$/i.test(normalized);
}

function isPreviewNumericCellValue(cell) {
  if (typeof cell === "number") {
    return !Number.isNaN(cell);
  }

  if (typeof cell !== "string") {
    return false;
  }

  const trimmed = cell.trim();
  if (!trimmed) {
    return false;
  }

  return (
    PREVIEW_NUMERIC_CELL_RE.test(trimmed) ||
    PREVIEW_NUMERIC_WITH_MARKER_SUFFIX_RE.test(trimmed)
  );
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
 * Flags row labels that look like a numeric range and were not classified as a
 * known metric type.
 *
 * Ordered category blocks (e.g. "20-29" / "30-39" / "40-49") are suppressed:
 * when at least one neighbor within ±2 rows also has a range-like label the
 * row is treated as part of a category block and no warning is emitted.
 *
 * Note: isolated range labels can still be valid for NPS scale rows or wave
 * labels — this check produces warnings, not errors.
 */
function checkNumericLikeLabels(rowDiagnostics) {
  const issues = [];

  for (let i = 0; i < rowDiagnostics.length; i++) {
    const row = rowDiagnostics[i];
    if (!row.label) continue;
    if (row.rowType !== "unknownText" && row.rowType !== "empty") continue;
    if (!looksLikeNumericRange(row.label)) continue;
    if (isInNumericRangeCategoryBlock(rowDiagnostics, i)) continue;

    issues.push({
      code: "NUMERIC_LIKE_LABEL",
      severity: "warning",
      message: `Row ${row.rowIndex + 1}: label "${row.label}" looks like a numeric range but was not matched to a known metric type. May be valid for NPS scales, age groups, or wave numbers.`,
      rowIndex: row.rowIndex,
      columnIndex: null,
      relatedRowIndexes: [],
      relatedColumnIndexes: [],
      evidence: { label: row.label, rowType: row.rowType },
    });
  }

  return issues;
}

/**
 * Flags row labels that look like a single numeric value (integer or decimal)
 * on rows that are not recognized service/metric types.
 *
 * Examples: "42", "3.5", "61,00" — may be uncoded values or export artifacts.
 * Fires for any non-service row type including "proportion", because the metric
 * detector may classify unknown rows as proportion rows by default.
 *
 * Does not fire when the row sits within a numeric category block — i.e. when
 * nearby rows also carry numeric range or single-numeric category labels
 * (NPS/rating scales such as 1 / 2 / 3 / 4 are suppressed this way).
 */
function checkSuspiciousNumericLabels(rowDiagnostics) {
  const issues = [];

  for (let i = 0; i < rowDiagnostics.length; i++) {
    const row = rowDiagnostics[i];
    if (!row.label) continue;
    if (row.rowType === "empty") continue;
    if (METRIC_SERVICE_ROW_TYPES.has(row.rowType)) continue;
    if (!looksLikeSingleNumericValue(row.label)) continue;
    if (isInNumericRangeCategoryBlock(rowDiagnostics, i)) continue;

    issues.push({
      code: "SUSPICIOUS_NUMERIC_LABEL",
      severity: "warning",
      message: `Row ${row.rowIndex + 1}: label "${row.label}" looks like a numeric value and may be an uncoded value or export/labeling issue.`,
      rowIndex: row.rowIndex,
      columnIndex: null,
      relatedRowIndexes: [],
      relatedColumnIndexes: [],
      evidence: { label: row.label, rowType: row.rowType },
    });
  }

  return issues;
}

/**
 * Flags rows where all or most non-empty numeric cells are 100% or equivalent,
 * unless the row is a recognized metric/service row type.
 *
 * Uses cellLooksLike100Percent() so that percent-string cells ("100%", "1%")
 * are evaluated in their display scale, not converted to a bare number first.
 *
 * Such rows may be service rows, test rows, or uncoded rows that should not
 * appear in client-facing tables.
 */
function checkSuspiciousAll100Rows(values, rowDiagnostics) {
  const issues = [];

  for (const row of rowDiagnostics) {
    if (row.rowType === "empty") continue;
    if (METRIC_SERVICE_ROW_TYPES.has(row.rowType)) continue;

    const rawRow = values && values[row.rowIndex];
    if (!rawRow) continue;

    let nonEmptyCount = 0;
    let all100Count = 0;

    for (const v of rawRow) {
      if (v === null || v === undefined || v === "") continue;
      const s = String(v).trim();
      const numStr = s.endsWith("%")
        ? s.slice(0, -1).trim().replace(",", ".")
        : s.replace(",", ".");
      if (Number.isNaN(Number(numStr))) continue;
      nonEmptyCount++;
      if (cellLooksLike100Percent(v)) all100Count++;
    }

    if (nonEmptyCount < ALL_100_MIN_COLUMNS) continue;
    if (all100Count / nonEmptyCount < ALL_100_MIN_FRACTION) continue;

    issues.push({
      code: "SUSPICIOUS_ALL_100_ROW",
      severity: "warning",
      message:
        `Row ${row.rowIndex + 1}: label "${row.label || "(no label)"}" contains 100% or equivalent ` +
        `across ${all100Count} of ${nonEmptyCount} column(s). ` +
        `May be a service, test, or uncoded row that should be checked before client delivery.`,
      rowIndex: row.rowIndex,
      columnIndex: null,
      relatedRowIndexes: [],
      relatedColumnIndexes: [],
      evidence: {
        label: row.label,
        all100Count,
        totalNonEmpty: nonEmptyCount,
        rowType: row.rowType,
      },
    });
  }

  return issues;
}

/** Returns true for labels that look like a numeric range: "20-29", "18–24", "25—34". */
function looksLikeNumericRange(label) {
  const s = String(label).trim();
  return /^\d+[\-–—]\d+$/.test(s);
}

/** Returns true for labels that look like a single numeric value: "42", "3.5", "61,00". */
function looksLikeSingleNumericValue(label) {
  const s = String(label).trim();
  if (/^\d+$/.test(s)) return true;
  if (/^\d+\.\d+$/.test(s)) return true;
  if (/^\d+,\d+$/.test(s)) return true;
  return false;
}

/**
 * Returns true if rawValue looks like 100%, respecting the storage scale.
 *
 * Percent-string cells (e.g. "100%", "100,0%"):
 *   Strip "%" and check the display number is approximately 100 (99.5–100.5).
 *   "1%" → display value 1 → false.  "100%" → display value 100 → true.
 *
 * Numeric values (plain numbers or numeric strings without "%"):
 *   Check percent scale (≈100) OR share scale (≈1.0).
 *   1 → true (Excel stores 100% as 1.0 in share/decimal scale).
 *   100 → true.  0.5 → false.  1 from "1%" is never reached here.
 */
function cellLooksLike100Percent(rawValue) {
  if (rawValue === null || rawValue === undefined || rawValue === "") return false;
  const s = String(rawValue).trim();
  if (s.endsWith("%")) {
    const n = Number(s.slice(0, -1).trim().replace(",", "."));
    return !Number.isNaN(n) && n >= 99.5 && n <= 100.5;
  }
  const n = Number(s.replace(",", "."));
  if (Number.isNaN(n)) return false;
  return (n >= 99.5 && n <= 100.5) || (n >= 0.995 && n <= 1.005);
}

/**
 * Returns true if the row at arrayIndex sits within a numeric category block.
 *
 * A category block is detected when at least one neighbor within ±2 positions
 * has a label that looks like either:
 * - a numeric range ("20-29", "30-39") — age/category ranges, or
 * - a single numeric value ("1", "2", "3.5") — NPS/rating scale rows.
 *
 * Used to suppress NUMERIC_LIKE_LABEL and SUSPICIOUS_NUMERIC_LABEL for labels
 * that are part of an ordered scale or category group.
 */
function isInNumericRangeCategoryBlock(rowDiagnostics, arrayIndex) {
  for (let delta = -2; delta <= 2; delta++) {
    if (delta === 0) continue;
    const ni = arrayIndex + delta;
    if (ni < 0 || ni >= rowDiagnostics.length) continue;
    const neighbor = rowDiagnostics[ni];
    if (!neighbor || !neighbor.label) continue;
    if (looksLikeNumericRange(neighbor.label)) return true;
    if (looksLikeSingleNumericValue(neighbor.label)) return true;
  }
  return false;
}

/**
 * Flags rows that contain at least 2 non-empty numeric cells but have no meaningful
 * row label — i.e. the label is empty, whitespace-only, or composed entirely of
 * symbol-only placeholder characters such as "-", "—", ".", "*".
 *
 * Skips recognized metric/service rows to avoid noise on intentionally label-free
 * support rows.
 */
function checkMissingRowLabelWithData(values, rowDiagnostics) {
  const issues = [];

  for (const row of rowDiagnostics) {
    if (METRIC_SERVICE_ROW_TYPES.has(row.rowType)) continue;
    if (!isEmptyOrSymbolOnlyLabel(row.primaryLabel)) continue;

    const rowNumbers = extractRowNumbers(values, row.rowIndex);
    if (rowNumbers.filter((v) => v !== null).length < 2) continue;

    const displayLabel = row.primaryLabel || row.label || "(empty)";
    issues.push({
      code: "MISSING_ROW_LABEL_WITH_DATA",
      severity: "warning",
      message: `Row ${row.rowIndex + 1}: row has data values but no meaningful label (label: "${displayLabel}"). Check that this row is not missing a category label.`,
      rowIndex: row.rowIndex,
      columnIndex: null,
      relatedRowIndexes: [],
      relatedColumnIndexes: [],
      evidence: { primaryLabel: row.primaryLabel, label: row.label, rowType: row.rowType },
    });
  }

  return issues;
}

/**
 * Flags rows whose label contains obvious spreadsheet formula errors (#N/A, #VALUE!,
 * etc.) or programming error strings (null, NaN, undefined, [object Object]).
 *
 * Applies to all non-empty rows regardless of rowType — an error artifact as a
 * label is always suspicious.
 */
function checkSuspiciousErrorLabels(rowDiagnostics) {
  const issues = [];

  for (const row of rowDiagnostics) {
    if (row.rowType === "empty") continue;
    if (!row.label) continue;

    const s = String(row.label).trim();
    if (!ERROR_LABEL_PATTERNS.some((p) => p.test(s))) continue;

    issues.push({
      code: "SUSPICIOUS_ERROR_LABEL",
      severity: "warning",
      message: `Row ${row.rowIndex + 1}: label "${row.label}" looks like a spreadsheet or programming error value and is almost certainly not a client-facing category.`,
      rowIndex: row.rowIndex,
      columnIndex: null,
      relatedRowIndexes: [],
      relatedColumnIndexes: [],
      evidence: { label: row.label, rowType: row.rowType },
    });
  }

  return issues;
}

/**
 * Flags rows whose label matches a known placeholder or test-row keyword in
 * English or Russian.
 *
 * Uses word-boundary matching to avoid catching the keyword as part of a longer
 * legitimate label (e.g. "placeholder" fires, "placeholder value" also fires,
 * but "temperature" does not fire on "temp").
 *
 * Skips recognized metric/service rows.
 */
function checkSuspiciousPlaceholderLabels(rowDiagnostics) {
  const issues = [];

  for (const row of rowDiagnostics) {
    if (row.rowType === "empty") continue;
    if (!row.label) continue;
    if (METRIC_SERVICE_ROW_TYPES.has(row.rowType)) continue;

    const s = String(row.label).trim();
    if (!PLACEHOLDER_LABEL_PATTERNS.some((p) => p.test(s))) continue;

    issues.push({
      code: "SUSPICIOUS_PLACEHOLDER_LABEL",
      severity: "warning",
      message: `Row ${row.rowIndex + 1}: label "${row.label}" looks like a test or placeholder row. Check that this row has not been accidentally left in the client table.`,
      rowIndex: row.rowIndex,
      columnIndex: null,
      relatedRowIndexes: [],
      relatedColumnIndexes: [],
      evidence: { label: row.label, rowType: row.rowType },
    });
  }

  return issues;
}

/**
 * Flags rows whose label looks like a raw variable or code name rather than a
 * human-readable category label.
 *
 * Heuristic:
 * - label contains an underscore (strong code-naming signal), OR
 * - label starts with 1–4 letters followed by 2+ digits (e.g. q12, var005), OR
 * - label follows an alternating letter-digit-letter-digit pattern (e.g. d1r3).
 * - label must have no spaces (code names do not have spaces).
 *
 * Skips recognized metric/service rows (their labels are already validated by
 * the detector) and rows with empty labels.
 */
function checkSuspiciousCodeLikeLabels(rowDiagnostics) {
  const issues = [];

  for (const row of rowDiagnostics) {
    if (row.rowType === "empty") continue;
    if (METRIC_SERVICE_ROW_TYPES.has(row.rowType)) continue;
    if (!row.label) continue;
    if (!looksLikeCodeLabel(row.label)) continue;

    issues.push({
      code: "SUSPICIOUS_CODE_LIKE_LABEL",
      severity: "warning",
      message: `Row ${row.rowIndex + 1}: label "${row.label}" looks like a variable or code name rather than a client-facing category. Check that this row label has not been left uncoded.`,
      rowIndex: row.rowIndex,
      columnIndex: null,
      relatedRowIndexes: [],
      relatedColumnIndexes: [],
      evidence: { label: row.label, rowType: row.rowType },
    });
  }

  return issues;
}

/**
 * Returns true when the label has no meaningful alphanumeric content.
 * Catches empty strings, whitespace-only strings, and symbol-only placeholders
 * such as "-", "—", ".", "*", "---", "?".
 */
function isEmptyOrSymbolOnlyLabel(label) {
  if (!label || !String(label).trim()) return true;
  return !/[a-zA-Zа-яА-ЯёЁ0-9]/.test(String(label).trim());
}

/**
 * Returns true when the label looks like a variable or code name.
 * No spaces are allowed (code names don't have spaces).
 *
 * Triggers on:
 * - underscore AND a code-like digit pattern:
 *     ends with _digits  → q1_1, Q12_3, var_005, brand_99
 *     starts with ≤3 letters + digits + underscore  → q1_, Q12_
 *   Underscore alone is NOT enough — labels like Top_2_Box, No_answer, Brand_A
 *   do not end with digits or start with a short letter+digit prefix.
 * - 1–4 letters + 2+ digits (no underscore): q12, var005, brand99
 * - alternating letter-digit-letter-digit: d1r3, b2c4
 */
function looksLikeCodeLabel(label) {
  const s = String(label).trim();
  if (/\s/.test(s)) return false;
  if (s.length < 2) return false;
  if (/_/.test(s)) {
    // Require a code-like digit pattern alongside the underscore.
    return /_\d+$/.test(s) || /^[a-zA-Z]{1,3}\d+_/.test(s);
  }
  if (/^[a-zA-Z]{1,4}\d{2,}/.test(s)) return true;
  if (/^[a-zA-Z]\d[a-zA-Z]\d/.test(s)) return true;
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
