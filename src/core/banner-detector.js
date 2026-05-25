/**
 * Banner structure detector.
 *
 * PURPOSE:
 * Detect column banner structure above selected data range.
 *
 * CURRENT STAGE:
 * Detection-only MVP.
 * Supports:
 * - one-level banner;
 * - repeated adjacent upper labels as group fallback;
 * - basic Total detection;
 * - status messages.
 *
 * DOES NOT:
 * - read Excel directly;
 * - write to Excel;
 * - calculate significance;
 * - depend on Office.js.
 */

import { BANNER_DICTIONARY } from "./config/dictionary.config";
import { normalizeLookupText, normalizeDisplayText } from "./string-utils";

const TOTAL_LABEL_KEYWORDS = BANNER_DICTIONARY.totalLabels;
const WAVE_GROUP_LABEL_KEYWORDS = BANNER_DICTIONARY.waveGroupLabels;
const TECHNICAL_WAVE_DESCRIPTOR_KEYWORDS = [
  ...WAVE_GROUP_LABEL_KEYWORDS,
  "quarter",
  "quarters",
  "квартал",
  "кварталы",
];

const TECHNICAL_WAVE_LABEL_DOMINANCE_THRESHOLD = 0.7;

const DEFAULT_GROUP_KEY = "group:default";
const DEFAULT_GROUP_LABEL = "Default";

const BANNER_MODE = {
  ONE_LEVEL: "oneLevel",
  TWO_LEVEL: "twoLevel",
  FALLBACK: "fallback",
};

const MESSAGE_SEVERITY = {
  DEBUG: "debug",
  INFO: "info",
  WARNING: "warning",
  ERROR: "error",
};

const TOTAL_TYPE = {
  LOCAL: "local",
  GLOBAL: "global",
};

const GROUP_SEMANTIC_TYPE = {
  DEFAULT: "default",
  WAVE: "wave",
};

const RECOMMENDED_COMPARISON_MODE = {
  DEFAULT: "default",
  PREVIOUS_COLUMN: "previousColumn",
  MIXED: "mixed",
};

/**
 * Main banner detection entry point.
 */
export function detectBannerStructure(bannerContext, settings = {}) {
  const selectedColumnCount = bannerContext.selectedColumnCount || 0;
  const lowerBannerRow = getLowerBannerRow(bannerContext);
  const upperScanRows = bannerContext.upperScanRows || [];

  const messages = [];

  if (!selectedColumnCount || selectedColumnCount < 1) {
    return createFallbackBannerStructure(selectedColumnCount, [
      createBannerMessage(
        MESSAGE_SEVERITY.WARNING,
        "BANNER_EMPTY_SELECTION",
        "Баннер: не удалось определить количество колонок выделения."
      ),
    ]);
  }

  if (!lowerBannerRow || lowerBannerRow.length === 0) {
    return createFallbackBannerStructure(selectedColumnCount, [
      createBannerMessage(
        MESSAGE_SEVERITY.WARNING,
        "BANNER_LOWER_LEVEL_NOT_FOUND",
        "Баннер: строка непосредственно над выделением не найдена. Используется fallback-группа."
      ),
    ]);
  }

  const normalizedLowerBannerRow = normalizeBannerRowLength(lowerBannerRow, selectedColumnCount);

  const groupLevelResult = detectMeaningfulGroupLevel(
    upperScanRows,
    normalizedLowerBannerRow,
    selectedColumnCount
  );

  if (groupLevelResult.message) {
    messages.push(groupLevelResult.message);
  }

  const columnDescriptors = buildColumnDescriptors({
    selectedColumnCount,
    lowerBannerRow: normalizedLowerBannerRow,
    groupLevel: groupLevelResult.groupLevel,
  });

  markUpperLevelTotalsForSparseLowerLabels(columnDescriptors, upperScanRows);

  const globalTotalColumnIndex = detectGlobalTotalColumnIndex(columnDescriptors);

  markGlobalTotalColumn(columnDescriptors, globalTotalColumnIndex);

  const nestedWaveGroupKeys =
    settings && settings.autoDetectWaveBanners
      ? detectGroupKeysWithNestedWaveDimension({
          columnDescriptors,
          lowerBannerRow: normalizedLowerBannerRow,
          upperScanRows,
          groupLevelRowOffset: groupLevelResult.groupLevel
            ? groupLevelResult.groupLevel.rowOffset
            : null,
        })
      : new Set();

  const groups = buildGroupsFromColumnDescriptors(
    columnDescriptors,
    settings,
    nestedWaveGroupKeys
  );

  const waveGroups = groups.filter(
    (group) => group.recommendedComparisonMode === RECOMMENDED_COMPARISON_MODE.PREVIOUS_COLUMN
  );

  if (waveGroups.length > 0) {
    messages.push(
      createBannerMessage(
        MESSAGE_SEVERITY.INFO,
        "BANNER_WAVE_GROUPS_DETECTED",
        `Баннер: обнаружены волновые группы: ${waveGroups.map((group) => group.label).join(", ")}.`
      )
    );
  }

  const totalColumnIndexes = columnDescriptors
    .filter((descriptor) => descriptor.isTotal)
    .map((descriptor) => descriptor.columnIndex);

  const mode = groupLevelResult.groupLevel ? BANNER_MODE.TWO_LEVEL : BANNER_MODE.ONE_LEVEL;

  messages.push(
    createBannerMessage(
      MESSAGE_SEVERITY.DEBUG,
      "BANNER_DETECTED",
      formatBannerDetectedMessage({
        mode,
        selectedColumnCount,
        groups,
        totalColumnIndexes,
      })
    )
  );

  return {
    isDetected: true,
    mode,
    columnDescriptors,
    groups,
    globalTotalColumnIndex,
    totalColumnIndexes,
    hasWaveGroups: waveGroups.length > 0,
    recommendedComparisonMode:
      waveGroups.length > 0
        ? RECOMMENDED_COMPARISON_MODE.MIXED
        : RECOMMENDED_COMPARISON_MODE.DEFAULT,
    messages,
  };
}

/**
 * Returns lower banner row from context.
 */
function getLowerBannerRow(bannerContext) {
  if (!bannerContext) {
    return [];
  }

  if (bannerContext.lowerBannerRow) {
    return bannerContext.lowerBannerRow;
  }

  if (bannerContext.lowerBannerRows && bannerContext.lowerBannerRows.length > 0) {
    return bannerContext.lowerBannerRows[0];
  }

  return [];
}

/**
 * Builds column descriptors.
 */
function buildColumnDescriptors({ selectedColumnCount, lowerBannerRow, groupLevel }) {
  const descriptors = [];

  for (let columnIndex = 0; columnIndex < selectedColumnCount; columnIndex++) {
    const lowerLabel = normalizeRawBannerCellValue(lowerBannerRow[columnIndex]);
    const normalizedLowerLabel = normalizeBannerLabel(lowerLabel);

    const groupLabel = groupLevel
      ? normalizeRawBannerCellValue(groupLevel.labels[columnIndex])
      : DEFAULT_GROUP_LABEL;

    const normalizedGroupLabel = normalizeBannerLabel(groupLabel);

    const comparisonGroupKey = groupLevel
      ? buildGroupKey(groupLabel, groupLevel.spansByColumnIndex[columnIndex])
      : DEFAULT_GROUP_KEY;

    const isTotal = isTotalBannerLabel(normalizedLowerLabel);

    descriptors.push({
      columnIndex,

      lowerLabel,
      normalizedLowerLabel,

      bannerPath: groupLevel ? [groupLabel, lowerLabel] : [lowerLabel],
      displayLabel: groupLevel ? `${groupLabel} / ${lowerLabel}` : lowerLabel,

      comparisonGroupKey,
      comparisonGroupLabel: groupLevel ? groupLabel : DEFAULT_GROUP_LABEL,

      isTotal,
      totalType: isTotal ? TOTAL_TYPE.LOCAL : null,

      isGlobalTotal: false,
      isLocalTotal: isTotal,

      source: {
        lowerLevelRowOffset: 0,
        groupLevelRowOffset: groupLevel ? groupLevel.rowOffset : null,
        mergeArea: null,
      },
    });
  }

  return descriptors;
}

/**
 * Builds group objects from descriptors.
 */
function buildGroupsFromColumnDescriptors(
  columnDescriptors,
  calculationSettings = {},
  nestedWaveGroupKeys = new Set()
) {
  const groupsByKey = new Map();

  for (const descriptor of columnDescriptors) {
    if (!groupsByKey.has(descriptor.comparisonGroupKey)) {
      const shouldAutoDetectWaveBanners =
        calculationSettings && calculationSettings.autoDetectWaveBanners;

      const isWaveGroup =
        shouldAutoDetectWaveBanners &&
        (isWaveGroupLabel(descriptor.comparisonGroupLabel) ||
          nestedWaveGroupKeys.has(descriptor.comparisonGroupKey));

      groupsByKey.set(descriptor.comparisonGroupKey, {
        groupKey: descriptor.comparisonGroupKey,
        label: descriptor.comparisonGroupLabel,
        bannerPath:
          descriptor.comparisonGroupLabel === DEFAULT_GROUP_LABEL
            ? []
            : [descriptor.comparisonGroupLabel],
        columnIndexes: [],
        localTotalColumnIndexes: [],
        hasLocalTotal: false,

        semanticType: isWaveGroup ? GROUP_SEMANTIC_TYPE.WAVE : GROUP_SEMANTIC_TYPE.DEFAULT,

        recommendedComparisonMode: isWaveGroup
          ? RECOMMENDED_COMPARISON_MODE.PREVIOUS_COLUMN
          : RECOMMENDED_COMPARISON_MODE.DEFAULT,
      });
    }

    const group = groupsByKey.get(descriptor.comparisonGroupKey);

    group.columnIndexes.push(descriptor.columnIndex);

    if (descriptor.isTotal && descriptor.totalType !== TOTAL_TYPE.GLOBAL) {
      group.localTotalColumnIndexes.push(descriptor.columnIndex);
      group.hasLocalTotal = true;
    }
  }

  return Array.from(groupsByKey.values());
}

/**
 * Detects meaningful group level in upper scan rows.
 *
 * Detection strategies:
 * 1. Reconstructed spans:
 *    non-empty upper label + following empty cells,
 *    constrained by continuous non-empty lower banner area.
 * 2. Repeated adjacent labels fallback:
 *    Gender | Gender | Gender.
 */
function detectMeaningfulGroupLevel(upperScanRows, lowerBannerRow, selectedColumnCount) {
  if (!upperScanRows || upperScanRows.length === 0) {
    return {
      groupLevel: null,
      message: createBannerMessage(
        MESSAGE_SEVERITY.DEBUG,
        "BANNER_ONE_LEVEL_FALLBACK",
        "Баннер: верхний уровень групп не найден. Используется один уровень баннера."
      ),
    };
  }

  const candidates = [];

  for (let rowIndex = 0; rowIndex < upperScanRows.length; rowIndex++) {
    const row = normalizeBannerRowLength(upperScanRows[rowIndex], selectedColumnCount);

    const reconstructedSpanResult = detectReconstructedSpanGroupLevel(
      row,
      lowerBannerRow,
      selectedColumnCount,
      rowIndex
    );

    if (reconstructedSpanResult.groupLevel) {
      candidates.push(reconstructedSpanResult);
      continue;
    }

    const repeatedLabelResult = detectRepeatedLabelGroupLevelInRow(
      row,
      selectedColumnCount,
      rowIndex
    );

    if (repeatedLabelResult.groupLevel) {
      candidates.push(repeatedLabelResult);
    }
  }

  if (candidates.length > 0) {
    return selectBestGroupLevelCandidate(candidates);
  }

  return {
    groupLevel: null,
    message: createBannerMessage(
      MESSAGE_SEVERITY.INFO,
      "BANNER_ONE_LEVEL_FALLBACK",
      "Баннер: верхний уровень групп не найден. Используется один уровень баннера."
    ),
  };
}

/**
 * Picks the best upper banner level.
 *
 * Rows like "Wave" / "Wave (quarter)" are technical descriptors for the
 * lower banner values. If a higher semantic group row exists, use that higher
 * row as the comparison-group level instead of fragmenting every wave pair
 * into its own group.
 */
function selectBestGroupLevelCandidate(candidates) {
  const firstCandidate = candidates[0];

  if (!isTechnicalWaveDescriptorGroupLevel(firstCandidate.groupLevel)) {
    return firstCandidate;
  }

  const semanticCandidate = candidates.find(
    (candidate) => !isTechnicalWaveDescriptorGroupLevel(candidate.groupLevel)
  );

  return semanticCandidate || firstCandidate;
}

/**
 * Temporary diagnostic helper for Excel smoke debugging.
 *
 * Mirrors detectMeaningfulGroupLevel candidate discovery and explains which
 * upper rows were candidate group levels before the final detector result is
 * consumed by taskpane banner-letter writing.
 */
export function buildBannerDetectionDebugSummary(bannerContext) {
  const selectedColumnCount = bannerContext ? bannerContext.selectedColumnCount || 0 : 0;
  const lowerBannerRow = normalizeBannerRowLength(
    getLowerBannerRow(bannerContext),
    selectedColumnCount
  );
  const upperScanRows = bannerContext && bannerContext.upperScanRows ? bannerContext.upperScanRows : [];

  const rows = [];
  const candidates = [];

  for (let rowIndex = 0; rowIndex < upperScanRows.length; rowIndex++) {
    const row = normalizeBannerRowLength(upperScanRows[rowIndex], selectedColumnCount);
    const reconstructedSpanResult = detectReconstructedSpanGroupLevel(
      row,
      lowerBannerRow,
      selectedColumnCount,
      rowIndex
    );

    let candidateResult = null;

    if (reconstructedSpanResult.groupLevel) {
      candidateResult = reconstructedSpanResult;
    } else {
      const repeatedLabelResult = detectRepeatedLabelGroupLevelInRow(
        row,
        selectedColumnCount,
        rowIndex
      );

      if (repeatedLabelResult.groupLevel) {
        candidateResult = repeatedLabelResult;
      }
    }

    if (candidateResult) {
      candidates.push(candidateResult);
    }

    const groupLevel = candidateResult ? candidateResult.groupLevel : null;
    const spans = groupLevel && groupLevel.spans ? groupLevel.spans : [];
    const labeledSpans = spans.filter((span) => span.label);
    const isTechnicalWaveDescriptor = isTechnicalWaveDescriptorGroupLevel(groupLevel);

    rows.push({
      rowIndex,
      bottomUpLevel: rowIndex + 2,
      row,
      isCandidate: Boolean(candidateResult),
      detectionMethod: groupLevel ? groupLevel.detectionMethod : null,
      isTechnicalWaveDescriptor,
      spanCount: spans.length,
      groupCount: labeledSpans.length,
      score: groupLevel ? (isTechnicalWaveDescriptor ? 0 : 1) : null,
      sampleSpans: labeledSpans.slice(0, 20).map((span) => ({
        label: span.label,
        startColumnIndex: span.startColumnIndex,
        endColumnIndex: span.endColumnIndex,
        columnIndexes: span.columnIndexes,
      })),
      messageCode: candidateResult && candidateResult.message ? candidateResult.message.code : null,
    });
  }

  const selectedCandidate = candidates.length > 0 ? selectBestGroupLevelCandidate(candidates) : null;
  const selectedGroupLevel = selectedCandidate ? selectedCandidate.groupLevel : null;

  return {
    selectedColumnCount,
    lowerBannerRow,
    upperScanRows,
    upperScanRowsByBottomUpLevel: upperScanRows.map((row, index) => ({
      level: index + 2,
      row,
    })),
    candidateRows: rows,
    selectedCandidate: selectedGroupLevel
      ? {
          bottomUpLevel: Math.abs(selectedGroupLevel.rowOffset) + 1,
          rowOffset: selectedGroupLevel.rowOffset,
          detectionMethod: selectedGroupLevel.detectionMethod,
          isTechnicalWaveDescriptor: isTechnicalWaveDescriptorGroupLevel(selectedGroupLevel),
          spanCount: selectedGroupLevel.spans ? selectedGroupLevel.spans.length : 0,
          sampleSpans: (selectedGroupLevel.spans || []).slice(0, 20).map((span) => ({
            label: span.label,
            startColumnIndex: span.startColumnIndex,
            endColumnIndex: span.endColumnIndex,
            columnIndexes: span.columnIndexes,
          })),
        }
      : null,
  };
}

/**
 * Detects merged-like group level from upper row.
 *
 * Pattern:
 *   Age | "" | ""
 * Lower:
 *   Total | 18-24 | 25-34
 *
 * Result:
 *   Age spans columns 0..2.
 */
function detectReconstructedSpanGroupLevel(
  upperRow,
  lowerBannerRow,
  selectedColumnCount,
  rowIndex
) {
  const rawSpans = buildReconstructedSpansFromUpperRow(upperRow, lowerBannerRow, selectedColumnCount);

  // Merge consecutive adjacent single-column wave-value spans whose lower-
  // banner cell is blank.  Without this, quarter labels like "2025Q4" / "2026Q1"
  // that appear without an explicit parent-group label (e.g., when the label
  // column was stripped from a partial selection) each create a separate
  // one-column group, which makes both data columns receive local label "a".
  // After merging, the run forms one multi-column span whose columns get "a"
  // and "b" respectively.
  const spans = mergeAdjacentWaveValueSpans(rawSpans, lowerBannerRow);

  const meaningfulSpans = spans.filter((span) => span.label && span.columnIndexes.length >= 2);

  if (meaningfulSpans.length === 0) {
    return {
      groupLevel: null,
      message: null,
    };
  }

  const hasMergedDownSingleColumnSpan = spans.some(
    (span) =>
      span.columnIndexes.length === 1 &&
      !normalizeRawBannerCellValue(lowerBannerRow[span.startColumnIndex]) &&
      !isTechnicalWaveOrValueLabel(span.label)
  );

  if (hasMergedDownSingleColumnSpan) {
    return {
      groupLevel: null,
      message: null,
    };
  }

  const labels = Array(selectedColumnCount).fill("");
  const spansByColumnIndex = {};

  for (const span of spans) {
    for (const columnIndex of span.columnIndexes) {
      labels[columnIndex] = span.label;
      spansByColumnIndex[columnIndex] = span;
    }
  }

  return {
    groupLevel: {
      rowOffset: -(rowIndex + 1),
      labels,
      spans,
      spansByColumnIndex,
      detectionMethod: "reconstructedSpan",
    },
    message: createBannerMessage(
      MESSAGE_SEVERITY.INFO,
      "BANNER_RECONSTRUCTED_SPAN_LEVEL_FOUND",
      `Баннер: найден верхний уровень групп по merged-like span на строке ${rowIndex + 1} над нижним уровнем.`
    ),
  };
}

/**
 * Builds adjacent same-label spans.
 */
function buildRepeatedLabelSpans(row) {
  const spans = [];

  let currentSpan = null;

  for (let columnIndex = 0; columnIndex < row.length; columnIndex++) {
    const label = normalizeRawBannerCellValue(row[columnIndex]);
    const normalizedLabel = normalizeBannerLabel(label);

    if (currentSpan && normalizedLabel && normalizedLabel === currentSpan.normalizedLabel) {
      currentSpan.columnIndexes.push(columnIndex);
      continue;
    }

    if (currentSpan) {
      spans.push(currentSpan);
    }

    currentSpan = {
      label,
      normalizedLabel,
      columnIndexes: [columnIndex],
      startColumnIndex: columnIndex,
      endColumnIndex: columnIndex,
    };
  }

  if (currentSpan) {
    spans.push(currentSpan);
  }

  for (const span of spans) {
    span.endColumnIndex = span.columnIndexes[span.columnIndexes.length - 1];
  }

  return spans;
}

/**
 * Creates fallback banner structure.
 */
function createFallbackBannerStructure(selectedColumnCount, messages = []) {
  const columnDescriptors = [];

  for (let columnIndex = 0; columnIndex < selectedColumnCount; columnIndex++) {
    columnDescriptors.push({
      columnIndex,
      lowerLabel: "",
      normalizedLowerLabel: "",
      bannerPath: [],
      displayLabel: `Column ${columnIndex + 1}`,

      comparisonGroupKey: DEFAULT_GROUP_KEY,
      comparisonGroupLabel: DEFAULT_GROUP_LABEL,

      isTotal: false,
      totalType: null,

      isGlobalTotal: false,
      isLocalTotal: false,

      source: {
        lowerLevelRowOffset: null,
        groupLevelRowOffset: null,
        mergeArea: null,
      },
    });
  }

  return {
    isDetected: false,
    mode: BANNER_MODE.FALLBACK,
    columnDescriptors,
    groups: [
      {
        groupKey: DEFAULT_GROUP_KEY,
        label: DEFAULT_GROUP_LABEL,
        bannerPath: [],
        columnIndexes: Array.from({ length: selectedColumnCount }, (_, index) => index),
        localTotalColumnIndexes: [],
        hasLocalTotal: false,
      },
    ],
    globalTotalColumnIndex: null,
    totalColumnIndexes: [],
    messages,
  };
}

/**
 * Returns true if normalized label is Total-like.
 */
function isTotalBannerLabel(normalizedLabel) {
  if (!normalizedLabel) {
    return false;
  }

  const tokens = normalizedLabel.split(" ").filter(Boolean);

  return TOTAL_LABEL_KEYWORDS.some(
    (keyword) => normalizedLabel === keyword || tokens.includes(keyword)
  );
}

/**
 * Normalizes banner labels for matching.
 */
function normalizeBannerLabel(rawLabel) {
  return normalizeLookupText(rawLabel);
}

/**
 * Normalizes raw cell value but keeps user-facing text.
 */
function normalizeRawBannerCellValue(rawValue) {
  return normalizeDisplayText(rawValue);
}

/**
 * Ensures banner row has exactly selectedColumnCount items.
 */
function normalizeBannerRowLength(row, selectedColumnCount) {
  const normalizedRow = Array.isArray(row) ? [...row] : [];

  while (normalizedRow.length < selectedColumnCount) {
    normalizedRow.push("");
  }

  return normalizedRow.slice(0, selectedColumnCount);
}

/**
 * Creates banner message.
 */
function createBannerMessage(severity, code, text) {
  return {
    severity,
    code,
    text,
  };
}

/**
 * Formats short detection summary message.
 */
function formatBannerDetectedMessage({ mode, selectedColumnCount, groups, totalColumnIndexes }) {
  const groupCount = groups ? groups.length : 0;
  const totalCount = totalColumnIndexes ? totalColumnIndexes.length : 0;

  return `Баннер: режим ${mode}, колонок: ${selectedColumnCount}, групп: ${groupCount}, Тоталов: ${totalCount}.`;
}

/**
 * Builds stable group key from group label and span position.
 *
 * PURPOSE:
 * Group identity must include position/span, not only label,
 * because the same group label can appear multiple times in one banner.
 */
function buildGroupKey(groupLabel, span) {
  const normalizedGroupLabel = normalizeBannerLabel(groupLabel);

  if (!span) {
    return `group:${normalizedGroupLabel || "unknown"}:unknown`;
  }

  return [
    "group",
    normalizedGroupLabel || "unknown",
    span.startColumnIndex,
    span.endColumnIndex,
  ].join(":");
}

/**
 * Merges consecutive adjacent single-column wave-value spans that have an
 * empty lower-banner cell into a single multi-column span.
 *
 * PURPOSE:
 * When a wave pair (e.g. "2025Q4" / "2026Q1") appears at the start of a
 * stripped partial-selection banner without an explicit parent-group label,
 * the upper scan row contains the quarter labels as consecutive single-column
 * spans with blank lower-banner cells beneath them.
 *
 * Without merging, each column gets its own group key and is separately
 * assigned local label "a", producing duplicate markers.  After merging, the
 * run forms one multi-column span so its columns receive distinct labels
 * ("a" and "b") exactly as wave pairs under an explicit "Волна (квартал)"
 * parent do.
 *
 * RULE: a run is merged when every span in the run satisfies all of:
 *   - exactly 1 column;
 *   - label is a technical wave value label (e.g. "2025Q4", "2026Q1");
 *   - the lower-banner cell at that column is blank;
 *   - the span is immediately adjacent (no gap) to the previous one in the run.
 * Runs of length 1 are left unchanged so the existing
 * `hasMergedDownSingleColumnSpan` guard can still handle them.
 */
function mergeAdjacentWaveValueSpans(spans, lowerBannerRow) {
  if (!spans || spans.length === 0) {
    return spans;
  }

  const result = [];
  let i = 0;

  while (i < spans.length) {
    const span = spans[i];

    const isWaveValueSingleColBlankLower =
      span.columnIndexes.length === 1 &&
      isTechnicalWaveOrValueLabel(span.label) &&
      !normalizeRawBannerCellValue((lowerBannerRow || [])[span.startColumnIndex]);

    if (!isWaveValueSingleColBlankLower) {
      result.push(span);
      i++;
      continue;
    }

    // Collect the full run of adjacent wave-value single-column spans.
    const runSpans = [span];
    let j = i + 1;

    while (j < spans.length) {
      const next = spans[j];

      const nextIsWaveValueSingleColBlankLower =
        next.columnIndexes.length === 1 &&
        isTechnicalWaveOrValueLabel(next.label) &&
        !normalizeRawBannerCellValue((lowerBannerRow || [])[next.startColumnIndex]);

      const isAdjacent =
        next.startColumnIndex === runSpans[runSpans.length - 1].startColumnIndex + 1;

      if (!nextIsWaveValueSingleColBlankLower || !isAdjacent) {
        break;
      }

      runSpans.push(next);
      j++;
    }

    if (runSpans.length === 1) {
      // Single span — keep as-is so hasMergedDownSingleColumnSpan can handle it.
      result.push(span);
    } else {
      // Merge the run into one multi-column span.  Use the first span's label
      // (position-based group key makes each merged run unique regardless of
      // label choice; the label only affects the comparisonGroupLabel display).
      const merged = {
        label: runSpans[0].label,
        normalizedLabel: runSpans[0].normalizedLabel,
        startColumnIndex: runSpans[0].startColumnIndex,
        endColumnIndex: runSpans[runSpans.length - 1].endColumnIndex,
        columnIndexes: runSpans.flatMap((s) => s.columnIndexes),
      };
      result.push(merged);
    }

    i = j;
  }

  return result;
}

/**
 * Builds reconstructed spans from an upper banner row.
 *
 * RULE:
 * A non-empty upper cell starts a span.
 * Following empty cells may belong to that span.
 * The span is constrained by continuous non-empty lower banner labels.
 */
function buildReconstructedSpansFromUpperRow(upperRow, lowerBannerRow, selectedColumnCount) {
  const spans = [];

  let currentSpan = null;

  for (let columnIndex = 0; columnIndex < selectedColumnCount; columnIndex++) {
    const cellText = normalizeRawBannerCellValue(upperRow[columnIndex]);

    if (cellText) {
      if (currentSpan) {
        currentSpan.endColumnIndex = columnIndex - 1;
        refineReconstructedSpanByLowerBannerRow(currentSpan, lowerBannerRow);
        finalizeCoreSpan(currentSpan);
        spans.push(currentSpan);
      }

      currentSpan = {
        label: cellText,
        normalizedLabel: normalizeBannerLabel(cellText),
        startColumnIndex: columnIndex,
        endColumnIndex: columnIndex,
        columnIndexes: [columnIndex],
      };

      continue;
    }

    if (currentSpan) {
      currentSpan.endColumnIndex = columnIndex;
    }
  }

  if (currentSpan) {
    refineReconstructedSpanByLowerBannerRow(currentSpan, lowerBannerRow);
    finalizeCoreSpan(currentSpan);
    spans.push(currentSpan);
  }

  return spans.filter((span) => span.label && span.columnIndexes.length > 0);
}

/**
 * Refines reconstructed span right boundary using lower banner row.
 *
 * Prevents:
 *   Age | "" | "" | "" | ""
 * from stretching beyond:
 *   Total | 18-24 | 25-34 | "" | ""
 */
function refineReconstructedSpanByLowerBannerRow(span, lowerBannerRow) {
  if (!lowerBannerRow || lowerBannerRow.length === 0) {
    return;
  }

  const lowerStartText = normalizeRawBannerCellValue(lowerBannerRow[span.startColumnIndex]);

  if (!lowerStartText) {
    return;
  }

  let lowerAreaEndColumnIndex = span.startColumnIndex;

  for (
    let columnIndex = span.startColumnIndex + 1;
    columnIndex < lowerBannerRow.length;
    columnIndex++
  ) {
    const lowerCellText = normalizeRawBannerCellValue(lowerBannerRow[columnIndex]);

    if (!lowerCellText) {
      break;
    }

    lowerAreaEndColumnIndex = columnIndex;
  }

  span.endColumnIndex = Math.min(span.endColumnIndex, lowerAreaEndColumnIndex);
}

/**
 * Finalizes span column indexes.
 */
function finalizeCoreSpan(span) {
  span.columnIndexes = [];

  for (let columnIndex = span.startColumnIndex; columnIndex <= span.endColumnIndex; columnIndex++) {
    span.columnIndexes.push(columnIndex);
  }
}

/**
 * Detects repeated adjacent labels in a single upper row.
 *
 * Example:
 *   Gender | Gender | Gender | Age | Age | Age
 */
function detectRepeatedLabelGroupLevelInRow(row, selectedColumnCount, rowIndex) {
  const spans = buildRepeatedLabelSpans(row);

  const meaningfulSpans = spans.filter((span) => span.label && span.columnIndexes.length >= 2);

  if (meaningfulSpans.length === 0) {
    return {
      groupLevel: null,
      message: null,
    };
  }

  const labels = [];

  for (let columnIndex = 0; columnIndex < selectedColumnCount; columnIndex++) {
    const span = spans.find((candidate) => candidate.columnIndexes.includes(columnIndex));

    labels[columnIndex] = span ? span.label : "";
  }

  const spansByColumnIndex = {};

  for (const span of spans) {
    for (const columnIndex of span.columnIndexes) {
      spansByColumnIndex[columnIndex] = span;
    }
  }

  return {
    groupLevel: {
      rowOffset: -(rowIndex + 1),
      labels,
      spans,
      spansByColumnIndex,
      detectionMethod: "repeatedLabels",
    },
    message: createBannerMessage(
      MESSAGE_SEVERITY.INFO,
      "BANNER_REPEATED_GROUP_LEVEL_FOUND",
      `Баннер: найден верхний уровень групп по повторяющимся значениям на строке ${rowIndex + 1} над нижним уровнем.`
    ),
  };
}

/**
 * Detects global Total column.
 *
 * MVP RULE:
 * The first selected column is treated as global Total if:
 * - its lower label is Total-like and its group label is also Total-like;
 * - OR its group label is Total-like.
 *
 * This covers common two-level / vertical-merge-like structures:
 *
 *   Global Total | Gender | Gender | Age | Age
 *   Global Total | Total  | Male   | Total | 18-24
 */
function detectGlobalTotalColumnIndex(columnDescriptors) {
  if (!columnDescriptors || columnDescriptors.length === 0) {
    return null;
  }

  const firstDescriptor = columnDescriptors[0];

  if (!firstDescriptor) {
    return null;
  }

  const lowerLabelIsTotal = isTotalBannerLabel(firstDescriptor.normalizedLowerLabel);

  const normalizedGroupLabel = normalizeBannerLabel(firstDescriptor.comparisonGroupLabel);
  const groupLabelIsTotal = isTotalBannerLabel(normalizedGroupLabel);

  if (groupLabelIsTotal) {
    if (countDescriptorsInSameGroup(columnDescriptors, firstDescriptor) > 1) {
      return null;
    }

    return firstDescriptor.columnIndex;
  }

  // First column is a stand-alone Total/Всего with no parent group span
  // covering it at the picked banner level. Treat it as global so non-total
  // sibling groups can be compared against it.
  if (
    lowerLabelIsTotal &&
    !normalizedGroupLabel &&
    countDescriptorsInSameGroup(columnDescriptors, firstDescriptor) === 1
  ) {
    return firstDescriptor.columnIndex;
  }

  if (lowerLabelIsTotal && groupLabelIsTotal) {
    return firstDescriptor.columnIndex;
  }

  return null;
}

function countDescriptorsInSameGroup(columnDescriptors, descriptor) {
  return columnDescriptors.filter(
    (candidate) => candidate.comparisonGroupKey === descriptor.comparisonGroupKey
  ).length;
}

/**
 * Promotes columns with empty lower banner labels to local Total when the
 * nearest visible upper banner cell in that column is total-like.
 *
 * Mirrors the writer's "nearest non-empty cell above" rule so that a Total
 * label living in a sparse/merged upper banner row is not silently treated
 * as a normal comparable column.
 */
function markUpperLevelTotalsForSparseLowerLabels(columnDescriptors, upperScanRows) {
  if (!columnDescriptors || columnDescriptors.length === 0) {
    return;
  }

  if (!Array.isArray(upperScanRows) || upperScanRows.length === 0) {
    return;
  }

  for (const descriptor of columnDescriptors) {
    if (descriptor.isTotal) {
      continue;
    }

    if (descriptor.normalizedLowerLabel) {
      continue;
    }

    for (let rowOffset = 0; rowOffset < upperScanRows.length; rowOffset++) {
      const upperRow = upperScanRows[rowOffset] || [];
      const upperText = normalizeRawBannerCellValue(upperRow[descriptor.columnIndex]);

      if (!upperText) {
        continue;
      }

      if (isTotalBannerLabel(normalizeBannerLabel(upperText))) {
        descriptor.isTotal = true;
        descriptor.isLocalTotal = true;
        descriptor.totalType = TOTAL_TYPE.LOCAL;

        attachSparseTotalToAdjacentGroup(descriptor, columnDescriptors);
      }

      break;
    }
  }
}

/**
 * Attaches a sparse upper-level local Total descriptor to the adjacent
 * named group so that comparison-pair builders treat it as that group's
 * local Total reference.
 *
 * Only descriptors whose own group identity is empty (no upper-level group
 * span at the picked group level) are reattached. The nearest non-Total
 * descriptor with a non-empty group label is preferred to the right; the
 * left side is used as a fallback for trailing sparse Totals.
 */
function attachSparseTotalToAdjacentGroup(totalDescriptor, columnDescriptors) {
  if (totalDescriptor.comparisonGroupLabel) {
    return;
  }

  const adjacent =
    findAdjacentNamedGroupDescriptor(totalDescriptor, columnDescriptors, 1) ||
    findAdjacentNamedGroupDescriptor(totalDescriptor, columnDescriptors, -1);

  if (!adjacent) {
    return;
  }

  totalDescriptor.comparisonGroupKey = adjacent.comparisonGroupKey;
  totalDescriptor.comparisonGroupLabel = adjacent.comparisonGroupLabel;
}

function findAdjacentNamedGroupDescriptor(totalDescriptor, columnDescriptors, step) {
  for (
    let index = totalDescriptor.columnIndex + step;
    index >= 0 && index < columnDescriptors.length;
    index += step
  ) {
    const candidate = columnDescriptors[index];

    if (!candidate) {
      continue;
    }

    if (candidate.isTotal) {
      continue;
    }

    if (!candidate.comparisonGroupLabel) {
      continue;
    }

    return candidate;
  }

  return null;
}

/**
 * Marks one descriptor as global Total and removes local Total role from it.
 */
function markGlobalTotalColumn(columnDescriptors, globalTotalColumnIndex) {
  if (globalTotalColumnIndex === null || globalTotalColumnIndex === undefined) {
    return;
  }

  const descriptor = columnDescriptors.find(
    (candidate) => candidate.columnIndex === globalTotalColumnIndex
  );

  if (!descriptor) {
    return;
  }

  descriptor.isTotal = true;
  descriptor.totalType = TOTAL_TYPE.GLOBAL;

  descriptor.isGlobalTotal = true;
  descriptor.isLocalTotal = false;
}

/**
 * Detects which comparison-group keys host a nested/repeated wave dimension
 * inside the group rather than at the parent group label.
 *
 * Two signals promote a group to wave-aware:
 * - lower banner labels in the group's columns look like concrete wave values
 *   (e.g. 2025Q4, 2026Q1) for at least two columns;
 * - any upper banner row other than the chosen group level row contains a
 *   wave / period / quarter descriptor label (e.g. "Волна (квартал)") for at
 *   least two of the group's columns.
 *
 * Single-column groups are never promoted because previous-column wave
 * comparison only makes sense across two or more adjacent wave columns.
 */
function detectGroupKeysWithNestedWaveDimension({
  columnDescriptors,
  lowerBannerRow,
  upperScanRows,
  groupLevelRowOffset,
}) {
  const groupColumnsByKey = new Map();

  for (const descriptor of columnDescriptors) {
    if (!groupColumnsByKey.has(descriptor.comparisonGroupKey)) {
      groupColumnsByKey.set(descriptor.comparisonGroupKey, []);
    }

    groupColumnsByKey.get(descriptor.comparisonGroupKey).push(descriptor.columnIndex);
  }

  const groupLevelRowIndex =
    typeof groupLevelRowOffset === "number" ? -groupLevelRowOffset - 1 : null;

  const waveGroupKeys = new Set();

  for (const [groupKey, columnIndexes] of groupColumnsByKey) {
    if (columnIndexes.length < 2) {
      continue;
    }

    if (countWaveValueLabelsInColumns(columnIndexes, lowerBannerRow) >= 2) {
      waveGroupKeys.add(groupKey);
      continue;
    }

    if (
      hasWaveDescriptorInIntermediateUpperRow(columnIndexes, upperScanRows, groupLevelRowIndex)
    ) {
      waveGroupKeys.add(groupKey);
    }
  }

  return waveGroupKeys;
}

function countWaveValueLabelsInColumns(columnIndexes, row) {
  if (!row) {
    return 0;
  }

  let count = 0;

  for (const columnIndex of columnIndexes) {
    const cellText = normalizeRawBannerCellValue(row[columnIndex]);

    if (cellText && isTechnicalWaveValueLabel(cellText)) {
      count++;
    }
  }

  return count;
}

function hasWaveDescriptorInIntermediateUpperRow(
  columnIndexes,
  upperScanRows,
  groupLevelRowIndex
) {
  if (!Array.isArray(upperScanRows) || upperScanRows.length === 0) {
    return false;
  }

  for (let rowIndex = 0; rowIndex < upperScanRows.length; rowIndex++) {
    if (rowIndex === groupLevelRowIndex) {
      continue;
    }

    const row = upperScanRows[rowIndex];

    if (!row) {
      continue;
    }

    let descriptorCount = 0;

    for (const columnIndex of columnIndexes) {
      const cellText = normalizeRawBannerCellValue(row[columnIndex]);

      if (cellText && isTechnicalWaveDescriptorLabel(cellText)) {
        descriptorCount++;
      }
    }

    if (descriptorCount >= 2) {
      return true;
    }
  }

  return false;
}

/**
 * Returns true if group label looks like a wave / period / measurement group.
 *
 * MVP RULE:
 * We use group-level label only.
 *
 * We intentionally do not treat plain numeric labels like "1, 2, 3"
 * as wave labels because those can be segment/object numbers.
 */
function isWaveGroupLabel(rawLabel) {
  const normalizedLabel = normalizeBannerLabel(rawLabel);

  if (!normalizedLabel) {
    return false;
  }

  const tokens = normalizedLabel.split(" ").filter(Boolean);

  return WAVE_GROUP_LABEL_KEYWORDS.some(
    (keyword) => normalizedLabel === keyword || tokens.includes(keyword)
  );
}

/**
 * Returns true when a detected group level is dominated by technical wave /
 * period labels, not the semantic grouping row users expect comparisons to
 * follow.
 */
function isTechnicalWaveDescriptorGroupLevel(groupLevel) {
  if (!groupLevel || !groupLevel.spans || groupLevel.spans.length === 0) {
    return false;
  }

  const labeledSpans = groupLevel.spans.filter((span) => span.label);

  if (labeledSpans.length === 0) {
    return false;
  }

  const technicalWaveSpanCount = labeledSpans.filter((span) =>
    isTechnicalWaveOrValueLabel(span.label)
  ).length;

  return technicalWaveSpanCount / labeledSpans.length >= TECHNICAL_WAVE_LABEL_DOMINANCE_THRESHOLD;
}

/**
 * Returns true if a label is either a generic wave descriptor or a concrete
 * wave / quarter value.
 */
function isTechnicalWaveOrValueLabel(rawLabel) {
  return isTechnicalWaveDescriptorLabel(rawLabel) || isTechnicalWaveValueLabel(rawLabel);
}

/**
 * Returns true if a label is a generic wave/period/quarter descriptor.
 */
function isTechnicalWaveDescriptorLabel(rawLabel) {
  const normalizedLabel = normalizeBannerLabel(rawLabel);

  if (!normalizedLabel) {
    return false;
  }

  const tokens = normalizedLabel.split(" ").filter(Boolean);

  return TECHNICAL_WAVE_DESCRIPTOR_KEYWORDS.some((keyword) => {
    const normalizedKeyword = normalizeBannerLabel(keyword);

    return (
      normalizedLabel === normalizedKeyword ||
      tokens.includes(normalizedKeyword) ||
      normalizedLabel.includes(normalizedKeyword)
    );
  });
}

/**
 * Returns true if a label looks like a concrete wave / quarter value.
 */
function isTechnicalWaveValueLabel(rawLabel) {
  const normalizedLabel = normalizeBannerLabel(rawLabel);

  if (!normalizedLabel) {
    return false;
  }

  const compactLabel = normalizedLabel.replace(/\s+/g, "");

  return (
    /^\d{4}q[1-4]$/.test(compactLabel) ||
    /^q[1-4]\d{4}$/.test(compactLabel) ||
    /^\d{4}кв[1-4]$/.test(compactLabel) ||
    /^кв[1-4]\d{4}$/.test(compactLabel) ||
    /^\d{4}квартал[1-4]$/.test(compactLabel) ||
    /^квартал[1-4]\d{4}$/.test(compactLabel)
  );
}
