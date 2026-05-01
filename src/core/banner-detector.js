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

const TOTAL_LABEL_KEYWORDS = ["total", "итого", "всего", "all", "overall"];
const WAVE_GROUP_LABEL_KEYWORDS = [
  "wave",
  "waves",
  "волна",
  "волны",
  "period",
  "periods",
  "период",
  "периоды",
  "замер",
  "замеры",
];

const DEFAULT_GROUP_KEY = "group:default";
const DEFAULT_GROUP_LABEL = "Default";

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
        "warning",
        "BANNER_EMPTY_SELECTION",
        "Баннер: не удалось определить количество колонок выделения."
      ),
    ]);
  }

  if (!lowerBannerRow || lowerBannerRow.length === 0) {
    return createFallbackBannerStructure(selectedColumnCount, [
      createBannerMessage(
        "warning",
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

  const globalTotalColumnIndex = detectGlobalTotalColumnIndex(columnDescriptors);

  markGlobalTotalColumn(columnDescriptors, globalTotalColumnIndex);

  const groups = buildGroupsFromColumnDescriptors(columnDescriptors, settings);

  const waveGroups = groups.filter((group) => group.recommendedComparisonMode === "previousColumn");

  if (waveGroups.length > 0) {
    messages.push(
      createBannerMessage(
        "info",
        "BANNER_WAVE_GROUPS_DETECTED",
        `Баннер: обнаружены волновые группы: ${waveGroups.map((group) => group.label).join(", ")}.`
      )
    );
  }

  const totalColumnIndexes = columnDescriptors
    .filter((descriptor) => descriptor.isTotal)
    .map((descriptor) => descriptor.columnIndex);

  const mode = groupLevelResult.groupLevel ? "twoLevel" : "oneLevel";

  messages.push(
    createBannerMessage(
      "debug",
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
    recommendedComparisonMode: waveGroups.length > 0 ? "mixed" : "default",
    messages,
  };
}

/**
 * Formats banner detection diagnostics for status output.
 */
export function formatBannerDetectionDiagnostics(bannerStructure) {
  if (!bannerStructure || !bannerStructure.isDetected) {
    return "Баннер: структура не обнаружена.";
  }

  const lines = [];

  lines.push("Баннер:");
  lines.push(`- Режим: ${bannerStructure.mode}`);

  if (bannerStructure.globalTotalColumnIndex !== null) {
    lines.push(`- Глобальный Тотал: колонка ${bannerStructure.globalTotalColumnIndex + 1}`);
  }

  if (bannerStructure.totalColumnIndexes && bannerStructure.totalColumnIndexes.length > 0) {
    lines.push(
      `- Найденные Тоталы: ${bannerStructure.totalColumnIndexes
        .map((columnIndex) => columnIndex + 1)
        .join(", ")}`
    );
  } else {
    lines.push("- Найденные Тоталы: нет");
  }

  if (bannerStructure.groups && bannerStructure.groups.length > 0) {
    lines.push("- Группы:");

    for (const group of bannerStructure.groups) {
      const columnLabels = group.columnIndexes
        .map((columnIndex) => {
          const descriptor = bannerStructure.columnDescriptors.find(
            (item) => item.columnIndex === columnIndex
          );

          return descriptor ? descriptor.lowerLabel || `Column ${columnIndex + 1}` : "";
        })
        .filter(Boolean)
        .join(", ");

      lines.push(`  - ${group.label}: ${columnLabels}`);
    }
  }

  const bannerMessages = bannerStructure.messages || [];

  if (bannerMessages.length > 0) {
    lines.push("- Сообщения:");

    for (const message of bannerMessages) {
      lines.push(`  - ${message.text}`);
    }
  }

  return lines.join("\n");
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
      totalType: isTotal ? "local" : null,

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
function buildGroupsFromColumnDescriptors(columnDescriptors, calculationSettings = {}) {
  const groupsByKey = new Map();

  for (const descriptor of columnDescriptors) {
    if (!groupsByKey.has(descriptor.comparisonGroupKey)) {
      const shouldAutoDetectWaveBanners =
        calculationSettings && calculationSettings.autoDetectWaveBanners;

      const isWaveGroup =
        shouldAutoDetectWaveBanners && isWaveGroupLabel(descriptor.comparisonGroupLabel);

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

        semanticType: isWaveGroup ? "wave" : "default",
        recommendedComparisonMode: isWaveGroup ? "previousColumn" : "default",
      });
    }

    const group = groupsByKey.get(descriptor.comparisonGroupKey);

    group.columnIndexes.push(descriptor.columnIndex);

    if (descriptor.isTotal && descriptor.totalType !== "global") {
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
        "debug",
        "BANNER_ONE_LEVEL_FALLBACK",
        "Баннер: верхний уровень групп не найден. Используется один уровень баннера."
      ),
    };
  }

  for (let rowIndex = 0; rowIndex < upperScanRows.length; rowIndex++) {
    const row = normalizeBannerRowLength(upperScanRows[rowIndex], selectedColumnCount);

    const reconstructedSpanResult = detectReconstructedSpanGroupLevel(
      row,
      lowerBannerRow,
      selectedColumnCount,
      rowIndex
    );

    if (reconstructedSpanResult.groupLevel) {
      return reconstructedSpanResult;
    }

    const repeatedLabelResult = detectRepeatedLabelGroupLevelInRow(
      row,
      selectedColumnCount,
      rowIndex
    );

    if (repeatedLabelResult.groupLevel) {
      return repeatedLabelResult;
    }
  }

  return {
    groupLevel: null,
    message: createBannerMessage(
      "info",
      "BANNER_ONE_LEVEL_FALLBACK",
      "Баннер: верхний уровень групп не найден. Используется один уровень баннера."
    ),
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
  const spans = buildReconstructedSpansFromUpperRow(upperRow, lowerBannerRow, selectedColumnCount);

  const meaningfulSpans = spans.filter((span) => span.label && span.columnIndexes.length >= 2);

  if (meaningfulSpans.length === 0) {
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
      "info",
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
    mode: "fallback",
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
  if (rawLabel === null || rawLabel === undefined) {
    return "";
  }

  return String(rawLabel)
    .toLowerCase()
    .replace(/ё/g, "е")
    .replace(/[.,:;()]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Normalizes raw cell value but keeps user-facing text.
 */
function normalizeRawBannerCellValue(rawValue) {
  if (rawValue === null || rawValue === undefined) {
    return "";
  }

  return String(rawValue).trim();
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
      "info",
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

  const groupLabelIsTotal = isTotalBannerLabel(
    normalizeBannerLabel(firstDescriptor.comparisonGroupLabel)
  );

  if (groupLabelIsTotal) {
    return firstDescriptor.columnIndex;
  }

  if (lowerLabelIsTotal && groupLabelIsTotal) {
    return firstDescriptor.columnIndex;
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
  descriptor.totalType = "global";

  descriptor.isGlobalTotal = true;
  descriptor.isLocalTotal = false;
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
