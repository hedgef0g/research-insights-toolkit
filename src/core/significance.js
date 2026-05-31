import {
  normalizeShare,
  normalizeVariance,
  normalizeNpsValue,
  normalizeNpsSpread,
} from "./normalizers";

import { getZThresholdForConfidence, getTThresholdForConfidence } from "./stat-thresholds";

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
  const zThreshold = getZThresholdForConfidence(confidenceLevel, {
    oneTailedTest: calculationSettings.oneTailedTest,
  });

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
    calculationSettings.excludedColumnIndexes || new Set(),
    calculationSettings.bannerStructure || null
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
      groupKey: comparisonPair.groupKey,
      groupLabel: comparisonPair.groupLabel,
      firstValue,
      secondValue,
      firstBase,
      secondBase,
      result: significanceResult,
      totalReferenceType: comparisonPair.totalReferenceType,
      autoPreviousColumnFromBanner: comparisonPair.autoPreviousColumnFromBanner,
    });
  }

  return rowComparisons;
}

// Latin marker letters, excluding t / T (which are reserved for Total markers).
const LATIN_LOWERCASE_MARKERS = "abcdefghijklmnopqrsuvwxyz"; // a-z without t.
const LATIN_UPPERCASE_MARKERS = "ABCDEFGHIJKLMNOPQRSUVWXYZ"; // A-Z without T.

// Allowed Cyrillic marker letters.
//
// Starting from the full Cyrillic alphabet we remove:
// - т / Т — reserved for Total markers (pre-existing exclusion);
// - visually confusable look-alikes of Latin markers (issue #312):
//     а, А, В, с, С, е, Е, К, М, Н, о, О, р, Р, у, х, Х
//
// Note: lowercase Cyrillic `у` is excluded because it reads like Latin `y`,
// but uppercase Cyrillic `У` is kept because it is visually distinct from
// Latin `Y`.
const CYRILLIC_LOWERCASE_MARKERS = "бвгджзийклмнпфцчшщъыьэюя";
const CYRILLIC_UPPERCASE_MARKERS = "БГДЖЗИЙЛПУФЦЧШЩЪЫЬЭЮЯ";

// Base alphabet used to build multi-character overflow markers (aa, ab, ac...).
// Always lowercase Latin (without t) regardless of the Cyrillic setting so the
// extended markers stay predictable and easy to recognise/clean up.
const MULTI_CHARACTER_MARKER_BASE = LATIN_LOWERCASE_MARKERS.split("");

// Every single-character token RIT may have written as a marker across any
// historical alphabet (the pre-#312 Latin + full-Cyrillic set, minus t/Т).
// Used for recognising/removing banner markers so legacy markers — including
// now-excluded Cyrillic look-alikes — are still cleaned up on re-runs.
const RECOGNIZED_SINGLE_CHARACTER_MARKERS = new Set(
  (
    "abcdefghijklmnopqrsuvwxyz" +
    "ABCDEFGHIJKLMNOPQRSUVWXYZ" +
    "абвгдежзийклмнопрсуфхцчшщъыьэюя" +
    "АБВГДЕЖЗИЙКЛМНОПРСУФХЦЧШЩЪЫЬЭЮЯ"
  ).split("")
);

/**
 * Returns the ordered single-character significance marker alphabet.
 *
 * @param {object} [options]
 * @param {boolean} [options.useCyrillicMarkers] When true, allowed Cyrillic
 *   markers are appended after the Latin markers. Defaults to false so that,
 *   for global release safety, generated markers are Latin-only unless the
 *   user explicitly opts in.
 */
export function getSignificanceMarkerAlphabet(options = {}) {
  const { useCyrillicMarkers = false } =
    options && typeof options === "object" ? options : {};

  const latinLabels = (LATIN_LOWERCASE_MARKERS + LATIN_UPPERCASE_MARKERS).split("");

  if (!useCyrillicMarkers) {
    return latinLabels;
  }

  const cyrillicLabels = (CYRILLIC_LOWERCASE_MARKERS + CYRILLIC_UPPERCASE_MARKERS).split("");

  return [...latinLabels, ...cyrillicLabels];
}

/**
 * Builds `count` multi-character overflow marker labels (aa, ab, ac, ...).
 *
 * Labels are produced as an odometer over MULTI_CHARACTER_MARKER_BASE, shortest
 * length first: all 2-letter combinations, then 3-letter, and so on.
 */
function buildMultiCharacterMarkerLabels(count) {
  const labels = [];
  const base = MULTI_CHARACTER_MARKER_BASE;

  for (let length = 2; labels.length < count; length++) {
    const combinationCount = Math.pow(base.length, length);

    for (let n = 0; n < combinationCount && labels.length < count; n++) {
      let remainder = n;
      let label = "";

      for (let position = 0; position < length; position++) {
        label = base[remainder % base.length] + label;
        remainder = Math.floor(remainder / base.length);
      }

      labels.push(label);
    }
  }

  return labels;
}

/**
 * Generates column significance labels.
 *
 * PURPOSE:
 * Create readable column labels for significance notation.
 *
 * IMPORTANT:
 * - Latin t / T and Cyrillic т / Т are excluded (reserved for Total markers).
 * - Cyrillic markers appear only when `useCyrillicMarkers` is enabled, and then
 *   exclude the visually confusable Latin look-alikes (issue #312).
 * - When the single-character alphabet is exhausted and
 *   `allowMultiCharacterMarkers` is enabled, the sequence is extended with
 *   multi-character markers (aa, ab, ac, ...).
 *
 * @param {object} [options]
 * @param {boolean} [options.useCyrillicMarkers]
 * @param {boolean} [options.allowMultiCharacterMarkers]
 * @param {number} [options.minimumCount] Ensure the returned array has at least
 *   this many labels (only extends past the single-character alphabet when
 *   `allowMultiCharacterMarkers` is true).
 *
 * OUTPUT:
 * Array of labels.
 */
export function generateSignificanceLabels(options = {}) {
  const {
    useCyrillicMarkers = false,
    allowMultiCharacterMarkers = false,
    minimumCount = 0,
  } = options && typeof options === "object" ? options : {};

  const singleCharacterLabels = getSignificanceMarkerAlphabet({ useCyrillicMarkers });

  if (!allowMultiCharacterMarkers || minimumCount <= singleCharacterLabels.length) {
    return singleCharacterLabels;
  }

  const extraLabels = buildMultiCharacterMarkerLabels(
    minimumCount - singleCharacterLabels.length
  );

  return [...singleCharacterLabels, ...extraLabels];
}

/**
 * Builds generation options for {@link generateSignificanceLabels} from the
 * current calculation settings, ensuring at least `minimumCount` labels.
 */
function significanceLabelOptionsFromSettings(calculationSettings = {}, minimumCount = 0) {
  return {
    useCyrillicMarkers: Boolean(calculationSettings.useCyrillicMarkers),
    allowMultiCharacterMarkers: Boolean(calculationSettings.allowMultiCharacterMarkers),
    minimumCount,
  };
}

/**
 * Returns true when `label` is a token RIT could have written as a significance
 * marker (one of the recognised single characters, or a multi-character
 * overflow marker built from the lowercase Latin base).
 *
 * Used by banner marker recognition/removal so significance markers are
 * detected without misreading ordinary parenthesised banner text such as
 * "Wave (quarter)" or "Region (NE)".
 */
export function isSignificanceMarkerLabel(label) {
  if (typeof label !== "string" || label.length === 0) {
    return false;
  }

  if (label.length === 1) {
    return RECOGNIZED_SINGLE_CHARACTER_MARKERS.has(label);
  }

  for (const character of label) {
    if (!MULTI_CHARACTER_MARKER_BASE.includes(character)) {
      return false;
    }
  }

  return true;
}

/**
 * Number of single-character markers available for the current settings.
 */
export function getSignificanceMarkerCapacity(calculationSettings = {}) {
  return getSignificanceMarkerAlphabet({
    useCyrillicMarkers: Boolean(calculationSettings.useCyrillicMarkers),
  }).length;
}

/**
 * Computes how many distinct single-character significance labels the given
 * table would need to assign, so marker overflow can be detected before any
 * results are written.
 *
 * - Banner mode: the largest per-group count of non-Total segment columns.
 * - firstColumnIsTotal: every column except the Total column.
 * - Otherwise: every column.
 */
export function computeRequiredSignificanceLabelCount(
  columnCount,
  calculationSettings = {},
  bannerStructure = null
) {
  if (!Number.isFinite(columnCount) || columnCount < 1) {
    return 0;
  }

  if (
    calculationSettings.respectBannerStructure &&
    bannerStructure &&
    bannerStructure.groups &&
    bannerStructure.groups.length > 0
  ) {
    const totalColumnIndexes = new Set(bannerStructure.totalColumnIndexes || []);
    const globalTotalColumnIndex =
      bannerStructure.globalTotalColumnIndex === undefined
        ? null
        : bannerStructure.globalTotalColumnIndex;

    let maxSegmentCount = 0;

    for (const group of bannerStructure.groups) {
      let segmentCount = 0;

      for (const columnIndex of group.columnIndexes || []) {
        if (columnIndex === globalTotalColumnIndex || totalColumnIndexes.has(columnIndex)) {
          continue;
        }
        segmentCount++;
      }

      maxSegmentCount = Math.max(maxSegmentCount, segmentCount);
    }

    return maxSegmentCount;
  }

  return calculationSettings.firstColumnIsTotal ? Math.max(0, columnCount - 1) : columnCount;
}

/**
 * Returns true when the table needs more single-character significance markers
 * than the current alphabet provides.
 */
export function detectSignificanceMarkerOverflow(
  columnCount,
  calculationSettings = {},
  bannerStructure = null
) {
  const requiredCount = computeRequiredSignificanceLabelCount(
    columnCount,
    calculationSettings,
    bannerStructure
  );

  return requiredCount > getSignificanceMarkerCapacity(calculationSettings);
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
  excludedColumnIndexes = new Set(),
  bannerStructure = null
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

  if (
    calculationSettings.respectBannerStructure &&
    bannerStructure &&
    bannerStructure.groups &&
    bannerStructure.groups.length > 0
  ) {
    return buildBannerStructureComparisonPairs(
      columnCount,
      calculationSettings,
      excludedColumnIndexes,
      bannerStructure
    );
  }

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
 * Builds ordinary comparison pairs only inside detected banner groups.
 */
function buildBannerGroupComparisonPairs(
  columnCount,
  bannerStructure,
  excludedColumnIndexes = new Set(),
  options = {}
) {
  const pairs = [];

  const extraExcludePredicate = options.excludedColumnIndexesExtra;

  const isExcluded = (columnIndex) => {
    if (excludedColumnIndexes.has(columnIndex)) {
      return true;
    }

    if (typeof extraExcludePredicate === "function") {
      return extraExcludePredicate(columnIndex);
    }

    return false;
  };

  for (const group of bannerStructure.groups || []) {
    const groupColumnIndexes = (group.columnIndexes || [])
      .filter((columnIndex) => columnIndex >= 0 && columnIndex < columnCount)
      .filter((columnIndex) => !isExcluded(columnIndex));

    for (let firstIndex = 0; firstIndex < groupColumnIndexes.length; firstIndex++) {
      const firstColumnIndex = groupColumnIndexes[firstIndex];

      for (
        let secondIndex = firstIndex + 1;
        secondIndex < groupColumnIndexes.length;
        secondIndex++
      ) {
        const secondColumnIndex = groupColumnIndexes[secondIndex];

        pairs.push({
          firstColumnIndex,
          secondColumnIndex,
          comparisonType: "bannerGroup",
          groupKey: group.groupKey,
          groupLabel: group.label,
        });
      }
    }
  }

  return pairs;
}

/**
 * Builds mixed banner group pairs.
 *
 * RULES:
 * - wave groups use automatic previous-column mode;
 * - non-wave groups use ordinary group comparisons;
 * - Total columns are excluded from both ordinary and auto previous-column pairs;
 * - Total comparison pairs are still built separately.
 */
function buildMixedBannerGroupComparisonPairs(
  columnCount,
  bannerStructure,
  excludedColumnIndexes = new Set(),
  totalColumnIndexes = new Set()
) {
  const pairs = [];
  const autoPreviousColumnGroupLabels = [];

  for (const group of bannerStructure.groups || []) {
    if (shouldUseAutoPreviousColumnForBannerGroup(group)) {
      const groupPairs = buildBannerPreviousColumnPairsForGroup(
        group,
        excludedColumnIndexes,
        totalColumnIndexes,
        { autoApplied: true }
      );

      if (groupPairs.length > 0) {
        autoPreviousColumnGroupLabels.push(group.label);
      }

      pairs.push(...groupPairs);
      continue;
    }

    const groupPairs = buildBannerGroupComparisonPairs(
      columnCount,
      {
        ...bannerStructure,
        groups: [group],
      },
      excludedColumnIndexes,
      {
        excludedColumnIndexesExtra: (columnIndex) =>
          excludedColumnIndexes.has(columnIndex) || totalColumnIndexes.has(columnIndex),
      }
    );

    pairs.push(...groupPairs);
  }

  if (autoPreviousColumnGroupLabels.length > 0) {
    addBannerStructureMessageOnce(
      bannerStructure,
      "info",
      "BANNER_AUTO_PREVIOUS_COLUMN_APPLIED",
      `Баннер: для волновых групп автоматически применён режим “Сравнение с предыдущей колонкой”: ${autoPreviousColumnGroupLabels.join(
        ", "
      )}.`
    );
  }

  return pairs;
}

/**
 * Builds comparison pairs using detected banner structure.
 *
 * RULES:
 * - ordinary comparisons are created only inside banner groups;
 * - Total comparisons follow banner Total hierarchy;
 * - excludeTotalFromComparisons removes all detected Total columns from all comparisons;
 * - compareOnlyWithTotal keeps only Total comparison pairs.
 */
function buildBannerStructureComparisonPairs(
  columnCount,
  calculationSettings,
  excludedColumnIndexes = new Set(),
  bannerStructure
) {
  const pairs = [];

  const totalColumnIndexes = new Set(bannerStructure.totalColumnIndexes || []);
  const globalTotalColumnIndex =
    bannerStructure.globalTotalColumnIndex === undefined
      ? null
      : bannerStructure.globalTotalColumnIndex;

  const excludeTotalFromComparisons = calculationSettings.excludeTotalFromComparisons;
  const compareOnlyWithTotal = calculationSettings.compareOnlyWithTotal;
  const compareWithPreviousColumn = calculationSettings.compareWithPreviousColumn;

  if (compareWithPreviousColumn) {
    return buildBannerPreviousColumnComparisonPairs(
      columnCount,
      bannerStructure,
      excludedColumnIndexes,
      {
        excludeTotalFromComparisons,
        totalColumnIndexes,
      }
    );
  }

  if (!excludeTotalFromComparisons) {
    const totalPairs = buildBannerTotalComparisonPairs(
      columnCount,
      bannerStructure,
      excludedColumnIndexes
    );

    pairs.push(...totalPairs);

    if (globalTotalColumnIndex !== null && globalTotalColumnIndex !== undefined) {
      addBannerStructureMessageOnce(
        bannerStructure,
        "info",
        "GLOBAL_TOTAL_USED",
        "Найден глобальный Тотал. Сравнение с Тоталом выполняется только относительно глобального Тотала. Локальные Тоталы не используются как внутригрупповые референсы и сравнивались с глобальным Тоталом как обычные колонки."
      );
    }
  }

  if (compareOnlyWithTotal) {
    if (pairs.length === 0) {
      addBannerStructureMessageOnce(
        bannerStructure,
        "warning",
        "BANNER_TOTAL_ONLY_NO_TOTAL_PAIRS",
        "Режим “Сравнивать только с Тотал” включён, но в выделенном баннере не найден глобальный или локальный Тотал. Сравнения не выполнены."
      );
    }

    return pairs;
  }

  const groupPairs = buildMixedBannerGroupComparisonPairs(
    columnCount,
    bannerStructure,
    excludedColumnIndexes,
    totalColumnIndexes
  );

  pairs.push(...groupPairs);

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
  if (calculationSettings.firstColumnIsTotal) {
    if (columnIndex === 0) {
      return "";
    }

    const labelIndex = columnIndex - 1;
    const labels = generateSignificanceLabels(
      significanceLabelOptionsFromSettings(calculationSettings, labelIndex + 1)
    );

    return labels[labelIndex] || "";
  }

  const labels = generateSignificanceLabels(
    significanceLabelOptionsFromSettings(calculationSettings, columnIndex + 1)
  );

  return labels[columnIndex] || "";
}

/**
 * Returns significance label for a column.
 *
 * If banner structure is enabled, labels are local to the detected banner group.
 * Otherwise, labels follow the existing global column indexing rules.
 */
function getSignificanceLabelForComparisonColumn(
  columnIndex,
  calculationSettings = {},
  comparison = null
) {
  if (
    calculationSettings.respectBannerStructure &&
    calculationSettings.bannerStructure &&
    comparison &&
    comparison.groupKey
  ) {
    return getBannerGroupLocalSignificanceLabel(
      columnIndex,
      calculationSettings.bannerStructure,
      comparison.groupKey,
      calculationSettings
    );
  }

  return getSignificanceLabelForColumnIndex(columnIndex, calculationSettings);
}

/**
 * Returns group-local significance label.
 *
 * RULES:
 * - all Total columns are skipped;
 * - Total columns do not consume label indexes;
 * - labels are assigned only to non-Total columns inside the group.
 *
 * Example:
 * Group columns: [Total, Male, Female]
 * Total -> ""
 * Male -> a
 * Female -> b
 */
function getBannerGroupLocalSignificanceLabel(
  columnIndex,
  bannerStructure,
  groupKey,
  calculationSettings = {}
) {
  const group = (bannerStructure.groups || []).find((candidate) => candidate.groupKey === groupKey);

  if (!group || !group.columnIndexes) {
    return getSignificanceLabelForColumnIndex(columnIndex, calculationSettings);
  }

  const totalColumnIndexes = new Set(bannerStructure.totalColumnIndexes || []);
  const globalTotalColumnIndex =
    bannerStructure.globalTotalColumnIndex === undefined
      ? null
      : bannerStructure.globalTotalColumnIndex;

  if (columnIndex === globalTotalColumnIndex || totalColumnIndexes.has(columnIndex)) {
    return "";
  }

  const groupSegmentColumnIndexes = group.columnIndexes.filter(
    (candidateColumnIndex) =>
      candidateColumnIndex !== undefined &&
      candidateColumnIndex !== null &&
      candidateColumnIndex !== globalTotalColumnIndex &&
      !totalColumnIndexes.has(candidateColumnIndex)
  );

  const localIndex = groupSegmentColumnIndexes.indexOf(columnIndex);

  if (localIndex < 0) {
    return "";
  }

  const labels = generateSignificanceLabels(
    significanceLabelOptionsFromSettings(calculationSettings, groupSegmentColumnIndexes.length)
  );

  return labels[localIndex] || "";
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
 *
 * When `useSpaceSeparator` is true (multi-character overflow mode), markers are
 * separated with spaces so adjacent multi-character markers stay distinct, e.g.
 * "aa ab ac" rather than the ambiguous "aaabac".
 */
function appendSignificanceMarkerToCellResult(cellResult, marker, useSpaceSeparator = false) {
  if (!marker) {
    return;
  }

  if (useSpaceSeparator && cellResult.markers) {
    cellResult.markers += ` ${marker}`;
  } else {
    cellResult.markers += marker;
  }

  applyFillReasonToCellResult(cellResult, CELL_FILL_REASONS.SIGNIFICANT);
}

/**
 * Prepends Total marker and applies corresponding fill.
 *
 * T = segment is significantly higher than Total.
 * t = segment is significantly lower than Total.
 */
function prependTotalMarkerToCellResult(cellResult, totalMarker, useSpaceSeparator = false) {
  const markerText = cellResult.markers || "";
  const markerTextWithoutOldTotalMarker = markerText.replace(/[tT]/g, "").trim();

  const separator =
    useSpaceSeparator && markerTextWithoutOldTotalMarker ? " " : "";

  cellResult.markers = `${totalMarker}${separator}${markerTextWithoutOldTotalMarker}`;

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

  // A marker token is a run of marker characters (and previous-column arrows).
  // Match one trailing token plus any further space-separated tokens so that
  // multi-character overflow markers ("42% aa ab ac") are removed in full, not
  // just the last group.
  const markerToken = `[${markerCharacters}↑↓]+`;
  const markerSuffixPattern = new RegExp(`\\s*${markerToken}(?:\\s+${markerToken})*$`);

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
  const tThreshold = getTThresholdForConfidence(confidenceLevel, degreesOfFreedom, {
    oneTailedTest: calculationSettings.oneTailedTest,
  });
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
    calculationSettings.excludedColumnIndexes || new Set(),
    calculationSettings.bannerStructure || null
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
      groupKey: comparisonPair.groupKey,
      groupLabel: comparisonPair.groupLabel,
      totalReferenceType: comparisonPair.totalReferenceType,

      firstValue: meanRow[firstColumnIndex],
      secondValue: meanRow[secondColumnIndex],
      firstSpread: spreadRow[firstColumnIndex],
      secondSpread: spreadRow[secondColumnIndex],
      firstBase: baseRow[firstColumnIndex],
      secondBase: baseRow[secondColumnIndex],

      result: significanceResult,
      autoPreviousColumnFromBanner: comparisonPair.autoPreviousColumnFromBanner,
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
    calculationSettings.excludedColumnIndexes || new Set(),
    calculationSettings.bannerStructure || null
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
      groupKey: comparisonPair.groupKey,
      groupLabel: comparisonPair.groupLabel,
      totalReferenceType: comparisonPair.totalReferenceType,

      firstValue: npsRow[firstColumnIndex],
      secondValue: npsRow[secondColumnIndex],
      firstPromoters: promotersRow[firstColumnIndex],
      secondPromoters: promotersRow[secondColumnIndex],
      firstDetractors: detractorsRow[firstColumnIndex],
      secondDetractors: detractorsRow[secondColumnIndex],
      firstBase: baseRow[firstColumnIndex],
      secondBase: baseRow[secondColumnIndex],

      result: significanceResult,
      autoPreviousColumnFromBanner: comparisonPair.autoPreviousColumnFromBanner,
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
    calculationSettings.excludedColumnIndexes || new Set(),
    calculationSettings.bannerStructure || null
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
      groupKey: comparisonPair.groupKey,
      groupLabel: comparisonPair.groupLabel,
      totalReferenceType: comparisonPair.totalReferenceType,

      firstValue: npsRow[firstColumnIndex],
      secondValue: npsRow[secondColumnIndex],
      firstSpread: spreadRow[firstColumnIndex],
      secondSpread: spreadRow[secondColumnIndex],
      firstBase: baseRow[firstColumnIndex],
      secondBase: baseRow[secondColumnIndex],

      result: significanceResult,
      autoPreviousColumnFromBanner: comparisonPair.autoPreviousColumnFromBanner,
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
  // In multi-character overflow mode, separate markers with spaces so adjacent
  // multi-character markers stay distinct ("aa ab ac", not "aaabac").
  const useSpaceSeparator = Boolean(calculationSettings.allowMultiCharacterMarkers);

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

      if (
        comparison.comparisonType === "bannerTotal" ||
        (calculationSettings.firstColumnIsTotal && comparison.comparisonType === "total")
      ) {
        applyTotalComparisonMarkerToFullCellResultMatrix(
          fullCellResultMatrix,
          valueRowIndex,
          comparison,
          useSpaceSeparator
        );

        continue;
      }

      if (calculationSettings.firstColumnIsTotal) {
        if (firstColumnIndex === 0 || secondColumnIndex === 0) {
          continue;
        }
      }

      const firstLabel = getSignificanceLabelForComparisonColumn(
        firstColumnIndex,
        calculationSettings,
        comparison
      );

      const secondLabel = getSignificanceLabelForComparisonColumn(
        secondColumnIndex,
        calculationSettings,
        comparison
      );

      if (!firstLabel || !secondLabel) {
        continue;
      }

      if (comparison.result.direction === "first_higher") {
        appendSignificanceMarkerToCellResult(
          fullCellResultMatrix[valueRowIndex][firstColumnIndex],
          secondLabel,
          useSpaceSeparator
        );
      }

      if (comparison.result.direction === "second_higher") {
        appendSignificanceMarkerToCellResult(
          fullCellResultMatrix[valueRowIndex][secondColumnIndex],
          firstLabel,
          useSpaceSeparator
        );
      }
    }
  }
}

/**
 * Applies Total comparison marker.
 *
 * RULES:
 * - firstColumnIndex is treated as Total reference for Total comparison pairs.
 * - marker is always written into compared/target column.
 * - "T" means target is significantly higher than Total reference.
 * - "t" means target is significantly lower than Total reference.
 */
function applyTotalComparisonMarkerToFullCellResultMatrix(
  fullCellResultMatrix,
  valueRowIndex,
  comparison,
  useSpaceSeparator = false
) {
  const totalReferenceColumnIndex = comparison.firstColumnIndex;
  const targetColumnIndex = comparison.secondColumnIndex;

  if (targetColumnIndex === totalReferenceColumnIndex) {
    return;
  }

  const targetIsHigher = comparison.result.direction === "second_higher";

  const targetIsLower = comparison.result.direction === "first_higher";

  if (!targetIsHigher && !targetIsLower) {
    return;
  }

  const totalMarker = targetIsHigher ? "T" : "t";

  prependTotalMarkerToCellResult(
    fullCellResultMatrix[valueRowIndex][targetColumnIndex],
    totalMarker,
    useSpaceSeparator
  );
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

  const shouldApplyPreviousColumnFill =
    calculationSettings.applyPreviousColumnFill || comparison.autoPreviousColumnFromBanner;

  if (comparison.result.direction === "second_higher") {
    cellResult.previousColumnArrow = "↑";
    cellResult.previousColumnArrowDirection = "up";

    if (shouldApplyPreviousColumnFill) {
      applyFillReasonToCellResult(cellResult, CELL_FILL_REASONS.SIGNIFICANT);
    }

    return;
  }

  if (comparison.result.direction === "first_higher") {
    cellResult.previousColumnArrow = "↓";
    cellResult.previousColumnArrowDirection = "down";

    if (shouldApplyPreviousColumnFill) {
      applyFillReasonToCellResult(cellResult, CELL_FILL_REASONS.LOWER_THAN_TOTAL);
    }
  }
}

/**
 * Builds banner-aware Total comparison pairs.
 *
 * If global Total exists:
 * - all other non-excluded columns are compared with global Total;
 * - local Totals are ordinary columns and can be compared with global Total.
 *
 * If no global Total exists:
 * - each group uses its local Total where available.
 */
function buildBannerTotalComparisonPairs(
  columnCount,
  bannerStructure,
  excludedColumnIndexes = new Set()
) {
  const pairs = [];

  const globalTotalColumnIndex =
    bannerStructure.globalTotalColumnIndex === undefined
      ? null
      : bannerStructure.globalTotalColumnIndex;

  const isExcluded = (columnIndex) => excludedColumnIndexes.has(columnIndex);

  if (
    globalTotalColumnIndex !== null &&
    globalTotalColumnIndex !== undefined &&
    !isExcluded(globalTotalColumnIndex)
  ) {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      if (columnIndex === globalTotalColumnIndex) {
        continue;
      }

      if (isExcluded(columnIndex)) {
        continue;
      }

      pairs.push({
        firstColumnIndex: globalTotalColumnIndex,
        secondColumnIndex: columnIndex,
        comparisonType: "bannerTotal",
        totalReferenceType: "global",
      });
    }

    return pairs;
  }

  for (const group of bannerStructure.groups || []) {
    const localTotalColumnIndexes = group.localTotalColumnIndexes || [];

    if (localTotalColumnIndexes.length === 0) {
      continue;
    }

    if (localTotalColumnIndexes.length > 1) {
      addBannerStructureMessageOnce(
        bannerStructure,
        "error",
        "BANNER_MULTIPLE_LOCAL_TOTALS",
        `В группе “${group.label}” найдено несколько Тоталов. Расчёт остановлен: невозможно однозначно определить колонку для сравнения.`
      );

      continue;
    }

    const localTotalColumnIndex = localTotalColumnIndexes[0];

    if (isExcluded(localTotalColumnIndex)) {
      continue;
    }

    for (const columnIndex of group.columnIndexes || []) {
      if (columnIndex === localTotalColumnIndex) {
        continue;
      }

      if (columnIndex < 0 || columnIndex >= columnCount) {
        continue;
      }

      if (isExcluded(columnIndex)) {
        continue;
      }

      pairs.push({
        firstColumnIndex: localTotalColumnIndex,
        secondColumnIndex: columnIndex,
        comparisonType: "bannerTotal",
        totalReferenceType: "local",
        groupKey: group.groupKey,
        groupLabel: group.label,
      });
    }
  }

  return pairs;
}

/**
 * Adds banner message only once.
 */
function addBannerStructureMessageOnce(bannerStructure, severity, code, text) {
  if (!bannerStructure) {
    return;
  }

  if (!bannerStructure.messages) {
    bannerStructure.messages = [];
  }

  const alreadyExists = bannerStructure.messages.some(
    (message) => message.code === code && message.text === text
  );

  if (alreadyExists) {
    return;
  }

  bannerStructure.messages.push({
    severity,
    code,
    text,
  });
}

/**
 * Builds previous-column comparison pairs within detected banner groups.
 *
 * RULES:
 * - compare current column only with immediately previous selected column;
 * - both columns must belong to the same banner group;
 * - do not cross group boundaries;
 * - do not skip over excluded columns;
 * - if excludeTotalFromComparisons is enabled, Total columns are excluded too.
 */
function buildBannerPreviousColumnComparisonPairs(
  columnCount,
  bannerStructure,
  excludedColumnIndexes = new Set(),
  options = {}
) {
  const pairs = [];

  const totalColumnIndexes = options.totalColumnIndexes || new Set();

  const isExcluded = (columnIndex) => {
    if (excludedColumnIndexes.has(columnIndex)) {
      return true;
    }

    if (totalColumnIndexes.has(columnIndex)) {
      return true;
    }

    return false;
  };

  const groupKeyByColumnIndex = buildBannerGroupKeyByColumnIndex(bannerStructure);

  for (let columnIndex = 1; columnIndex < columnCount; columnIndex++) {
    const previousColumnIndex = columnIndex - 1;

    if (isExcluded(previousColumnIndex) || isExcluded(columnIndex)) {
      continue;
    }

    const previousGroupKey = groupKeyByColumnIndex.get(previousColumnIndex);
    const currentGroupKey = groupKeyByColumnIndex.get(columnIndex);

    if (!previousGroupKey || !currentGroupKey) {
      continue;
    }

    if (previousGroupKey !== currentGroupKey) {
      continue;
    }

    const group = (bannerStructure.groups || []).find(
      (candidate) => candidate.groupKey === currentGroupKey
    );

    pairs.push({
      firstColumnIndex: previousColumnIndex,
      secondColumnIndex: columnIndex,
      comparisonType: "previousColumn",
      groupKey: currentGroupKey,
      groupLabel: group ? group.label : "",
    });
  }

  return pairs;
}

/**
 * Builds previous-column pairs inside one banner group.
 *
 * RULES:
 * - only immediate selected-neighbor columns are compared;
 * - no cross-group comparisons;
 * - excluded columns are not skipped over;
 * - Total columns are excluded from this chain.
 */
function buildBannerPreviousColumnPairsForGroup(
  group,
  excludedColumnIndexes = new Set(),
  totalColumnIndexes = new Set(),
  options = {}
) {
  const pairs = [];

  const autoApplied = Boolean(options.autoApplied);

  const groupColumnIndexes = group.columnIndexes || [];

  const isExcluded = (columnIndex) => {
    if (excludedColumnIndexes.has(columnIndex)) {
      return true;
    }

    if (totalColumnIndexes.has(columnIndex)) {
      return true;
    }

    return false;
  };

  for (let index = 1; index < groupColumnIndexes.length; index++) {
    const previousColumnIndex = groupColumnIndexes[index - 1];
    const currentColumnIndex = groupColumnIndexes[index];

    if (isExcluded(previousColumnIndex) || isExcluded(currentColumnIndex)) {
      continue;
    }

    pairs.push({
      firstColumnIndex: previousColumnIndex,
      secondColumnIndex: currentColumnIndex,
      comparisonType: "previousColumn",
      groupKey: group.groupKey,
      groupLabel: group.label,
      autoPreviousColumnFromBanner: autoApplied,
    });
  }

  return pairs;
}

/**
 * Builds map: columnIndex -> banner group key.
 */
function buildBannerGroupKeyByColumnIndex(bannerStructure) {
  const groupKeyByColumnIndex = new Map();

  for (const group of bannerStructure.groups || []) {
    for (const columnIndex of group.columnIndexes || []) {
      groupKeyByColumnIndex.set(columnIndex, group.groupKey);
    }
  }

  return groupKeyByColumnIndex;
}

/**
 * Builds column label map for detected banner structure.
 *
 * PURPOSE:
 * Used by banner-aware header marker writing.
 *
 * RULES:
 * - labels are local to each banner group;
 * - all Total columns are always skipped;
 * - skipped columns do not consume label indexes;
 * - global Total is never labeled.
 */
export function buildBannerLocalSignificanceLabelMap(bannerStructure, calculationSettings = {}) {
  const labelMap = new Map();

  if (!bannerStructure || !bannerStructure.groups) {
    return labelMap;
  }

  // The widest banner group decides how many labels we may need; when overflow
  // markers are enabled this can extend past the single-character alphabet.
  const requiredLabelCount = computeRequiredSignificanceLabelCount(
    1,
    { ...calculationSettings, respectBannerStructure: true },
    bannerStructure
  );

  const labels = generateSignificanceLabels(
    significanceLabelOptionsFromSettings(calculationSettings, requiredLabelCount)
  );

  const totalColumnIndexes = new Set(bannerStructure.totalColumnIndexes || []);
  const globalTotalColumnIndex =
    bannerStructure.globalTotalColumnIndex === undefined
      ? null
      : bannerStructure.globalTotalColumnIndex;

  for (const group of bannerStructure.groups || []) {
    if (
      !calculationSettings.compareWithPreviousColumn &&
      group.recommendedComparisonMode === "previousColumn"
    ) {
      continue;
    }

    let localLabelIndex = 0;

    for (const columnIndex of group.columnIndexes || []) {
      if (columnIndex === globalTotalColumnIndex) {
        continue;
      }

      if (totalColumnIndexes.has(columnIndex)) {
        continue;
      }

      const label = labels[localLabelIndex];

      if (label) {
        labelMap.set(columnIndex, label);
      }

      localLabelIndex++;
    }
  }

  return labelMap;
}

/**
 * Returns true if banner group should use automatic previous-column mode.
 *
 * This is used only when global compareWithPreviousColumn is OFF.
 */
function shouldUseAutoPreviousColumnForBannerGroup(group) {
  return group && group.recommendedComparisonMode === "previousColumn";
}
