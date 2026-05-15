import { METRIC_DICTIONARY } from "./config/dictionary.config"; // Импортируем наш конфиг
import { normalizeLookupText } from "./string-utils";

export const LABEL_SCAN_COLUMNS_LEFT = 2;

export function normalizeLabelText(rawLabel) {
  return normalizeLookupText(rawLabel);
}

function doesKeywordMatchLabel(normalizedLabel, rawKeyword) {
  const normalizedKeyword = normalizeLabelText(rawKeyword);

  if (!normalizedLabel || !normalizedKeyword) {
    return false;
  }

  if (normalizedLabel === normalizedKeyword) {
    return true;
  }

  // Short abbreviations are too risky for substring matching.
  // Example: "ско" can match inside "мужской".
  if (normalizedKeyword.length <= 3) {
    return false;
  }

  return normalizedLabel.includes(normalizedKeyword);
}

/**
 * Checks whether normalized label matches any known keyword.
 *
 * Short abbreviations are matched only exactly.
 * This prevents false positives like:
 * - "мужской" matching "ско"
 */
function labelContainsAnyKeyword(normalizedLabel, keywords) {
  return keywords.some((keyword) => doesKeywordMatchLabel(normalizedLabel, keyword));
}

/**
 * Classifies one normalized row label based on the dictionary config.
 *
 * PURPOSE:
 * Detect what kind of data row this is:
 * proportion, mean, SD, variance, NPS, promoters, detractors, base, or unknown.
 */
export function classifyMetricLabel(rawLabel) {
  const normalizedLabel = normalizeLabelText(rawLabel);

  if (!normalizedLabel) {
    return {
      rowType: "empty",
      normalizedLabel,
    };
  }

  // Проходимся по словарю. Как только находим совпадение — возвращаем тип.
  for (const dictionaryEntry of METRIC_DICTIONARY) {
    if (labelContainsAnyKeyword(normalizedLabel, dictionaryEntry.keywords)) {
      const result = {
        rowType: dictionaryEntry.rowType,
        normalizedLabel,
      };
      if (dictionaryEntry.baseSubtype !== undefined) {
        result.baseSubtype = dictionaryEntry.baseSubtype;
      }
      return result;
    }
  }

  // Если ни одно слово из конфига не подошло
  return {
    rowType: "unknownText",
    normalizedLabel,
  };
}

/**
 * Extracts the best available label for one selected data row.
 *
 * PURPOSE:
 * We scan 1-2 cells to the left of the selected data row.
 * If several cells contain text, we prefer the closest non-empty text.
 *
 * INPUT:
 * leftLabelRowValues - array of values from cells left of one selected row.
 */
export function extractRowLabelFromLeftCells(leftLabelRowValues) {
  if (!leftLabelRowValues || leftLabelRowValues.length === 0) {
    return "";
  }

  for (
    let labelColumnIndex = leftLabelRowValues.length - 1;
    labelColumnIndex >= 0;
    labelColumnIndex--
  ) {
    const currentLabelValue = leftLabelRowValues[labelColumnIndex];

    if (isNumericLikeCellValue(currentLabelValue)) {
      continue;
    }

    const normalizedLabel = normalizeLabelText(currentLabelValue);

    if (normalizedLabel) {
      return String(currentLabelValue);
    }
  }

  return "";
}

/**
 * Detects row labels for a selected data range.
 *
 * PURPOSE:
 * This is the first diagnostic detector step.
 * It does not calculate significance yet.
 *
 * INPUT:
 * selectedValues - 2D array of selected data values.
 * leftLabelValues - 2D array of cells located to the left of selected data.
 *
 * OUTPUT:
 * Diagnostic object describing detected row labels.
 */
export function detectMetricRowsFromLeftLabels(selectedValues, leftLabelValues) {
  const rowDiagnostics = []; // One diagnostic item per selected data row.
  const selectedRowCount = selectedValues ? selectedValues.length : 0; // Number of selected rows.

  for (let rowIndex = 0; rowIndex < selectedRowCount; rowIndex++) {
    const leftLabelRowValues = leftLabelValues ? leftLabelValues[rowIndex] : []; // Labels left of current row.
    const rawLabel = extractRowLabelFromLeftCells(leftLabelRowValues); // Best label candidate.
    const classification = classifyMetricLabel(rawLabel); // Row type detection result.

    const rowDiagnostic = {
      rowIndex,
      displayRowNumber: rowIndex + 1,
      rawLabel,
      normalizedLabel: classification.normalizedLabel,
      rowType: classification.rowType,
    };
    if (classification.baseSubtype !== undefined) {
      rowDiagnostic.baseSubtype = classification.baseSubtype;
    }
    rowDiagnostics.push(rowDiagnostic);
  }

  return {
    rowDiagnostics,
  };
}

/**
 * Formats detector diagnostics for display in the Excel task pane.
 *
 * PURPOSE:
 * Temporary output for validating detection on real spreadsheet tables.
 */
export function formatMetricDetectionDiagnostics(detectionResult) {
  const outputLines = []; // Lines displayed in task pane.

  outputLines.push("Metric detection diagnostics");
  outputLines.push("");

  for (const rowDiagnostic of detectionResult.rowDiagnostics) {
    outputLines.push(
      `Row ${rowDiagnostic.displayRowNumber}: ` +
        `label="${rowDiagnostic.rawLabel || ""}", ` +
        `type=${rowDiagnostic.rowType}`
    );
  }

  return outputLines.join("\n");
}

/**
 * Searches pending proportion row indexes for the first row matching targetType.
 *
 * PURPOSE:
 * Used by extended NPS detection to locate Detractors and Promoters that were
 * already buffered as proportion rows before the NPS row was encountered.
 */
function findRowTypeInPending(rowDiagnostics, pendingRows, targetType) {
  for (const rowIndex of pendingRows) {
    const rowDiagnostic = rowDiagnostics[rowIndex];

    if (rowDiagnostic?.rowType === targetType) {
      return rowIndex;
    }
  }
  return null;
}

/**
 * Finds the nearest base row below a given row.
 *
 * PURPOSE:
 * In complex tables, a base row may be shared across several metrics.
 * Example:
 * % row
 * % row
 * Mean row
 * SD row
 * Base row
 *
 * In this case, the same base may apply both to proportions and means.
 */
function findNextBaseRowIndex(rowDiagnostics, startRowIndex) {
  for (let rowIndex = startRowIndex + 1; rowIndex < rowDiagnostics.length; rowIndex++) {
    if (rowDiagnostics[rowIndex].rowType === "base") {
      return rowIndex;
    }
  }

  return null;
}

/**
 * Checks whether a normalized label refers to the NPS Neutral/Passive segment.
 *
 * PURPOSE:
 * Identify the optional middle row in NPS-first format 2 (NPS / Promoters / Neutral / Detractors / BASE).
 * Neutral has no dictionary entry and classifies as unknownText, so we check the label directly
 * to avoid matching arbitrary unknownText rows.
 */
function isNeutralLabel(normalizedLabel) {
  if (!normalizedLabel) {
    return false;
  }

  const neutralKeywords = [
    "neutral",
    "neutrals",
    "пассивные",
    "пассивный",
    "нейтральные",
    "нейтральный",
  ];

  return neutralKeywords.some((keyword) => doesKeywordMatchLabel(normalizedLabel, keyword));
}

/**
 * Returns the baseSubtype for a detected base row, if any.
 */
function getBaseSubtype(rowDiagnostics, baseRowIndex) {
  return rowDiagnostics[baseRowIndex]?.baseSubtype;
}

/**
 * Attaches baseSubtype to a block object when the base row carries one.
 */
function attachBaseSubtype(block, rowDiagnostics, baseRowIndex) {
  const baseSubtype = getBaseSubtype(rowDiagnostics, baseRowIndex);
  if (baseSubtype !== undefined) {
    block.baseSubtype = baseSubtype;
  }
  return block;
}

/**
 * Attempts to parse NPS-first pattern (format 1 or 2) at rowIndex.
 *
 * FORMAT 1: NPS / Promoters / Detractors / BASE
 * FORMAT 2: NPS / Promoters / Neutral / Detractors / BASE
 *
 * RETURNS:
 * { npsIdx, promotersIdx, detractorsIdx, baseRowIndex, hasNeutral, neutralIdx? }
 * or null if pattern does not match.
 */
function tryParseNpsFirstBlock(rowDiagnostics, rowIndex) {
  const currentRowType = rowDiagnostics[rowIndex]?.rowType;
  if (currentRowType !== "nps") {
    return null;
  }

  // Format 1: NPS / Promoters / Detractors / BASE
  if (
    rowIndex + 3 < rowDiagnostics.length &&
    rowDiagnostics[rowIndex + 1].rowType === "promoters" &&
    rowDiagnostics[rowIndex + 2].rowType === "detractors" &&
    rowDiagnostics[rowIndex + 3].rowType === "base"
  ) {
    return {
      npsIdx: rowIndex,
      promotersIdx: rowIndex + 1,
      detractorsIdx: rowIndex + 2,
      baseRowIndex: rowIndex + 3,
      hasNeutral: false,
    };
  }

  // Format 2: NPS / Promoters / Neutral / Detractors / BASE
  if (
    rowIndex + 4 < rowDiagnostics.length &&
    rowDiagnostics[rowIndex + 1].rowType === "promoters" &&
    isNeutralLabel(rowDiagnostics[rowIndex + 2].normalizedLabel) &&
    rowDiagnostics[rowIndex + 3].rowType === "detractors" &&
    rowDiagnostics[rowIndex + 4].rowType === "base"
  ) {
    return {
      npsIdx: rowIndex,
      promotersIdx: rowIndex + 1,
      neutralIdx: rowIndex + 2,
      detractorsIdx: rowIndex + 3,
      baseRowIndex: rowIndex + 4,
      hasNeutral: true,
    };
  }

  return null;
}
/**
 * Returns a numeric priority for a base row. Lower value = higher priority.
 * Priority: Effective (0) > Unweighted (1) > plain Base (2) > Weighted (3).
 */
function getBasePriorityValue(rowDiagnostic) {
  const subtype = rowDiagnostic?.baseSubtype;
  if (subtype === "effective") return 0;
  if (subtype === "unweighted") return 1;
  if (subtype === "weighted") return 3;
  return 2; // plain Base
}

/**
 * Selects the highest-priority base row from a consecutive run of base rows
 * starting at firstBaseIndex.
 *
 * PURPOSE:
 * When a block is followed by multiple base rows (e.g. Weighted Base then
 * Effective Base), pick the one best suited for significance testing rather
 * than always taking the first.
 */
function selectBestFromConsecutiveBases(rowDiagnostics, firstBaseIndex) {
  let bestIndex = firstBaseIndex;
  let bestPriority = getBasePriorityValue(rowDiagnostics[firstBaseIndex]);

  for (let i = firstBaseIndex + 1; i < rowDiagnostics.length; i++) {
    if (rowDiagnostics[i].rowType !== "base") break;
    const p = getBasePriorityValue(rowDiagnostics[i]);
    if (p < bestPriority) {
      bestPriority = p;
      bestIndex = i;
    }
  }

  return bestIndex;
}

/**
 * Finds the nearest base row below startRowIndex and selects the best
 * from consecutive base rows at that position.
 */
function findBestBaseRowIndex(rowDiagnostics, startRowIndex) {
  const firstBase = findNextBaseRowIndex(rowDiagnostics, startRowIndex);
  if (firstBase === null) return null;
  return selectBestFromConsecutiveBases(rowDiagnostics, firstBase);
}

/**
 * Builds calculation blocks from detected row labels.
 *
 * PURPOSE:
 * Support complex tables where proportions, means, and NPS can appear
 * in one selected range in different combinations.
 */
export function buildCalculationBlocks(detectionResult) {
  const rowDiagnostics = detectionResult.rowDiagnostics; // Classified rows.
  const calculationBlocks = []; // Final list of calculation blocks.
  const pendingProportionRows = []; // Proportion rows waiting for the next available base.

  let rowIndex = 0; // Current row scanner position.

  while (rowIndex < rowDiagnostics.length) {
    const currentRowType = rowDiagnostics[rowIndex].rowType; // Current detected row type.

    // 1. Проценты собираем в буфер
    if (isProportionValueRowType(currentRowType)) {
      pendingProportionRows.push(rowIndex);
      rowIndex++;
      continue;
    }

    // 2. Строка Базы: закрываем висящие проценты, если они есть
    if (currentRowType === "base") {
      if (pendingProportionRows.length > 0) {
        const bestBaseIndex = selectBestFromConsecutiveBases(rowDiagnostics, rowIndex);
        calculationBlocks.push(
          attachBaseSubtype(
            { metricType: "proportion", valueRowIndexes: [...pendingProportionRows], baseRowIndex: bestBaseIndex },
            rowDiagnostics, bestBaseIndex
          )
        );
        pendingProportionRows.length = 0;
      }
      rowIndex++;
      continue;
    }

    // 3. Блок Mean (Средние + Разброс)
    if (
      currentRowType === "mean" &&
      rowIndex + 1 < rowDiagnostics.length &&
      (rowDiagnostics[rowIndex + 1].rowType === "standardDeviation" ||
        rowDiagnostics[rowIndex + 1].rowType === "variance")
    ) {
      const spreadRowIndex = rowIndex + 1;
      const baseRowIndex = findBestBaseRowIndex(rowDiagnostics, spreadRowIndex);

      if (baseRowIndex !== null) {
        calculationBlocks.push(
          attachBaseSubtype(
            { metricType: "mean", valueRowIndex: rowIndex, spreadRowIndex, spreadType: rowDiagnostics[spreadRowIndex].rowType, baseRowIndex },
            rowDiagnostics, baseRowIndex
          )
        );

        if (pendingProportionRows.length > 0) {
          calculationBlocks.push(
            attachBaseSubtype(
              { metricType: "proportion", valueRowIndexes: [...pendingProportionRows], baseRowIndex },
              rowDiagnostics, baseRowIndex
            )
          );
          pendingProportionRows.length = 0;
        }

        // ИСПРАВЛЕНИЕ: Прыгаем только через строки текущего блока (Среднее + Разброс), а не к Базе
        rowIndex = spreadRowIndex + 1;
        continue;
      }
    }


    // 4. NPS-first (format 1 and 2, unified handler)
    const npsFirstBlock = tryParseNpsFirstBlock(rowDiagnostics, rowIndex);
    if (npsFirstBlock !== null) {
      const baseRowIndex = selectBestFromConsecutiveBases(rowDiagnostics, npsFirstBlock.baseRowIndex);

      if (pendingProportionRows.length > 0) {
        calculationBlocks.push(
          attachBaseSubtype(
            { metricType: "proportion", valueRowIndexes: [...pendingProportionRows], baseRowIndex },
            rowDiagnostics, baseRowIndex
          )
        );
        pendingProportionRows.length = 0;
      }

      const proportionValueIndexes = npsFirstBlock.hasNeutral
        ? [npsFirstBlock.promotersIdx, npsFirstBlock.neutralIdx, npsFirstBlock.detractorsIdx]
        : [npsFirstBlock.promotersIdx, npsFirstBlock.detractorsIdx];

      calculationBlocks.push(
        attachBaseSubtype(
          { metricType: "proportion", valueRowIndexes: proportionValueIndexes, baseRowIndex },
          rowDiagnostics, baseRowIndex
        )
      );

      calculationBlocks.push(
        attachBaseSubtype(
          { metricType: "npsStructure", valueRowIndex: npsFirstBlock.npsIdx, promotersRowIndex: npsFirstBlock.promotersIdx, detractorsRowIndex: npsFirstBlock.detractorsIdx, baseRowIndex },
          rowDiagnostics, baseRowIndex
        )
      );

      rowIndex = baseRowIndex + 1;
      continue;
    }

    // 5. Блок NPS Spread (NPS + Разброс)
    if (
      currentRowType === "nps" &&
      rowIndex + 1 < rowDiagnostics.length &&
      (rowDiagnostics[rowIndex + 1].rowType === "standardDeviation" ||
        rowDiagnostics[rowIndex + 1].rowType === "variance")
    ) {
      const spreadRowIndex = rowIndex + 1;
      const baseRowIndex = findBestBaseRowIndex(rowDiagnostics, spreadRowIndex);

      if (baseRowIndex !== null) {
        calculationBlocks.push(
          attachBaseSubtype(
            { metricType: "npsSpread", valueRowIndex: rowIndex, spreadRowIndex, spreadType: rowDiagnostics[spreadRowIndex].rowType, baseRowIndex },
            rowDiagnostics, baseRowIndex
          )
        );

        if (pendingProportionRows.length > 0) {
          calculationBlocks.push(
            attachBaseSubtype(
              { metricType: "proportion", valueRowIndexes: [...pendingProportionRows], baseRowIndex },
              rowDiagnostics, baseRowIndex
            )
          );
          pendingProportionRows.length = 0;
        }

        // ИСПРАВЛЕНИЕ: Прыгаем только через строки текущего блока
        rowIndex = spreadRowIndex + 1;
        continue;
      }
    }

    // 6. Extended NPS — NPS row follows Detractors and Promoters already buffered as proportion rows
    // Handles: 1–10, Bottom-3, Mid-4, Top-3, Detractors, [Neutral], Promoters, NPS, BASE
    if (currentRowType === "nps") {
      const detractorsIdx = findRowTypeInPending(
        rowDiagnostics,
        pendingProportionRows,
        "detractors"
      );
      const promotersIdx = findRowTypeInPending(rowDiagnostics, pendingProportionRows, "promoters");

      if (detractorsIdx !== null && promotersIdx !== null) {
        const baseRowIndex = findBestBaseRowIndex(rowDiagnostics, rowIndex);

        if (baseRowIndex !== null) {
          // All buffered rows, including Detractors and Promoters, receive proportion markers.
          calculationBlocks.push(
            attachBaseSubtype(
              { metricType: "proportion", valueRowIndexes: [...pendingProportionRows], baseRowIndex },
              rowDiagnostics, baseRowIndex
            )
          );
          pendingProportionRows.length = 0;

          // NPS row uses NPS significance logic; Detractors and Promoters are its support inputs.
          calculationBlocks.push(
            attachBaseSubtype(
              { metricType: "npsStructure", valueRowIndex: rowIndex, promotersRowIndex: promotersIdx, detractorsRowIndex: detractorsIdx, baseRowIndex },
              rowDiagnostics, baseRowIndex
            )
          );

          rowIndex++;
          continue;
        }
      }
    }

    rowIndex++;
  }

  return calculationBlocks;
}

/**
 * Checks whether row can be treated as a proportion value row.
 *
 * PURPOSE:
 * Prevent service rows like SD, Variance, and Base from being calculated
 * as ordinary proportions.
 *
 * Promoters and Detractors are treated as ordinary proportion rows and can also
 * serve as support rows for NPS calculations.
 */
function isProportionValueRowType(rowType) {
  return (
    rowType === "proportion" ||
    rowType === "promoters" ||
    rowType === "detractors" ||
    rowType === "empty" ||
    rowType === "unknownText"
  );
}

/**
 * Returns row indexes where significance markers are allowed.
 *
 * PURPOSE:
 * Prevent service rows from receiving significance letters.
 *
 * Markers are allowed only in actual metric value rows:
 * - proportion value rows
 * - mean row
 * - NPS row
 *
 * Markers are NOT allowed in:
 * - base rows
 * - SD / variance rows
 */
export function getAllowedMarkerRowIndexes(calculationBlocks) {
  const allowedMarkerRows = new Set(); // Rows where marker letters may be written.

  for (const calculationBlock of calculationBlocks) {
    if (calculationBlock.metricType === "proportion") {
      for (const valueRowIndex of calculationBlock.valueRowIndexes) {
        allowedMarkerRows.add(valueRowIndex);
      }
    }

    if (
      calculationBlock.metricType === "mean" ||
      calculationBlock.metricType === "npsStructure" ||
      calculationBlock.metricType === "npsSpread"
    ) {
      allowedMarkerRows.add(calculationBlock.valueRowIndex);
    }
  }

  return allowedMarkerRows;
}

/**
 * Checks whether a cell value looks like a numeric data value.
 *
 * PURPOSE:
 * Prevent numeric columns located between labels and selected range
 * from being treated as row labels.
 */
function isNumericLikeCellValue(rawValue) {
  if (rawValue === null || rawValue === undefined || rawValue === "") {
    return false;
  }

  if (typeof rawValue === "number") {
    return true;
  }

  const textValue = String(rawValue).trim().replace("%", "").replace(",", ".");

  if (textValue === "") {
    return false;
  }

  return !Number.isNaN(Number(textValue));
}
