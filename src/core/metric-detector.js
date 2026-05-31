import { METRIC_DICTIONARY } from "./config/dictionary.config"; // Импортируем наш конфиг
import { normalizeLookupText } from "./string-utils";
import { isGeneratedSignificanceFootnoteRow } from "./significance-footnote";

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

  // Collect all non-empty, non-numeric text cells left-to-right so that
  // split two-column labels (e.g. ["Base", "weighted"]) are concatenated
  // and detected as a single label ("Base weighted").
  const parts = [];

  for (let i = 0; i < leftLabelRowValues.length; i++) {
    const cellValue = leftLabelRowValues[i];

    if (isNumericLikeCellValue(cellValue)) {
      continue;
    }

    const normalizedLabel = normalizeLabelText(cellValue);

    if (normalizedLabel) {
      parts.push(String(cellValue));
    }
  }

  return parts.join(" ");
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
 *
 * When preferredBase is set (not "auto"), try to find a base with that
 * subtype in the consecutive run. If not found, fall back to auto priority.
 */
function selectBestFromConsecutiveBases(rowDiagnostics, firstBaseIndex, { preferredBase = "auto" } = {}) {
  // Collect all consecutive base rows starting at firstBaseIndex.
  const consecutiveIndexes = [firstBaseIndex];
  for (let i = firstBaseIndex + 1; i < rowDiagnostics.length; i++) {
    if (rowDiagnostics[i].rowType !== "base") break;
    consecutiveIndexes.push(i);
  }

  if (preferredBase !== "auto") {
    // Try to find a base row matching the preferred subtype.
    for (const idx of consecutiveIndexes) {
      const subtype = rowDiagnostics[idx]?.baseSubtype;
      // "plain" means no baseSubtype (undefined).
      const isMatch =
        preferredBase === "plain" ? subtype === undefined : subtype === preferredBase;
      if (isMatch) {
        return idx;
      }
    }
    // Preferred subtype not found — fall through to auto priority.
  }

  // Auto priority: Effective (0) > Unweighted (1) > plain Base (2) > Weighted (3).
  let bestIndex = firstBaseIndex;
  let bestPriority = getBasePriorityValue(rowDiagnostics[firstBaseIndex]);
  for (let i = 1; i < consecutiveIndexes.length; i++) {
    const idx = consecutiveIndexes[i];
    const p = getBasePriorityValue(rowDiagnostics[idx]);
    if (p < bestPriority) {
      bestPriority = p;
      bestIndex = idx;
    }
  }

  return bestIndex;
}

/**
 * Finds the nearest base row below startRowIndex and selects the best
 * from consecutive base rows at that position.
 */
function findBestBaseRowIndex(rowDiagnostics, startRowIndex, options) {
  const firstBase = findNextBaseRowIndex(rowDiagnostics, startRowIndex);
  if (firstBase === null) return null;
  return selectBestFromConsecutiveBases(rowDiagnostics, firstBase, options);
}

/**
 * Returns true when a row is a generated RIT row (e.g. a significance-settings
 * footnote) rather than real table content. Such rows act as hard table
 * boundaries and must never be crossed by the upward Base fallback.
 */
function isGeneratedBoundaryRow(rowDiagnostic) {
  return isGeneratedSignificanceFootnoteRow(rowDiagnostic?.rawLabel);
}

/**
 * Silent upward fallback: finds a usable Base row ABOVE a metric block when no
 * Base was found below it.
 *
 * PURPOSE:
 * Some valid layouts place the Base row above the metric rows, e.g.
 *   Base
 *   Agree / Disagree        (proportions)
 * or
 *   Base
 *   Mean / SD               (mean + spread)
 * The primary below-block detection misses these, so the block would be skipped.
 *
 * BOUNDARIES (the search stops without a result when it would cross any of):
 * - a blank separator row;
 * - a generated RIT row (significance footnote / backlink);
 * - any other value/metric row — these belong to the current or a previous
 *   block, so we never tunnel past them looking for a distant Base;
 * - a Base row already consumed by a previous below-block (would steal a Base
 *   from a previous table).
 *
 * Only a Base row sitting directly above the block (optionally as part of a
 * consecutive run of Base rows) is accepted. When a run is found, the same
 * priority rules as below-detection pick the best Base in the run.
 *
 * @param {Array} rowDiagnostics
 * @param {number} blockTopRowIndex - index of the block's first (topmost) row
 * @param {Set<number>} consumedBaseRows - base indexes already used by a below-block
 * @param {object} options - { preferredBase }
 * @returns {number|null} best above-Base index, or null when none is usable
 */
function findBaseAboveBlockFallback(rowDiagnostics, blockTopRowIndex, consumedBaseRows, options) {
  let scanIndex = blockTopRowIndex - 1;

  while (scanIndex >= 0) {
    const rowDiagnostic = rowDiagnostics[scanIndex];

    // Generated rows and blank separators are hard table boundaries.
    if (isGeneratedBoundaryRow(rowDiagnostic) || rowDiagnostic.rowType === "empty") {
      return null;
    }

    if (rowDiagnostic.rowType === "base") {
      // A Base already used by a previous block belongs to that table.
      if (consumedBaseRows.has(scanIndex)) {
        return null;
      }

      // Walk up to the top of this consecutive run of (unconsumed) Base rows so
      // the existing priority rules can choose the best one in the run.
      let runTopIndex = scanIndex;
      while (
        runTopIndex - 1 >= 0 &&
        rowDiagnostics[runTopIndex - 1].rowType === "base" &&
        !consumedBaseRows.has(runTopIndex - 1)
      ) {
        runTopIndex--;
      }

      return selectBestFromConsecutiveBases(rowDiagnostics, runTopIndex, options);
    }

    // Any other row type is a value/metric row; do not cross it.
    return null;
  }

  return null;
}

/**
 * Builds calculation blocks from detected row labels.
 *
 * PURPOSE:
 * Support complex tables where proportions, means, and NPS can appear
 * in one selected range in different combinations.
 */
export function buildCalculationBlocks(detectionResult, { preferredBase = "auto" } = {}) {
  const rowDiagnostics = detectionResult.rowDiagnostics; // Classified rows.
  const calculationBlocks = []; // Final list of calculation blocks.
  const pendingProportionRows = []; // Proportion rows waiting for the next available base.
  const baseOptions = { preferredBase };
  const consumedBaseRows = new Set(); // Base rows already claimed by a below-block.

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
        const bestBaseIndex = selectBestFromConsecutiveBases(rowDiagnostics, rowIndex, baseOptions);
        consumedBaseRows.add(bestBaseIndex);
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
      const baseRowIndex = findBestBaseRowIndex(rowDiagnostics, spreadRowIndex, baseOptions);

      if (baseRowIndex !== null) {
        consumedBaseRows.add(baseRowIndex);
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

      // Silent fallback: no Base below the mean block — try a Base directly above
      // it within the same table context (issue #310). Statistical handling is
      // unchanged; only the Base row source differs.
      const aboveBaseRowIndex = findBaseAboveBlockFallback(
        rowDiagnostics,
        rowIndex,
        consumedBaseRows,
        baseOptions
      );

      if (aboveBaseRowIndex !== null) {
        consumedBaseRows.add(aboveBaseRowIndex);
        calculationBlocks.push(
          attachBaseSubtype(
            { metricType: "mean", valueRowIndex: rowIndex, spreadRowIndex, spreadType: rowDiagnostics[spreadRowIndex].rowType, baseRowIndex: aboveBaseRowIndex },
            rowDiagnostics, aboveBaseRowIndex
          )
        );

        rowIndex = spreadRowIndex + 1;
        continue;
      }
    }


    // 4. NPS-first (format 1 and 2, unified handler)
    const npsFirstBlock = tryParseNpsFirstBlock(rowDiagnostics, rowIndex);
    if (npsFirstBlock !== null) {
      const baseRowIndex = selectBestFromConsecutiveBases(rowDiagnostics, npsFirstBlock.baseRowIndex, baseOptions);
      consumedBaseRows.add(baseRowIndex);

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
      const baseRowIndex = findBestBaseRowIndex(rowDiagnostics, spreadRowIndex, baseOptions);

      if (baseRowIndex !== null) {
        consumedBaseRows.add(baseRowIndex);
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
        const baseRowIndex = findBestBaseRowIndex(rowDiagnostics, rowIndex, baseOptions);

        if (baseRowIndex !== null) {
          consumedBaseRows.add(baseRowIndex);
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

  // Silent fallback for a proportion block whose Base sits ABOVE its rows
  // (issue #310). Proportions are normally closed by a Base row below them; if
  // any remain buffered at the end, no Base was found below. Try a Base directly
  // above the block within the same table context before skipping it.
  if (pendingProportionRows.length > 0) {
    // Blank and generated rows are buffered with proportions, so anchor the
    // upward search on the first real value row. A blank or generated row
    // between that row and the Base above is then correctly treated as a table
    // boundary.
    const blockTopRowIndex = pendingProportionRows.find(
      (idx) =>
        rowDiagnostics[idx].rowType !== "empty" && !isGeneratedBoundaryRow(rowDiagnostics[idx])
    );
    const aboveBaseRowIndex =
      blockTopRowIndex === undefined
        ? null
        : findBaseAboveBlockFallback(rowDiagnostics, blockTopRowIndex, consumedBaseRows, baseOptions);

    if (aboveBaseRowIndex !== null) {
      consumedBaseRows.add(aboveBaseRowIndex);
      calculationBlocks.push(
        attachBaseSubtype(
          { metricType: "proportion", valueRowIndexes: [...pendingProportionRows], baseRowIndex: aboveBaseRowIndex },
          rowDiagnostics, aboveBaseRowIndex
        )
      );
      pendingProportionRows.length = 0;
    }
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
