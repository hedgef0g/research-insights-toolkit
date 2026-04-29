import { METRIC_DICTIONARY } from "./config/dictionary.config"; // Импортируем наш конфиг

export const LABEL_SCAN_COLUMNS_LEFT = 2;

export function normalizeLabelText(rawLabel) {
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
 * Checks whether normalized label contains any known keyword.
 */
function labelContainsAnyKeyword(normalizedLabel, keywords) {
  return keywords.some((keyword) => normalizedLabel.includes(keyword));
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
      return {
        rowType: dictionaryEntry.rowType,
        normalizedLabel,
      };
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

  /**
   * We scan from right to left because the closest label to the selected
   * data range is usually the most relevant one.
   */
  for (
    let labelColumnIndex = leftLabelRowValues.length - 1;
    labelColumnIndex >= 0;
    labelColumnIndex--
  ) {
    const currentLabelValue = leftLabelRowValues[labelColumnIndex]; // Candidate label cell value.
    const normalizedLabel = normalizeLabelText(currentLabelValue); // Normalized candidate.

    if (normalizedLabel && Number.isNaN(Number(normalizedLabel))) {
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

    rowDiagnostics.push({
      rowIndex,
      displayRowNumber: rowIndex + 1,
      rawLabel,
      normalizedLabel: classification.normalizedLabel,
      rowType: classification.rowType,
    });
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
  for (
    let rowIndex = startRowIndex + 1;
    rowIndex < rowDiagnostics.length;
    rowIndex++
  ) {
    if (rowDiagnostics[rowIndex].rowType === "base") {
      return rowIndex;
    }
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
        calculationBlocks.push({
          metricType: "proportion",
          valueRowIndexes: [...pendingProportionRows],
          baseRowIndex: rowIndex,
        });
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
      const baseRowIndex = findNextBaseRowIndex(rowDiagnostics, spreadRowIndex);

      if (baseRowIndex !== null) {
        calculationBlocks.push({
          metricType: "mean",
          valueRowIndex: rowIndex,
          spreadRowIndex,
          spreadType: rowDiagnostics[spreadRowIndex].rowType,
          baseRowIndex,
        });

        if (pendingProportionRows.length > 0) {
          calculationBlocks.push({
            metricType: "proportion",
            valueRowIndexes: [...pendingProportionRows],
            baseRowIndex,
          });
          pendingProportionRows.length = 0;
        }

        // ИСПРАВЛЕНИЕ: Прыгаем только через строки текущего блока (Среднее + Разброс), а не к Базе
        rowIndex = spreadRowIndex + 1;
        continue;
      }
    }

    // 4. Блок NPS Структура (NPS + Промоутеры + Детракторы)
    if (
      currentRowType === "nps" &&
      rowIndex + 2 < rowDiagnostics.length &&
      rowDiagnostics[rowIndex + 1].rowType === "promoters" &&
      rowDiagnostics[rowIndex + 2].rowType === "detractors"
    ) {
      const promotersRowIndex = rowIndex + 1;
      const detractorsRowIndex = rowIndex + 2;
      const baseRowIndex = findNextBaseRowIndex(rowDiagnostics, detractorsRowIndex);

      if (baseRowIndex !== null) {
        calculationBlocks.push({
          metricType: "npsStructure",
          valueRowIndex: rowIndex,
          promotersRowIndex,
          detractorsRowIndex,
          baseRowIndex,
        });

        if (pendingProportionRows.length > 0) {
          calculationBlocks.push({
            metricType: "proportion",
            valueRowIndexes: [...pendingProportionRows],
            baseRowIndex,
          });
          pendingProportionRows.length = 0;
        }

        // ИСПРАВЛЕНИЕ: Прыгаем только через строки текущего блока (NPS + Пром + Детр)
        rowIndex = detractorsRowIndex + 1;
        continue;
      }
    }

    // 5. Блок NPS Spread (NPS + Разброс)
    if (
      currentRowType === "nps" &&
      rowIndex + 1 < rowDiagnostics.length &&
      (rowDiagnostics[rowIndex + 1].rowType === "standardDeviation" ||
        rowDiagnostics[rowIndex + 1].rowType === "variance")
    ) {
      const spreadRowIndex = rowIndex + 1;
      const baseRowIndex = findNextBaseRowIndex(rowDiagnostics, spreadRowIndex);

      if (baseRowIndex !== null) {
        calculationBlocks.push({
          metricType: "npsSpread",
          valueRowIndex: rowIndex,
          spreadRowIndex,
          spreadType: rowDiagnostics[spreadRowIndex].rowType,
          baseRowIndex,
        });

        if (pendingProportionRows.length > 0) {
          calculationBlocks.push({
            metricType: "proportion",
            valueRowIndexes: [...pendingProportionRows],
            baseRowIndex,
          });
          pendingProportionRows.length = 0;
        }

        // ИСПРАВЛЕНИЕ: Прыгаем только через строки текущего блока
        rowIndex = spreadRowIndex + 1;
        continue;
      }
    }

    rowIndex++;
  }

  // Fallback
  if (calculationBlocks.length === 0 && rowDiagnostics.length >= 2) {
    calculationBlocks.push({
      metricType: "proportion",
      valueRowIndexes: Array.from(
        { length: rowDiagnostics.length - 1 },
        (_, index) => index
      ),
      baseRowIndex: rowDiagnostics.length - 1,
    });
  }

  return calculationBlocks;
}

/**
 * Checks whether row can be treated as a proportion value row.
 *
 * PURPOSE:
 * Prevent service rows like Promoters, Detractors, SD, Variance, and Base
 * from being calculated as ordinary proportions.
 */
function isProportionValueRowType(rowType) {
  return (
    rowType === "proportion" ||
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
 * - promoters / detractors rows
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