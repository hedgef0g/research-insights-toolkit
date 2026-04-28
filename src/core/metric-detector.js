/**
 * Number of columns to scan to the left of the selected data range.
 *
 * PURPOSE:
 * Labels may be located immediately to the left of data,
 * or one extra column further left.
 */
export const LABEL_SCAN_COLUMNS_LEFT = 2;

/**
 * Normalizes a label string before classification.
 *
 * PURPOSE:
 * Spreadsheet labels may contain dots, extra spaces, different casing,
 * Russian ё/е variations, etc.
 */
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
 *
 * PURPOSE:
 * Central helper for row type classification.
 */
function labelContainsAnyKeyword(normalizedLabel, keywords) {
  return keywords.some((keyword) => normalizedLabel.includes(keyword));
}

/**
 * Classifies one normalized row label.
 *
 * PURPOSE:
 * Detect what kind of data row this is:
 * proportion, mean, SD, variance, NPS, promoters, detractors, base, or unknown.
 */
export function classifyMetricLabel(rawLabel) {
  const normalizedLabel = normalizeLabelText(rawLabel); // Clean label text.

  if (!normalizedLabel) {
    return {
      rowType: "empty",
      normalizedLabel,
    };
  }

  const proportionKeywords = [
    "%",
    "percent",
    "percentage",
    "share",
    "доля",
    "процент",
    "проценты",
  ];

  const meanKeywords = [
    "mean",
    "average",
    "avg",
    "среднее",
    "средняя",
    "срзнач",
    "ср знач",
    "ср значение",
    "среднее значение",
  ];

  const standardDeviationKeywords = [
    "sd",
    "std",
    "std dev",
    "standard deviation",
    "stdev",
    "st dev",
    "стандартное отклонение",
    "ст отклонение",
    "ст откл",
    "среднеквадратическое отклонение",
    "ско",
  ];

  const varianceKeywords = [
    "variance",
    "var",
    "dispersion",
    "дисперсия",
    "дисп",
  ];

  const npsKeywords = [
    "nps",
    "нпс",
    "net promoter score",
    "индекс лояльности",
  ];

  const promotersKeywords = [
    "promoters",
    "promoter",
    "промоутеры",
    "промоутер",
    "сторонники",
    "лояльные",
  ];

  const detractorsKeywords = [
    "detractors",
    "detractor",
    "детракторы",
    "детрактор",
    "критики",
    "недовольные",
  ];

  const baseKeywords = [
    "base",
    "база",
    "основание",
    "выборка",
    "количество",
    "кол во",
  ];

  if (labelContainsAnyKeyword(normalizedLabel, npsKeywords)) {
    return {
      rowType: "nps",
      normalizedLabel,
    };
  }

  if (labelContainsAnyKeyword(normalizedLabel, promotersKeywords)) {
    return {
      rowType: "promoters",
      normalizedLabel,
    };
  }

  if (labelContainsAnyKeyword(normalizedLabel, detractorsKeywords)) {
    return {
      rowType: "detractors",
      normalizedLabel,
    };
  }

  if (labelContainsAnyKeyword(normalizedLabel, standardDeviationKeywords)) {
    return {
      rowType: "standardDeviation",
      normalizedLabel,
    };
  }

  if (labelContainsAnyKeyword(normalizedLabel, varianceKeywords)) {
    return {
      rowType: "variance",
      normalizedLabel,
    };
  }

  if (labelContainsAnyKeyword(normalizedLabel, meanKeywords)) {
    return {
      rowType: "mean",
      normalizedLabel,
    };
  }

  if (labelContainsAnyKeyword(normalizedLabel, baseKeywords)) {
    return {
      rowType: "base",
      normalizedLabel,
    };
  }

  if (labelContainsAnyKeyword(normalizedLabel, proportionKeywords)) {
    return {
      rowType: "proportion",
      normalizedLabel,
    };
  }

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
 * Builds auto-calculation plan from detected row labels.
 *
 * PURPOSE:
 * First supported smart mode:
 * Mean + SD/Variance + Base
 *
 * If pattern is not found, fallback to proportions.
 */
export function buildAutoCalculationPlan(detectionResult) {
  const detectedRows = detectionResult.rowDiagnostics; // All classified rows.

  if (detectedRows.length >= 3) {
    const firstRowType = detectedRows[0].rowType;
    const secondRowType = detectedRows[1].rowType;
    const thirdRowType = detectedRows[2].rowType;

    const isMeanWithStandardDeviation =
      firstRowType === "mean" &&
      secondRowType === "standardDeviation" &&
      thirdRowType === "base";

    if (isMeanWithStandardDeviation) {
      return {
        metricType: "mean",
        spreadType: "standardDeviation",
        baseRowIndex: 2,
      };
    }

    const isMeanWithVariance =
      firstRowType === "mean" &&
      secondRowType === "variance" &&
      thirdRowType === "base";

    if (isMeanWithVariance) {
      return {
        metricType: "mean",
        spreadType: "variance",
        baseRowIndex: 2,
      };
    }
  }

  const isNpsWithPromotersDetractors =
  detectedRows.length >= 4 &&
  detectedRows[0].rowType === "nps" &&
  detectedRows[1].rowType === "promoters" &&
  detectedRows[2].rowType === "detractors" &&
  detectedRows[3].rowType === "base";

  if (isNpsWithPromotersDetractors) {
    return {
      metricType: "npsStructure",
      baseRowIndex: 3,
    };
  }

    const isNpsWithStandardDeviation =
    detectedRows.length >= 3 &&
    detectedRows[0].rowType === "nps" &&
    detectedRows[1].rowType === "standardDeviation" &&
    detectedRows[2].rowType === "base";

  if (isNpsWithStandardDeviation) {
    return {
      metricType: "npsSpread",
      spreadType: "standardDeviation",
      baseRowIndex: 2,
    };
  }

  const isNpsWithVariance =
    detectedRows.length >= 3 &&
    detectedRows[0].rowType === "nps" &&
    detectedRows[1].rowType === "variance" &&
    detectedRows[2].rowType === "base";

  if (isNpsWithVariance) {
    return {
      metricType: "npsSpread",
      spreadType: "variance",
      baseRowIndex: 2,
    };
  }

  return {
    metricType: "proportion",
  };
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
 *
 * KEY LOGIC:
 * Proportion rows may not have their own immediate Base row.
 * If a proportion block is followed by Mean or NPS, the next Base row may be shared.
 */
export function buildCalculationBlocks(detectionResult) {
  const rowDiagnostics = detectionResult.rowDiagnostics; // Classified rows.
  const calculationBlocks = []; // Final list of calculation blocks.
  const pendingProportionRows = []; // Proportion rows waiting for the next available base.

  let rowIndex = 0; // Current row scanner position.

  while (rowIndex < rowDiagnostics.length) {
    const currentRowType = rowDiagnostics[rowIndex].rowType; // Current detected row type.

    /**
     * Proportion-like row:
     * Store it for later. We will attach it to the next Base row.
     */
    if (isProportionValueRowType(currentRowType)) {
      pendingProportionRows.push(rowIndex);
      rowIndex++;
      continue;
    }

    /**
     * Base row:
     * If there are pending proportion rows, close them using this base.
     */
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

    /**
     * Mean block:
     * Mean
     * SD or Variance
     * Base
     */
    if (
      currentRowType === "mean" &&
      rowIndex + 1 < rowDiagnostics.length &&
      (rowDiagnostics[rowIndex + 1].rowType === "standardDeviation" ||
        rowDiagnostics[rowIndex + 1].rowType === "variance")
    ) {
      const spreadRowIndex = rowIndex + 1; // Row with SD or variance.
      const baseRowIndex = findNextBaseRowIndex(rowDiagnostics, spreadRowIndex); // Nearest base below spread row.

      if (baseRowIndex !== null) {
        calculationBlocks.push({
          metricType: "mean",
          valueRowIndex: rowIndex,
          spreadRowIndex,
          spreadType: rowDiagnostics[spreadRowIndex].rowType,
          baseRowIndex,
        });

        /**
         * If proportion rows appeared before this mean block and had no base yet,
         * use this same base for them.
         */
        if (pendingProportionRows.length > 0) {
          calculationBlocks.push({
            metricType: "proportion",
            valueRowIndexes: [...pendingProportionRows],
            baseRowIndex,
          });

          pendingProportionRows.length = 0;
        }

        rowIndex = baseRowIndex + 1;
        continue;
      }
    }

    /**
     * NPS structure block:
     * NPS
     * Promoters
     * Detractors
     * Base
     */
    if (
      currentRowType === "nps" &&
      rowIndex + 2 < rowDiagnostics.length &&
      rowDiagnostics[rowIndex + 1].rowType === "promoters" &&
      rowDiagnostics[rowIndex + 2].rowType === "detractors"
    ) {
      const promotersRowIndex = rowIndex + 1; // Promoter share row.
      const detractorsRowIndex = rowIndex + 2; // Detractor share row.
      const baseRowIndex = findNextBaseRowIndex(
        rowDiagnostics,
        detractorsRowIndex
      ); // Nearest base below detractors row.

      if (baseRowIndex !== null) {
        calculationBlocks.push({
          metricType: "npsStructure",
          valueRowIndex: rowIndex,
          promotersRowIndex,
          detractorsRowIndex,
          baseRowIndex,
        });

        /**
         * If proportion rows appeared before this NPS block and had no base yet,
         * use this same base for them.
         */
        if (pendingProportionRows.length > 0) {
          calculationBlocks.push({
            metricType: "proportion",
            valueRowIndexes: [...pendingProportionRows],
            baseRowIndex,
          });

          pendingProportionRows.length = 0;
        }

        rowIndex = baseRowIndex + 1;
        continue;
      }
    }

    /**
     * NPS spread block:
     * NPS
     * SD or Variance
     * Base
     */
    if (
      currentRowType === "nps" &&
      rowIndex + 1 < rowDiagnostics.length &&
      (rowDiagnostics[rowIndex + 1].rowType === "standardDeviation" ||
        rowDiagnostics[rowIndex + 1].rowType === "variance")
    ) {
      const spreadRowIndex = rowIndex + 1; // Row with NPS SD or variance.
      const baseRowIndex = findNextBaseRowIndex(rowDiagnostics, spreadRowIndex); // Nearest base below spread row.

      if (baseRowIndex !== null) {
        calculationBlocks.push({
          metricType: "npsSpread",
          valueRowIndex: rowIndex,
          spreadRowIndex,
          spreadType: rowDiagnostics[spreadRowIndex].rowType,
          baseRowIndex,
        });

        /**
         * If proportion rows appeared before this NPS spread block and had no base yet,
         * use this same base for them.
         */
        if (pendingProportionRows.length > 0) {
          calculationBlocks.push({
            metricType: "proportion",
            valueRowIndexes: [...pendingProportionRows],
            baseRowIndex,
          });

          pendingProportionRows.length = 0;
        }

        rowIndex = baseRowIndex + 1;
        continue;
      }
    }

    rowIndex++;
  }

  /**
   * Fallback:
   * If no labelled blocks were detected, use the old assumption:
   * all rows except the last one are proportions, last row is base.
   */
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