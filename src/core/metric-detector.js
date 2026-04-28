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

  return {
    metricType: "proportion",
  };
}