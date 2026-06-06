import {
  collectStatisticTypeLabels,
  buildSignificanceFootnoteCellValue,
  resolveFootnoteSpan,
} from "../core/significance-footnote";

/**
 * Builds a footnote job for a single processed table, or null when the footnote
 * setting is off / geometry is unusable.
 *
 * Geometry is sheet-absolute and 0-based:
 *   - dataStartRowIndex / dataStartColIndex: top-left of the data body.
 *   - dataRowCount / dataColCount: data body dimensions.
 *   - leftLabelValues: label columns immediately left of the data body. Their
 *     width sets how far left the merged footnote extends (the "full table width
 *     including label columns").
 *
 * Returns null when the footnote setting is off. The feature is explicitly
 * incompatible with "labels on left side" mode (labels are not adjacent to the
 * data, so a footnote cannot span the table correctly): readCalculationSettings-
 * FromPanel already forces addTableFootnote=false in that mode, and this extra
 * guard makes the rule independent of the caller.
 *
 * processedScopeSuffix is an optional, already-formatted detail (e.g.
 * " Обработано: B12:F34.") appended to the visible footnote text. It is
 * supplied only by Manual Run; auto-run callers omit it so their footnote text
 * is unchanged.
 */
export function buildSignificanceFootnoteJob({
  sheetName,
  dataStartRowIndex,
  dataStartColIndex,
  dataRowCount,
  dataColCount,
  leftLabelValues,
  adjacentLabelColumnCount,
  calculationBlocks,
  calculationSettings,
  processedScopeSuffix,
}) {
  if (!calculationSettings.addTableFootnote) return null;
  if (calculationSettings.labelsOnLeftSide) return null;
  if (!Number.isFinite(dataStartRowIndex) || !Number.isFinite(dataStartColIndex)) return null;
  if (!(dataRowCount > 0) || !(dataColCount > 0)) return null;

  const labelColumns =
    Array.isArray(leftLabelValues) && Array.isArray(leftLabelValues[0]) ? leftLabelValues[0].length : 0;

  const { tableLeftColIndex, tableRightColIndex } = resolveFootnoteSpan({
    dataStartColIndex,
    dataColCount,
    labelColumns,
    adjacentLabelColumnCount,
  });
  const tableBottomRowIndex = dataStartRowIndex + dataRowCount - 1;

  const footnoteCellValue = buildSignificanceFootnoteCellValue({
    confidenceLevel: calculationSettings.confidenceLevel,
    oneTailedTest: calculationSettings.oneTailedTest,
    statisticLabels: collectStatisticTypeLabels(calculationBlocks),
    scopeDetail: processedScopeSuffix,
  });

  return {
    sheetName,
    tableBottomRowIndex,
    tableLeftColIndex,
    tableRightColIndex,
    dataStartColIndex,
    footnoteCellValue,
  };
}
