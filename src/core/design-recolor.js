/**
 * Optional design recolor helpers for Research Insights Toolkit (issue #306).
 *
 * Pure, Office.js-free geometry helpers shared by the taskpane run flows. The
 * taskpane builds a "design recolor job" for each successfully processed table
 * and applies the resulting rectangles with batched, range-level fill writes.
 *
 * The feature recolors two regions with a single shared color:
 *   1. the detected banner/header rows directly above the data body, spanning
 *      the data columns;
 *   2. the adjacent row-label columns immediately left of the data body,
 *      spanning the data-body rows.
 *
 * The banner-row × label-column corner is intentionally left untouched so each
 * region is a clean rectangle (banner = data columns, labels = data rows). This
 * matches the issue specification and keeps every write rectangular.
 *
 * Label width is NOT re-guessed here from raw geometry: the caller passes the
 * adjacent label column count derived from the selected-range interpretation
 * (leftLabelValues / normalized label columns), so a blank gap column between
 * the labels and the data body never gets recolored.
 */

function isBlankCell(value) {
  return value === null || value === undefined || String(value).trim() === "";
}

/**
 * Counts the contiguous non-blank label columns immediately adjacent to the
 * data body, scanning from the rightmost (data-adjacent) column leftward and
 * stopping at the first fully-blank column.
 *
 * `leftLabelValues` is a rows × columns matrix whose columns are ordered
 * left-to-right, so the rightmost column is the one touching the data body.
 *
 *   - 1 real label column (further-left column blank) → 1
 *   - 2 real label columns                            → 2
 *   - no label columns / empty input                  → 0
 *   - a blank gap column adjacent to the data body    → 0 (gap not crossed)
 *
 * @param {Array<Array<*>>} leftLabelValues
 * @returns {number}
 */
export function countAdjacentLabelColumns(leftLabelValues) {
  if (!Array.isArray(leftLabelValues) || leftLabelValues.length === 0) return 0;
  const width = Array.isArray(leftLabelValues[0]) ? leftLabelValues[0].length : 0;
  if (width === 0) return 0;

  let count = 0;
  for (let col = width - 1; col >= 0; col--) {
    let hasContent = false;
    for (let row = 0; row < leftLabelValues.length; row++) {
      const rowValues = leftLabelValues[row];
      if (rowValues && !isBlankCell(rowValues[col])) {
        hasContent = true;
        break;
      }
    }
    if (hasContent) {
      count++;
    } else {
      break;
    }
  }
  return count;
}

/**
 * Builds a design recolor job for a single processed table, or null when the
 * feature is off / geometry is unusable.
 *
 * Geometry is sheet-absolute and 0-based:
 *   - dataStartRowIndex / dataStartColIndex: top-left of the data body.
 *   - dataRowCount / dataColCount:           data body dimensions.
 *   - bannerRowCount:                        banner/header rows directly above
 *                                            the data body (0 when none).
 *   - adjacentLabelColumnCount:              row-label columns immediately left
 *                                            of the data body (0 when none).
 *
 * Returns null when the feature setting is off. The feature is explicitly
 * incompatible with "labels on left side" mode (labels are far-left and not
 * adjacent to the data, so the recolor span could cross unrelated columns):
 * readCalculationSettingsFromPanel already forces recolorBannerAndLabels=false
 * in that mode, and this extra guard makes the rule independent of the caller.
 *
 * The returned job contains rectangles (sheet-absolute, 0-based) so the caller
 * applies one fill per rectangle — never per cell.
 *
 * @returns {{ sheetName: string, color: string, rects: Array<{rowIndex:number,columnIndex:number,rowCount:number,columnCount:number}> } | null}
 */
export function buildDesignRecolorJob({
  sheetName,
  dataStartRowIndex,
  dataStartColIndex,
  dataRowCount,
  dataColCount,
  adjacentLabelColumnCount,
  bannerRowCount,
  calculationSettings,
}) {
  if (!calculationSettings || !calculationSettings.recolorBannerAndLabels) return null;
  if (calculationSettings.labelsOnLeftSide) return null;
  if (!Number.isFinite(dataStartRowIndex) || !Number.isFinite(dataStartColIndex)) return null;
  if (!(dataRowCount > 0) || !(dataColCount > 0)) return null;

  const color = calculationSettings.bannerLabelFillColor;
  if (typeof color !== "string" || color.trim() === "") return null;

  const rects = [];

  // Banner: rows directly above the data body, spanning the data columns only.
  const safeBannerRows =
    Number.isFinite(bannerRowCount) && bannerRowCount > 0 ? bannerRowCount : 0;
  const bannerTopRowIndex = dataStartRowIndex - safeBannerRows;
  if (safeBannerRows > 0 && bannerTopRowIndex >= 0) {
    rects.push({
      rowIndex: bannerTopRowIndex,
      columnIndex: dataStartColIndex,
      rowCount: safeBannerRows,
      columnCount: dataColCount,
    });
  }

  // Row labels: columns immediately left of the data body, spanning data rows.
  // Capped at dataStartColIndex so the span never runs off the left edge.
  const labelColCount = Math.min(
    Number.isFinite(adjacentLabelColumnCount) && adjacentLabelColumnCount > 0
      ? adjacentLabelColumnCount
      : 0,
    dataStartColIndex
  );
  if (labelColCount > 0) {
    rects.push({
      rowIndex: dataStartRowIndex,
      columnIndex: dataStartColIndex - labelColCount,
      rowCount: dataRowCount,
      columnCount: labelColCount,
    });
  }

  if (rects.length === 0) return null;

  return { sheetName, color, rects };
}
