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
 * Counts the row-label columns to recolor for a label matrix the interpreter has
 * already accepted as adjacent to the data body.
 *
 * `leftLabelValues` is a rows × columns matrix whose columns are ordered
 * left-to-right, so the rightmost column is the one touching the data body. The
 * label area spans from the **leftmost column that has any content** through the
 * rightmost (data-adjacent) column inclusive.
 *
 * A blank data-adjacent (rightmost) column is therefore NOT treated as a gap:
 * in valid two-column label layouts the text can live only in the left cell or
 * in a cell merged across both columns, leaving the right cell visually empty.
 * Recolor should still cover the whole label area that calculation reads.
 *
 * Real blank helper gaps between the labels and the data body are filtered
 * upstream by the selected-range interpreter (leadingEmptyCols /
 * additionalLeadingEmptyCols / labelsOnLeftSide force the count to 0 before this
 * is ever called), so this function never has to re-detect them.
 *
 * For a two-column matrix (when the interpreter has not forced 0):
 *   - both columns have content                         → 2
 *   - only the right/data-adjacent column has content   → 1
 *   - only the left column has content (right blank)    → 2
 *   - both columns blank                                → 0
 *
 * @param {Array<Array<*>>} leftLabelValues
 * @returns {number}
 */
export function countAdjacentLabelColumns(leftLabelValues) {
  if (!Array.isArray(leftLabelValues) || leftLabelValues.length === 0) return 0;
  const width = Array.isArray(leftLabelValues[0]) ? leftLabelValues[0].length : 0;
  if (width === 0) return 0;

  // The leftmost column carrying any content marks the left edge of the label
  // area; the area always extends rightward to the data body. Leading blank
  // columns (further from the data) are excluded; a blank rightmost column is
  // kept (merged / left-stored two-column labels).
  for (let col = 0; col < width; col++) {
    let hasContent = false;
    for (let row = 0; row < leftLabelValues.length; row++) {
      const rowValues = leftLabelValues[row];
      if (rowValues && !isBlankCell(rowValues[col])) {
        hasContent = true;
        break;
      }
    }
    if (hasContent) {
      return width - col;
    }
  }

  return 0;
}

/**
 * Resolves the adjacent row-label column count — the shared source of truth used
 * by both Check (label-column report) and design recolor — from selected-range
 * interpretation data.
 *
 * This is the meaningful, recolorable adjacent label AREA width, which can be
 * wider than the normalizer's text-only `labelColCount`. In a normalized
 * selection EVERY column to the left of the data body is part of the row-label /
 * structure area (label text columns plus any merged or blank label-area columns
 * a two-column label header may leave empty), so the area width is exactly the
 * data-body column offset. This is what fixes Mean + SD/Variance + Base tables
 * whose labels are merged across two columns or stored only in the left column:
 * the blank data-adjacent label column is no longer dropped.
 *
 * Genuine gaps are still excluded:
 *   - `labelsOnLeftSide`            → 0 (far-left labels, recolor disabled);
 *   - pass-through `leadingEmptyColumns > 0` → 0 (external labels separated from
 *     the data body by a blank helper/spacer column — a real gap).
 *
 * @param {object} args
 * @param {"normalized"|"passThrough"|"blocked"} args.state
 * @param {boolean} args.labelsOnLeftSide
 * @param {number}  args.dataColumnOffset   - columns left of the data body inside
 *        the selection (normalizer label columns + embedded + stripped blanks).
 *        Used for the normalized state only.
 * @param {number}  args.leadingEmptyColumns - pass-through leading blank helper
 *        columns between an external label column and the data body.
 * @param {Array<Array<*>>} args.leftLabelValues - label matrix (rightmost column
 *        adjacent to the data body) for externally/embedded-loaded labels.
 * @returns {number}
 */
export function resolveAdjacentLabelColumnCount({
  state,
  labelsOnLeftSide,
  dataColumnOffset = 0,
  leadingEmptyColumns = 0,
  leftLabelValues,
} = {}) {
  if (labelsOnLeftSide) return 0;
  if (state === "blocked") return 0;

  if (state === "normalized") {
    // Every column left of the data body in a normalized table is label/structure
    // (label text + merged/blank label-area columns). Width = data column offset.
    if (dataColumnOffset > 0) return dataColumnOffset;
    // No in-selection label columns: labels were loaded externally, immediately
    // left of the data body — fall back to scanning the loaded matrix.
    return countAdjacentLabelColumns(leftLabelValues);
  }

  // pass-through: external/embedded labels immediately left of the data body.
  // A leading empty helper column means external labels are separated from the
  // data body by a real gap, so there is no adjacent label area to recolor.
  if (leadingEmptyColumns > 0) return 0;
  return countAdjacentLabelColumns(leftLabelValues);
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
