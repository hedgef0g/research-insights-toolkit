import { removeSignificanceMarkersFromText } from "./significance";

/**
 * Writes significance markers into selected range.
 *
 * PURPOSE:
 * Shared writer for proportions, means, and NPS blocks.
 * It also protects cells from Excel auto-formatting values as time.
 */
export function writeMarkersToSelectedRange(
  selectedRange,
  selectedText,
  markerMatrix
) {
  const rowCount = markerMatrix.length; // Number of rows in marker matrix.
  const columnCount = markerMatrix[0] ? markerMatrix[0].length : 0; // Number of columns.

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      const markers = markerMatrix[rowIndex][columnIndex]; // Marker letters for current cell.

      if (!markers) {
        continue;
      }

      const currentCell = selectedRange.getCell(rowIndex, columnIndex); // Target cell.
      const displayedValueWithoutMarkers = removeSignificanceMarkersFromText(
        selectedText[rowIndex][columnIndex]
      );

      currentCell.numberFormat = [["@"]]; // Force text format to prevent Excel time conversion.

      currentCell.values = [
        [`${displayedValueWithoutMarkers} ${markers}`.trim()],
      ];

      currentCell.format.font.bold = true;
      currentCell.format.fill.color = "#E2F0D9";
    }
  }
}