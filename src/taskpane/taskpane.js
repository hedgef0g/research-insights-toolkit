/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

import {
  compareAllRowsUsingBottomBases,
  buildSignificanceMarkerMatrix,
} from "../core/significance";

/**
 * Reads selected Excel range, calculates pairwise significance,
 * and writes significance letters directly into value cells.
 *
 * MVP v0.3:
 * - Last selected row is treated as bases.
 * - All rows above are treated as values.
 * - Significant higher values receive labels of lower columns.
 */
async function runSignificanceFromSelection() {
  await Excel.run(async (context) => {
    const selectedRange = context.workbook.getSelectedRange(); // Current selected Excel range.

    selectedRange.load(["values", "text", "rowCount", "columnCount"]); // Load raw and displayed cell content.

    // Center all cells in the selected range after macro execution.
    selectedRange.format.horizontalAlignment = "Center";
    selectedRange.format.verticalAlignment = "Center";

    await context.sync();

    const selectedValues = selectedRange.values; // Raw values used for calculations.
    const selectedText = selectedRange.text; // Displayed values used for visible output.

    const outputElement = document.getElementById("significance-result"); // Result block in task pane.

    if (
      !selectedValues ||
      selectedValues.length < 2 ||
      selectedValues[0].length < 2
    ) {
      outputElement.textContent =
        "Please select at least 2 columns and 2 rows. Last row must contain bases.";
      return;
    }

    const allResults = compareAllRowsUsingBottomBases(selectedValues);

    if (allResults === null) {
      outputElement.textContent = "Could not process selected range.";
      return;
    }

    const markerMatrix = buildSignificanceMarkerMatrix(allResults);

    const valueRowCount = allResults.baseRowIndex; // Number of rows above base row.
    const columnCount = selectedValues[0].length; // Number of selected columns.

    for (let rowIndex = 0; rowIndex < valueRowCount; rowIndex++) {
      for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
        const markers = markerMatrix[rowIndex][columnIndex]; // Letters to add to current cell.

        if (!markers) {
          continue;
        }

        const currentCell = selectedRange.getCell(rowIndex, columnIndex); // Cell that receives significance letters.
        const displayedValue = selectedText[rowIndex][columnIndex]; // Current visible cell text.

        currentCell.values = [[`${displayedValue}${markers}`]];

        // Format cells where the value is significantly higher than at least one other column.
        currentCell.format.font.bold = true; // Make the whole cell text bold.
        currentCell.format.fill.color = "#E2F0D9"; // Pale green fill.
      }
    }

    await context.sync();

    outputElement.textContent = "Significance markers added to selected cells.";
  });
}

/**
 * Initializes task pane events after Office is ready.
 *
 * PURPOSE:
 * Connect the visible button in the Excel panel with our calculation logic.
 */
Office.onReady(() => {
  const calculateButton = document.getElementById("calculate-significance"); // Button in taskpane.html.

  calculateButton.addEventListener("click", runSignificanceFromSelection);
});

/**
 * Formats all pairwise comparison results into readable text for the task pane.
 *
 * PURPOSE:
 * Temporary MVP output.
 * Later we will replace this with table markers, colors, or letters.
 *
 * INPUT:
 * allResults - object returned by compareAllRowsUsingBottomBases().
 *
 * OUTPUT:
 * Multiline text for display in the Excel task pane.
 */
function formatAllComparisonsForDisplay(allResults) {
  const outputLines = []; // Final text lines for the task pane.

  outputLines.push("Pairwise significance results");
  outputLines.push(`Base row: ${allResults.baseRowIndex + 1}`);
  outputLines.push("");

  for (const comparisonRow of allResults.comparisonRows) {
    const displayedRowNumber = comparisonRow.valueRowIndex + 1; // Human-readable row number inside selection.

    outputLines.push(`Value row ${displayedRowNumber}:`);

    for (const comparison of comparisonRow.rowComparisons) {
      const firstColumnNumber = comparison.firstColumnIndex + 1; // Human-readable column number inside selection.
      const secondColumnNumber = comparison.secondColumnIndex + 1; // Human-readable column number inside selection.

      if (comparison.result === null) {
        outputLines.push(
          `  Col ${firstColumnNumber} vs Col ${secondColumnNumber}: skipped`
        );
        continue;
      }

      outputLines.push(
        `  Col ${firstColumnNumber} vs Col ${secondColumnNumber}: ` +
          `z=${comparison.result.zScore.toFixed(3)}, ` +
          `sig=${comparison.result.isSignificant ? "YES" : "NO"}, ` +
          `direction=${comparison.result.direction}`
      );
    }

    outputLines.push("");
  }

  return outputLines.join("\n");
}