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

import { compareAllRowsUsingBottomBases } from "../core/significance";

/**
 * Reads selected Excel range and compares all columns pairwise for each value row.
 *
 * PURPOSE:
 * Excel-specific wrapper for MVP v0.2.
 *
 * EXPECTED SELECTION:
 * Row 1..N-1: values
 * Last row: bases
 *
 * Example:
 * 42   35   50
 * 30   33   29
 * 200  180  210
 */
async function runSignificanceFromSelection() {
  await Excel.run(async (context) => {
    const selectedRange = context.workbook.getSelectedRange(); // Current selected Excel range.

    selectedRange.load("values"); // Load cell values from Excel.

    await context.sync(); // Execute Excel API request.

    const selectedValues = selectedRange.values; // 2D array of selected cell values.
    const outputElement = document.getElementById("significance-result"); // Output block in task pane.

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

    outputElement.textContent = formatAllComparisonsForDisplay(allResults);
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