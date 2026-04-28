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

import { calculateProportionSignificance } from "../core/significance";

/**
 * Reads selected Excel range and runs significance calculation.
 *
 * PURPOSE:
 * This is the Excel-specific wrapper.
 * It knows how to read Excel selection and send values to the shared core.
 *
 * MVP EXPECTED SELECTION:
 * Row 1: value 1, value 2
 * Row 2: base 1,  base 2
 *
 * FUTURE EXTENSIONS:
 * - Detect bases above values.
 * - Support weighted base row.
 * - Ignore text cells intelligently.
 * - Support more than 2 comparison points.
 */
async function runSignificanceFromSelection() {
  await Excel.run(async (context) => {
    const selectedRange = context.workbook.getSelectedRange(); // Current selected Excel range.

    selectedRange.load("values"); // Ask Excel to load cell values from selected range.

    await context.sync(); // Execute pending Excel API commands.

    const selectedValues = selectedRange.values; // Two-dimensional array from Excel.

    const outputElement = document.getElementById("significance-result"); // Result block in the task pane.

    if (
      !selectedValues ||
      selectedValues.length < 2 ||
      selectedValues[0].length < 2
    ) {
      outputElement.textContent =
        "Please select at least 4 cells: two values above two bases.";
      return;
    }

    const firstValue = selectedValues[0][0]; // First value, top-left cell.
    const secondValue = selectedValues[0][1]; // Second value, top-right cell.

    const firstBase = selectedValues[1][0]; // First base, bottom-left cell.
    const secondBase = selectedValues[1][1]; // Second base, bottom-right cell.

    const significanceResult = calculateProportionSignificance(
      firstValue,
      firstBase,
      secondValue,
      secondBase
    );

    if (significanceResult === null) {
      outputElement.textContent =
        "Could not calculate significance. Check that values and bases are numeric.";
      return;
    }

    outputElement.textContent =
      `First value: ${significanceResult.firstProportion}\n` +
      `Second value: ${significanceResult.secondProportion}\n` +
      `First base: ${significanceResult.firstBase}\n` +
      `Second base: ${significanceResult.secondBase}\n` +
      `z-score: ${significanceResult.zScore.toFixed(3)}\n` +
      `Significant at 95%: ${
        significanceResult.isSignificant ? "YES" : "NO"
      }\n` +
      `Direction: ${significanceResult.direction}`;
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