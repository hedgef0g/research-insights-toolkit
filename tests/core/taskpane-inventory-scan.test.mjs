import { describe, it } from "node:test";
import assert from "node:assert/strict";

import {
  INVENTORY_SCAN_EMERGENCY_CELL_LIMIT,
  inventoryScanErrorMessage,
  formatInventorySkippedSheetLine,
  buildInventoryContentSkippedRow,
} from "../../src/taskpane/taskpane-inventory-scan.js";

describe("taskpane-inventory-scan", () => {
  it("formats the emergency limit workbook message as a circuit breaker", () => {
    const line = formatInventorySkippedSheetLine({
      sheetName: "Wide sheet",
      reason: "emergencyLimit",
      rowCount: 2000,
      columnCount: 3000,
      cellCount: 6000000,
    });

    assert.match(line, /Wide sheet: аварийно пропущен/);
    assert.match(line, new RegExp(String(INVENTORY_SCAN_EMERGENCY_CELL_LIMIT)));
    assert.doesNotMatch(line, /слишком большой для сканирования/);
  });

  it("formats scan errors without falling back to the old too-large wording", () => {
    const line = formatInventorySkippedSheetLine({
      sheetName: "Polluted",
      reason: "scanError",
      message: "RequestPayloadSizeLimitExceeded",
    });

    assert.strictEqual(
      line,
      "Polluted: ошибка сканирования — RequestPayloadSizeLimitExceeded."
    );
  });

  it("builds content rows for scan errors with dimensions and message", () => {
    const row = buildInventoryContentSkippedRow({
      sheetName: "Sheet A",
      reason: "scanError",
      rowCount: 120,
      columnCount: 48,
      message: "host rejected values load",
    });

    assert.deepStrictEqual(row, [
      "Sheet A",
      "Skipped",
      "Ошибка сканирования листа",
      "120 строк, 48 колонок",
      "host rejected values load",
    ]);
  });

  it("normalizes Office errors into single-line messages", () => {
    const message = inventoryScanErrorMessage({
      message: "First line\nSecond line",
    });

    assert.strictEqual(message, "First line Second line");
  });
});
