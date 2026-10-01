import { describe, it } from "node:test";
import assert from "node:assert/strict";

import { shouldBlockResolverScan } from "../../src/taskpane/active-cell-resolver.js";
import { INVENTORY_SCAN_EMERGENCY_CELL_LIMIT } from "../../src/taskpane/taskpane-inventory-scan.js";

describe("active-cell-resolver", () => {
  it("does not block current-table resolution at the old 250k sheet-size guard", () => {
    const rowCount = 522;
    const columnCount = 516;
    const cellCount = rowCount * columnCount;

    assert.equal(cellCount, 269352);
    assert.equal(cellCount > 250000, true);
    assert.equal(cellCount < INVENTORY_SCAN_EMERGENCY_CELL_LIMIT, true);
    assert.equal(shouldBlockResolverScan(rowCount, columnCount), false);
  });

  it("still blocks truly pathological used ranges using the shared inventory emergency limit", () => {
    assert.equal(
      shouldBlockResolverScan(1, INVENTORY_SCAN_EMERGENCY_CELL_LIMIT),
      false
    );
    assert.equal(
      shouldBlockResolverScan(1, INVENTORY_SCAN_EMERGENCY_CELL_LIMIT + 1),
      true
    );
  });
});
