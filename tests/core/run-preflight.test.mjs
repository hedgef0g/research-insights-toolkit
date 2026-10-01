import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { preflightBatchMarkerOverflow } from "../../src/taskpane/run-preflight.js";

function createDecider(choice) {
  let calls = 0;
  return {
    get calls() {
      return calls;
    },
    async resolve() {
      calls += 1;
      return choice;
    },
  };
}

function createItemMap(columnCount) {
  return new Map([
    [
      "Sheet1!A1:ZZ10",
      {
        columnCount,
      },
    ],
  ]);
}

function createItemMapFromCounts(entries) {
  return new Map(
    entries.map(([sheetName, rangeAddress, columnCount]) => [
      `${sheetName}!${rangeAddress}`,
      { columnCount },
    ])
  );
}

const eligible = [{ sheetName: "Sheet1", rangeAddress: "A1:ZZ10" }];

describe("preflightBatchMarkerOverflow", () => {
  it("skips marker-capacity prompts in previous-column mode", async () => {
    const settings = { compareWithPreviousColumn: true };
    const decider = createDecider("stop");

    const stopped = await preflightBatchMarkerOverflow(
      eligible,
      createItemMap(200),
      settings,
      decider
    );

    assert.strictEqual(stopped, false);
    assert.strictEqual(decider.calls, 0);
    assert.strictEqual(settings.allowMultiCharacterMarkers, undefined);
  });

  it("does not prompt when the cheap column-count gate cannot overflow", async () => {
    const settings = {};
    const decider = createDecider("stop");

    const stopped = await preflightBatchMarkerOverflow(
      eligible,
      createItemMap(3),
      settings,
      decider
    );

    assert.strictEqual(stopped, false);
    assert.strictEqual(decider.calls, 0);
    assert.strictEqual(settings.allowMultiCharacterMarkers, undefined);
  });

  it("returns stopped when the user stops a non-banner overflow", async () => {
    const settings = {};
    const decider = createDecider("stop");

    const stopped = await preflightBatchMarkerOverflow(
      eligible,
      createItemMap(200),
      settings,
      decider
    );

    assert.strictEqual(stopped, true);
    assert.strictEqual(decider.calls, 1);
    assert.strictEqual(settings.allowMultiCharacterMarkers, undefined);
  });

  it("enables multi-character markers when the user continues a non-banner overflow", async () => {
    const settings = {};
    const decider = createDecider("continue");

    const stopped = await preflightBatchMarkerOverflow(
      eligible,
      createItemMap(200),
      settings,
      decider
    );

    assert.strictEqual(stopped, false);
    assert.strictEqual(decider.calls, 1);
    assert.strictEqual(settings.allowMultiCharacterMarkers, true);
  });

  it("prompts when any eligible non-banner table exceeds marker capacity", async () => {
    const settings = {};
    const decider = createDecider("continue");
    const mixedEligible = [
      { sheetName: "Sheet1", rangeAddress: "A1:C10" },
      { sheetName: "Sheet2", rangeAddress: "A1:ZZ10" },
    ];

    const stopped = await preflightBatchMarkerOverflow(
      mixedEligible,
      createItemMapFromCounts([
        ["Sheet1", "A1:C10", 3],
        ["Sheet2", "A1:ZZ10", 200],
      ]),
      settings,
      decider
    );

    assert.strictEqual(stopped, false);
    assert.strictEqual(decider.calls, 1);
    assert.strictEqual(settings.allowMultiCharacterMarkers, true);
  });

  it("ignores missing inventory metadata when the known candidates cannot overflow", async () => {
    const settings = {};
    const decider = createDecider("stop");
    const mixedEligible = [
      { sheetName: "Sheet1", rangeAddress: "A1:C10" },
      { sheetName: "Missing", rangeAddress: "A1:ZZ10" },
    ];

    const stopped = await preflightBatchMarkerOverflow(
      mixedEligible,
      createItemMapFromCounts([["Sheet1", "A1:C10", 3]]),
      settings,
      decider
    );

    assert.strictEqual(stopped, false);
    assert.strictEqual(decider.calls, 0);
    assert.strictEqual(settings.allowMultiCharacterMarkers, undefined);
  });
});
