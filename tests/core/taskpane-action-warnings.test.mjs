import assert from "node:assert/strict";
import test, { describe } from "node:test";

import { collectActionWarningKeys } from "../../src/taskpane/taskpane-action-warnings.js";

describe("taskpane action warnings", () => {
  test("returns no warnings outside Run and Autorun actions", () => {
    assert.deepStrictEqual(
      collectActionWarningKeys({
        action: "check",
        labelsOnLeftSide: true,
        addTableFootnoteRequested: true,
        recolorRequested: true,
      }),
      []
    );
  });

  test("shows manual Run recolor warning only for manual Run", () => {
    assert.deepStrictEqual(
      collectActionWarningKeys({
        action: "run",
        recolorRequested: true,
      }),
      ["manual-run-recolor"]
    );

    assert.deepStrictEqual(
      collectActionWarningKeys({
        action: "autorun",
        recolorRequested: true,
      }),
      []
    );
  });

  test("prioritizes far-left label incompatibility over manual Run recolor warning", () => {
    assert.deepStrictEqual(
      collectActionWarningKeys({
        action: "run",
        labelsOnLeftSide: true,
        recolorRequested: true,
      }),
      ["recolor-labels-left"]
    );
  });

  test("includes far-left footnote and recolor incompatibility warnings together", () => {
    assert.deepStrictEqual(
      collectActionWarningKeys({
        action: "autorun",
        labelsOnLeftSide: true,
        addTableFootnoteRequested: true,
        recolorRequested: true,
      }),
      ["footnote-labels-left", "recolor-labels-left"]
    );
  });
});
