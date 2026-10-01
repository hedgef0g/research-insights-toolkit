import { describe, it } from "node:test";
import assert from "node:assert/strict";

import { createMarkerOverflowDecider } from "../../src/taskpane/taskpane-dialogs.js";

describe("createMarkerOverflowDecider", () => {
  it("caches a continue decision and reuses the in-flight prompt", async () => {
    let calls = 0;
    let resolvePrompt;
    const decider = createMarkerOverflowDecider(() => {
      calls += 1;
      return new Promise((resolve) => {
        resolvePrompt = resolve;
      });
    });

    assert.equal(decider.decision, null);

    const first = decider.resolve();
    const second = decider.resolve();

    assert.equal(calls, 1);

    resolvePrompt(true);

    assert.equal(await first, "continue");
    assert.equal(await second, "continue");
    assert.equal(decider.decision, "continue");
    assert.equal(await decider.resolve(), "continue");
    assert.equal(calls, 1);
  });

  it("caches a stop decision", async () => {
    let calls = 0;
    const decider = createMarkerOverflowDecider(async () => {
      calls += 1;
      return false;
    });

    assert.equal(await decider.resolve(), "stop");
    assert.equal(decider.decision, "stop");
    assert.equal(await decider.resolve(), "stop");
    assert.equal(calls, 1);
  });
});
