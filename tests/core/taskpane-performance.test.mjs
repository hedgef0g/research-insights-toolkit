import { describe, it, beforeEach } from "node:test";
import assert from "node:assert/strict";

function createLocalStorage(initial = {}) {
  const store = new Map(Object.entries(initial));
  return {
    getItem(key) {
      return store.has(key) ? store.get(key) : null;
    },
    setItem(key, value) {
      store.set(key, String(value));
    },
    removeItem(key) {
      store.delete(key);
    },
  };
}

async function importPerformanceModule() {
  return import(
    new URL(
      `../../src/taskpane/taskpane-performance.js?test=${Date.now()}-${Math.random()}`,
      import.meta.url
    )
  );
}

describe("taskpane performance flags", () => {
  beforeEach(() => {
    delete globalThis.__RIT_PERF;
    delete globalThis.__RIT_BANNER_VALUES_READ_PROBE;
    globalThis.localStorage = createLocalStorage();
  });

  it("uses runtime flags before localStorage for the main perf flag", async () => {
    globalThis.localStorage = createLocalStorage({ RIT_PERF: "1" });
    globalThis.__RIT_PERF = false;

    const { perfEnabled } = await importPerformanceModule();

    assert.equal(perfEnabled(), false);
  });

  it("reads optional probe flags from runtime and localStorage", async () => {
    const { perfFlagEnabled } = await importPerformanceModule();

    assert.equal(
      perfFlagEnabled("__RIT_BANNER_VALUES_READ_PROBE", "RIT_BANNER_VALUES_READ_PROBE"),
      false
    );

    globalThis.localStorage.setItem("RIT_BANNER_VALUES_READ_PROBE", "1");
    assert.equal(
      perfFlagEnabled("__RIT_BANNER_VALUES_READ_PROBE", "RIT_BANNER_VALUES_READ_PROBE"),
      true
    );

    globalThis.__RIT_BANNER_VALUES_READ_PROBE = "0";
    assert.equal(
      perfFlagEnabled("__RIT_BANNER_VALUES_READ_PROBE", "RIT_BANNER_VALUES_READ_PROBE"),
      false
    );
  });
});
