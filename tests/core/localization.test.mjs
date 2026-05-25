import { describe, it } from "node:test";
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
  };
}

function createLeaf(i18nKey) {
  return {
    dataset: { i18n: i18nKey },
    children: [],
    textContent: "",
  };
}

function createHint(hintI18nKey) {
  return {
    dataset: { hintI18n: hintI18nKey, hintBase: "" },
  };
}

function createButton(lang, isActive = false) {
  const classes = new Set(isActive ? ["lang-btn", "is-active"] : ["lang-btn"]);
  return {
    dataset: { lang },
    classList: {
      toggle(name, force) {
        if (force) classes.add(name);
        else classes.delete(name);
      },
      contains(name) {
        return classes.has(name);
      },
    },
  };
}

function createDocument(elementsBySelector) {
  return {
    querySelectorAll(selector) {
      return elementsBySelector[selector] || [];
    },
  };
}

async function importLocalizationModule() {
  return import(new URL(`../../src/taskpane/localization.js?test=${Date.now()}-${Math.random()}`, import.meta.url));
}

describe("taskpane localization", () => {
  it("switches to English, updates translated content, and toggles the active button state", async () => {
    const ruButton = createButton("ru", true);
    const enButton = createButton("en");
    const label = createLeaf("tabs.run");
    const hint = createHint("hint.runCurrentTable");

    globalThis.localStorage = createLocalStorage();
    globalThis.document = createDocument({
      "[data-i18n]": [label],
      "[data-hint-i18n]": [hint],
      ".lang-btn[data-lang]": [ruButton, enButton],
    });

    const { getLanguage, setLanguage } = await importLocalizationModule();

    setLanguage("en");

    assert.strictEqual(getLanguage(), "en");
    assert.strictEqual(globalThis.localStorage.getItem("rit.ui.lang"), "en");
    assert.strictEqual(label.textContent, "Run");
    assert.strictEqual(hint.dataset.hintBase, "Scope: selected range");
    assert.equal(ruButton.classList.contains("is-active"), false);
    assert.equal(enButton.classList.contains("is-active"), true);
  });

  it("loads the saved language on a fresh module import and reapplies the active state", async () => {
    const ruButton = createButton("ru", true);
    const enButton = createButton("en");
    const label = createLeaf("tabs.content");

    globalThis.localStorage = createLocalStorage({ "rit.ui.lang": "en" });
    globalThis.document = createDocument({
      "[data-i18n]": [label],
      "[data-hint-i18n]": [],
      ".lang-btn[data-lang]": [ruButton, enButton],
    });

    const { applyI18n, getLanguage, loadSavedLanguage } = await importLocalizationModule();

    loadSavedLanguage();
    applyI18n();

    assert.strictEqual(getLanguage(), "en");
    assert.strictEqual(label.textContent, "Contents");
    assert.equal(ruButton.classList.contains("is-active"), false);
    assert.equal(enButton.classList.contains("is-active"), true);
  });
});
