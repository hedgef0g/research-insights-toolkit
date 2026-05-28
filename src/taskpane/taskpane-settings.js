/* global console, document */

const LOCAL_SETTINGS_STORAGE_KEY = "rit.settings.v1";
const SETTINGS_COLLAPSED_KEY = "rit.ui.settingsCollapsed";

export const SETTINGS_CONTROL_CONFIG = [
  { id: "confidence-level", type: "value", settingName: "confidenceLevel" },

  {
    id: "one-tailed-test",
    type: "checked",
    settingName: "oneTailedTest",
  },

  { id: "round-cell-values", type: "checked", settingName: "roundCellValues" },
  { id: "result-formatting-level", type: "value", settingName: "resultFormattingLevel" },

  {
    id: "compare-with-previous-column",
    type: "checked",
    settingName: "compareWithPreviousColumn",
  },
  {
    id: "apply-previous-column-fill",
    type: "checked",
    settingName: "applyPreviousColumnFill",
  },

  { id: "write-banner-letters", type: "checked", settingName: "writeBannerLetters" },
  { id: "respect-banner-structure", type: "checked", settingName: "respectBannerStructure" },
  {
    id: "auto-detect-wave-banners",
    type: "checked",
    settingName: "autoDetectWaveBanners",
  },
  { id: "labels-on-left-side", type: "checked", settingName: "labelsOnLeftSide" },

  { id: "compare-only-with-total", type: "checked", settingName: "compareOnlyWithTotal" },
  {
    id: "exclude-total-from-comparisons",
    type: "checked",
    settingName: "excludeTotalFromComparisons",
  },

  { id: "first-column-is-total", type: "checked", settingName: "firstColumnIsTotal" },
  { id: "total-in-each-banner", type: "checked", settingName: "totalInEachBanner" },

  { id: "significant-fill-color", type: "value", settingName: "significantFillColor" },
  {
    id: "lower-than-total-fill-color",
    type: "value",
    settingName: "lowerThanTotalFillColor",
  },

  {
    id: "fill-only-total-comparisons",
    type: "checked",
    settingName: "fillOnlyTotalComparisons",
  },

  {
    id: "exclude-small-bases",
    type: "checked",
    settingName: "excludeSmallBasesFromComparisons",
  },
  { id: "small-base-threshold", type: "number", settingName: "smallBaseThreshold" },
  { id: "small-base-fill-color", type: "value", settingName: "smallBaseFillColor" },

  { id: "preferred-base", type: "value", settingName: "preferredBase" },

  { id: "settings-storage-mode", type: "value", settingName: "settingsStorageMode" },
];

export const DEFAULT_CALCULATION_SETTINGS = {
  confidenceLevel: "95",
  oneTailedTest: false,

  roundCellValues: false,
  resultFormattingLevel: "full",

  compareWithPreviousColumn: false,
  applyPreviousColumnFill: false,

  writeBannerLetters: false,
  respectBannerStructure: false,
  autoDetectWaveBanners: false,
  labelsOnLeftSide: false,

  compareOnlyWithTotal: false,
  excludeTotalFromComparisons: false,

  firstColumnIsTotal: false,
  totalInEachBanner: false,

  significantFillColor: "#E2F0D9",
  lowerThanTotalFillColor: "#FCE4D6",

  fillOnlyTotalComparisons: false,

  excludeSmallBasesFromComparisons: false,
  smallBaseThreshold: 50,
  smallBaseFillColor: "#D0D0D0",

  preferredBase: "auto",

  settingsStorageMode: "none",
};

function applySettingsCollapseState(toggleBtn, panelBody, isCollapsed) {
  toggleBtn.setAttribute("aria-expanded", isCollapsed ? "false" : "true");
  panelBody.classList.toggle("is-collapsed", isCollapsed);
}

export function initializeSettingsCollapse() {
  const toggleBtn = document.getElementById("settings-toggle");
  const panelBody = document.getElementById("settings-panel-body");

  if (!toggleBtn || !panelBody) return;

  let isCollapsed = false;
  try {
    const saved = localStorage.getItem(SETTINGS_COLLAPSED_KEY);
    if (saved !== null) isCollapsed = saved === "true";
  } catch (_) {
    /* non-fatal */
  }

  applySettingsCollapseState(toggleBtn, panelBody, isCollapsed);

  toggleBtn.addEventListener("click", () => {
    const nextCollapsed = toggleBtn.getAttribute("aria-expanded") === "true";
    applySettingsCollapseState(toggleBtn, panelBody, nextCollapsed);
    try {
      localStorage.setItem(SETTINGS_COLLAPSED_KEY, String(nextCollapsed));
    } catch (_) {
      /* non-fatal */
    }
  });
}

/**
 * Loads saved settings into the panel if local saving was enabled.
 */
export function loadSavedSettingsIntoPanel() {
  const savedSettings = readSettingsFromLocalStorage();

  if (!savedSettings) {
    return;
  }

  if (savedSettings.settingsStorageMode !== "local") {
    return;
  }

  applySettingsToPanel(savedSettings);
}

/**
 * Reads saved settings from localStorage.
 */
export function readSettingsFromLocalStorage() {
  try {
    const rawSettings = localStorage.getItem(LOCAL_SETTINGS_STORAGE_KEY);

    if (!rawSettings) {
      return null;
    }

    return JSON.parse(rawSettings);
  } catch (error) {
    console.warn("Could not read saved RIT settings.", error);
    return null;
  }
}

/**
 * Saves settings to localStorage.
 */
export function saveSettingsToLocalStorage(settings) {
  try {
    localStorage.setItem(LOCAL_SETTINGS_STORAGE_KEY, JSON.stringify(settings));
  } catch (error) {
    console.warn("Could not save RIT settings.", error);
  }
}

/**
 * Clears saved local settings.
 */
export function clearSavedLocalSettings() {
  try {
    localStorage.removeItem(LOCAL_SETTINGS_STORAGE_KEY);
  } catch (error) {
    console.warn("Could not clear saved RIT settings.", error);
  }
}

/**
 * Applies saved settings to task pane controls.
 */
export function applySettingsToPanel(settings) {
  for (const controlConfig of SETTINGS_CONTROL_CONFIG) {
    const element = document.getElementById(controlConfig.id);

    if (!element) {
      continue;
    }

    const value = settings[controlConfig.settingName];

    if (value === undefined || value === null) {
      continue;
    }

    if (controlConfig.type === "checked") {
      element.checked = Boolean(value);
      continue;
    }

    if (controlConfig.type === "number") {
      element.value = String(value);
      continue;
    }

    element.value = String(value);
  }
}
