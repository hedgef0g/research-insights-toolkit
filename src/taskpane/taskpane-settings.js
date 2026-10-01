/* global console, document, window */

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
  { id: "add-table-footnote", type: "checked", settingName: "addTableFootnote" },

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
  { id: "use-cyrillic-markers", type: "checked", settingName: "useCyrillicMarkers" },
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

  {
    id: "recolor-banner-and-labels",
    type: "checked",
    settingName: "recolorBannerAndLabels",
  },
  { id: "banner-label-fill-color", type: "value", settingName: "bannerLabelFillColor" },

  { id: "preferred-base", type: "value", settingName: "preferredBase" },

  { id: "settings-storage-mode", type: "value", settingName: "settingsStorageMode" },
];

export const DEFAULT_CALCULATION_SETTINGS = {
  confidenceLevel: "95",
  oneTailedTest: false,

  roundCellValues: false,
  resultFormattingLevel: "full",
  addTableFootnote: false,

  compareWithPreviousColumn: false,
  applyPreviousColumnFill: false,

  writeBannerLetters: false,
  respectBannerStructure: false,
  autoDetectWaveBanners: false,
  labelsOnLeftSide: false,

  // Default-off for global release safety (issue #312).
  useCyrillicMarkers: false,

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

  recolorBannerAndLabels: false,
  bannerLabelFillColor: "#FFF2CC",

  preferredBase: "auto",

  settingsStorageMode: "none",
};

export const SETTINGS_TOOLTIPS = {
  "confidence-level":
    "Выберите уровень значимости для статистических тестов. Чем выше уровень, тем строже проверка и тем меньше отличий будут признаны значимыми.",

  "one-tailed-test":
    "Использовать односторонний тест вместо двустороннего. При том же уровне значимости такой тест легче находит отличия, но предполагает проверку различия в одном направлении.",

  "round-cell-values":
    "Округлять отображаемые значения перед добавлением маркеров. Расчёты при этом выполняются по исходным очищенным значениям, а не по округлённым.",

  "compare-with-previous-column":
    "Сравнивать каждую колонку только с колонкой слева: колонка 2 с колонкой 1, колонка 3 с колонкой 2 и так далее. Вместо букв используются стрелки вверх или вниз.",

  "apply-previous-column-fill":
    "Применять заливку к ячейкам со значимыми отличиями в режиме сравнения с предыдущей колонкой. Для роста используется обычная заливка значимости, для снижения — цвет “ниже Total”.",

  "write-banner-letters":
    "Добавлять буквенные индексы колонок в строку над выделенным диапазоном. Например: Segment 1 (a), Segment 2 (b). В режиме учёта структуры баннера буквы ставятся локально внутри групп.",

  "respect-banner-structure":
    "Анализировать структуру баннера над выделенным диапазоном. Это позволяет сравнивать колонки только внутри групп, определять локальные и глобальные Total, а также распознавать волновые баннеры.",

  "auto-detect-wave-banners":
    "Автоматически распознавать волновые группы в баннере, например Wave, Period, Волна, Период, и применять к ним сравнение с предыдущей колонкой. Обычные группы при этом продолжают сравниваться внутри группы обычным способом.",

  "labels-on-left-side":
    "Искать лейблы строк не рядом с выделенным диапазоном, а в самых левых колонках листа. Полезно для широких таблиц, где данные выделены справа, а названия строк находятся далеко слева.",

  "compare-only-with-total":
    "Сравнивать каждую колонку только с колонкой Total. Обычные попарные сравнения между сегментами выполняться не будут.",

  "exclude-total-from-comparisons":
    "Исключить Total из расчётов. Total не будет использоваться как база сравнения и не будет сравниваться с другими колонками.",

  "first-column-is-total":
    "Считать первую колонку выделенного диапазона Total. Она будет использоваться как референс для сравнения с остальными колонками.",

  "total-in-each-banner":
    "Считать, что Total находится внутри каждой группы баннера. При включённом учёте структуры баннера расположение Total определяется автоматически.",

  "significant-fill-color":
    "Цвет заливки для ячеек, которые статистически значимо выше другой сравниваемой ячейки или Total.",

  "lower-than-total-fill-color":
    "Цвет заливки для ячеек, которые статистически значимо ниже Total или ниже предыдущей колонки в режиме сравнения с предыдущей колонкой.",

  "fill-only-total-comparisons":
    "Применять обычную зелёную заливку только к ячейкам, которые значимо выше Total. Отличия между обычными сегментами будут отмечаться буквами, но без зелёной заливки.",

  "exclude-small-bases":
    "Исключать из расчётов колонки, где база меньше заданного порога. Такие колонки не участвуют в статистических сравнениях и получают отдельную заливку.",

  "small-base-threshold":
    "Минимальный допустимый размер базы. Если база колонки меньше этого значения, колонка исключается из расчётов.",

  "small-base-fill-color":
    "Цвет заливки для колонок с маленькой базой. Эта заливка имеет самый высокий приоритет и перекрывает остальные типы заливки.",

  "recolor-banner-and-labels":
    "Перекрашивать баннер над данными и примыкающие колонки с лейблами строк одним цветом для обработанных таблиц. Не затрагивает заливки значимости, подписи под таблицей и листы отчётов.",

  "banner-label-fill-color":
    "Общий цвет заливки для баннера и лейблов строк, когда включена перекраска баннера и лейблов.",

  "preferred-base":
    "Выберите тип базы для расчёта значимости. «Авто» использует приоритет: Effective → Unweighted → Base → Weighted. Если выбранный тип базы не найден в таблице, используется автоматический приоритет.",

  "settings-storage-mode":
    "Выберите, сохранять ли настройки панели. Локальное сохранение работает только на этом устройстве и в этом браузере/Excel WebView.",

  "reset-settings":
    "Сбросить все настройки к значениям по умолчанию и удалить локально сохранённые настройки.",
};

function applySettingsCollapseState(toggleBtn, panelBody, isCollapsed) {
  toggleBtn.setAttribute("aria-expanded", isCollapsed ? "false" : "true");
  panelBody.classList.toggle("is-collapsed", isCollapsed);
}

export function initializeSettingsPanel({
  initializePreviousColumnComparisonSettings,
  initializeSettingsResetButton,
  initializeBannerStructureSettings,
} = {}) {
  initializeSettingsCollapse();
  initializeSettingsTooltips();
  initializePreviousColumnTotalWarningPlacement();

  bindMutuallyExclusiveCheckboxes("compare-only-with-total", "exclude-total-from-comparisons");
  bindMutuallyExclusiveCheckboxes("first-column-is-total", "total-in-each-banner");

  if (initializePreviousColumnComparisonSettings) {
    initializePreviousColumnComparisonSettings();
  }

  if (initializeSettingsResetButton) {
    initializeSettingsResetButton();
  }

  initializeSettingsTabs();

  if (initializeBannerStructureSettings) {
    initializeBannerStructureSettings();
  }

  const helpLink = document.getElementById("help-link");

  if (helpLink) {
    helpLink.addEventListener("click", (event) => {
      event.preventDefault();
      window.open(
        "https://hedgef0g.github.io/research-insights-toolkit/assets/rit-help-ru.html",
        "_blank"
      );
    });
  }
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

export function initializePreviousColumnTotalWarningPlacement() {
  const warningElement = document.getElementById("previous-column-total-warning");
  const totalInEachBannerCheckbox = document.getElementById("total-in-each-banner");

  if (!warningElement || !totalInEachBannerCheckbox || !totalInEachBannerCheckbox.parentElement) {
    return;
  }

  totalInEachBannerCheckbox.parentElement.insertAdjacentElement("afterend", warningElement);
}

/**
 * Adds Russian tooltips to settings controls.
 *
 * PURPOSE:
 * Keep taskpane.html clean and define all setting explanations in one place.
 */
export function initializeSettingsTooltips() {
  for (const [elementId, tooltipText] of Object.entries(SETTINGS_TOOLTIPS)) {
    const element = document.getElementById(elementId);

    if (!element) {
      continue;
    }

    element.title = tooltipText;
    element.setAttribute("aria-label", tooltipText);

    const explicitLabel = document.querySelector(`label[for="${elementId}"]`);

    if (explicitLabel) {
      explicitLabel.title = tooltipText;
      continue;
    }

    const wrappingLabel = element.closest("label");

    if (wrappingLabel) {
      wrappingLabel.title = tooltipText;
      continue;
    }

    const parent = element.parentElement;

    if (parent) {
      const labelLikeText = parent.querySelector("span, .checkbox-text, .label-text");

      if (labelLikeText) {
        labelLikeText.title = tooltipText;
      }
    }
  }
}

/**
 * Makes two checkboxes mutually exclusive.
 *
 * PURPOSE:
 * If one checkbox is enabled, the other one is disabled automatically.
 */
export function bindMutuallyExclusiveCheckboxes(firstCheckboxId, secondCheckboxId) {
  const firstCheckbox = document.getElementById(firstCheckboxId);
  const secondCheckbox = document.getElementById(secondCheckboxId);

  if (!firstCheckbox || !secondCheckbox) {
    return;
  }

  firstCheckbox.addEventListener("change", () => {
    if (firstCheckbox.checked) {
      secondCheckbox.checked = false;
    }
  });

  secondCheckbox.addEventListener("change", () => {
    if (secondCheckbox.checked) {
      firstCheckbox.checked = false;
    }
  });
}

export function initializeSettingsTabs() {
  const navContainer = document.getElementById("settings-tab-nav");

  if (!navContainer) {
    return;
  }

  navContainer.addEventListener("click", (e) => {
    const tab = e.target.closest("[data-settings-tab]");
    if (!tab) return;

    const targetPanel = tab.dataset.settingsTab;

    navContainer.querySelectorAll(".settings-tab").forEach((t) => {
      const active = t.dataset.settingsTab === targetPanel;
      t.classList.toggle("is-active", active);
      t.setAttribute("aria-selected", active ? "true" : "false");
    });

    document.querySelectorAll(".settings-tab-panel").forEach((panel) => {
      panel.style.display = panel.dataset.settingsPanel === targetPanel ? "" : "none";
    });
  });
}
