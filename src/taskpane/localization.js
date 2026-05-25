/* global localStorage, document */

/**
 * Lightweight RU/EN localization module for the taskpane UI.
 *
 * - Supported languages: ru, en
 * - Default language: ru
 * - t(key) returns the translated string; falls back to Russian, then the key itself
 * - applyI18n() updates all [data-i18n] elements and [data-hint-i18n] data attributes
 * - Language selection is persisted to localStorage independently of settings storage mode
 */

const LANG_STORAGE_KEY = "rit.ui.lang";

const STRINGS = {
  ru: {
    "lang.label": "Язык",

    "tabs.run": "Расчёт",
    "tabs.autorun": "Автозапуск",
    "tabs.check": "Проверка",
    "tabs.content": "Оглавление",

    "scope.currentTable": "Текущая таблица",
    "scope.currentSheet": "Текущий лист",
    "scope.wholeWorkbook": "Вся книга",

    "btn.calculate": "Рассчитать",
    "btn.clearSignificance": "Очистить значимости",
    "btn.checkSelection": "Проверить выделение",
    "btn.checkTable": "Проверить таблицу",
    "btn.checkSheet": "Проверить лист",
    "btn.checkWorkbook": "Проверить книгу",
    "btn.createContents": "Создать оглавление",

    "hint.runCurrentTable": "Область: выделенный диапазон",
    "hint.autorunCurrentTable": "Область: активная ячейка → таблица",
    "hint.checkCurrentTable": "Проверяет таблицу вокруг активной ячейки.",
    "hint.checkCurrentSheet": "Проверяет таблицы на текущем листе.",
    "hint.checkWholeWorkbook": "Проверяет таблицы по всей книге.",
    "hint.checkWriteMode": " С записью на лист.",
    "hint.checkReadOnly": " Только чтение.",

    "content.format": "Формат оглавления",
    "content.formatClient": "Для клиента",
    "content.formatMinimalCheck": "С минимальной проверкой",
    "content.formatFullCheck": "С полной проверкой",
    "content.addBacklinks": "Ставить обратные ссылки",
    "content.hintText":
      "Сканирует все листы книги и создаёт или обновляет лист «Content» со списком таблиц. Данные таблиц не изменяются.",

    "control.addDiagnosticSheet": "Добавить диагностический лист",
    "control.writeResult": "Записать результат",

    "panel.status": "Статус",
    "panel.check": "Проверка",
    "panel.inventory": "Оглавление",

    "settings.tab.basic": "Основные",
    "settings.tab.comparisons": "Сравнения",
    "settings.tab.banner": "Баннер",
    "settings.tab.bases": "Базы",
    "settings.tab.design": "Дизайн",

    "settings.confidenceLevel": "Уровень значимости",
    "settings.preferredBase": "База для расчёта",
    "settings.saveSettings": "Сохранять настройки...",
    "settings.reset": "Сброс",

    "storage.none": "Не сохранять",
    "storage.local": "Локально",
    "storage.cloud": "В облаке",

    "status.nonContiguousSelection":
      "Выделение состоит из нескольких несмежных областей. " +
      "Для этой операции выберите один непрерывный диапазон или поставьте курсор внутри одной таблицы.",
  },

  en: {
    "lang.label": "Language",

    "tabs.run": "Run",
    "tabs.autorun": "Autorun",
    "tabs.check": "Check",
    "tabs.content": "Contents",

    "scope.currentTable": "Current table",
    "scope.currentSheet": "Current sheet",
    "scope.wholeWorkbook": "Whole workbook",

    "btn.calculate": "Calculate",
    "btn.clearSignificance": "Clear significance",
    "btn.checkSelection": "Check selection",
    "btn.checkTable": "Check table",
    "btn.checkSheet": "Check sheet",
    "btn.checkWorkbook": "Check workbook",
    "btn.createContents": "Create contents",

    "hint.runCurrentTable": "Scope: selected range",
    "hint.autorunCurrentTable": "Scope: active cell → table",
    "hint.checkCurrentTable": "Checks the table around the active cell.",
    "hint.checkCurrentSheet": "Checks tables on the current sheet.",
    "hint.checkWholeWorkbook": "Checks tables across the whole workbook.",
    "hint.checkWriteMode": " Writing to sheet.",
    "hint.checkReadOnly": " Read only.",

    "content.format": "Contents format",
    "content.formatClient": "Client view",
    "content.formatMinimalCheck": "With minimal check",
    "content.formatFullCheck": "With full check",
    "content.addBacklinks": "Add backlinks",
    "content.hintText":
      "Scans all workbook sheets and creates or updates the Content sheet with a list of tables. Table data is not modified.",

    "control.addDiagnosticSheet": "Add diagnostic sheet",
    "control.writeResult": "Write result",

    "panel.status": "Status",
    "panel.check": "Check",
    "panel.inventory": "Contents",

    "settings.tab.basic": "General",
    "settings.tab.comparisons": "Comparisons",
    "settings.tab.banner": "Banner",
    "settings.tab.bases": "Bases",
    "settings.tab.design": "Design",

    "settings.confidenceLevel": "Significance level",
    "settings.preferredBase": "Calculation base",
    "settings.saveSettings": "Save settings...",
    "settings.reset": "Reset",

    "storage.none": "Don't save",
    "storage.local": "Local",
    "storage.cloud": "Cloud",

    "status.nonContiguousSelection":
      "Selection contains multiple non-contiguous areas. " +
      "For this operation, select a single continuous range or place the cursor inside one table.",
  },
};

export const SUPPORTED_LANGUAGES = ["ru", "en"];
export const DEFAULT_LANGUAGE = "ru";

let _currentLanguage = DEFAULT_LANGUAGE;

export function getLanguage() {
  return _currentLanguage;
}

/**
 * Sets the active language, persists the choice, and re-applies all translations.
 */
export function setLanguage(lang) {
  if (!SUPPORTED_LANGUAGES.includes(lang)) return;
  _currentLanguage = lang;
  try {
    localStorage.setItem(LANG_STORAGE_KEY, lang);
  } catch (_) {
    /* non-fatal */
  }
  applyI18n();
}

/**
 * Loads the previously saved language from localStorage (if any).
 * Call once during taskpane initialization before applyI18n().
 */
export function loadSavedLanguage() {
  try {
    const saved = localStorage.getItem(LANG_STORAGE_KEY);
    if (saved && SUPPORTED_LANGUAGES.includes(saved)) {
      _currentLanguage = saved;
    }
  } catch (_) {
    /* non-fatal */
  }
}

/**
 * Returns the translated string for key in the current language.
 * Falls back to the default language, then returns the key itself — never throws.
 *
 * @param {string} key
 * @param {Object} [params] - Optional {placeholder: value} substitution map
 */
export function t(key, params) {
  const dict = STRINGS[_currentLanguage] || STRINGS[DEFAULT_LANGUAGE];
  let str = dict[key];
  if (str === undefined) {
    str = STRINGS[DEFAULT_LANGUAGE][key];
  }
  if (str === undefined) {
    return key;
  }
  if (params) {
    for (const [k, v] of Object.entries(params)) {
      str = str.replace(`{${k}}`, String(v));
    }
  }
  return str;
}

/**
 * Applies translations to all [data-i18n] elements and refreshes [data-hint-i18n]
 * data-hint-base attributes so that updateCheckHints() picks up the new language.
 *
 * Safe to call repeatedly on language switch — does not create or destroy elements.
 */
export function applyI18n() {
  // Update text content for leaf elements tagged with data-i18n.
  // Only update elements with no child elements to avoid destroying nested HTML
  // (e.g. labels that wrap a checkbox <input>).
  document.querySelectorAll("[data-i18n]").forEach((el) => {
    if (el.children.length === 0) {
      el.textContent = t(el.dataset.i18n);
    }
  });

  // Update the data-hint-base attribute for check workspace hints so that
  // updateCheckHints() (in taskpane.js) can append the correct mode suffix.
  document.querySelectorAll("[data-hint-i18n]").forEach((el) => {
    el.dataset.hintBase = t(el.dataset.hintI18n);
  });

  // Keep the language selector in sync with the current language.
  const langSelect = document.getElementById("language-selector");
  if (langSelect && langSelect.value !== _currentLanguage) {
    langSelect.value = _currentLanguage;
  }
}
