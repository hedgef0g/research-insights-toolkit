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

    "settings.panelTitle": "Настройки",

    "settings.tab.basic": "Основные",
    "settings.tab.comparisons": "Сравнения",
    "settings.tab.banner": "Баннер",
    "settings.tab.bases": "Базы",
    "settings.tab.design": "Дизайн",

    "settings.tab.basic.short": "Основные",
    "settings.tab.comparisons.short": "Сравнения",
    "settings.tab.banner.short": "Баннер",
    "settings.tab.bases.short": "Базы",
    "settings.tab.design.short": "Дизайн",

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

    "status.multiTableGapAutorun":
      "Выделение содержит несколько таблиц или блоков данных, разделённых пустыми строками. " +
      "Для «Текущей таблицы» поставьте курсор в одну таблицу или используйте «Автозапуск → Текущий лист».",
    "status.multiTableGapCheck":
      "Выделение содержит несколько таблиц или блоков данных, разделённых пустыми строками. " +
      "Для «Текущей таблицы» поставьте курсор в одну таблицу или используйте «Проверить лист».",
    "status.multiTableGapCheckSelection":
      "В диапазоне обнаружено несколько блоков данных, разделённых пустыми строками. " +
      "Перейдите в ячейку внутри одной таблицы или используйте «Проверить лист».",
    "status.noDataInRange": "Нет данных в диапазоне.",

    "status.autorunTableNotResolved": "Автозапуск — Текущая таблица: не удалось определить таблицу.",
    "status.autorunCalcError": "Автозапуск — Текущая таблица: ошибка расчёта — {msg}",
    "status.autorunTableDetectError": "Автозапуск — Текущая таблица: ошибка при определении таблицы — {msg}",
    "status.autorunProcessed": "Автозапуск — Текущая таблица: выполнен. {sheet}!{range}. Блоков: {count}.",
    "status.autorunSkipped": "Автозапуск — Текущая таблица: {msg}.",
    "status.autorunClearError": "Автозапуск — Текущая таблица: ошибка очистки — {msg}",
    "status.autorunCleared": "Автозапуск — Текущая таблица: очищено. {sheet}!{range}.",
    "status.autorunClearSkipped": "Автозапуск — Текущая таблица: очистка пропущена — {msg}.",
    "status.autorunReportWriteError": "[Отчёт: ошибка записи — {msg}]",

    "status.settingsReset": "Настройки сброшены к значениям по умолчанию.",

    "status.runDone": "Расчёт выполнен. Обработано блоков: {count}.",
    "status.clearDone": "Значимости очищены.",

    "status.checkDone":
      "Проверка завершена. {sheet}!{range}. Строк: {rows}. Блоков: {blocks}. Баз: {bases}. Предупреждений: {warnings}. Критических: {critical}.",
    "status.checkSelectionDone":
      "Проверка выделения завершена. {sheet}!{range}. Строк: {rows}. Блоков: {blocks}. Баз: {bases}. Предупреждений: {warnings}. Критических: {critical}.",

    "status.resolverNoTable":
      "Активная ячейка не находится внутри ни одного кандидата. Перейдите в ячейку внутри таблицы.",
    "status.resolverGeneratedSheet":
      "Лист создан надстройкой и не содержит исследовательских таблиц.",
    "status.resolverAmbiguousBoundary":
      "Активная ячейка входит в несколько перекрывающихся кандидатов. Уточните позицию курсора.",
    "status.resolverBlocked": "Лист слишком большой для сканирования.",
    "status.resolverFallback": "Не удалось определить таблицу под активной ячейкой.",

    "settings.comparePrevCol": "Сравнение с предыдущей колонкой",
    "settings.applyPrevColFill": "Применить заливку",
    "settings.compareWithTotal": "Сравнение с Тоталом",
    "settings.legendCompareMode": "Режим сравнения с Тоталом",
    "settings.compareOnlyWithTotal": "Сравнивать только с Тотал",
    "settings.excludeTotalFromComparisons": "Не сравнивать с Тотал",
    "settings.legendTotalPosition": "Расположение Тотала",
    "settings.firstColumnIsTotal": "Первая колонка — Тотал",
    "settings.warningTotalPreviousCol":
      "WARNING: При выбранных настройках Тотал будет обрабатываться как обычная предыдущая " +
      "колонка, если не включить «Не сравнивать с Тотал»",
    "settings.totalInEachBanner": "Тотал в каждом баннере",

    "settings.writeBannerLetters": "Проставлять буквы в баннере",
    "settings.respectBannerStructure": "Учитывать структуру баннера",
    "settings.autoDetectWaveBanners": "Автоматически определять волны",
    "settings.labelsOnLeftSide": "Лейблы значений слева листа",

    "settings.smallBases": "Маленькие базы",
    "settings.excludeSmallBases": "Не сравнивать маленькие базы",
    "settings.smallBaseThreshold": "База <",

    "settings.roundCellValues": "Округлять значения в ячейках",
    "settings.fillsGroup": "Заливки",
    "settings.significantFillColor": "Обычная значимость",
    "settings.lowerThanTotalFillColor": "Заливка < Тотал",
    "settings.fillOnlyTotalComparisons": "Заливка только для Тотала",
    "settings.smallBaseFillColor": "Заливка маленькой базы",

    "settings.oneTailedTest": "Односторонний тест",
    "settings.preferredBaseAuto": "Авто",

    "help.link": "Справка об использовании",
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

    "settings.panelTitle": "Settings",

    "settings.tab.basic": "General",
    "settings.tab.comparisons": "Comparisons",
    "settings.tab.banner": "Banner",
    "settings.tab.bases": "Bases",
    "settings.tab.design": "Design",

    "settings.tab.basic.short": "General",
    "settings.tab.comparisons.short": "Comparisons",
    "settings.tab.banner.short": "Banner",
    "settings.tab.bases.short": "Bases",
    "settings.tab.design.short": "Design",

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

    "status.multiTableGapAutorun":
      "Selection contains multiple tables or data blocks separated by empty rows. " +
      "For Current Table, place the cursor inside one table or use Autorun → Current sheet.",
    "status.multiTableGapCheck":
      "Selection contains multiple tables or data blocks separated by empty rows. " +
      "For Current Table, place the cursor inside one table or use Check sheet.",
    "status.multiTableGapCheckSelection":
      "Multiple data blocks separated by empty rows detected in the range. " +
      "Navigate to a cell inside a single table or use Check sheet.",
    "status.noDataInRange": "No data in range.",

    "status.autorunTableNotResolved": "Autorun — Current table: could not identify the table.",
    "status.autorunCalcError": "Autorun — Current table: calculation error — {msg}",
    "status.autorunTableDetectError": "Autorun — Current table: error detecting table — {msg}",
    "status.autorunProcessed": "Autorun — Current table: done. {sheet}!{range}. Blocks: {count}.",
    "status.autorunSkipped": "Autorun — Current table: {msg}.",
    "status.autorunClearError": "Autorun — Current table: clear error — {msg}",
    "status.autorunCleared": "Autorun — Current table: cleared. {sheet}!{range}.",
    "status.autorunClearSkipped": "Autorun — Current table: clear skipped — {msg}.",
    "status.autorunReportWriteError": "[Report: write error — {msg}]",

    "status.settingsReset": "Settings were reset to defaults.",

    "status.runDone": "Calculation complete. Blocks processed: {count}.",
    "status.clearDone": "Significance markers removed.",

    "status.checkDone":
      "Check complete. {sheet}!{range}. Rows: {rows}. Blocks: {blocks}. Bases: {bases}. Warnings: {warnings}. Critical: {critical}.",
    "status.checkSelectionDone":
      "Selection check complete. {sheet}!{range}. Rows: {rows}. Blocks: {blocks}. Bases: {bases}. Warnings: {warnings}. Critical: {critical}.",

    "status.resolverNoTable":
      "Active cell is not inside any candidate table. Navigate to a cell inside a table.",
    "status.resolverGeneratedSheet":
      "This sheet was created by the add-in and does not contain research tables.",
    "status.resolverAmbiguousBoundary":
      "Active cell falls within multiple overlapping candidates. Refine the cursor position.",
    "status.resolverBlocked": "Sheet is too large to scan.",
    "status.resolverFallback": "Could not identify the table under the active cell.",

    "settings.comparePrevCol": "Compare with previous column",
    "settings.applyPrevColFill": "Apply fill",
    "settings.compareWithTotal": "Compare with Total",
    "settings.legendCompareMode": "Total comparison mode",
    "settings.compareOnlyWithTotal": "Compare only with Total",
    "settings.excludeTotalFromComparisons": "Exclude Total from comparisons",
    "settings.legendTotalPosition": "Total position",
    "settings.firstColumnIsTotal": "First column is Total",
    "settings.warningTotalPreviousCol":
      "WARNING: With the current settings Total will be treated as a regular previous column " +
      "unless you enable 'Exclude Total from comparisons'",
    "settings.totalInEachBanner": "Total in each banner",

    "settings.writeBannerLetters": "Write letters in banner",
    "settings.respectBannerStructure": "Respect banner structure",
    "settings.autoDetectWaveBanners": "Auto-detect wave banners",
    "settings.labelsOnLeftSide": "Value labels on sheet left",

    "settings.smallBases": "Small bases",
    "settings.excludeSmallBases": "Exclude small bases",
    "settings.smallBaseThreshold": "Base <",

    "settings.roundCellValues": "Round cell values",
    "settings.fillsGroup": "Fills",
    "settings.significantFillColor": "Regular significance",
    "settings.lowerThanTotalFillColor": "Fill < Total",
    "settings.fillOnlyTotalComparisons": "Fill only for Total comparisons",
    "settings.smallBaseFillColor": "Small base fill",

    "settings.oneTailedTest": "One-tailed test",
    "settings.preferredBaseAuto": "Auto",

    "help.link": "User guide",
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

  // Update title attributes for elements that declare data-i18n-title
  // (e.g. action tabs that may be ellipsized at narrow widths).
  document.querySelectorAll("[data-i18n-title]").forEach((el) => {
    el.title = t(el.dataset.i18nTitle);
  });

  // Update the data-hint-base attribute for check workspace hints so that
  // updateCheckHints() (in taskpane.js) can append the correct mode suffix.
  document.querySelectorAll("[data-hint-i18n]").forEach((el) => {
    el.dataset.hintBase = t(el.dataset.hintI18n);
  });

  // Keep flag buttons in sync with the current language.
  document.querySelectorAll(".lang-btn[data-lang]").forEach((btn) => {
    btn.classList.toggle("is-active", btn.dataset.lang === _currentLanguage);
  });
}
