# Research Insights Toolkit — Roadmap

Рабочий roadmap проекта. Это не публичное обещание сроков, а список направлений развития от текущего MVP к более стабильному продукту.

## Completed/stabilized
- [x] read-only table preview model exists as a core model but is not yet wired into UI
- [x] shared text normalization utilities exist
- [x] banner-aware stabilization
- [x] numeric output preservation
- [x] Clear significance numeric restoration
- [x] taskpane primary action layout improvement
- [x] warning-only selected range guardrails

## High-priority future/spec work
- [ ] Design selected range normalization for full-table selections.
- [ ] Decide whether and how to wire table preview UI.
- [ ] Continue base placement product/spec work.
- [ ] Multi-column row labels and weighted/effective bases remain discovery/spec first.
- [ ] Custom modal dialogs remain separate and should not be mixed into selection normalization.

## Фаза 1: Очистка и стабильность (Текущий этап)
- [ ] **Рефакторинг UI-слоя**: Выделение `ui-controller.js` для работы с DOM и настройками. Очистка `taskpane.js`.
- [ ] **Кастомные диалоги**: Замена нестабильных `window.confirm/alert` на HTML-модальные окна для консистентной работы в Excel Desktop и Web.
- [ ] **Общие утилиты**: Создание `matrix-helpers.js` для DRY-логики обхода таблиц в детекторе метрик и баннеров.
- [ ] **Оптимизация записи**: Усиление батчинга (группировки) операций форматирования в `excel-writer.js` для повышения производительности на больших таблицах.

## Фаза 2: Улучшение парсинга (Smart Detection)
- [ ] **Поиск баз вне выделения**: Алгоритм автоматического поиска строки Total, если она находится непосредственно над выделенным диапазоном.
- [ ] **Изоляция эвристик баннера**: Вынос логики распознавания сложных шапок (Report Title, Merged Spans) в отдельные тестируемые модули.
- [ ] **Валидация перед расчетом**: Режим «Проверить таблицу», визуализирующий, как надстройка поняла структуру данных перед внесением изменений.

## Фаза 3: Расширение аналитики
- [ ] **Взвешенные базы (Weighted bases)**: Поддержка корректного расчета значимости для данных с весами.
- [ ] **Зависимые выборки**: Реализация тестов на пересекающиеся аудитории (Overlap tests).
- [ ] **Дополнительные метрики**: Поддержка Median, Multiple Response групп.

## Фаза 4: Масштабирование
- [ ] **Google Sheets Support**: Портирование платформонезависимого ядра на Google Apps Script.
- [ ] **Облачное хранение настроек**: Синхронизация предпочтений пользователя (цвета, пороги) между разными книгами и устройствами.


## Deferred
- [ ] worksheet/workbook auto-scan
- [ ] row-wise comparisons
- [ ] weighted base calculation support
- [ ] broad taskpane redesign
