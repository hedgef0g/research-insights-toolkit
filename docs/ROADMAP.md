# Research Insights Toolkit — Roadmap

Рабочий roadmap проекта. Это не публичное обещание сроков, а список направлений развития от текущего MVP к более стабильному продукту.

## Phase 1 — Manual workflow stabilization (Completed)

The manual selected-range workflow is stable. Run and Clear both go through
the shared selected-range normalization path with banner-aware and
wave-aware behavior.

- [x] read-only table preview model exists as a core model but is not yet wired into UI
- [x] shared text normalization utilities
- [x] banner-aware stabilization
- [x] numeric output preservation
- [x] Clear significance numeric restoration
- [x] taskpane primary action layout improvement
- [x] warning-only selected range guardrails
- [x] selected-range normalization for full-table selections, used by Run and Clear
- [x] repeat full-table normalization and extended NPS labels (#93)
- [x] sparse/merged banner marker placement (#94)
- [x] Clear full-table body-only restoration (#96)
- [x] Clear banner/header marker clearing (#97)
- [x] banner-aware Total detection (#98)
- [x] nested wave-aware detection (#100)
- [x] manual smoke checklist refresh (#101)

## Current focus — Check table / preview foundation

Expose the normalized interpretation model to the user before Excel
mutation, so the user can verify how the add-in understood the table.
The intent is to reuse the existing selected-range normalization output
as the source of truth for the Check-table view rather than building a
separate parsing path.

- [ ] **Валидация перед расчетом**: Режим «Проверить таблицу», визуализирующий, как надстройка поняла структуру данных перед внесением изменений.
- [ ] Decide whether and how to wire the existing read-only table preview model into UI on top of the normalized interpretation output.

## High-priority future/spec work

- [ ] Continue base placement product/spec work.
- [ ] Multi-column row labels and weighted/effective bases remain discovery/spec first.
- [ ] Custom modal dialogs remain separate and should not be mixed into preview wiring.

## Code health backlog (was Фаза 1)
- [ ] **Рефакторинг UI-слоя**: Выделение `ui-controller.js` для работы с DOM и настройками. Очистка `taskpane.js`.
- [ ] **Кастомные диалоги**: Замена нестабильных `window.confirm/alert` на HTML-модальные окна для консистентной работы в Excel Desktop и Web.
- [ ] **Общие утилиты**: Создание `matrix-helpers.js` для DRY-логики обхода таблиц в детекторе метрик и баннеров.
- [ ] **Оптимизация записи**: Усиление батчинга (группировки) операций форматирования в `excel-writer.js` для повышения производительности на больших таблицах.

## Smart detection backlog
- [ ] **Поиск баз вне выделения**: Алгоритм автоматического поиска строки Total, если она находится непосредственно над выделенным диапазоном.
- [ ] **Изоляция эвристик баннера**: Вынос логики распознавания сложных шапок (Report Title, Merged Spans) в отдельные тестируемые модули.

## Automatic runner direction
- [ ] Automatic runner should reuse the normalized interpretation / preview model.
- [ ] Do not implement auto-runner as repeated blind calls to the manual Run path.

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
