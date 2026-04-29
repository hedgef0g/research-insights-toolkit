# Decisions

## 2026-04-28

The first significance MVP assumes:
- input values are proportions/percentages;
- bases are below values;
- selection is a 2x2 range;
- significance is calculated using pooled z-test;
- confidence level is fixed at 95%;
- Excel wrapper is only responsible for reading selection and displaying result;
- statistical logic lives in src/core/significance.js.

# Latest Decisions

## Product Logic

- The main user flow is a single unified button with automatic metric detection.
- Separate explicit buttons remain available for testing and fallback use.
- Real-world tables are prioritized over idealized academic layouts.
- Mixed tables with several metric types must be supported.

## Detection Architecture

- Single-plan detection was deprecated in favor of block-plan detection.
- A selected range may contain multiple independent calculation blocks.
- Base rows may be:
  - dedicated to one metric
  - shared across several metrics
- Proportion rows may wait for the next available base row.

## Output Rules

- Significance markers may appear only in value rows:
  - proportions
  - mean
  - NPS
- Service rows must never receive markers:
  - Base
  - SD
  - Variance
  - Promoters
  - Detractors

## UX

- One-click startup inside VS Code is required.
- Auto-start on folder open enabled for faster iteration.
- Reduce manual terminal work wherever possible.

## Strategic Direction

- Build an intelligent spreadsheet insights tool, not just a macro pack.
- Priority is Excel first, Google Sheets second.


## UI / Settings Panel

- Панель настроек становится основой для дальнейшего развития продукта.
- Все новые пользовательские режимы должны по возможности добавляться через настройки панели, а не через отдельные кнопки.
- Основная кнопка расчёта называется **"Запустить"**.
- Блок **"Настройки"** должен быть сворачиваемым, чтобы панель оставалась компактной.
- Настройки, заложенные под будущую логику, могут временно быть `disabled`, если соответствующая расчётная логика ещё не реализована.

## Banner Logic

- Подстановка букв в баннер управляется отдельной настройкой.
- Маркер в баннере записывается в формате `(a)`, `(b)`, `(c)` и т.д.
- Маркер добавляется в ячейку над соответствующим столбцом выделенного диапазона.
- Если в ячейке баннера уже есть маркер в этом формате, он заменяется актуальным.
- Если актуальный маркер уже стоит в конце ячейки, повторное добавление не выполняется.
- Если выделенный диапазон начинается с первой строки листа и включена подстановка букв в баннер, расчёт должен быть остановлен с предупреждением. Автоматическую вставку строки решено отложить до отдельной UI-итерации.

## Label Detection

- Поведение по умолчанию: искать лейблы непосредственно слева от выделенного диапазона.
- Дополнительный режим: искать лейблы в самых левых колонках листа.
- Числовые значения между лейблами и выделенным диапазоном не должны распознаваться как лейблы.
- Это позволяет пользователю временно исключать колонку Total из выделения, сохраняя корректное распознавание метрик.

## Rounding

- Округление выполняется до подстановки букв в ячейку.
- Округление применяется на этапе подготовки отображаемого значения, а не на этапе статистического расчёта.
- Статистические расчёты должны использовать исходные очищенные значения, а не округлённые отображаемые значения.
- По умолчанию:
  - доли, NPS, Promoters и Detractors округляются до 1 знака;
  - средние, стандартные отклонения и дисперсии округляются до 2 знаков.
- При включённой настройке **"Округлять значения в ячейках"**:
  - доли, NPS, Promoters и Detractors округляются до целого;
  - средние, стандартные отклонения и дисперсии округляются до 1 знака.

## Known Limitations

- Нативный `window.confirm()` в Office Add-in WebView может работать ненадёжно, поэтому для сценария “Добавить строку / Отменить” потребуется отдельный HTML warning panel или modal.
- Логика сравнения с Total и полноценная работа со структурой баннера пока заложены в UI, но не реализованы в расчётном ядре.