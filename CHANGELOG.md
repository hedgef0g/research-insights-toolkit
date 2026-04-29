# Changelog

## [Current Development Stage]

#### 28.04.2026 

### Added
- Unified significance engine with automatic metric detection.
- Support for significance testing of proportions.
- Support for significance testing of means using:
  - Standard deviation
  - Variance
- Support for significance testing of NPS using:
  - Promoters / Detractors structure
  - Standard deviation
  - Variance
- Automatic pairwise all-vs-all comparison across selected columns.
- Automatic significance letters appended directly into Excel cells.
- Pale green highlight + bold formatting for significant cells.
- Automatic center alignment for selected ranges.
- Separate button for clearing significance markers.
- Metric type diagnostics tool.
- Shared-base support across mixed metric tables.
- Support for mixed tables containing:
  - proportions
  - means
  - NPS
  - any combination of them

### Improved
- Replaced single-plan detection with block-plan detector.
- Smarter handling of tables where one base row serves several metric blocks.
- Better compatibility with real research tables.
- Added protection against Excel converting values into time format after adding letters.
- Improved project launch workflow with automatic VS Code startup tasks.

### Fixed
- NPS significance routing issues in auto mode.
- Incorrect significance letters appearing in Promoters / Detractors rows.
- Branch execution conflicts between mean / NPS / fallback logic.
- Repeated marker formatting inconsistencies.


#### 29.04.2026 

### Первый проход рефакторинга
Универсальное автоопределение метрик (buildCalculationBlocks) работает стабильно, поэтому старые ручные режимы расчета для конкретных типов данных (и жестко привязанные к ним парсеры) больше не нужны. Кодовая база очищена от неиспользуемых функций для упрощения дальнейшей поддержки.

Основные изменения по файлам:
- taskpane.js: удалены обработчики (runMeanSignificance..., runNpsSignificance...) и слушатели событий для кнопок явного расчета средних и NPS. Очищены неиспользуемые импорты.
- metric-detector.js: удален устаревший планировщик buildAutoCalculationPlan (полностью заменен на актуальный buildCalculationBlocks).
- significance.js: удалена легаси-функция из MVP v0.2 compareAllRowsUsingBottomBases, а также функции-обертки для отпавших ручных расчетов (compareMeansUsingSpreadAndBaseRows, compareNpsUsingStructureRows, compareNpsUsingSpreadAndBaseRows).

Примечание: математическое ядро расчетов значимости (z-тесты, t-тесты) и функционал диагностической кнопки оставлены без изменений.

### Второй проход рефакторинга
Устранение технического долга: разделение ответственности (Separation of Concerns) и избавление от дублирования кода (DRY) для повышения читаемости и масштабируемости проекта.

Основные архитектурные изменения:
- Вынос Excel-рендера: функция writeMarkersToSelectedRange перемещена из taskpane.js в новый модуль core/excel-writer.js. taskpane.js теперь выступает исключительно в роли контроллера.
- Изоляция математического ядра: логика парсинга и очистки "грязных" табличных данных (пробелы, символы %, запятые) вынесена из significance.js в core/normalizers.js. Создана единая утилита parseRawCellValue. В significance.js остались только чистые статистические функции.
- Config-driven детектор: хардкод массивов с ключевыми словами вырезан из metric-detector.js. Создан конфигурационный файл core/dictionary.config.js (METRIC_DICTIONARY), что позволяет легко добавлять новые термины и языки без изменения логики алгоритма поиска.

### Дополнение функционала №1

## Добавлено

- Настраиваемый уровень значимости через интерфейс панели:
  - 99%
  - 95%
  - 90%
  - 80%
  - 66.6%

- Выбранный уровень значимости теперь применяется ко всем поддерживаемым типам метрик:
  - пропорции
  - средние значения
  - NPS по стандартному отклонению / дисперсии
  - NPS по структуре (Promoters / Detractors)

## Улучшено

- Перезапуск расчёта теперь полностью очищает старые маркеры значимости и форматирование перед новым пересчётом.

- Улучшен block-plan detector для комплексных таблиц:
  - корректная работа при общей базе для нескольких метрик
  - устранён пропуск следующих блоков после Mean / NPS
  - улучшена обработка смешанных таблиц с пропорциями, средними и NPS


### Дополнение функционала №2

## Добавлено

- Добавлен сворачиваемый блок **"Настройки"** в панели надстройки.
- Добавлена основа UI для будущей работы с баннерами:
  - проставление букв в баннере;
  - учёт структуры баннера;
  - поиск лейблов в самых левых колонках листа;
  - настройки сравнения с Total;
  - настройки расположения Total;
  - настройки заливок;
  - настройки маленьких баз;
  - выбор режима сохранения настроек.
- Добавлена настройка **"Округлять значения в ячейках"**.
- Добавлена подстановка буквенных маркеров в строку над выделенным диапазоном:
  - формат маркера: `(a)`, `(b)`, `(c)` и т.д.;
  - если в ячейке баннера уже есть старый маркер, он заменяется новым;
  - если нужный маркер уже стоит в конце ячейки, повторно он не добавляется.
- Добавлена возможность искать лейблы метрик не только непосредственно слева от выделенного диапазона, но и в самых левых колонках листа.
- Добавлено предупреждение для случая, когда включена подстановка букв в баннер, но выделенный диапазон начинается с первой строки листа.

## Улучшено

- Кнопка расчёта переименована в **"Запустить"**.
- Значения в ячейках теперь могут округляться перед подстановкой букв:
  - по умолчанию доли, NPS, Promoters и Detractors округляются до 1 знака;
  - по умолчанию средние, стандартные отклонения и дисперсии округляются до 2 знаков;
  - при включённой настройке округления доли, NPS, Promoters и Detractors округляются до целого;
  - при включённой настройке округления средние, стандартные отклонения и дисперсии округляются до 1 знака.
- Улучшен поиск лейблов слева от выделенного диапазона:
  - промежуточные числовые колонки больше не принимаются за текстовые лейблы;
  - это позволяет выделять данные без колонки Total, сохраняя корректное распознавание строк.

## Исправлено

- Исправлена ошибка, при которой поиск лейблов мог ошибочно использовать числовое значение из промежуточной колонки вместо настоящего лейбла.