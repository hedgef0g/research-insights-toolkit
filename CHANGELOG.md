# Changelog

## [Current Development Stage]

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


### 29.04.2026

Универсальное автоопределение метрик (buildCalculationBlocks) работает стабильно, поэтому старые ручные режимы расчета для конкретных типов данных (и жестко привязанные к ним парсеры) больше не нужны. Кодовая база очищена от неиспользуемых функций для упрощения дальнейшей поддержки.

Основные изменения по файлам:
- taskpane.js: удалены обработчики (runMeanSignificance..., runNpsSignificance...) и слушатели событий для кнопок явного расчета средних и NPS. Очищены неиспользуемые импорты.
- metric-detector.js: удален устаревший планировщик buildAutoCalculationPlan (полностью заменен на актуальный buildCalculationBlocks).
- significance.js: удалена легаси-функция из MVP v0.2 compareAllRowsUsingBottomBases, а также функции-обертки для отпавших ручных расчетов (compareMeansUsingSpreadAndBaseRows, compareNpsUsingStructureRows, compareNpsUsingSpreadAndBaseRows).

Примечание: математическое ядро расчетов значимости (z-тесты, t-тесты) и функционал диагностической кнопки оставлены без изменений.