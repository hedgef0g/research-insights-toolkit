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