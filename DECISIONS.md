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