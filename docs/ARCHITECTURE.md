# Architecture Overview

## Primary Architecture

- `taskpane/taskpane.js` — Excel/UI controller;
- `core/metric-detector.js` — metric row detection and block-plan construction;
- `core/dictionary.config.js` — config-driven metric dictionary;
- `core/normalizers.js` — normalization of Excel values;
- `core/significance.js` — statistical tests, comparison routing, markers, fill reasons;
- `core/stat-thresholds.js` — statistical critical thresholds;
- `core/banner-detector.js` — platform-independent banner structure detection;
- `core/excel-writer.js` — writing values and formatting back to Excel.

## Architectural Constraints

- No statistical calculations in `taskpane.js` (UI orchestration only).
- No Excel range writing in `significance.js`.
- No UI DOM access in core modules.

## Core Development Rules

- Do not break the manual selected-range workflow.
- Do not rewrite `significance.js` wholesale.
- Do not return `excel-writer.js` to per-cell writes.
- Keep changes small and isolated.
- Prefer one issue = one PR.
- Do not mix refactor and feature work in the same PR.
- Preserve current UX status-message rules.
- If broader changes seem needed, stop and describe them instead of implementing.

## High-Risk Files

- `core/significance.js`
- `core/banner-detector.js`
- `core/excel-writer.js`
- `taskpane/taskpane.js`
