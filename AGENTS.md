# AI Agent Instructions for Research Insights Toolkit

## 1. Project purpose
Research Insights Toolkit is an Excel-first Office Add-in for research tables and statistical significance.

## 2. Current stable workflow
The current stable workflow is manual selected-range calculation: the user selects a table range manually and clicks Run.

## 3. Core development rules
- Do not break the manual selected-range workflow.
- Do not rewrite significance.js wholesale.
- Do not return excel-writer.js to per-cell writes.
- Keep changes small and isolated.
- Prefer one issue = one PR.
- Do not mix refactor and feature work in the same PR.
- Preserve current UX status-message rules.
- If broader changes seem needed, stop and describe them instead of implementing.

## 4. High-risk files
- core/significance.js
- core/banner-detector.js
- core/excel-writer.js
- taskpane/taskpane.js

## 5. Smoke-test checklist
Mention these scenarios:
- basic proportions
- Total
- local Total banner
- global Total banner
- previous-column
- wave auto off/on
- small bases
- means/NPS

## 6. Current recommended safe task
shared text normalization utilities.
