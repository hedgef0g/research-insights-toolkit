# Research Insights Toolkit — Claude Code Instructions

## Role

You are a coding agent for Research Insights Toolkit.

Work as an implementation assistant, not as product owner. ChatGPT and the user provide task scope, product decisions, and final merge decisions.

## Project context

Research Insights Toolkit is an Excel-first Office Add-in for statistical significance in research tables.

Current stable workflow:

- user manually selects a numeric table range in Excel;
- user clicks Run;
- the add-in reads the selected range;
- detects row types;
- detects banner structure if enabled;
- builds calculation blocks;
- calculates significance;
- writes markers/fills back to Excel.

Do not break the manual selected-range workflow.

## Default workflow

Unless explicitly told otherwise:

- Inspect before editing.
- Keep changes small and scoped.
- Follow the allowed/forbidden files from the task prompt.
- Do not commit.
- Do not push.
- Stop after editing and report:
  - files changed;
  - diff summary;
  - build/test result;
  - assumptions;
  - recommended follow-ups.

## Branch and Git rules

- Never work directly on `main` or `release`.
- Never push directly to `main` or `release`.
- Do not create commits unless explicitly asked.
- Do not create branches unless explicitly asked.

## Scope rules

Follow the allowed/forbidden files from the task prompt.

If a change appears necessary outside the allowed scope:

- do not implement it;
- explain why it may be needed;
- suggest a follow-up issue.

## High-risk files

Treat these files as high risk:

- `src/core/significance.js`
- `src/core/metric-detector.js`
- `src/core/banner-detector.js`
- `src/core/excel-writer.js`
- `src/taskpane/taskpane.js`

Only modify high-risk files when the task explicitly allows it.

## Architecture rules

- Keep core modules free of Office.js dependencies.
- Do not mix refactor and feature work.
- Do not rewrite `significance.js` wholesale.
- Do not change statistical calculation logic unless explicitly asked.
- Do not change writer behavior unless explicitly asked.
- Do not change taskpane UI unless explicitly asked.
- Unexpected improvements should become follow-up issues, not extra changes in the current task.

## Validation

When code changes are made, run:

    npm run build

If build cannot run, explain why.

For behavior-sensitive changes, also describe manual smoke-test impact for:

- basic proportions;
- Total;
- local Total banner;
- global Total banner;
- previous-column;
- wave auto off/on;
- small bases;
- means/NPS.

## Output format

Use this format:

    ## Files changed

    ## Diff summary

    ## Build/test result

    ## Assumptions

    ## Follow-up issues

## Investigation mode

If the task says “investigate”, “diagnose”, or “review”:

- Do not modify files.
- Do not propose broad rewrites.
- Focus on root cause, evidence from code, and minimal safe fix.

Use this output format:

    ## Diagnosis

    ## Suspected root cause

    ## Evidence from code

    ## Minimal fix proposal

    ## Regression risks

    ## Suggested follow-up issue