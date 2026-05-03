# Research Insights Toolkit — Agent Instructions

## Purpose

This file contains default instructions for AI agents working on Research Insights Toolkit.

Use this file to avoid repeating long project context in every task prompt.

Task-specific instructions, GitHub Issues, and direct user instructions override this file.

## Agent roles

- ChatGPT: Product Lead, task architect, roadmap prioritization, issue drafting, PR review, merge/no-merge recommendation.
- Claude Code: primary coding agent.
- Codex: coding/review fallback and regression-focused reviewer.
- Jules: documentation and process PRs.
- Human owner: final merge decision and Excel smoke testing.

Agents must not decide that a task is complete without human review.

## Product context

Research Insights Toolkit is an Excel-first Office Add-in for statistical significance in research tables.

The current stable workflow is manual selected-range:

- user manually selects the numeric data range in Excel;
- user clicks Run;
- add-in reads the selected range;
- detects row types;
- detects banner structure if enabled;
- builds calculation blocks;
- calculates significance;
- writes markers/fills back to Excel.

Do not break this workflow.

Automatic worksheet/workbook scanning is not implemented and should not be added unless explicitly requested.

## Core product rules

- The user is expected to select the numeric data area.
- Row labels are read from cells to the left of the selection.
- Banner/header rows are read from rows above the selection.
- Current Run behavior must remain stable unless the task explicitly changes it.
- Preview/check-table features must remain read-only unless explicitly wired.
- Warning-only guardrails must not block Run unless explicitly requested.
- Do not auto-trim selected ranges unless explicitly requested.

## High-risk files

Treat these files as high risk:

- `src/core/significance.js`
- `src/core/metric-detector.js`
- `src/core/banner-detector.js`
- `src/core/excel-writer.js`
- `src/taskpane/taskpane.js`

Only modify high-risk files when the task or issue explicitly allows it.

## Strict architecture rules

- Keep core modules free of Office.js dependencies.
- Do not mix refactor and feature work in one PR.
- Do not rewrite `significance.js` wholesale.
- Do not change statistical calculation logic unless explicitly requested.
- Do not change Excel writer behavior unless explicitly requested.
- Do not change taskpane UI unless explicitly requested.
- Do not add worksheet/workbook auto-scan unless explicitly requested.
- Do not implement row-wise comparisons unless explicitly requested.
- Do not implement weighted base calculation support unless explicitly requested.
- Unexpected improvements should become follow-up issues, not extra changes.

## Excel writer and numeric output rules

- Cells with appended significance markers may be text, for example `21% b`.
- Cells without marker text should preserve numeric Excel values where possible.
- Preserve visible display conventions:
  - `28` should remain plain numeric display;
  - `28%` should remain percent display;
  - `0.28` should remain decimal-share display.
- Do not convert all proportion values to one display scale.
- Do not convert the whole selected range to text.
- Clear significance should restore numeric-looking cleaned values where possible.
- Do not change `src/core/excel-writer.js` without an explicit writer/output issue.

## Banner rules

- Banner-aware comparisons must preserve the manual selected-range model.
- Multi-row and merged-like banners are supported through banner detection and taskpane banner-letter placement logic.
- Banner letters are separate from data-cell markers.
- Do not change banner detection, banner-letter writing, or comparison pair logic unless the task explicitly allows it.

## NPS expectations

NPS-first:

- NPS row receives NPS significance markers.
- Promoters / Detractors receive ordinary proportion markers.
- Base receives no marker.

NPS-first with Neutral:

- NPS receives NPS significance markers.
- Promoters / Neutral / Detractors receive ordinary proportion markers.
- Base receives no marker.

Extended NPS:

- Scale rows, buckets, and support rows receive ordinary proportion markers.
- NPS receives NPS significance markers.
- Base receives no marker.

NPS + SD/Base or variance/Base:

- NPS receives NPS spread marker.
- SD/variance and Base receive no markers.

## Git and PR rules

Unless explicitly instructed otherwise:

- Do not push to `main` or `release`.
- Do not merge PRs.
- Keep one issue → one agent → one branch → one PR.
- Keep PRs small and scoped.
- Commit and push only if the task explicitly allows it.
- Opening a PR is allowed only if the task explicitly allows it.
- Always report:
  - PR link;
  - files changed;
  - build/test result;
  - assumptions or limitations.

If working in an agent-created worktree, report the path and branch.

For investigation-only tasks:

- Do not modify files.
- Do not commit.
- Do not push.
- A clean agent worktree is acceptable if it is based on latest `main`.

## Validation

For code changes, run:

    npm run build

If PowerShell blocks `npm.ps1`, use:

    npm.cmd run build

If build cannot run, explain why.

For behavior-sensitive changes, include relevant manual smoke notes.

Common smoke areas:

- basic proportions;
- Total;
- local Total banner;
- global Total banner;
- banner letters;
- previous-column mode;
- wave auto off/on;
- small bases;
- means / SD / variance;
- NPS;
- Run → Clear significance;
- numeric output and display conventions.

## Task modes

### Investigation mode

Use when the task says investigate, diagnose, review, or spec.

Rules:

- Do not modify files.
- Do not propose broad rewrites.
- Identify root cause and minimal safe fix.
- Prefer exact files/functions.

Output:

    ## Diagnosis
    ## Suspected root cause
    ## Evidence from code
    ## Minimal fix proposal
    ## Regression risks
    ## Suggested follow-up issue

### Coding mode

Use when the task explicitly allows code changes.

Rules:

- Follow issue scope.
- Modify only allowed files.
- Keep patch minimal.
- Do not add unrelated cleanup.
- Run build.
- Open PR only if explicitly allowed.

Output:

    ## Files changed
    ## Diff summary
    ## Build/test result
    ## PR
    ## Assumptions / limitations

### Documentation mode

Use for docs/help/process tasks.

Rules:

- Do not modify runtime code.
- Do not update docs for behavior that is not merged and tested.
- Keep docs user-facing and accurate.
- Avoid documenting speculative roadmap as current behavior.

Output:

    ## Files changed
    ## Summary
    ## PR
    ## Assumptions / limitations

## Default prompt contract

A short task prompt may rely on this file.

Example:

    Use AGENTS.md.
    Implement issue #36.
    You may commit, push, and open a PR.
    Do not merge.
    Report PR link, files changed, build result, and assumptions.

If the prompt conflicts with this file, follow the prompt.