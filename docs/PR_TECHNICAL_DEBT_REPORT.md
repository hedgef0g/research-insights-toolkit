# PR Technical Debt Report Template

**How to use this template:**
Copy this file into your PR description or link to a completed report from the PR.
Fill in every section. If a section is genuinely empty, write "None" and explain why
(e.g. "No fragile dependencies -- this PR only adds a docs file").
Claiming "no technical debt" without justification is not acceptable.

**Reference documents:**
- docs/TABLE_STRUCTURE_MATRIX.md  -- coverage statuses, fragile areas, PARTIAL/FUTURE flags
- docs/GOLD_STANDARD_TEST_SUITE.md -- GST case IDs, validation expectations, agent workflow

---

## 1. PR Summary

    PR:           #<number>  (link)
    Issue:        #<number>  (link)
    Change type:  [ ] Feature  [ ] Bug fix  [ ] Refactor  [ ] Documentation  [ ] Process
    Branch:       <branch name>
    Affected files:
      - <file path>
      - ...
    Product areas:
      - <e.g. Proportion calculation, Banner detection, NPS, Excel writer, Taskpane UI>

---

## 2. Shortcuts Taken

List every deliberate simplification made to keep this PR scoped or to meet a deadline.
Include: heuristics used instead of robust logic, simplified assumptions, partial
support for a structure, fallback behavior, hardcoded patterns, deferred validation,
and skipped edge cases.

    Example:
    - Used first-matching Base row instead of implementing the full base priority
      ladder (Effective > Unweighted > plain > Weighted). Remaining ladder steps
      are captured in docs/TABLE_STRUCTURE_MATRIX.md section 4.
    - NPS ordering assumed Promoters-before-Detractors; arbitrary ordering is
      documented as PARTIAL in TABLE_STRUCTURE_MATRIX.md section 1.3.

If none: "None. Justify: <reason>"

---

## 3. Known Limitations

List structures, inputs, or scenarios that are NOT handled by this PR and are not
addressed by an existing follow-up issue. Include: unhandled table structures,
untested edge cases, platform gaps, user-facing limitations, unsupported inputs.

Cross-reference TABLE_STRUCTURE_MATRIX.md status where relevant
(PARTIAL / UNSUPPORTED / FUTURE).

    Example:
    - Three-or-more-level banners remain UNSUPPORTED (matrix section 2.1).
    - Effective Base row type not implemented; FUTURE in matrix section 4.
    - Google Sheets: not applicable to this PR; excluded per FUTURE status.

If none: "None. Justify: <reason>"

---

## 4. Fragile Areas

List any high-risk files touched or indirectly affected by this PR.
For each, state what could break and what regression checks cover it.

High-risk files per AGENTS.md:
  src/core/significance.js
  src/core/metric-detector.js
  src/core/banner-detector.js
  src/core/excel-writer.js
  src/taskpane/taskpane.js

    Example:
    - significance.js: pooled z-test formula unchanged; only block-formation input
      changed. Regression risk: incorrect base value passed to test.
      Covered by: GST-001, GST-002 manual smoke.
    - metric-detector.js: added keyword; existing keywords untouched.
      Regression risk: new keyword causes false positive on unlabeled rows.
      Covered by: GST-007 (out-of-range proportion case).

If none touched: "None. Justify: <reason>"

---

## 5. Validation Performed

State exactly what was run and what the result was.

### Build

    npm run build result: [ ] PASS  [ ] FAIL  [ ] Not run (reason: <...>)

### Manual smoke tests

List each smoke area checked from AGENTS.md. Mark PASS, SKIP (not affected), or FAIL.

    [ ] Basic proportions
    [ ] Total (first-column-is-Total)
    [ ] Local Total banner
    [ ] Global Total banner
    [ ] Banner letters
    [ ] Previous-column mode
    [ ] Wave auto off/on
    [ ] Small bases
    [ ] Means / SD / variance
    [ ] NPS
    [ ] Run -> Clear significance
    [ ] Numeric output and display conventions

### Gold Standard Test Suite cases checked

List GST case IDs verified against this PR. If no expected-output files exist yet,
note "PLANNED -- manual check only."

    Example:
    - GST-001: PASS (manual)
    - GST-020: PLANNED -- manual check only
    - GST-040: not affected

If no GST cases are relevant: "None affected by this PR."

---

## 6. Follow-up Issues

List any TODOs, design questions, missing tests, deferred cases, or documentation
gaps created or uncovered by this PR. Each item should become a GitHub issue.

    Example:
    - [ ] File issue: add data-quality warning for out-of-range proportion values
          (matrix section 1.1, 6 -- currently PARTIAL, no guard).
    - [ ] File issue: harden Total-outside-selection edge case
          (matrix section 2.2 -- PARTIAL, noted in TABLE_STRUCTURE_MATRIX.md notes).
    - [ ] File issue: GST-062 workbook -- create Weighted Base test fixture once
          weighted-base path is implemented.

If none: "None. Justify: <reason>"

---

*Template version: 2026-05-07. Maintained in docs/PR_TECHNICAL_DEBT_REPORT.md.*
*Refer to AGENTS.md for agent workflow rules and validation requirements.*
