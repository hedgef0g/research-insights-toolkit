# Architecture Decision Records

This file records significant architecture decisions for Research Insights Toolkit.
Each entry names the decision, its context, and the reasoning so future contributors
understand why the code is shaped the way it is.

---

## ADR-001 — Keep selected-range interpretation in the taskpane adapter, not in core

**Date:** 2026-05-17
**Status:** Active
**Related PR:** #120, #122

### Context

PR #120 extracted shared selected-range interpretation for Run and Check into
`src/taskpane/selected-range-interpreter.js`. Before that PR, both Run and Check
each loaded labels, banner rows, and write-target ranges independently, leading
to subtle drift between what Run calculated and what Check previewed.

### Decision

`selected-range-interpreter.js` is an **Excel/taskpane adapter**, not a core
module. It lives in `src/taskpane/` and is allowed to depend on Office.js
context and Range objects. It is not placed in `src/core/`.

### Rationale

The adapter still needs Office.js to load left-label values and banner rows from
worksheet cells adjacent to the selection. Separating that I/O into a pure-core
module would require passing fully loaded arrays in, which means the caller (taskpane)
would have to know which rows and columns to pre-load — reintroducing the coupling
the extraction was meant to remove.

The right future decomposition is:

1. Extract a pure platform-neutral function that accepts pre-loaded arrays and
   returns normalization, embedded-label detection, and column-stripping results.
2. Keep a thin Excel adapter that performs the Office.js I/O and then calls (1).

That refactor should happen only after the adapter contract is stable and the
invariants are covered by tests.

### Invariants this decision protects

- `valuesForCalculation` width equals `writeTargetRange.columnCount`.
- `valuesForCalculation` width equals `bannerContext.selectedColumnCount` when banner context is present.
- Label, unit, and header columns are excluded from both `valuesForCalculation` and `writeTargetRange`.
- Run and Check consume the same interpreted values, labels, and banner context — they diverge only in what they do after interpretation.

### Implications for future platform ports

**Google Sheets:** The adapter pattern makes a Sheets port tractable. A
`sheets-range-interpreter.js` would implement the same `interpretSelectedRange()`
contract, loading adjacent cells via the Sheets API instead of Office.js. Core
detection and significance modules would not need changes because they consume
the platform-neutral interpretation object.

**VBA / COM add-in:** A VBA bridge would produce the same interpretation object
shape (as a JSON payload, for example) and hand it to the same core pipeline.
The adapter boundary is the porting seam.

Duplicating selection interpretation logic across platforms should be treated as
a defect. New Run/Check/Auto Runner work should route through the adapter rather
than reload labels and banner rows independently.
