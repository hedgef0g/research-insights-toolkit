# STATUS.md

Technical health and status report for Research Insights Toolkit.

For the user-facing feature overview, see [../README.md](../README.md).
For structural coverage (supported / partial / unsupported table shapes), see
[TABLE_STRUCTURE_MATRIX.md](TABLE_STRUCTURE_MATRIX.md).
For the manual smoke validation checklist, see [TEST_CASES.md](TEST_CASES.md).

## Current phase

Phase 1 — manual workflow stabilization — is complete. The manual
selected-range workflow is stable: Run and Clear both go through the shared
selected-range normalization path with banner-aware and wave-aware behavior.

Recent stabilization work landed in PRs #93, #94, #96, #97, #98, #100, #101.

## Next focus

Check table / preview foundation: expose the normalized interpretation model
to the user before Excel mutation, reusing the existing selected-range
normalization output rather than introducing a separate parsing path. See
[ROADMAP.md](ROADMAP.md) and [SELECTED_RANGE_NORMALIZATION.md](SELECTED_RANGE_NORMALIZATION.md).

## Engine health

- Selected-range normalization is the shared entry point for Run and Clear.
- Banner detection, banner-letter writing, and wave-aware behavior are
  considered stable for the structures listed as supported in
  [TABLE_STRUCTURE_MATRIX.md](TABLE_STRUCTURE_MATRIX.md).
- Numeric output preservation and Clear-significance numeric restoration are
  in place. Cells without marker text remain numeric where possible; display
  conventions (`28`, `28%`, `0.28`) are preserved.
- Statistical engine (pooled z-test for proportions, Welch's t-test for
  means, NPS structure, NPS spread) is unchanged and remains under project
  control; external libraries are used only for threshold quantiles.

## Validation surface

- [TEST_CASES.md](TEST_CASES.md) — manual smoke checklist, refreshed in #101.
- [TABLE_STRUCTURE_MATRIX.md](TABLE_STRUCTURE_MATRIX.md) — structure coverage source of truth.
- [GOLD_STANDARD_TEST_SUITE.md](GOLD_STANDARD_TEST_SUITE.md) — validation planning source for non-trivial changes.

## Known limitations

- Google Sheets support is not implemented.
- Cloud settings storage is not implemented.
- Multi-level banner support is implemented for supported research-table shapes, but broader arbitrary header layouts remain out of scope.
- Report-title detection and broader table-boundary detection are not fully implemented.
- Total outside selection is specified but may need additional edge-case hardening.
- Some development-only diagnostic helpers remain in `taskpane.js` behind no active user-facing call path.
- Runtime implementation remains Excel-first; new core behavior should still be designed platform-neutral where possible.
