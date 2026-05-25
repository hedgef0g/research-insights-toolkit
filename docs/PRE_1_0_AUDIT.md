# Pre-1.0 Technical Debt Baseline

Verified audit of the `main` branch as of 2026-05-25.  
Branch: `claude/kind-lumiere-1ec99e`.  
Do not modify source files based on this document alone — each section maps to a
proposed follow-up issue listed in §7.

---

## 1. Command Results

### Available npm scripts (verified from `package.json`)

| Script | Command |
|--------|---------|
| `build` | `webpack --mode production` |
| `build:dev` | `webpack --mode development` |
| `dev-server` | `webpack serve --mode development` |
| `lint` | `office-addin-lint check` |
| `lint:fix` | `office-addin-lint fix` |
| `prettier` | `office-addin-lint prettier` |
| `test` | `node --no-warnings --loader ./tests/core/loader.mjs --test tests/core/*.test.mjs` |
| `validate` / `validate:local` | `office-addin-manifest validate manifest.xml` |
| `start` / `start:local` | `office-addin-debugging start manifest.xml` |
| `stop` / `stop:local` | `office-addin-debugging stop manifest.xml` |
| `signin` / `signout` | `office-addin-dev-settings m365-account login/logout` |
| `watch` | `webpack --mode development --watch` |

There is **no `test:integration`** or `test:e2e` script. All automated tests are
unit-level, exercised in Node without Office.js.

### `npm test`

```
302 tests pass, 0 fail — 50 suites — 886 ms
```

All pass. No skipped or cancelled tests.

### `npm run build`

Exit 0. Three webpack performance warnings, no errors.

```
WARNING: asset size limit
  taskpane.js  379 KiB  (threshold 244 KiB)

WARNING: entrypoint size limit
  taskpane (379 KiB)

WARNING: webpack performance recommendations — consider import() or require.ensure
```

Minified outputs:

| Asset | Size |
|-------|------|
| `taskpane.js` | 379 KiB |
| `polyfill.js` | 229 KiB |
| `commands.js` | 367 bytes |
| `2c75a5d5…css` | 18.8 KiB |

### `npm run lint`

Exit 1. **411 problems (399 errors, 12 warnings)**.

| Rule | Count | Severity | Notes |
|------|-------|----------|-------|
| `prettier/prettier` | ~370 | error | Formatting only; 363 auto-fixable with `--fix` |
| `@typescript-eslint/no-unused-vars` | 21 | error | See §4 |
| `office-addins/no-navigational-load` | 7 | warning | `isNullObject` load on nav properties |
| `office-addins/load-object-before-read` | 7 | error | `bannerScanRange` and `selectedRange.worksheet` |
| `office-addins/call-sync-before-read` | 1 | error | `selectedRange.worksheet` at taskpane.js:3406 |

The bulk (≥90 %) are `prettier/prettier` formatting divergences with no semantic
impact. The remaining 29 issues carry substantive risk — see §3 and §5.

### `git diff --check`

Exit 0. No trailing-whitespace or CRLF conflicts in the working tree.

---

## 2. High-risk File Size Map

Files are grouped by AGENTS.md "high-risk" designation and by actual line count.

### High-risk files (per AGENTS.md)

| File | Lines |
|------|-------|
| `src/taskpane/taskpane.js` | **6,417** |
| `src/core/significance.js` | 2,102 |
| `src/core/banner-detector.js` | 1,445 |
| `src/core/metric-detector.js` | 658 |
| `src/core/excel-writer.js` | 374 |

`taskpane.js` defines **124 functions** and contains every Office.js interaction
path, all UI event wiring, settings persistence, run-report writing, content
sheet writing, banner-letter writing, and several development-only diagnostic
helpers. It has the highest change-risk of any file in the project.

### Other large runtime files

| File | Lines |
|------|-------|
| `src/core/table-preview-model.js` | 1,531 |
| `src/core/range-normalizer.js` | 1,177 |
| `src/taskpane/selected-range-interpreter.js` | 941 |
| `src/core/table-inventory-scanner.js` | 722 |
| `src/taskpane/localization.js` | 395 |
| `src/taskpane/taskpane.html` | 433 |
| `src/taskpane/active-cell-resolver.js` | 210 |
| `src/core/active-cell-table-finder.js` | 160 |
| `src/core/action-constants.js` | 126 |
| `src/core/config/dictionary.config.js` | 114 |
| `src/core/batch-candidate-filter.js` | 102 |
| `src/core/normalizers.js` | 81 |
| `src/core/stat-thresholds.js` | 80 |
| `src/core/string-utils.js` | 42 |
| `src/commands/commands.js` | 35 |

---

## 3. Lint / Build / Test Baseline

| Check | Result | Exit |
|-------|--------|------|
| `npm test` | 302/302 pass | 0 |
| `npm run build` | 0 errors, 3 perf warnings | 0 |
| `npm run lint` | 399 errors, 12 warnings | 1 |
| `git diff --check` | Clean | 0 |

### Lint breakdown by file (substantive issues only)

Prettier formatting errors are present in most `src/` files modified recently.
The non-formatting issues concentrate in `taskpane.js`:

**`src/taskpane/taskpane.js`**
- `call-sync-before-read` at line 3406: `selectedRange.worksheet` is accessed
  without a preceding `context.sync()` after the load call. This is an
  Office.js correctness violation; the worksheet reference may be stale.
- `load-object-before-read` at lines 5008, 5013, 5274–5276, 6344: all in the
  dead diagnostic code cluster (§4). Five of the seven load-before-read errors
  would be resolved by removing that dead code.
- `no-navigational-load` (7 warnings) at lines 1248, 1370, 3801, 3878, 4485,
  4824 in `taskpane.js` and line 86 in `active-cell-resolver.js`: loading
  `isNullObject` on navigation properties. This is a performance advisory; it
  does not block execution.
- `no-unused-vars` (11 of the 21 occurrences): `_selectionErr` ×4,
  `_` ×4, `error` ×1, `calculationSettings` ×1 (selected-range-interpreter.js),
  `Excel` ×1 (import in commands-adjacent file).

**`src/core/significance.js`**
- `no-unused-vars`: `applyTotalComparisonMarkerToFullMarkerMatrix` at line 1513.

**`src/core/banner-detector.js`**
- `no-unused-vars`: `normalizedGroupLabel` at line 282, `formatBannerDetectionDiagnostics` export at line 193 (imported into `taskpane.js` but never called).

**`src/core/table-inventory-scanner.js`**
- `no-unused-vars`: `LABEL_SCAN_COLUMNS_LEFT` at line 30.

**`src/core/metric-detector.js`**
- `no-unused-vars`: `options` parameter at line 909, `_` ×2 at lines 324 and 340.

---

## 4. Dead-Code Candidates

All candidates below are confirmed by the `no-unused-vars` lint pass or by
manual call-site search. None are reached by any currently wired user action.

### Cluster A — Banner diagnostic infrastructure (~450 lines)

| Symbol | File | Line | Evidence |
|--------|------|------|----------|
| `ENABLE_BANNER_SPAN_DIAGNOSTICS` | taskpane.js | 353 | Const, always `false`, never read |
| `loadBannerSpanDiagnosticsForSelectedRange` | taskpane.js | 4956 | Lint `no-unused-vars`; no call site |
| `loadBannerMergeDiagnosticsForSelectedRange` | taskpane.js | 5224 | Lint `no-unused-vars`; no call site |
| `reconstructHorizontalSpansFromRowText` | taskpane.js | 5057 | Called only from `loadBannerSpanDiagnostics…` |
| `finalizeDiagnosticSpanSelection` | taskpane.js | 5128 | Called only from `loadBannerSpanDiagnostics…` |
| `refineDiagnosticSpanRightBoundaryByLowerBannerRow` | taskpane.js | 5156 | Called only from `loadBannerSpanDiagnostics…` |
| `normalizeBannerDiagnosticCellText` | taskpane.js | 5208 | Called only from diagnostic cluster |
| `getCellMergeDiagnosticInfo` | taskpane.js | 5300 | Called only from `loadBannerMergeDiagnostics…` |
| `formatBannerDetectionDiagnostics` | banner-detector.js | 193 | Imported at taskpane.js:35 but lint confirms unused |

### Cluster B — Metric detection diagnostic button (~60 lines)

| Symbol | File | Line | Evidence |
|--------|------|------|----------|
| `runMetricDetectionDiagnostics` | taskpane.js | 3042 | Wired to `#detect-metric-type` at line 399–400, but that element is absent from `taskpane.html`; the `if (detectMetricTypeButton)` guard silently prevents the listener from ever registering |

### Isolated dead symbols

| Symbol | File | Line | Evidence |
|--------|------|------|----------|
| `applyTotalComparisonMarkerToFullMarkerMatrix` | significance.js | 1513 | Lint `no-unused-vars`; `applyTotalComparisonMarkerToFullCellResultMatrix` (line 1476) is called instead — likely a naming divergence during refactor |
| `LABEL_SCAN_COLUMNS_LEFT` | table-inventory-scanner.js | 30 | Lint `no-unused-vars`; const `2` never consumed |
| `normalizedGroupLabel` | banner-detector.js | 282 | Lint `no-unused-vars`; assigned, never read |

### `console.warn` / debug output

Three `console.warn` calls exist in `taskpane.js` at lines 6110, 6122, 6133 (settings load/save/clear error paths). These are intentional error-path logging, not development artifacts, and are not candidates for removal.

No `console.log`, `console.debug`, or `debugger` statements were found in `src/`.

### TODO / FIXME / HACK markers

Zero occurrences of `TODO`, `FIXME`, `HACK`, or `XXX` in `src/`. The absence
of these markers is consistent with in-line comment style — known limitations
are tracked in AGENTS.md, STATUS.md, and TABLE_STRUCTURE_MATRIX.md rather than
inline source annotations.

---

## 5. Performance-Risk Candidates

### `context.sync()` call density

| File | `context.sync()` calls |
|------|----------------------|
| `src/taskpane/taskpane.js` | **77** |
| `src/taskpane/selected-range-interpreter.js` | 7 |
| `src/taskpane/active-cell-resolver.js` | 4 |

Total across all source: **88** `context.sync()` calls. Most are correct and
necessary (Excel mutations require sync), but two functions show notable
density:

**`writeRunReportSheet`** (taskpane.js:1366–1439): 7–10 syncs per invocation
depending on whether the tier-2 sheet-positioning fallback fires. The pattern
is load → sync → mutate → sync → verify → sync (×2 tier-1 + up to 3 more
tier-2). This is intentional defensive code for cross-platform tab-order
reliability, documented in inline comments. The trade-off is latency on slow
Office hosts.

**`writeInventoryContentSheet`** (taskpane.js:4821–4896): identical 7–10 sync
pattern for the same reason. Same defensive justification.

These two functions together can consume up to 20 round-trips in worst-case
(tier-2 fallback for both). They fire only during Content / Run report sheet
creation, not on every Run.

### Office.js performance advisories

7 `no-navigational-load` warnings: all call `load("isNullObject")` on the
result of `getUsedRangeOrNullObject()`. The lint rule flags this as a
navigation-property load that slows the add-in. The pattern is used because
`isNullObject` is the intended way to check for null objects in Office.js
without throwing. An alternative is `try { range.address } catch {}` but that
has its own caveats. These are low-priority.

1 `call-sync-before-read` error at taskpane.js:3406: `selectedRange.worksheet`
is accessed without ensuring a sync after the range load. This is an active
code path (not dead code) and is a correctness risk, not just a performance
advisory. Worksheet access may return a stale or invalid reference.

### Bundle size

`taskpane.js` at 379 KiB (minified, production) exceeds the webpack threshold
(244 KiB) by 55%. Contributing factors:

- All application logic in one entry point (no code splitting)
- `@stdlib/stats-base-dists-*` quantile libraries: 848 KiB of source compiled
  into the bundle — these contribute disproportionately relative to their
  narrow use (threshold lookup only)

Since the add-in runs inside Excel (sideloaded or AppSource), network latency
is typically low. The performance impact is felt on initial load, not on
subsequent operations.

---

## 6. Documentation Gaps

Comparison of `docs/USER_GUIDE.md` content against the current `taskpane.html`
UI and `src/taskpane/localization.js` string keys.

### Covered in USER_GUIDE.md

- Расчёт / Run tab: Run, Clear, Check selection buttons ✅
- Автозапуск / Autorun tab: all three scopes ✅
- Проверка / Check tab: all three scopes ✅
- Оглавление / Contents tab ✅
- Run report sheet ✅
- Language switcher (RU/EN) ✅
- Known limitations section ✅
- Troubleshooting section ✅

### Not covered or partially covered

**Settings panel (20+ settings across 5 tabs)**  
USER_GUIDE.md makes no mention of the settings panel. The following settings
have no user documentation:

- Confidence level selector (99 / 95 / 90 / 80 / 66.6%)
- One-tailed test checkbox
- Preferred base selector (Auto / Effective / Unweighted / Base / Weighted)
- Compare with previous column + apply fill
- Compare only with Total / Exclude Total from comparisons (mutually exclusive)
- First column is Total / Total in each banner
- Write banner letters, Respect banner structure, Auto-detect wave banners
- Labels on left side of sheet
- Exclude small bases + threshold input
- Round cell values
- Fill colors: significant, lower-than-total, small-base
- Fill only total comparisons checkbox

**Settings storage and persistence**  
The storage mode selector (Не сохранять / Локально / В облаке [disabled]) and
Reset button in the utility footer are undocumented. The fact that cloud storage
is UI-visible but non-functional is not communicated to users.

**Content output format**  
The `content-output-mode` select (`minimal` / `client` / `full-check`) in the
Оглавление tab is not documented. Users have no guidance on when to use each
format.

**Run report checkbox scope**  
USER_GUIDE.md describes the Run report checkbox for Проверка but does not
mention the equivalent Расчёт > `run-add-report` checkbox.

**English-language help**  
The UI has full English localization (EN mode via language switcher), but
`USER_GUIDE.md` and the inline help file (`assets/rit-help-ru.html`) are
Russian-only. There is no English help resource.

**STATUS.md currency**  
`STATUS.md` does not reflect:
- Scope selector (Текущая таблица / Текущий лист / Вся книга) implementation
- Settings panel collapse feature
- Localization foundation (RU/EN)
- Content output mode feature

These were shipped in the most recent PRs but STATUS.md still describes the
"Next focus" as the check-table preview foundation, which is now implemented.

---

## 7. Proposed Follow-up Issues

Maximum 10 issues, ordered by risk/impact.

### Issue A — Fix `call-sync-before-read` at taskpane.js:3406

**Risk: Medium-high (correctness)**  
`selectedRange.worksheet` is read at line 3406 without a `context.sync()` after
the range was loaded. Confirmed by `office-addins/call-sync-before-read` lint
error. This is in an active code path (non-dead). Worksheet property access on
a stale proxy can return undefined or throw on some Office hosts.

Scope: `src/taskpane/taskpane.js` line 3406 only.

---

### Issue B — Remove banner diagnostic dead code cluster

**Risk: Low (dead code removal, ~450 lines)**  
Remove the following from `taskpane.js`:
- `loadBannerSpanDiagnosticsForSelectedRange` and its 5 helper functions
  (taskpane.js:4956–5339)
- `loadBannerMergeDiagnosticsForSelectedRange` and `getCellMergeDiagnosticInfo`
  (taskpane.js:5224–5339)
- `ENABLE_BANNER_SPAN_DIAGNOSTICS` const (taskpane.js:353)
- The `detect-metric-type` button stub and `runMetricDetectionDiagnostics`
  (taskpane.js:377–400, 3042–3102)

Remove from `banner-detector.js`:
- `formatBannerDetectionDiagnostics` export (banner-detector.js:193)

Remove from `taskpane.js` imports:
- `formatBannerDetectionDiagnostics` from the import at line 35

This would resolve 6 of 7 `load-object-before-read` lint errors and 2 of 21
`no-unused-vars` errors.

---

### Issue C — Resolve `applyTotalComparisonMarkerToFullMarkerMatrix` in significance.js

**Risk: Low (isolated dead function)**  
`applyTotalComparisonMarkerToFullMarkerMatrix` (significance.js:1513) is never
called. The neighboring `applyTotalComparisonMarkerToFullCellResultMatrix`
(line 1476) is called. Confirm whether the unused variant is a vestigial API
from a previous refactor, then either remove it or document why it is retained.

Do not touch the statistical logic in either function without an explicit issue.

---

### Issue D — Run `prettier --fix` pass

**Risk: Minimal (formatting only)**  
363 of the 399 lint errors are `prettier/prettier` formatting issues, all
auto-fixable. A single `npm run lint:fix` pass followed by `npm run build` and
`npm test` verification is the entire scope. No logic changes.

This unblocks future PRs from noisy diff output.

---

### Issue E — Remove isolated dead constants and unused parameters

**Risk: Low**  
- `LABEL_SCAN_COLUMNS_LEFT` (table-inventory-scanner.js:30) — remove const
- `normalizedGroupLabel` (banner-detector.js:282) — remove assignment
- `options` parameter (metric-detector.js:909) — remove from function signature
- `calculationSettings` parameter (selected-range-interpreter.js:525) — remove
  if confirmed unused throughout the function body

Each is a single-line removal. Verify no external callers before removing
function parameter signatures.

---

### Issue F — Document settings panel in USER_GUIDE.md

**Risk: None (docs only)**  
Add a "Settings" section to `USER_GUIDE.md` describing each of the 5 settings
tabs and their controls. Include:
- What each setting does
- Default values
- Which settings interact (e.g., Compare only with Total vs. Exclude Total are
  mutually exclusive; Banner letters requires Respect banner structure)
- That Cloud storage is not yet functional

---

### Issue G — Document storage persistence and Content output format

**Risk: None (docs only)**  
Two UI features with no user documentation:
1. The storage mode selector (footer) and its Local mode behavior
2. The Content output format selector (minimal / client / full-check)

These can be a single small PR appending to USER_GUIDE.md.

---

### Issue H — Create English user guide or bilingual USER_GUIDE.md

**Risk: None (docs only)**  
The UI ships with full RU/EN localization. There is no English help resource.
At minimum, add English section headers and key-phrase translations to
`USER_GUIDE.md`, or create a parallel `USER_GUIDE_EN.md`. The inline help
(`assets/rit-help-ru.html`) is out of scope here unless the help link in the
footer is also updated.

---

### Issue I — Acknowledge or suppress webpack bundle size warning

**Risk: Low**  
Add `performance: { hints: false }` to `webpack.config.js` (production mode
only) if the 379 KiB bundle is accepted as intentional for this deployment
model. Alternatively, document a code-splitting plan for a later milestone.

This suppresses noise in CI and `npm run build` output without changing the
bundle.

---

### Issue J — Review `no-navigational-load` warnings in taskpane.js

**Risk: Low (performance advisory)**  
7 occurrences of `load("isNullObject")` on navigation properties (
`getUsedRangeOrNullObject()` results). Per Office.js best practice, loading
`isNullObject` on a null-object return is a valid pattern, but the lint tool
flags it. Evaluate whether the `isNullObject` property is the correct approach
for each call site or whether the null-object pattern can be avoided upstream.

This is a performance advisory, not a correctness issue.

---

## Appendix: `context.sync` line index in `taskpane.js`

For reference when planning batch-sync consolidation:

| Lines | Count | Function / area |
|-------|-------|-----------------|
| 638–992 | 7 | `runSignificanceFromSelection` area |
| 1245–1249 | 1 | `ensureRunReportWorksheet` |
| 1371–1436 | 7–9 | `writeRunReportSheet` (2-tier sheet positioning) |
| 1552–1850 | 6 | `runAutoSignificance` area |
| 1912–2012 | 3 | `runAutoCurrentTable` area |
| 2202–2402 | 3 | Clear paths |
| 2810–2896 | 2 | `clearSignificanceFromSelection` |
| 2939–3108 | 5 | `clearBannerMarkersAboveRange` |
| 3224–3408 | 2 | Check paths |
| 3795–3909 | 6 | `collectWorkbookInventoryResults` area |
| 4313–4487 | 5 | Backlink writing |
| 4826–4893 | 7–9 | `writeInventoryContentSheet` (2-tier sheet positioning) |
| 4963–4995 | 2 | `runTableInventory` |
| 5229–5315 | 3 | Diagnostic cluster (dead) |
| 5589–5876 | 8 | `writeBannerMarkersAboveSelectedRange` area |
| 6339–6369 | 2 | `clearStaleBannerMarkersLeftOfWriteRange` |
