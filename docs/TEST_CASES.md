# Research Insights Toolkit — Manual Test Cases

Last updated: 2026-05-12

## Purpose

Manual Excel smoke and regression checklist for RIT. Use these cases to validate Run, Clear, banner-aware behavior, and recent stabilization fixes before release.

This file is the practical checklist. For structural coverage see `TABLE_STRUCTURE_MATRIX.md`, and for the planned validation suite see `GOLD_STANDARD_TEST_SUITE.md`.

## General assumptions

Unless a case states otherwise:

- User selects only the numeric data area, labels are immediately to the left.
- Confidence level is 95%, two-tailed.
- Small-base threshold is 50.
- Old markers and fills are cleared before each Run.
- `respect-banner-structure` is off unless the case turns it on.

---

# A. Release-like smoke pack

Short pack to run before any release. Covers Run, Clear, banner, NPS, previous-column, and output formatting in 10–12 minutes.

| # | Case | Expected |
|---|---|---|
| A1 | All-vs-all proportions, two columns, 95% | Higher cell gets the other column's letter; Base unmarked; normal fill applied. |
| A2 | Mean + SD + Base, three columns | Welch t-test; markers only on Mean row. |
| A3 | NPS-first (NPS / Promoters / Detractors / Base) | NPS gets NPS markers; Promoters / Detractors get ordinary markers; Base unmarked. |
| A4 | Extended NPS (scale 1–10 + buckets + NPS + Base), full-table selection | NPS gets NPS markers; scale and bucket rows get ordinary markers; Base unmarked. |
| A5 | First column is Total, segments compared with Total | `T` / `t` on segments; Total column never receives markers. |
| A6 | Previous-column mode, three waves | `↑` / `↓` written only into right/current column; no banner letters. |
| A7 | Banner letters on, one-level banner | `(a)`, `(b)`, ... in banner row; Total banner cell unmarked when first-column-is-Total is on. |
| A8 | Wave auto previous-column (banner with named group + Wave group) | Named group uses ordinary letters; Wave group uses arrows automatically; status mentions auto wave mode. |
| A9 | Full-table Run on a table already containing markers (repeat Run) | Stale markers cleared first; new markers applied only to data body; banner labels like `Wave (quarter)` are preserved. |
| A10 | Run → Clear significance on a calculated full table | All RIT markers, bold, and fill removed from data body and banner/header cells; ordinary parenthesized labels preserved. |
| A11 | Numeric output: `28`, `28%`, `0.28` in mixed cells | Display conventions preserved; only marker-bearing cells become text; unmarked cells keep their numeric value and format. |
| A12 | Broad multi-table selection (full-table Run) | Run is blocked with a status message; no writes occur. |

Pass criteria: every row above passes without manual cleanup between cases.

---

# B. Basic proportions and means

## B1. All-vs-all proportions

Input: `| 50% | 70% | 51% |`, base 500 across.

Expected: A vs B, A vs C, B vs C all computed. Higher significant cell receives the lower column's letter. Base row unmarked. Significant cells get normal fill.

## B2. Re-run cleanup

Run B1, then change confidence or values so previous differences are no longer significant. Re-run.

Expected: Old markers, fills, and bold removed. New output reflects only the current run.

## B3. Mean + SD + Base

Input: Mean 3.5 / 4.2 / 3.6, SD 1.1 / 1.0 / 1.1, Base 300 each.

Expected: Welch t-test used. Markers only on Mean row. SD and Base unmarked.

---

# C. NPS

## C1. NPS-first

Rows: NPS, Promoters, Detractors, Base.

Expected: NPS recalculated from Promoters − Detractors and treated as mean of {+1, 0, −1}. NPS row gets NPS markers. Promoters and Detractors get ordinary proportion markers. Base unmarked.

## C2. NPS-first with Neutral

Rows: NPS, Promoters, Neutral, Detractors, Base.

Expected: Same as C1; Neutral also gets ordinary proportion markers.

## C3. Extended NPS

Rows: scale 1–10 rows, buckets (e.g. Top-3, Bottom-3), Promoters, Neutral, Detractors, NPS, Base.

Expected:
- Scale rows, buckets, and support rows get ordinary proportion markers.
- NPS row gets NPS markers.
- Base unmarked.

Regression note (PR `1198b36`): when the selected range is a full table that already contains markers from a prior Run, the extended NPS label fallback must still recognize 1..10 scale rows with NPS support/Base labels in the label column.

## C4. NPS + SD/Base and NPS + variance/Base

Expected: NPS spread logic used. NPS row gets NPS spread markers. SD or Variance and Base unmarked.

---

# D. Total comparison

## D1. First-column-is-Total indexing

Input: Total, Segment 1–3.

Expected: Total receives no letter index. Segment 1 = `a`, Segment 2 = `b`, Segment 3 = `c`. Banner markers (when banner letters are on) appear only on segment banner cells, never on the Total banner cell.

## D2. Total comparison markers

Settings: first-column-is-Total on, no other Total toggles.

Expected: Segment higher than Total → `T`; lower → `t`. Total marker comes first when combined with segment letter: `Ta`, `tb`. Total column never receives markers.

## D3. Compare only with Total

Expected: Only segment-vs-Total pairs. Output is only `T` / `t`. No ordinary letter markers.

## D4. Compare only with Total without Total location

Expected: Run stops. Status: user must specify Total location.

## D5. Exclude Total from comparisons

Expected: Only segment-vs-segment pairs. Total column has no marker; no `T`/`t`.

## D6. Lower-than-Total fill

Expected: Cells below Total receive lower-than-Total fill color; cells above Total receive normal fill.

## D7. Fill only for Total comparisons

Expected: Normal fill applied only to cells with `T`. Cells with only segment letters keep their letter and bold but no normal fill. Lower-than-Total fill still applies to `t`.

---

# E. Small bases

## E1. Small base in a proportion block

Segment with base 30 (threshold 50) is excluded from comparisons in the block, receives small-base fill in % and Base rows, no markers. Total and other segments still compared normally.

## E2. Small base in mean block

Same behavior in Mean, SD, Base rows of the mean block.

## E3. Small base in NPS structure block

Same behavior in NPS, Promoters, Detractors, Base rows of the NPS block.

## E4. Small base in Total column

Expected: Run stops. Status warns about Total base. Note: prior markers may already be cleared before the warning fires — acceptable.

## E5. Previous-column mode + small base

Excluded small-base column is not skipped over: pairs that touch the excluded column are simply not generated.

---

# F. Previous-column mode

## F1. Basic previous-column

Settings: `compare-with-previous-column = true`.

Expected: Each column (except the first) is compared with the one immediately to its left. Arrows written only into the right column: `↑` if higher, `↓` if lower. No banner letters. No fill unless `apply-previous-column-fill` is on.

## F2. Previous-column fill

`apply-previous-column-fill = true`: `↑` cells get normal fill, `↓` cells get lower-than-Total fill color. Small-base fill still wins.

## F3. First column is Total + previous-column

Without exclude-Total: Total is treated as ordinary previous column; segment 1 compared against it; status warns. With exclude-Total: Total is skipped, chain starts at segment 2 vs segment 1.

## F4. Mutually exclusive settings

- `compare-only-with-total` is disabled while previous-column mode is active.
- `write-banner-letters` is disabled while previous-column mode is active.
- `fill-only-total-comparisons` is disabled while previous-column mode is active.
- `exclude-total-from-comparisons` is disabled unless first-column-is-Total is on.

Defensive validation must still stop Run if any invalid combination is reached.

---

# G. Banner detection and banner-aware comparisons

## G1. One-level banner, fallback group

Settings: `respect-banner-structure = true` on a single-level banner.

Expected: Calculation works using fallback group. Status: `Расчёт выполнен. Обработано блоков: 1.` No technical banner diagnostics.

## G2. Repeated-label two-level banner

Banner upper row repeats group label (`Gender Gender Gender Age Age Age`); lower row has labels and Totals.

Expected: Two groups detected (Gender, Age). No cross-group comparisons. Marker indexing is group-local: in each group Male/Female and 18–24/25–34 each start from `a`.

## G3. Merged-like reconstructed span

Upper row is sparse (`Age` only in first cell of group span). Expected: Reconstructed span detects Age group covering all its columns. No technical span diagnostics.

## G4. Local Total inside a group

Group has Total + segments. Expected: Local Total is the reference for its group. Local Total receives no ordinary letter and is excluded from ordinary group comparisons. `T`/`t` is used for local Total comparisons.

## G5. Group without Total

Group has only segments. Expected: Ordinary group comparisons work. No Total comparisons generated. Not an error.

## G6. Multiple Totals in one group

Banner: `Total Male Total Female` under one group.

Expected: Run stops. Status: multiple Totals found in one group. No technical dump.

## G7. Compare-only-with-Total but no Total found

Expected: No pairs generated. Status reports that compare-only-with-Total is on but no Total exists.

## G8. Compare-only-with-Total / exclude Total under banner structure

`compare-only-with-total = true` → only local-Total pairs, only `T`/`t`. `exclude-total-from-comparisons = true` → all detected Totals excluded; no `T`/`t`; only intra-group segment comparisons.

---

# H. Banner-aware Total detection (PR #98)

Recent regression-prone cases:

## H1. Sparse upper local Total

Upper row in a group has `Total` (or `Всего`) only in the first cell of the group's span, lower row label is blank for that cell. Expected: sparse upper label is detected as the local Total reference; sibling columns compare against it; local Total cell receives no banner letter.

## H2. Standalone first-column `Всего` promoted to global Total

The first banner column is a standalone `Total` / `Всего` with no parent group; the rest of the banner has named groups (Gender, Age, etc.).

Expected:
- That column is promoted to global Total.
- Global Total is the only Total reference.
- Local Totals in other groups (if any) are compared with global Total like ordinary target columns and may receive `T`/`t`.
- Global Total banner cell receives no letter.

## H3. Repeated local Totals stay local

Two named groups, each with its own `Total` column. Expected: both Totals remain local (one per group). Neither is promoted to a global Total.

## H4. Non-Total guard

A banner column labeled `Total spend` or `Total awareness` (compound label) must not be detected as a Total. Expected: column is treated as an ordinary segment.

---

# I. Wave auto previous-column and nested wave (PR #100)

## I1. Mixed named group + wave group (auto wave on)

Banner upper: `Gender Gender Wave Wave Wave`; lower: `Male Female W1 W2 W3`. Settings: `respect-banner-structure = true`, `compare-with-previous-column = false`, `auto-detect-wave-banners = true`.

Expected:
- Gender uses ordinary group comparisons (`a` / `b`).
- Wave group uses previous-column automatically: W2 vs W1, W3 vs W2.
- Wave arrows are filled even if the previous-column fill checkbox is off.
- UI previous-column checkbox stays off.
- Status mentions auto wave mode.
- Banner letters are written only for Gender; wave columns receive no letters.

## I2. Non-wave group does not auto-switch

Banner: `Gender Gender Gender / Male Female Other`. Expected: ordinary group comparisons. No arrows. No auto-wave status message.

## I3. Nested wave inside named groups

Banner upper row contains category groups; nested below each category is a wave dimension such as `Волна (квартал)` with `2025Q4` / `2026Q1` (or similar) repeated per category.

Expected (auto wave on):
- Nested wave dimension is detected inside each category group.
- Within every category, the quarters are compared previous-column (Q4 → Q1 → ...).
- No banner significance letters are written in auto-wave mode.
- No cross-category comparisons.

## I4. Multi-column `Всего` wave group is not promoted to global Total

A wave group whose label happens to be `Всего` with multiple columns must remain a local group, not a single global Total.

## I5. Auto wave off — ordinary banner-aware behavior preserved

`auto-detect-wave-banners = false`: I1 / I3 fall back to ordinary banner-aware comparisons (letters, no arrows).

## I6. Manual previous-column overrides mixed mode

`compare-with-previous-column = true` (manual) with a mixed banner: previous-column applied inside all groups (Female vs Male, W2 vs W1, W3 vs W2). Auto-wave status message not required.

## I7. Compare-only-with-Total suppresses auto wave

`respect-banner-structure = true`, `compare-only-with-total = true` over a Wave banner that contains a Total column: only Total comparisons, no auto-wave arrows, no auto-wave status message.

## I8. Wave descriptor row below semantic group row

Banner has a real semantic group row (e.g. category names) above and a technical wave/quarter descriptor row (`2025Q4`, `2026Q1`, ...) immediately below. Expected: the higher semantic row is used as the meaningful group level; banner letters are not written across the wave descriptor row.

---

# J. Banner letters

## J1. Banner letters without Total

One-level banner, `write-banner-letters = true`. Banner cells receive `(a)`, `(b)`, `(c)`. Pre-existing markers at the end of the cell are replaced or preserved per current marker update rules.

## J2. Banner letters with first column as Total

Total banner cell receives no marker; if it previously had a RIT marker, it is removed. Segments receive `(a)`, `(b)`, ...

## J3. Banner first row

Selected range starts in worksheet row 1 with banner letters on. Expected: Run stops. Status asks user to add a row above the selection. No automatic row insertion.

## J4. Banner-local letters under banner structure

`respect-banner-structure = true`, two groups (Gender, Age), each with Total + segments. Expected lower banner row: `Total / Male (a) / Female (b) / Total / 18–24 (a) / 25–34 (b)`. Upper row unchanged. Old trailing markers on Total columns removed.

## J5. Banner letters with global Total

Global Total column receives no letter. Local Totals receive no letters. Segment labels are group-local.

## J6. Sparse / vertically merged banner cells, respect-structure OFF (PR #94)

`respect-banner-structure = false`, `write-banner-letters = true`, full-table Run, banner labels live in sparse or vertically merged rows above the data body.

Expected:
- Writer resolves the visible label cell upward (sparse / merged-like target).
- RIT banner markers are appended / replaced / removed on the resolved target.
- Ordinary parenthesized text such as `Wave (quarter)`, `Волна (квартал)`, `Brand (new)` is preserved.
- On repeat Run with first-column-is-Total, stale markers on the visible Total label are cleared.

## J7. Parenthesized label preservation

Banner labels like `Wave (quarter)`, `Волна (квартал)`, `Brand (new)`, `Всё покупаю сам(а)` survive both Run (banner letters on) and Clear without being mistaken for RIT markers.

Known limitation: a user-authored banner header ending with a single-letter parenthesized token that exactly matches a RIT significance label, e.g. `Foo (a)`, may still be treated as a RIT banner marker by Clear.

---

# K. Run normalization (full-table)

## K1. Strict numeric-only selection (legacy path)

User selects only the numeric data area. Expected: behavior unchanged from prior releases; no normalization needed.

## K2. Full-table selection narrows to data body

User selects the full table including labels and banner rows. Expected: Run normalizes the selection to the detected data body before cleanup, calculation, and writes. Labels and banner are used for context only.

## K3. Repeat full-table Run on a table with existing markers (PR #93)

Full-table selection on a table whose cells already contain RIT markers from a prior Run.

Expected:
- Numeric-looking cells with trailing significance markers are still detected as numeric.
- Labels like `BASE`, `Bottom-3`, `Wave (quarter)` are preserved.
- Extended NPS label-column fallback recognizes 1..10 scale rows with NPS support/Base labels.
- Output is correct on the second and subsequent Runs.

## K4. Broad / multi-table selection

User selects a range spanning multiple tables, blank gaps, or unrelated content.

Expected: Run is blocked before any Excel mutation. Status explains why. No clearing, no writes.

---

# L. Clear significance

## L1. Strict selected-range Clear (legacy)

Run, then Clear with the same strict numeric selection.

Expected: All RIT markers, bold, and fill are removed from the selected range. Numeric values and formats are restored where possible.

## L2. Full-table Clear normalizes to data body (PR #96)

Full-table selection, then Clear.

Expected:
- Clear resolves the target to the detected data body before any mutation.
- Marker removal, number-format restore, bold clear, and fill clear apply only to the data-body subrange.
- Banner / header significance letters are handled separately (see L3).
- If the selection is broad / multi-table, Clear shows the normalizer blocking message and writes nothing.

## L3. Clear removes banner / header significance markers (PR #97)

Run with banner letters on, then Clear.

Expected:
- Data body markers cleared as in L2.
- Banner / header cells above the data body are scanned (immediate row above + up to 5 rows above, matching Run's scan window).
- Only recognized trailing RIT banner markers are removed.
- Ordinary parenthesized text is preserved: `Wave (quarter)`, `Волна (квартал)`, `Brand (new)`, `Всё покупаю сам(а)`.

Known limitation: see J7.

## L4. Clear with first-column-is-Total stale markers

Run with first-column-is-Total on banner letters, then change the data or settings and run Clear. Expected: stale Total-row markers in banner are removed; Total banner cell ends with no marker.

## L5. Clear with broad multi-table selection

Selecting multiple tables and clicking Clear: blocked with status message, no writes.

---

# M. Numeric output and display conventions

## M1. Percent preservation

Input cell `42%`. After Run with a significant result: displays as `42.0% a` (or `42% a` with rounding on). Percent sign retained. Comma decimal input handled safely.

## M2. Rounding off

Share-like rows (proportions, NPS, Promoters, Detractors): 1 decimal. Mean-like rows (Mean, SD, Variance): 2 decimals.

## M3. Rounding on

Share-like: 0 decimals. Mean-like: 1 decimal. Statistics always computed from normalized original values, never rounded display values.

## M4. Mixed numeric conventions: 28, 28%, 0.28

A table containing plain numbers `28`, percent-formatted `28%`, and decimal-share `0.28`.

Expected:
- Unmarked cells preserve their original Excel numeric values and display formats.
- Marker-bearing cells maintain the visible numeric string convention before the appended marker: `28% b`, not `0.28 b`.
- The whole selected range is not converted to text.

## M5. Selected-range guardrail is warning-only

Manually select a range that includes title / question text / a stray header outside the data area, then Run.

Expected: Calculation proceeds. Status shows a warning that non-data rows may be included. Run is not blocked and the selection is not auto-trimmed.

---

# N. UI

## N1. Default status panel

On open: status panel is hidden, no Microsoft placeholder text visible.

## N2. Status after Run / Clear

Status panel appears, message wraps within taskpane width, no horizontal overflow.

## N3. Primary action hierarchy

`Запустить` is visually dominant. `Очистить значимости` is secondary.

## N4. Settings persistence (local)

Change confidence, one-tailed, several checkboxes, colors → restart Excel. Settings restored, dependent UI states refreshed, invalid combinations normalized.

## N5. Reset settings

Press Reset. Defaults restored, local storage cleared, status reports reset, defaults persist across Excel restart.

---

# O. Status messages

| Scenario | Expected status |
|---|---|
| Simple one-level banner, respect on | `Расчёт выполнен. Обработано блоков: 1.` No technical banner dump. |
| Wave auto previous-column | Success + `Баннер: для волновых групп автоматически применён режим “Сравнение с предыдущей колонкой”: Wave.` |
| Global Total | Success + brief global-Total explanation. No detailed banner dump. |
| Banner error (e.g. multiple local Totals) | Run stops. Concise user-facing message. No technical dump. |
| Broad multi-table selection | Run / Clear blocked. Normalizer blocking message. |

---

# P. Label detection

## P1. Labels immediately left

Labels 1–2 columns left of the selected data with `labels-on-left-side = false`. Detector reads nearby labels; numeric columns between labels and data are skipped.

## P2. Labels in leftmost sheet columns

User selects a right-side slice of a wide table; `labels-on-left-side = true`. Detector reads labels from the leftmost sheet columns.

---

# Q. Mixed block detection

## Q1. Mixed table with shared base

Rows: %, %, NPS, Promoters, Detractors, Mean, SD, Base.

Expected: Detector builds proportions block, NPS block, and mean block. Shared Base used appropriately. Detector does not skip later blocks after the NPS/mean blocks. Service rows do not receive markers (Promoters and Detractors do).

---

# R. Regression checklist (release gate)

Verify before release:

- [ ] All-vs-all proportions still work.
- [ ] Means use Welch's t-test; markers only on Mean.
- [ ] NPS-first and NPS-first-with-Neutral apply NPS markers to NPS only.
- [ ] Extended NPS labels detected on repeat full-table Run.
- [ ] NPS spread (SD / variance + Base) marks NPS only.
- [ ] Confidence selector affects all metric types.
- [ ] Old markers, bold, and fills cleared on rerun.
- [ ] Percent signs preserved; mixed `28` / `28%` / `0.28` display preserved.
- [ ] First-column-is-Total indexing correct; Total never receives markers.
- [ ] `T` / `t` markers and fill-only-Total behavior correct.
- [ ] Lower-than-Total fill applied to `t` cells.
- [ ] Small-base filtering runs before calculations; small-base fill has highest priority.
- [ ] Small-base Total stops calculation with clear status.
- [ ] Previous-column writes arrows only into the right column; no banner letters; incompatible settings disabled.
- [ ] Previous-column does not skip over excluded columns.
- [ ] Banner one-level fallback works without technical diagnostics.
- [ ] Repeated-label and sparse-merged banners detected correctly.
- [ ] Banner-aware Total detection: sparse upper local Total; first-column standalone `Всего` promoted to global Total; repeated local Totals stay local; compound labels like `Total spend` not detected as Total.
- [ ] Wave auto previous-column: mixed named + wave banners use arrows for wave group only.
- [ ] Nested wave dimension inside named groups: previous-column inside each category; no banner letters in auto-wave; no cross-category comparisons.
- [ ] Multi-column `Всего` wave group is not promoted to global Total.
- [ ] Wave descriptor row below semantic group row: letters not written across descriptor row.
- [ ] Auto-wave off restores ordinary banner-aware behavior.
- [ ] Banner letters with sparse/merged banner cells (respect-structure off) write to the resolved label cell and preserve parenthesized text.
- [ ] Full-table Run narrows to data body; broad multi-table selection blocked.
- [ ] Full-table Clear narrows to data body; banner / header markers removed; parenthesized labels preserved.
- [ ] Status panel hidden by default, wraps long text, `Запустить` remains visually primary.
- [ ] Settings persistence and reset behavior work.
