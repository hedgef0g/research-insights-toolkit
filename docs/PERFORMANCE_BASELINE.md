# Performance Baseline

Pre-1.0 performance baseline for the dev-only `RIT_PERF` instrumentation added in PR #254.

This document is for measurement only.

- Do not change product behavior while collecting this baseline.
- Do not optimize yet.
- Do not refactor as part of baseline collection.

## Purpose

Use this baseline to capture current timing characteristics before 1.0 so later optimization work is based on measured bottlenecks rather than assumptions.

## Enable `RIT_PERF`

Enable the dev-only instrumentation from the taskpane devtools console:

```js
localStorage.setItem('RIT_PERF', '1')
```

Reload the taskpane after enabling the flag so the current session picks up the setting.

To disable it later:

```js
localStorage.removeItem('RIT_PERF')
```

## Measurement Rules

- Run each scenario at least 3 times.
- If the first run is obviously distorted by Excel or taskpane warm-up, ignore that run and continue until you still have at least 3 measured runs.
- Record the median result, not the best result.
- Keep workbook state and settings stable across repeated runs for the same scenario.
- Treat this as a measurement pass, not a tuning pass.

## Scenarios To Measure

Measure the current flows that the instrumentation now exposes:

| Scenario | What to run | Notes |
|---------|-------------|-------|
| Selected-range Run on a small table | Manual `Run` on a small representative selected range | Use the stable manual selected-range workflow |
| Selected-range Run on a large table | Manual `Run` on a large representative selected range | Prefer a realistic heavy table, not a synthetic stress case unless noted |
| Current-table autorun | `Autorun -> Current table` | Active cell should be inside the target table |
| Current-sheet autorun | `Autorun -> Current sheet` | Record how many tables are scanned on the sheet |
| Workbook autorun | `Autorun -> Whole workbook` | Record workbook-wide table count |
| Content generation | `Contents` flow that creates or refreshes `Content` | Capture content-sheet generation timing separately |
| Run report writing | Any flow that creates or refreshes `Run report` | Capture report-writing timing separately from scanning where possible |

## Recording Template

Use one row per measured scenario/workbook combination.

| Workbook / scenario | Table count | Approximate range size | Flow | scanMs | loopMs / writeMs | contentWriteMs | totalMs | Notes |
|---------------------|-------------|------------------------|------|--------|------------------|----------------|---------|-------|
| Example: Tracker Q2 / selected-range Run small | 1 | ~40 rows x 8 cols | Manual selected-range Run |  |  |  |  | Median of 3 warm runs |
|  |  |  |  |  |  |  |  |  |
|  |  |  |  |  |  |  |  |  |
|  |  |  |  |  |  |  |  |  |
|  |  |  |  |  |  |  |  |  |
|  |  |  |  |  |  |  |  |  |
|  |  |  |  |  |  |  |  |  |

## Field Guidance

- `Workbook / scenario`: Identify the workbook and the exact scenario variant being measured.
- `Table count`: Record tables involved in the flow. For manual selected-range Run this is usually `1`.
- `Approximate range size`: Use an estimate such as `~40 rows x 8 cols` when exact dimensions are not important.
- `Flow`: Use a short label such as `Manual selected-range Run`, `Autorun current sheet`, `Workbook autorun`, `Content generation`, or `Run report writing`.
- `scanMs`: Time spent discovering or scanning candidate tables when the instrumentation exposes it.
- `loopMs / writeMs`: Record whichever value the flow emits. For some flows this is iteration/processing time; for others it is write time.
- `contentWriteMs`: Fill only when the `Content` sheet write path is part of the scenario.
- `totalMs`: Record the end-to-end median for the scenario.
- `Notes`: Capture warm-up exclusions, unusual workbook shape, banner complexity, merged-like headers, or anything else that may explain outliers.

## How To Interpret The Baseline

- Prefer medians over individual runs.
- Compare similar flows against each other before comparing different workbook shapes.
- Separate scan-heavy issues from write-heavy issues when the instrumentation makes that visible.
- Treat a single slow run as noise unless it repeats.

## Follow-up Rules

- Do not optimize unless a measured bottleneck is identified.
- Do not introduce per-cell writes.
- Prefer shared execution improvements over manual-only or autorun-only forks unless the measurement clearly shows a flow-specific bottleneck.

## Suggested Capture Notes

When recording a baseline, note any factors that could matter later:

- banner depth or merged-like banner structure;
- number of tables on the sheet or workbook;
- whether `Content` or `Run report` had to be created vs updated;
- whether the workbook was freshly opened or already warm;
- whether the scenario was manual selected-range or detected-table based.
