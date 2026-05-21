import { describe, it } from "node:test";
import assert from "node:assert/strict";

import {
  BATCH_SKIP_REASONS,
  filterWorkbookCandidates,
} from "../../src/core/batch-candidate-filter.js";

// ─── Helpers ──────────────────────────────────────────────────────────────────

function makeItem(overrides) {
  return {
    rangeAddress: "A1:D10",
    resolvedRangeAddress: null,
    candidateStatus: "available",
    canRunCheckTable: true,
    title: "",
    resolvedTitle: null,
    ...overrides,
  };
}

function makeSheet(sheetName, items) {
  return { sheetName, items };
}

function makeInventory(...sheetResults) {
  return { sheetResults };
}

// ─── BATCH_SKIP_REASONS ───────────────────────────────────────────────────────

describe("BATCH_SKIP_REASONS", () => {
  it("is frozen", () => {
    assert.ok(Object.isFrozen(BATCH_SKIP_REASONS));
  });

  it("has stable string values", () => {
    assert.strictEqual(BATCH_SKIP_REASONS.MISSING_RANGE, "missing_range");
    assert.strictEqual(BATCH_SKIP_REASONS.CANDIDATE_UNCERTAIN, "candidate_uncertain");
    assert.strictEqual(BATCH_SKIP_REASONS.CANDIDATE_REJECTED, "candidate_rejected");
    assert.strictEqual(BATCH_SKIP_REASONS.UNKNOWN_STATUS, "unknown_status");
  });
});

// ─── filterWorkbookCandidates — eligible ──────────────────────────────────────

describe("filterWorkbookCandidates — eligible candidates", () => {
  it("includes available + canRunCheckTable + rangeAddress candidate", () => {
    const item = makeItem({ candidateStatus: "available", canRunCheckTable: true, rangeAddress: "B2:E8" });
    const { eligible, skipped } = filterWorkbookCandidates(makeInventory(makeSheet("Sheet1", [item])));
    assert.strictEqual(eligible.length, 1);
    assert.strictEqual(eligible[0].sheetName, "Sheet1");
    assert.strictEqual(eligible[0].rangeAddress, "B2:E8");
    assert.strictEqual(skipped.length, 0);
  });

  it("uses resolvedRangeAddress over rangeAddress when both present", () => {
    const item = makeItem({
      candidateStatus: "available",
      canRunCheckTable: true,
      rangeAddress: "A1:D10",
      resolvedRangeAddress: "A2:D11",
    });
    const { eligible } = filterWorkbookCandidates(makeInventory(makeSheet("Sheet1", [item])));
    assert.strictEqual(eligible[0].rangeAddress, "A2:D11");
  });

  it("includes title from resolvedTitle when present", () => {
    const item = makeItem({
      candidateStatus: "available",
      canRunCheckTable: true,
      title: "Raw",
      resolvedTitle: "Resolved",
    });
    const { eligible } = filterWorkbookCandidates(makeInventory(makeSheet("Sheet1", [item])));
    assert.strictEqual(eligible[0].title, "Resolved");
  });

  it("falls back to title when resolvedTitle is absent", () => {
    const item = makeItem({ candidateStatus: "available", canRunCheckTable: true, title: "Raw", resolvedTitle: null });
    const { eligible } = filterWorkbookCandidates(makeInventory(makeSheet("Sheet1", [item])));
    assert.strictEqual(eligible[0].title, "Raw");
  });

  it("title is empty string when both title and resolvedTitle are absent", () => {
    const item = makeItem({ candidateStatus: "available", canRunCheckTable: true, title: "", resolvedTitle: null });
    const { eligible } = filterWorkbookCandidates(makeInventory(makeSheet("Sheet1", [item])));
    assert.strictEqual(eligible[0].title, "");
  });

  it("does not include available candidate when canRunCheckTable is false", () => {
    const item = makeItem({ candidateStatus: "available", canRunCheckTable: false });
    const { eligible, skipped } = filterWorkbookCandidates(makeInventory(makeSheet("Sheet1", [item])));
    assert.strictEqual(eligible.length, 0);
    assert.strictEqual(skipped.length, 1);
    assert.strictEqual(skipped[0].reason, BATCH_SKIP_REASONS.UNKNOWN_STATUS);
  });
});

// ─── filterWorkbookCandidates — skipped ──────────────────────────────────────

describe("filterWorkbookCandidates — skipped candidates", () => {
  it("skips uncertain candidate with CANDIDATE_UNCERTAIN reason", () => {
    const item = makeItem({ candidateStatus: "uncertain", rangeAddress: "A1:D10" });
    const { eligible, skipped } = filterWorkbookCandidates(makeInventory(makeSheet("Sheet1", [item])));
    assert.strictEqual(eligible.length, 0);
    assert.strictEqual(skipped.length, 1);
    assert.strictEqual(skipped[0].reason, BATCH_SKIP_REASONS.CANDIDATE_UNCERTAIN);
    assert.strictEqual(skipped[0].status, "uncertain");
    assert.strictEqual(skipped[0].sheetName, "Sheet1");
    assert.strictEqual(skipped[0].rangeAddress, "A1:D10");
  });

  it("skips rejected candidate with CANDIDATE_REJECTED reason", () => {
    const item = makeItem({ candidateStatus: "rejected", rangeAddress: "C3:F12" });
    const { skipped } = filterWorkbookCandidates(makeInventory(makeSheet("Sheet2", [item])));
    assert.strictEqual(skipped[0].reason, BATCH_SKIP_REASONS.CANDIDATE_REJECTED);
    assert.strictEqual(skipped[0].status, "rejected");
    assert.strictEqual(skipped[0].rangeAddress, "C3:F12");
  });

  it("skips candidate with no range address using MISSING_RANGE reason", () => {
    const item = makeItem({ candidateStatus: "available", rangeAddress: null, resolvedRangeAddress: null });
    const { skipped } = filterWorkbookCandidates(makeInventory(makeSheet("Sheet1", [item])));
    assert.strictEqual(skipped[0].reason, BATCH_SKIP_REASONS.MISSING_RANGE);
    assert.strictEqual(skipped[0].rangeAddress, null);
  });

  it("stores null rangeAddress when both resolvedRangeAddress and rangeAddress are absent", () => {
    const item = makeItem({ candidateStatus: "available", rangeAddress: null, resolvedRangeAddress: null });
    const { skipped } = filterWorkbookCandidates(makeInventory(makeSheet("Sheet1", [item])));
    assert.strictEqual(skipped[0].reason, BATCH_SKIP_REASONS.MISSING_RANGE);
    assert.strictEqual(skipped[0].rangeAddress, null);
  });

  it("skips unknown status with UNKNOWN_STATUS reason", () => {
    const item = makeItem({ candidateStatus: "future_status", rangeAddress: "A1:D10" });
    const { skipped } = filterWorkbookCandidates(makeInventory(makeSheet("Sheet1", [item])));
    assert.strictEqual(skipped[0].reason, BATCH_SKIP_REASONS.UNKNOWN_STATUS);
    assert.strictEqual(skipped[0].status, "future_status");
  });

  it("skips null candidateStatus with UNKNOWN_STATUS and null status field", () => {
    const item = makeItem({ candidateStatus: null, rangeAddress: "A1:D10" });
    const { skipped } = filterWorkbookCandidates(makeInventory(makeSheet("Sheet1", [item])));
    assert.strictEqual(skipped[0].reason, BATCH_SKIP_REASONS.UNKNOWN_STATUS);
    assert.strictEqual(skipped[0].status, null);
  });
});

// ─── filterWorkbookCandidates — Content sheet ─────────────────────────────────

describe("filterWorkbookCandidates — Content sheet exclusion", () => {
  it("ignores the Content sheet entirely — items not counted toward skipped", () => {
    const contentItem = makeItem({ candidateStatus: "available" });
    const { eligible, skipped } = filterWorkbookCandidates(
      makeInventory(makeSheet("Content", [contentItem]))
    );
    assert.strictEqual(eligible.length, 0);
    assert.strictEqual(skipped.length, 0);
  });

  it("excludes Content sheet while processing other sheets normally", () => {
    const contentItem = makeItem({ candidateStatus: "available", rangeAddress: "A1:D10" });
    const normalItem = makeItem({ candidateStatus: "available", rangeAddress: "B2:E8" });
    const { eligible, skipped } = filterWorkbookCandidates(
      makeInventory(makeSheet("Content", [contentItem]), makeSheet("Data", [normalItem]))
    );
    assert.strictEqual(eligible.length, 1);
    assert.strictEqual(eligible[0].sheetName, "Data");
    assert.strictEqual(skipped.length, 0);
  });

  it("respects a custom contentSheetName option", () => {
    const item = makeItem({ candidateStatus: "available", rangeAddress: "A1:D10" });
    const { eligible, skipped } = filterWorkbookCandidates(
      makeInventory(makeSheet("Содержание", [item])),
      { contentSheetName: "Содержание" }
    );
    assert.strictEqual(eligible.length, 0);
    assert.strictEqual(skipped.length, 0);
  });

  it("uses 'Content' as the default contentSheetName", () => {
    const item = makeItem({ candidateStatus: "available", rangeAddress: "A1:D10" });
    const { eligible } = filterWorkbookCandidates(makeInventory(makeSheet("Content", [item])));
    assert.strictEqual(eligible.length, 0);
  });
});

// ─── filterWorkbookCandidates — multi-sheet and mixed ────────────────────────

describe("filterWorkbookCandidates — multi-sheet and mixed results", () => {
  it("aggregates eligible and skipped across multiple sheets", () => {
    const availableItem = makeItem({ candidateStatus: "available", rangeAddress: "A1:D10" });
    const uncertainItem = makeItem({ candidateStatus: "uncertain", rangeAddress: "A1:D5" });
    const rejectedItem = makeItem({ candidateStatus: "rejected", rangeAddress: "A1:D3" });
    const inv = makeInventory(
      makeSheet("Sheet1", [availableItem]),
      makeSheet("Sheet2", [uncertainItem, rejectedItem])
    );
    const { eligible, skipped } = filterWorkbookCandidates(inv);
    assert.strictEqual(eligible.length, 1);
    assert.strictEqual(skipped.length, 2);
  });

  it("returns empty arrays for empty sheetResults", () => {
    const { eligible, skipped } = filterWorkbookCandidates({ sheetResults: [] });
    assert.strictEqual(eligible.length, 0);
    assert.strictEqual(skipped.length, 0);
  });

  it("returns empty arrays for sheet with no items", () => {
    const { eligible, skipped } = filterWorkbookCandidates(makeInventory(makeSheet("Sheet1", [])));
    assert.strictEqual(eligible.length, 0);
    assert.strictEqual(skipped.length, 0);
  });

  it("preserves order: eligible items appear in sheet/item traversal order", () => {
    const a = makeItem({ candidateStatus: "available", rangeAddress: "A1:D5", title: "First" });
    const b = makeItem({ candidateStatus: "available", rangeAddress: "A6:D10", title: "Second" });
    const { eligible } = filterWorkbookCandidates(makeInventory(makeSheet("Sheet1", [a, b])));
    assert.strictEqual(eligible[0].title, "First");
    assert.strictEqual(eligible[1].title, "Second");
  });
});
