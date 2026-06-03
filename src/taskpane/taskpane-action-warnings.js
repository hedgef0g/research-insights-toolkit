/* global document */

export function collectActionWarningKeys({
  action,
  labelsOnLeftSide = false,
  addTableFootnoteRequested = false,
  recolorRequested = false,
} = {}) {
  if (action !== "run" && action !== "autorun") {
    return [];
  }

  const warningKeys = [];

  if (labelsOnLeftSide && addTableFootnoteRequested) {
    warningKeys.push("footnote-labels-left");
  }

  if (labelsOnLeftSide && recolorRequested) {
    warningKeys.push("recolor-labels-left");
  } else if (action === "run" && recolorRequested) {
    warningKeys.push("manual-run-recolor");
  }

  return warningKeys;
}

export function renderActionWarnings(warningBlock, warningKeys) {
  if (!warningBlock) {
    return;
  }

  const activeWarnings = new Set(warningKeys);
  let hasWarnings = false;

  warningBlock.querySelectorAll("[data-action-warning]").forEach((line) => {
    const show = activeWarnings.has(line.dataset.actionWarning);
    line.style.display = show ? "" : "none";
    hasWarnings = hasWarnings || show;
  });

  warningBlock.style.display = hasWarnings ? "" : "none";
}

export function refreshActionWarnings({
  documentRef = typeof document === "undefined" ? null : document,
  action,
  labelsOnLeftSide = false,
  addTableFootnoteRequested = false,
  recolorRequested = false,
} = {}) {
  if (!documentRef) {
    return;
  }

  const warningBlock = documentRef.getElementById("action-settings-warnings");
  renderActionWarnings(
    warningBlock,
    collectActionWarningKeys({
      action,
      labelsOnLeftSide,
      addTableFootnoteRequested,
      recolorRequested,
    })
  );
}
