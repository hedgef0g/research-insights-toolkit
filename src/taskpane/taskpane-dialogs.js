/* global document */

const MARKER_OVERFLOW_DIALOG_ID = "marker-overflow-dialog";
const MARKER_OVERFLOW_CONTINUE_ID = "marker-overflow-continue";
const MARKER_OVERFLOW_STOP_ID = "marker-overflow-stop";

/**
 * Shows the in-taskpane marker-overflow dialog and resolves the user's choice.
 *
 * window.confirm is not supported in the Office add-in webview, so this drives a
 * small custom modal defined in taskpane.html. Returns a Promise:
 * - true  -> continue with multi-character markers;
 * - false -> stop the calculation.
 *
 * Esc, clicking the backdrop, or missing markup all resolve to false (stop) so
 * the safe choice (no writes) is the default.
 */
export function confirmMarkerOverflowDialog() {
  return new Promise((resolve) => {
    const overlay = document.getElementById(MARKER_OVERFLOW_DIALOG_ID);
    const continueButton = document.getElementById(MARKER_OVERFLOW_CONTINUE_ID);
    const stopButton = document.getElementById(MARKER_OVERFLOW_STOP_ID);

    if (!overlay || !continueButton || !stopButton) {
      // Fail safe: without the dialog markup, stop rather than write blindly.
      resolve(false);
      return;
    }

    const finish = (shouldContinue) => {
      overlay.style.display = "none";
      continueButton.removeEventListener("click", onContinue);
      stopButton.removeEventListener("click", onStop);
      overlay.removeEventListener("mousedown", onBackdrop);
      document.removeEventListener("keydown", onKeydown);
      resolve(shouldContinue);
    };

    const onContinue = () => finish(true);
    const onStop = () => finish(false);
    const onBackdrop = (event) => {
      // Clicking outside the dialog body counts as Stop.
      if (event.target === overlay) finish(false);
    };
    const onKeydown = (event) => {
      if (event.key === "Escape") finish(false);
    };

    continueButton.addEventListener("click", onContinue);
    stopButton.addEventListener("click", onStop);
    overlay.addEventListener("mousedown", onBackdrop);
    document.addEventListener("keydown", onKeydown);

    overlay.style.display = "flex";
    continueButton.focus();
  });
}

/**
 * Per-operation marker-overflow decision.
 *
 * The dialog is shown at most once per Run; the user's choice is then reused for
 * every table processed in the same operation (so batch runs do not re-prompt
 * for each table). `resolve()` is async and returns "continue" or "stop". A
 * shared in-flight promise guards against showing two dialogs if `resolve()` is
 * awaited from more than one place before the first choice is made.
 */
export function createMarkerOverflowDecider(confirmDialog = confirmMarkerOverflowDialog) {
  let decision = null; // null | "continue" | "stop"
  let pending = null;

  return {
    async resolve() {
      if (decision !== null) {
        return decision;
      }

      if (!pending) {
        pending = confirmDialog().then((shouldContinue) => {
          decision = shouldContinue ? "continue" : "stop";
          return decision;
        });
      }

      return pending;
    },
    get decision() {
      return decision;
    },
  };
}
