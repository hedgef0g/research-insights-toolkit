// Development-only performance instrumentation for RIT taskpane flows.
//
// Enable:  localStorage.setItem('RIT_PERF', '1')   (in browser DevTools)
// Disable: localStorage.removeItem('RIT_PERF')
//
// When enabled, each instrumented flow emits a single console.debug entry:
//   [RIT perf] <flowName>  { phase1Ms, phase2Ms, ..., totalMs }
//
// All exported functions are no-ops when disabled, adding zero overhead to
// production runs.

function _perfEnabled() {
  try {
    return typeof localStorage !== "undefined" && localStorage.getItem("RIT_PERF") === "1";
  } catch (_) {
    return false;
  }
}

export function perfEnabled() {
  return _perfEnabled();
}

// Returns Date.now() when enabled, 0 otherwise.
export function perfNow() {
  return _perfEnabled() ? Date.now() : 0;
}

// Returns milliseconds elapsed since startMs when enabled, 0 otherwise.
export function perfElapsed(startMs) {
  return _perfEnabled() && startMs ? Date.now() - startMs : 0;
}

// Emits a console.debug entry when enabled. No-op otherwise.
export function perfLog(flowName, phases) {
  if (!_perfEnabled()) return;
  console.debug("[RIT perf]", flowName, phases);
}
