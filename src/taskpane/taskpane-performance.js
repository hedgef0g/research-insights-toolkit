// Development-only performance instrumentation for RIT taskpane flows.
//
// Enable from DevTools with either:
//   window.__RIT_PERF = true
//   localStorage.setItem('RIT_PERF', '1')
//
// Disable with either:
//   window.__RIT_PERF = false
//   localStorage.removeItem('RIT_PERF')
//
// The flag is read dynamically on every call, so no taskpane reload is needed.
//
// When enabled, each instrumented flow emits a single console.info entry:
//   [RIT perf] <flowName>  { phase1Ms, phase2Ms, ..., totalMs }
//
// All exported functions are no-ops when disabled, adding zero overhead to
// production runs.

function _perfWindowFlagEnabled() {
  try {
    if (typeof globalThis === "undefined") return false;
    return globalThis.__RIT_PERF === true;
  } catch (_) {
    return false;
  }
}

function _perfStorageFlagEnabled() {
  try {
    return typeof localStorage !== "undefined" && localStorage.getItem("RIT_PERF") === "1";
  } catch (_) {
    return false;
  }
}

function _perfEnabled() {
  return _perfWindowFlagEnabled() || _perfStorageFlagEnabled();
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

// Emits a console.info entry when enabled. No-op otherwise.
export function perfLog(flowName, phases) {
  if (!_perfEnabled()) return;
  console.info("[RIT perf]", flowName, phases);
}
