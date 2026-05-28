// Development-only performance instrumentation for RIT taskpane flows.
//
// Enable the main performance log from DevTools with either:
//   window.__RIT_PERF = true
//   window.__RIT_PERF = "1"
//   localStorage.setItem('RIT_PERF', '1')
//
// Optional nested probes can use their own runtime/storage flags. For example:
//   window.__RIT_BANNER_VALUES_READ_PROBE = true
//   localStorage.setItem('RIT_BANNER_VALUES_READ_PROBE', '1')
//
// Optional banner write host-flush profiling:
//   window.__RIT_BANNER_WRITE_PROFILE = true
//   localStorage.setItem('RIT_BANNER_WRITE_PROFILE', '1')
//
// This intentionally changes banner write flushing while enabled, splitting
// marker numberFormat and values into separate syncs so their host cost can be
// measured. It is disabled unless RIT_PERF is also enabled.
//
// Disable with either:
//   window.__RIT_PERF = false
//   localStorage.removeItem('RIT_PERF')
//
// The runtime window/globalThis flag takes priority over storage. This lets
// DevTools disable logging for the current taskpane session even if a stored
// RIT_PERF value exists.
//
// The flag is read dynamically on every call, so no taskpane reload is needed.
//
// When enabled, each instrumented flow emits a single console.info entry:
//   [RIT perf] <flowName>  { phase1Ms, phase2Ms, ..., totalMs }
//
// All exported functions are no-ops when disabled, adding zero overhead to
// production runs.

function _readRuntimeFlag(flagName) {
  try {
    if (typeof globalThis === "undefined") return undefined;
    return globalThis[flagName];
  } catch (_) {
    return undefined;
  }
}

function _runtimeFlagEnabled(flagName) {
  const runtimeFlag = _readRuntimeFlag(flagName);

  if (runtimeFlag === true || runtimeFlag === "1") return true;
  if (runtimeFlag === false || runtimeFlag === "0") return false;

  return undefined;
}

function _storageFlagEnabled(storageKey) {
  try {
    return typeof localStorage !== "undefined" && localStorage.getItem(storageKey) === "1";
  } catch (_) {
    return false;
  }
}

function _perfEnabled() {
  const runtimeEnabled = _runtimeFlagEnabled("__RIT_PERF");

  if (runtimeEnabled !== undefined) return runtimeEnabled;

  return _storageFlagEnabled("RIT_PERF");
}

function _bannerWriteProfileRuntimeFlagEnabled() {
  try {
    if (typeof globalThis === "undefined") return undefined;
    const runtimeFlag = globalThis.__RIT_BANNER_WRITE_PROFILE;
    if (runtimeFlag === true || runtimeFlag === "1") return true;
    if (runtimeFlag === false || runtimeFlag === "0") return false;
  } catch (_) {
    return undefined;
  }

  return undefined;
}

function _bannerWriteProfileStorageFlagEnabled() {
  try {
    return typeof localStorage !== "undefined" && localStorage.getItem("RIT_BANNER_WRITE_PROFILE") === "1";
  } catch (_) {
    return false;
  }
}

export function perfEnabled() {
  return _perfEnabled();
}

export function perfFlagEnabled(runtimeFlagName, storageKey = runtimeFlagName) {
  const runtimeEnabled = _runtimeFlagEnabled(runtimeFlagName);

  if (runtimeEnabled !== undefined) return runtimeEnabled;

  return _storageFlagEnabled(storageKey);
}

export function perfBannerWriteProfileEnabled() {
  if (!_perfEnabled()) return false;

  const runtimeEnabled = _bannerWriteProfileRuntimeFlagEnabled();
  if (runtimeEnabled !== undefined) return runtimeEnabled;

  return _bannerWriteProfileStorageFlagEnabled();
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
