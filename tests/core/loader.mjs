/**
 * ESM loader hooks for pure-core unit tests.
 *
 * resolve — adds .js extension to bare relative imports so that Node's native
 *            ESM resolver can find files that omit extensions (e.g. webpack
 *            convention): "./metric-detector" → "./metric-detector.js".
 *
 * load    — marks src/core source files as ES modules so they can be imported
 *            from .mjs test files without "type":"module" in package.json.
 *            Without this, Node would attempt to parse them as CommonJS and
 *            fail on the export keyword.
 */

import { existsSync } from "node:fs";
import { fileURLToPath, pathToFileURL } from "node:url";
import { dirname, join } from "node:path";

// Only treat these suffixes as "already has a real file extension".
// Bare path segments like "./config/dictionary.config" end in ".config"
// which must NOT be treated as a file extension here.
const JS_EXTENSIONS = /\.(js|mjs|cjs|json|node|ts|mts|cts)$/i;

export async function resolve(specifier, context, nextResolve) {
  if (specifier.startsWith(".") && context.parentURL && !JS_EXTENSIONS.test(specifier)) {
    try {
      const parentDir = dirname(fileURLToPath(context.parentURL));
      const candidate = join(parentDir, specifier + ".js");
      if (existsSync(candidate)) {
        return { shortCircuit: true, url: pathToFileURL(candidate).href };
      }
    } catch {
      // fall through to default resolver
    }
  }
  return nextResolve(specifier, context);
}

export async function load(url, context, nextLoad) {
  // Force ESM format for project source files.
  // src/core/*.js and src/taskpane/*.js files use export/import syntax but
  // the package has no "type":"module", so Node would otherwise parse them
  // as CommonJS.
  if (
    (url.includes("/src/core/") || url.includes("/src/taskpane/")) &&
    !url.includes("/node_modules/") &&
    url.endsWith(".js")
  ) {
    const result = await nextLoad(url, context);
    return { ...result, format: "module" };
  }
  return nextLoad(url, context);
}
