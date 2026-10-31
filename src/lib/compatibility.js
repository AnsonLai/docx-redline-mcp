import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

/**
 * The library release this server was developed and tested against. Before 1.0 the library does
 * not promise compatibility between releases, so the dependency is pinned to exactly this version
 * and a different installed release is reported rather than trusted. Bump both together, and only
 * after the compatibility tests pass.
 */
export const TESTED_LIBRARY_RELEASE = "0.11.1";

/**
 * Library capabilities this server relies on. They are documented in the library's CHANGELOG and
 * are all present in the tested release; they are listed so an upgrade review knows what to check.
 */
export const REQUIRED_LIBRARY_FEATURES = Object.freeze([
  "openDocx facade (inspect, preflight, applyOperations, resolveRevisions, deleteComments, resolveComment)",
  "package-scope revision tokens",
  "recovery envelope and retryPlan",
  "localized replacements",
  "header/footer part selector",
  "revision fidelity verification"
]);

export function readInstalledLibraryRelease() {
  const entry = fileURLToPath(import.meta.resolve("@ansonlai/docx-redline-js"));
  const manifest = JSON.parse(readFileSync(join(dirname(entry), "package.json"), "utf8"));
  return manifest.version;
}

export function getLibraryCompatibility() {
  const installed = readInstalledLibraryRelease();
  return {
    tested: TESTED_LIBRARY_RELEASE,
    installed,
    matches: installed === TESTED_LIBRARY_RELEASE
  };
}

let warned = false;

/** Writes one warning to stderr (never stdout, which carries the MCP protocol) on a mismatch. */
export function warnOnLibraryMismatch(stderr = process.stderr) {
  const compatibility = getLibraryCompatibility();
  if (!compatibility.matches && !warned) {
    warned = true;
    stderr.write(
      `[warn] docx-redline-mcp was tested with @ansonlai/docx-redline-js@${compatibility.tested} `
      + `but ${compatibility.installed} is installed. Treat edits as unverified until the compatibility tests pass.\n`
    );
  }
  return compatibility;
}
