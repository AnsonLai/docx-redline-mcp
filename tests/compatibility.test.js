import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  TESTED_LIBRARY_RELEASE,
  getLibraryCompatibility,
  warnOnLibraryMismatch
} from "../src/lib/compatibility.js";
import { createToolHandlers } from "../src/server.js";
import { writeDocxFixture } from "./helpers/fixtures.js";

test("the dependency is pinned to the library release the server declares as tested", async () => {
  const manifest = JSON.parse(await readFile(new URL("../package.json", import.meta.url), "utf8"));
  assert.equal(manifest.dependencies["@ansonlai/docx-redline-js"], TESTED_LIBRARY_RELEASE);
  assert.deepEqual(getLibraryCompatibility(), {
    tested: TESTED_LIBRARY_RELEASE,
    installed: TESTED_LIBRARY_RELEASE,
    matches: true
  });
});

test("a matching library is silent", () => {
  const chunks = [];
  const compatibility = warnOnLibraryMismatch({ write: chunk => chunks.push(chunk) });
  assert.equal(compatibility.matches, true);
  assert.deepEqual(chunks, []);
});

test("open_docx reports the library compatibility", async () => {
  const handlers = createToolHandlers();
  const opened = await handlers.openDocx({ path: await writeDocxFixture("compat.docx", {}) });
  assert.equal(opened.library.matches, true);
  assert.equal(opened.library.installed, TESTED_LIBRARY_RELEASE);
});
