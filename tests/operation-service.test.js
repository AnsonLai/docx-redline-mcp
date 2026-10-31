import test from "node:test";
import assert from "node:assert/strict";
import JSZip from "jszip";
import {
  applyOperationsToPackage,
  chooseActionableError,
  isMaintenanceOperation,
  resolveApplyOptions
} from "../src/lib/operation-service.js";
import { buildDocxBuffer } from "./helpers/fixtures.js";

async function sessionFor(options) {
  return { zip: await JSZip.loadAsync(await buildDocxBuffer(options)) };
}

async function documentXmlOf(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  return zip.file("word/document.xml").async("string");
}

test("applies a redline operation and returns the new package bytes", async () => {
  const session = await sessionFor({ paragraphs: ["Original text"] });

  const { payload, bytes } = await applyOperationsToPackage(session, [
    { type: "redline", target: "Original text", modified: "Updated text" }
  ], { author: "Roo" });

  assert.equal(payload.status, "ok");
  assert.equal(payload.completion, true);
  assert.equal(payload.hasChanges, true);
  assert.deepEqual(payload.authorsUsed, ["Roo"]);
  assert.equal(payload.results[0].status, "applied");
  assert.ok(payload.receipts[0].revisionItems.length > 0);
  assert.match(await documentXmlOf(bytes), /w:(ins|del)/);
});

test("does not modify the session when applying", async () => {
  const session = await sessionFor({ paragraphs: ["Original text"] });
  const before = await session.zip.file("word/document.xml").async("string");

  await applyOperationsToPackage(session, [
    { type: "redline", target: "Original text", modified: "Updated text" }
  ]);

  assert.equal(await session.zip.file("word/document.xml").async("string"), before);
});

test("resolveApplyOptions defaults to atomic strict targets and forwards library options", () => {
  assert.deepEqual(resolveApplyOptions(), { atomic: true, strictTargets: true });
  assert.deepEqual(
    resolveApplyOptions({ atomic: false, existingRevisions: "auto", structuredContent: false, ignored: 1 }),
    { atomic: false, strictTargets: true, existingRevisions: "auto", structuredContent: false }
  );
});

test("recognizes maintenance operations case-insensitively", () => {
  assert.equal(isMaintenanceOperation({ type: "Accept_All_Tracked_Changes" }), true);
  assert.equal(isMaintenanceOperation({ type: "resolve_comment" }), true);
  assert.equal(isMaintenanceOperation({ type: "redline" }), false);
});

test("chooseActionableError promotes the failing operation over the generic batch error", () => {
  const batch = { code: "BATCH_OPERATION_FAILED", message: "Atomic batch rolled back" };
  const ordinary = { code: "TARGET_TEXT_MISMATCH", message: "text differs", recovery: { action: "reinspect" } };
  const needsUser = {
    code: "COMMENTED_CONTENT_DELETE",
    message: "has comments",
    recovery: { action: "resolve_comments", requiresUserAuthorization: true }
  };

  assert.equal(chooseActionableError([], batch), batch, "no failing operation keeps the batch error");
  assert.equal(chooseActionableError([{ index: 1, status: "applied" }], null), null);

  const single = chooseActionableError([{ index: 2, status: "error", error: ordinary }], batch);
  assert.equal(single.code, "TARGET_TEXT_MISMATCH");
  assert.equal(single.operationIndex, 2);

  const several = chooseActionableError([
    { index: 1, status: "error", error: ordinary },
    { index: 3, status: "error", error: needsUser }
  ], batch);
  assert.equal(several.code, "COMMENTED_CONTENT_DELETE", "a refusal needing the user wins");
  assert.equal(several.operationIndex, 3);

  const first = chooseActionableError([
    { index: 1, status: "error", error: ordinary },
    { index: 2, status: "error", error: { ...ordinary, code: "OTHER" } }
  ], batch);
  assert.equal(first.code, "TARGET_TEXT_MISMATCH");
  assert.match(first.message, /2 operations failed/);
});
