import test from "node:test";
import assert from "node:assert/strict";
import { access, readFile } from "node:fs/promises";
import { basename, dirname, join } from "node:path";
import JSZip from "jszip";
import { createToolHandlers } from "../src/server.js";
import { OperationFailedError } from "../src/lib/errors.js";
import { writeDocxFixture } from "./helpers/fixtures.js";

async function openFixture(options, name = "agent.docx") {
  const handlers = createToolHandlers();
  const filePath = await writeDocxFixture(name, options);
  const opened = await handlers.openDocx({ path: filePath });
  return { handlers, handle: opened.handle, filePath, opened };
}

async function failure(promise) {
  const error = await promise.then(() => null, caught => caught);
  assert.ok(error instanceof OperationFailedError, `expected a structured failure, got ${error}`);
  return error.payload;
}

async function find(handlers, handle, search, extra = {}) {
  const inspection = await handlers.inspectDocument({ handle, search, ...extra });
  return inspection.paragraphs.find(paragraph => paragraph.exactText.includes(search));
}

test("edit_document applies handle-based edits and returns compact, readable results", async () => {
  const { handlers, handle } = await openFixture({
    paragraphs: ["1. The Supplier may terminate this Agreement.", "2. Payment is due in 30 days."]
  });
  const first = await find(handlers, handle, "terminate");
  const second = await find(handlers, handle, "Payment");

  const result = await handlers.editDocument({
    handle,
    author: "Roo",
    edits: [
      { target: first.target, replacements: [{ find: "may terminate", replace: "may only terminate for cause" }] },
      { target: second.target, comment: { content: "Confirm with finance", anchorText: "30 days" } }
    ]
  });

  assert.equal(result.status, "ok");
  assert.equal(result.completion, true);
  assert.equal(result.results, undefined, "raw operation results are replaced by compact edits");
  assert.deepEqual(result.edits.map(edit => [edit.index, edit.target, edit.intent, edit.status]), [
    [1, first.target, "replacements", "applied"],
    [2, second.target, "comment", "applied"]
  ]);
  assert.match(result.edits[0].replacements[0].afterExcerpt, /may only terminate for cause/);
  assert.deepEqual(result.authorsUsed, ["Roo"]);
  assert.ok(result.receipts.length === 2);
});

test("edit_document retires old handles and issues refreshed ones", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["Alpha clause.", "Beta clause."] });
  const alpha = await find(handlers, handle, "Alpha");

  const result = await handlers.editDocument({
    handle,
    edits: [{ target: alpha.target, desiredText: "Alpha clause, amended." }]
  });

  assert.equal(result.refreshedTargets.length, 1);
  const refreshed = result.refreshedTargets[0];
  assert.equal(refreshed.previousHandle, alpha.target);
  assert.notEqual(refreshed.handle, alpha.target);
  assert.equal(refreshed.exactText, "Alpha clause, amended.");

  const stale = await failure(handlers.editDocument({
    handle,
    edits: [{ target: alpha.target, desiredText: "Alpha again" }]
  }));
  assert.equal(stale.error.code, "STALE_TARGET_HANDLE");
  assert.equal(stale.error.recovery.requiresReinspection, true);
  assert.deepEqual(stale.retryPlan.failedIndexes, [1]);

  // The refreshed handle works without another inspection.
  const again = await handlers.editDocument({
    handle,
    edits: [{ target: refreshed.handle, replacements: [{ find: "amended", replace: "revised" }] }]
  });
  assert.equal(again.completion, true);
});

test("edit_document handles from before write_part are stale", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["Alpha clause."] });
  const alpha = await find(handlers, handle, "Alpha");
  const xml = (await handlers.readPart({ handle, partPath: "word/document.xml" })).text;
  await handlers.writePart({ handle, partPath: "word/document.xml", xml: xml.replace("Alpha", "Alpha!") });

  const stale = await failure(handlers.editDocument({
    handle,
    edits: [{ target: alpha.target, desiredText: "x" }]
  }));
  assert.equal(stale.error.code, "STALE_TARGET_HANDLE");
});

test("edit_document rejects unknown handles and malformed edits before touching the document", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["Alpha clause."] });
  const alpha = await find(handlers, handle, "Alpha");
  const before = (await handlers.readPart({ handle, partPath: "word/document.xml" })).text;

  const unknown = await failure(handlers.editDocument({
    handle,
    edits: [{ target: "T999", delete: true }]
  }));
  assert.equal(unknown.error.code, "TARGET_HANDLE_NOT_FOUND");

  const none = await failure(handlers.editDocument({ handle, edits: [{ target: alpha.target }] }));
  assert.equal(none.error.code, "INVALID_AGENT_REQUEST");
  assert.equal(none.error.recovery.action, "change_request");

  const two = await failure(handlers.editDocument({
    handle,
    edits: [{ target: alpha.target, delete: true, desiredText: "x" }]
  }));
  assert.equal(two.error.code, "INVALID_AGENT_REQUEST");
  assert.match(two.error.message, /exactly one/);

  const empty = await failure(handlers.editDocument({
    handle,
    edits: [{ target: alpha.target, replacements: [{ find: "", replace: "x" }] }]
  }));
  assert.equal(empty.error.field, "edits[0].replacements[0]");

  assert.equal((await handlers.readPart({ handle, partPath: "word/document.xml" })).text, before);
});

test("edit_document is atomic and reports which edit failed", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["Alpha clause.", "Beta clause."] });
  const alpha = await find(handlers, handle, "Alpha");
  const beta = await find(handlers, handle, "Beta");
  const before = (await handlers.readPart({ handle, partPath: "word/document.xml" })).text;

  const failed = await failure(handlers.editDocument({
    handle,
    edits: [
      { target: alpha.target, desiredText: "Alpha clause, amended." },
      { target: beta.target, replacements: [{ find: "text that is not there", replace: "x" }] }
    ]
  }));

  assert.equal(failed.status, "error");
  assert.equal(failed.completion, false);
  assert.equal(failed.rolledBack, true);
  assert.equal(failed.edits.length, 2);
  assert.equal(failed.edits[1].status, "error");
  assert.ok(failed.edits[1].error.code);
  assert.deepEqual(failed.retryPlan.failedIndexes, [2]);
  assert.equal((await handlers.readPart({ handle, partPath: "word/document.xml" })).text, before);
});

test("a refusal that needs user authorization is surfaced as such", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["Alpha clause.", "Beta clause."] });
  const beta = await find(handlers, handle, "Beta");
  await handlers.applyOperation({
    handle,
    operation: { type: "comment", targetHandle: beta.target, textToComment: "Beta", commentContent: "Open point" }
  });
  const fresh = await find(handlers, handle, "Beta");

  const failed = await failure(handlers.editDocument({
    handle,
    edits: [{ target: fresh.target, delete: true }]
  }));
  assert.equal(failed.error.code, "COMMENTED_CONTENT_DELETE", "the real cause is promoted over BATCH_OPERATION_FAILED");
  assert.equal(failed.error.recovery.requiresUserAuthorization, true);
  assert.equal(failed.error.operationIndex, 1);
});

test("apply_operations accepts targetHandle and targetEndHandle", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["Alpha clause.", "Beta clause."] });
  const alpha = await find(handlers, handle, "Alpha");

  const applied = await handlers.applyOperations({
    handle,
    operations: [{ type: "redline", targetHandle: alpha.target, modified: "Alpha clause, amended." }]
  });
  assert.equal(applied.completion, true);

  const stale = await failure(handlers.applyOperations({
    handle,
    operations: [{ type: "redline", targetHandle: alpha.target, modified: "again" }]
  }));
  assert.equal(stale.error.code, "STALE_TARGET_HANDLE");

  const preflight = await handlers.preflightOperations({
    handle,
    operations: [{ type: "redline", targetHandle: (await find(handlers, handle, "Beta")).target, modified: "Beta clause!" }]
  });
  assert.equal(preflight.valid, true);
});

test("restoreText needs a rejected-view handle, and plain edits refuse one", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["Keep this clause.", "Delete this clause.", "Closing clause."] });
  const doomed = await find(handlers, handle, "Delete this");
  await handlers.editDocument({ handle, author: "Other Party", edits: [{ target: doomed.target, delete: true }] });

  const accepted = await find(handlers, handle, "Keep");
  const wrongView = await failure(handlers.editDocument({
    handle,
    author: "Roo",
    edits: [{ target: accepted.target, restoreText: "Keep this clause." }]
  }));
  assert.equal(wrongView.error.code, "RESTORE_REQUIRES_REJECTED_VIEW");

  const deleted = await find(handlers, handle, "Delete this", { revisionView: "rejected" });
  const wrongIntent = await failure(handlers.editDocument({
    handle,
    edits: [{ target: deleted.target, desiredText: "Something else" }]
  }));
  assert.equal(wrongIntent.error.code, "EDIT_REQUIRES_ACCEPTED_VIEW");

  const restored = await handlers.editDocument({
    handle,
    author: "Roo",
    edits: [{ target: deleted.target, restoreText: "Delete this clause, but only after notice." }]
  });
  assert.equal(restored.completion, true);
  const inspection = await handlers.inspectDocument({ handle });
  assert.ok(inspection.paragraphs.some(paragraph => /only after notice/.test(paragraph.exactText)));
});

test("resolve_review refuses an unscoped request and asks for the user", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["Alpha clause."] });
  await handlers.applyOperation({
    handle,
    author: "Alice",
    operation: { type: "redline", target: "Alpha clause.", modified: "Alpha clause, amended." }
  });

  for (const request of [
    { action: "accept" },
    { action: "reject", author: "Alice", allAuthors: true },
    { action: "delete_comments" },
    { action: "resolve_comment" }
  ]) {
    const refused = await failure(handlers.resolveReview({ handle, ...request }));
    assert.equal(refused.error.code, "REVIEW_SCOPE_REQUIRED", JSON.stringify(request));
    assert.equal(refused.error.recovery.requiresUserAuthorization, true);
  }
  assert.deepEqual((await handlers.inspectDocument({ handle })).revisionAuthors, ["Alice"]);
});

test("resolve_review applies an explicit scope and retires handles", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["Alpha clause.", "Beta clause."] });
  await handlers.applyOperation({
    handle,
    author: "Alice",
    operation: { type: "redline", target: "Alpha clause.", modified: "Alpha clause, amended." }
  });
  await handlers.applyOperation({
    handle,
    author: "Bob",
    operation: { type: "comment", target: "Beta clause.", textToComment: "Beta", commentContent: "Bob's note" }
  });
  const alpha = await find(handlers, handle, "Alpha");

  const accepted = await handlers.resolveReview({ handle, action: "accept", author: "Alice" });
  assert.equal(accepted.completion, true);
  assert.equal(accepted.action, "accept");
  assert.deepEqual((await handlers.inspectDocument({ handle })).revisionAuthors, []);

  const stale = await failure(handlers.editDocument({
    handle,
    edits: [{ target: alpha.target, delete: true }]
  }));
  assert.equal(stale.error.code, "STALE_TARGET_HANDLE");

  const comment = (await handlers.inspectDocument({ handle })).comments[0];
  const resolved = await handlers.resolveReview({ handle, action: "resolve_comment", commentIds: [comment.id] });
  assert.equal(resolved.results[0].resolved, true);
  const reopened = await handlers.resolveReview({ handle, action: "reopen_comment", commentIds: [comment.id] });
  assert.equal(reopened.results[0].resolved, false);

  await handlers.resolveReview({ handle, action: "delete_comments", commentIds: [comment.id] });
  assert.equal((await handlers.inspectDocument({ handle })).comments.length, 0);
});

test("save_docx never replaces the source unless asked", async () => {
  const { handlers, handle, filePath } = await openFixture({ paragraphs: ["Alpha clause."] }, "contract.docx");
  const original = await readFile(filePath);
  await handlers.applyOperation({
    handle,
    operation: { type: "redline", target: "Alpha clause.", modified: "Alpha clause, amended." }
  });

  const first = await handlers.saveDocx({ handle });
  assert.equal(first.completion, true);
  assert.equal(first.replacedSource, false);
  assert.equal(basename(first.savedPath), "contract.reviewed.docx");
  assert.equal(dirname(first.savedPath), dirname(filePath));
  assert.equal(first.dirty, true, "the source file still differs from the working copy");
  assert.deepEqual(await readFile(filePath), original, "source is untouched");

  const second = await handlers.saveDocx({ handle });
  assert.equal(basename(second.savedPath), "contract.reviewed-2.docx", "an earlier copy is not overwritten");

  const explicit = await handlers.saveDocx({ handle, outputPath: join(dirname(filePath), "mine.docx") });
  assert.equal(basename(explicit.savedPath), "mine.docx");
  await access(explicit.savedPath);

  const zip = await JSZip.loadAsync(await readFile(first.savedPath));
  assert.match(await zip.file("word/document.xml").async("string"), /w:(ins|del)/);

  const replaced = await handlers.saveDocx({ handle, inPlace: true });
  assert.equal(replaced.replacedSource, true);
  assert.equal(replaced.dirty, false);
  assert.notDeepEqual(await readFile(filePath), original);
});

test("save_docx validates its arguments", async () => {
  const { handlers, handle, filePath } = await openFixture({ paragraphs: ["Alpha clause."] });
  const both = await failure(handlers.saveDocx({ handle, inPlace: true, outputPath: join(dirname(filePath), "x.docx") }));
  assert.equal(both.error.code, "INVALID_AGENT_REQUEST");
  const wrongExtension = await failure(handlers.saveDocx({ handle, outputPath: join(dirname(filePath), "x.txt") }));
  assert.equal(wrongExtension.error.field, "outputPath");
});

test("close_docx protects unsaved edits", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["Alpha clause."] });
  await handlers.applyOperation({
    handle,
    operation: { type: "redline", target: "Alpha clause.", modified: "Alpha clause, amended." }
  });

  const refused = await failure(handlers.closeDocx({ handle }));
  assert.equal(refused.error.code, "UNSAVED_CHANGES");
  assert.equal(refused.error.recovery.action, "save_docx");
  await handlers.inspectDocument({ handle });

  await handlers.saveDocx({ handle });
  const closed = await handlers.closeDocx({ handle });
  assert.equal(closed.closed, true);
  await assert.rejects(handlers.inspectDocument({ handle }), /Unknown handle/);

  const other = await openFixture({ paragraphs: ["Beta"] });
  assert.equal((await other.handlers.closeDocx({ handle: other.handle })).closed, true, "an unedited document closes freely");
});

test("close_docx can discard unsaved edits explicitly", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["Alpha clause."] });
  await handlers.applyOperation({
    handle,
    operation: { type: "redline", target: "Alpha clause.", modified: "Alpha clause, amended." }
  });
  const closed = await handlers.closeDocx({ handle, discardChanges: true });
  assert.equal(closed.discardedChanges, true);
});

test("a comment anchored across a tracked change is refused with the fix", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["Alpha clause.", "Beta clause."] });
  const alpha = await find(handlers, handle, "Alpha");
  const edited = await handlers.editDocument({
    handle,
    author: "Roo",
    edits: [{ target: alpha.target, replacements: [{ find: "Alpha", replace: "Gamma" }] }]
  });
  const refreshed = edited.refreshedTargets[0].handle;
  const before = (await handlers.readPart({ handle, partPath: "word/document.xml" })).text;

  const refused = await failure(handlers.editDocument({
    handle,
    author: "Roo",
    edits: [{ target: refreshed, comment: { content: "Check this" } }]
  }));
  assert.equal(refused.error.code, "COMMENT_ANCHOR_CROSSES_REVISION");
  assert.match(refused.error.message, /Anchor the comment to text outside the changed span/);
  assert.equal(refused.completion, false);
  assert.equal((await handlers.readPart({ handle, partPath: "word/document.xml" })).text, before);

  // Re-anchoring on the unchanged part of the paragraph works.
  const fresh = await find(handlers, handle, "clause");
  const fixed = await handlers.editDocument({
    handle,
    author: "Roo",
    edits: [{ target: fresh.target, comment: { content: "Check this", anchorText: "clause." } }]
  });
  assert.equal(fixed.completion, true);
});
