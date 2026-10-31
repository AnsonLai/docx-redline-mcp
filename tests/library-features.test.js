import test from "node:test";
import assert from "node:assert/strict";
import { createToolHandlers } from "../src/server.js";
import { OperationFailedError } from "../src/lib/operation-service.js";
import { writeDocxFixture } from "./helpers/fixtures.js";

async function openFixture(options, name = "fixture.docx") {
  const handlers = createToolHandlers();
  const filePath = await writeDocxFixture(name, options);
  const opened = await handlers.openDocx({ path: filePath });
  return { handlers, opened, handle: opened.handle, filePath };
}

async function documentXml(handlers, handle) {
  return (await handlers.readPart({ handle, partPath: "word/document.xml" })).text;
}

test("open_docx returns a package revision and document summary", async () => {
  const { opened } = await openFixture({ paragraphs: ["One.", "Two.", ""], header: "Running header" });

  assert.equal(opened.packageRevision.scope, "package");
  assert.equal(opened.summary.counts.paragraphs, 3);
  assert.deepEqual(opened.summary.revisionAuthors, []);
  assert.equal(opened.summary.headersFooters.length, 1);
  assert.equal(opened.summary.headersFooters[0].kind, "header");
});

test("open_docx summary covers the whole document, not just a page", async () => {
  const { handlers, handle, filePath } = await openFixture({ paragraphs: ["Alpha.", "Beta.", "Gamma."] });
  await handlers.applyOperation({
    handle,
    author: "Late Reviewer",
    operation: { type: "comment", target: "Gamma.", textToComment: "Gamma", commentContent: "Note" }
  });
  await handlers.applyOperation({
    handle,
    author: "Late Reviewer",
    operation: { type: "redline", target: "Beta.", modified: "Beta changed." }
  });
  await handlers.saveDocx({ handle });

  const reopened = await handlers.openDocx({ path: filePath });
  assert.deepEqual(reopened.summary.counts, { paragraphs: 3, revisedParagraphs: 1, comments: 1 });
  assert.deepEqual(reopened.summary.revisionAuthors, ["Late Reviewer"]);
  assert.deepEqual(reopened.summary.commentAuthors, ["Late Reviewer"]);
});

test("inspect_document returns compact, targetable paragraphs and filters by search", async () => {
  const { handlers, handle } = await openFixture({
    paragraphs: ["Introduction.", "The Supplier may terminate this Agreement.", "Closing remarks."]
  });

  const inspection = await handlers.inspectDocument({ handle, search: "terminate", around: 1 });

  assert.equal(inspection.status, "ok");
  assert.equal(inspection.packageRevision.scope, "package");
  assert.equal(inspection.selection.totalMatches, 1);
  assert.equal(inspection.paragraphs.length, 3);
  const match = inspection.paragraphs.find(paragraph => /terminate/.test(paragraph.exactText));
  assert.ok(match.fingerprint);
  assert.equal(match.segments, undefined, "segments are omitted for paragraphs without revisions");
  assert.equal(match.text, undefined, "text duplicating exactText is omitted");
  assert.deepEqual(inspection.headersFooters, []);

  const full = await handlers.inspectDocument({ handle, search: "terminate", compact: false });
  assert.ok(full.paragraphs[0].segments);
});

test("inspect_document pages results with a default limit", async () => {
  const paragraphs = Array.from({ length: 130 }, (_, index) => `Paragraph number ${index + 1}.`);
  const { handlers, handle } = await openFixture({ paragraphs });

  const first = await handlers.inspectDocument({ handle });
  assert.equal(first.paragraphs.length, 100);
  assert.equal(first.selection.truncated, true);

  const second = await handlers.inspectDocument({ handle, after: first.selection.nextAfter });
  assert.equal(second.paragraphs.length, 30);
});

test("inspect_document lists headers and footers and revision authors", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["Body."], header: "Confidential" });
  await handlers.applyOperation({
    handle,
    author: "Reviewer A",
    operation: { type: "redline", target: "Body.", modified: "Body text." }
  });

  const inspection = await handlers.inspectDocument({ handle });
  assert.deepEqual(inspection.revisionAuthors, ["Reviewer A"]);
  assert.equal(inspection.headersFooters[0].paragraphs[0].text, "Confidential");
  assert.equal(inspection.paragraphs[0].hasRevisions, true);
  assert.ok(inspection.paragraphs[0].segments.some(segment => segment.kind !== "baseline"));
});

test("preflight_operations resolves targets without changing the document", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["Original text", "Other text"] });
  const before = await documentXml(handlers, handle);

  const ok = await handlers.preflightOperations({
    handle,
    author: "Roo",
    operations: [{ type: "redline", target: "Original text", modified: "Updated text" }]
  });
  assert.equal(ok.valid, true);
  assert.equal(ok.results[0].status, "ready");
  assert.deepEqual(ok.authorsUsed, ["Roo"]);

  const missing = await handlers.preflightOperations({
    handle,
    operations: [{ type: "redline", target: "Does not exist", modified: "x" }]
  });
  assert.equal(missing.valid, false);
  assert.equal(missing.results[0].status, "error");

  assert.equal(await documentXml(handlers, handle), before);
  await assert.rejects(
    handlers.preflightOperations({ handle, operations: [{ type: "delete_all_comments" }] }),
    /maintenance operation/
  );
});

test("apply_operations reports per-operation results, receipts and the new revision", async () => {
  const { handlers, handle, opened } = await openFixture({ paragraphs: ["First clause.", "Second clause."] });

  const applied = await handlers.applyOperations({
    handle,
    author: "Roo",
    operations: [
      { type: "redline", target: "First clause.", modified: "First amended clause." },
      { type: "comment", target: "Second clause.", textToComment: "Second", commentContent: "Check this" }
    ]
  });

  assert.equal(applied.status, "ok");
  assert.equal(applied.complete, true);
  assert.equal(applied.dirty, true);
  assert.equal(applied.results.length, 2);
  assert.deepEqual(applied.results.map(result => result.index), [1, 2]);
  assert.deepEqual(applied.receipts.map(receipt => receipt.operationIndex), [1, 2]);
  assert.ok(applied.artifactsChanged.includes("word/comments.xml"));
  assert.notEqual(applied.packageRevision.value, opened.packageRevision.value);

  const inspection = await handlers.inspectDocument({ handle });
  assert.equal(inspection.packageRevision.value, applied.packageRevision.value);
  assert.equal(inspection.comments.length, 1);
});

test("apply_operations is atomic by default and surfaces recovery details", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["First clause.", "Second clause."] });
  const before = await documentXml(handlers, handle);

  const error = await handlers.applyOperations({
    handle,
    operations: [
      { type: "redline", target: "First clause.", modified: "First amended clause." },
      { type: "redline", target: "Missing clause.", modified: "Anything" }
    ]
  }).then(() => null, caught => caught);

  assert.ok(error instanceof OperationFailedError);
  assert.equal(error.payload.status, "error");
  assert.equal(error.payload.rolledBack, true);
  assert.ok(error.payload.error.code);
  assert.equal(error.payload.retryPlan.base, "original");
  assert.deepEqual(error.payload.retryPlan.failedIndexes, [2]);
  assert.equal(await documentXml(handlers, handle), before, "nothing is applied");
});

test("apply_operations with atomic false keeps the operations that succeeded", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["First clause.", "Second clause."] });

  const applied = await handlers.applyOperations({
    handle,
    options: { atomic: false },
    operations: [
      { type: "redline", target: "First clause.", modified: "First amended clause." },
      { type: "redline", target: "Missing clause.", modified: "Anything" }
    ]
  });

  assert.equal(applied.status, "partial");
  assert.equal(applied.complete, false);
  assert.equal(applied.hasChanges, true);
  assert.deepEqual(applied.results.map(result => result.status), ["applied", "error"]);
  assert.match(await documentXml(handlers, handle), /w:(ins|del)/);
});

test("expectedRevision rejects edits against a document that changed since inspection", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["First clause.", "Second clause."] });
  const inspected = await handlers.inspectDocument({ handle });

  await handlers.applyOperation({
    handle,
    operation: { type: "redline", target: "First clause.", modified: "First amended clause." }
  });

  const error = await handlers.applyOperation({
    handle,
    expectedRevision: inspected.packageRevision,
    operation: { type: "redline", target: "Second clause.", modified: "Second amended clause." }
  }).then(() => null, caught => caught);

  assert.ok(error instanceof OperationFailedError);
  assert.equal(error.payload.error.code, "REVISION_MISMATCH");
  assert.equal(error.payload.error.recovery.requiresReinspection, true);

  const fresh = await handlers.inspectDocument({ handle });
  const applied = await handlers.applyOperation({
    handle,
    expectedRevision: fresh.packageRevision,
    operation: { type: "redline", target: "Second clause.", modified: "Second amended clause." }
  });
  assert.equal(applied.status, "ok");
});

test("operations can target inspected paragraphs by id and fingerprint", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["First clause.", "Second clause."] });
  const inspection = await handlers.inspectDocument({ handle, search: "Second" });
  const paragraph = inspection.paragraphs[0];

  const applied = await handlers.applyOperation({
    handle,
    operation: {
      type: "redline",
      target: { exactText: paragraph.exactText, fingerprint: paragraph.fingerprint, index: paragraph.index },
      replacements: [{ find: "Second", replace: "Third" }]
    }
  });

  assert.equal(applied.status, "ok");
  const xml = await documentXml(handlers, handle);
  assert.match(xml, /Third/);
  assert.match(xml, /Second/);
});

test("operations can edit a header through the part selector", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["Body."], header: "Draft version" });

  const applied = await handlers.applyOperation({
    handle,
    author: "Roo",
    operation: {
      type: "redline",
      part: { kind: "header", type: "default" },
      target: { exactText: "Draft version" },
      modified: "Final version"
    }
  });

  assert.equal(applied.status, "ok");
  assert.ok(applied.artifactsChanged.includes("word/header1.xml"));
  const header = await handlers.readPart({ handle, partPath: "word/header1.xml" });
  assert.match(header.text, /w:(ins|del)/);
});

test("accept and reject can be limited to one author", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["Alpha text.", "Beta text."] });
  await handlers.applyOperation({
    handle,
    author: "Alice",
    operation: { type: "redline", target: "Alpha text.", modified: "Alpha changed." }
  });
  await handlers.applyOperation({
    handle,
    author: "Bob",
    operation: { type: "redline", target: "Beta text.", modified: "Beta changed." }
  });

  const accepted = await handlers.applyOperation({
    handle,
    operation: { type: "accept_all_tracked_changes", author: "Alice" }
  });
  assert.equal(accepted.results[0].status, "applied");
  assert.equal(accepted.results[0].author, "Alice");

  const inspection = await handlers.inspectDocument({ handle });
  assert.deepEqual(inspection.revisionAuthors, ["Bob"]);

  await handlers.applyOperation({ handle, operation: { type: "reject_all_tracked_changes" } });
  const afterReject = await handlers.inspectDocument({ handle });
  assert.deepEqual(afterReject.revisionAuthors, []);
  assert.equal(afterReject.paragraphs[1].exactText, "Beta text.");
});

test("delete_all_comments can target one author or specific ids, and comments can be resolved", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["Alpha text.", "Beta text."] });
  await handlers.applyOperation({
    handle,
    author: "Alice",
    operation: { type: "comment", target: "Alpha text.", textToComment: "Alpha", commentContent: "From Alice" }
  });
  await handlers.applyOperation({
    handle,
    author: "Bob",
    operation: { type: "comment", target: "Beta text.", textToComment: "Beta", commentContent: "From Bob" }
  });

  const before = await handlers.inspectDocument({ handle });
  assert.equal(before.comments.length, 2);
  const bobComment = before.comments.find(comment => comment.author === "Bob");

  const resolved = await handlers.applyOperation({
    handle,
    operation: { type: "resolve_comment", commentId: bobComment.id }
  });
  assert.equal(resolved.results[0].status, "applied");
  assert.equal(resolved.results[0].resolved, true);
  const afterResolve = await handlers.inspectDocument({ handle });
  assert.equal(afterResolve.comments.find(comment => comment.author === "Bob").done, true);

  await handlers.applyOperation({ handle, operation: { type: "delete_all_comments", author: "Alice" } });
  const afterAuthor = await handlers.inspectDocument({ handle });
  assert.deepEqual(afterAuthor.comments.map(comment => comment.author), ["Bob"]);

  await handlers.applyOperation({
    handle,
    operation: { type: "delete_all_comments", ids: [bobComment.id] }
  });
  assert.equal((await handlers.inspectDocument({ handle })).comments.length, 0);
});

test("maintenance and edit operations can share one batch", async () => {
  const { handlers, handle } = await openFixture({ paragraphs: ["Alpha text.", "Beta text."] });
  await handlers.applyOperation({
    handle,
    author: "Alice",
    operation: { type: "redline", target: "Alpha text.", modified: "Alpha changed." }
  });

  const applied = await handlers.applyOperations({
    handle,
    author: "Bob",
    operations: [
      { type: "accept_all_tracked_changes" },
      { type: "redline", target: "Beta text.", modified: "Beta changed." }
    ]
  });

  assert.equal(applied.status, "ok");
  assert.deepEqual(applied.results.map(result => result.index), [1, 2]);
  const inspection = await handlers.inspectDocument({ handle });
  assert.deepEqual(inspection.revisionAuthors, ["Bob"]);
  assert.equal(inspection.paragraphs[0].exactText, "Alpha changed.");
});
