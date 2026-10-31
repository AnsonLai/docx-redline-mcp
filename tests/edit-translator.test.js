import test from "node:test";
import assert from "node:assert/strict";
import { expandOperationHandles, expectedTextAfter, translateEdits } from "../src/lib/edit-translator.js";
import { createTargetRegistry, registerParagraph, resolveHandle } from "../src/lib/target-registry.js";

function registryWith(paragraphs, revision = "rev1") {
  const registry = createTargetRegistry();
  const handles = paragraphs.map(paragraph => registerParagraph(registry, paragraph, revision));
  return { registry, handles, resolve: name => resolveHandle(registry, name, revision) };
}

const paragraph = (index, exactText, extra = {}) => ({
  index,
  exactText,
  fingerprint: `fnv1a32:${index}`,
  revisionView: "accepted",
  inTable: false,
  humanReference: exactText.slice(0, 10),
  ...extra
});

test("registers paragraphs once per identity and binds handles to a revision", () => {
  const { registry, handles, resolve } = registryWith([paragraph(1, "One"), paragraph(2, "Two")]);
  assert.deepEqual(handles, ["T1", "T2"]);
  assert.equal(registerParagraph(registry, paragraph(1, "One"), "rev1"), "T1", "same paragraph keeps its handle");
  assert.equal(resolve("T2").entry.descriptor.exactText, "Two");
  assert.equal(resolveHandle(registry, "T1", "rev2").error.code, "STALE_TARGET_HANDLE");
  assert.equal(resolveHandle(registry, "T9", "rev1").error.code, "TARGET_HANDLE_NOT_FOUND");
});

test("a new revision retires older handles but still recognizes them as stale", () => {
  const { registry } = registryWith([paragraph(1, "One")]);
  const next = registerParagraph(registry, paragraph(1, "One!"), "rev2");
  assert.equal(next, "T2");
  assert.equal(resolveHandle(registry, "T1", "rev2").error.code, "STALE_TARGET_HANDLE");
  assert.ok(resolveHandle(registry, "T2", "rev2").entry);
});

test("descriptors use the paragraph id when there is one and the index otherwise", () => {
  const { resolve } = registryWith([
    paragraph(1, "One", { paragraphId: "AB12" }),
    paragraph(2, "Two"),
    paragraph(3, "Three", { revisionView: "rejected" })
  ]);
  assert.equal(resolve("T1").entry.descriptor.paragraphId, "AB12");
  assert.equal(resolve("T1").entry.descriptor.index, undefined);
  assert.equal(resolve("T2").entry.descriptor.index, 2);
  assert.equal(resolve("T3").entry.descriptor.revisionView, "rejected");
});

test("translates each intent into a canonical operation", () => {
  const { resolve } = registryWith([paragraph(1, "The Supplier may terminate.")]);
  const { operations, sources } = translateEdits([
    { target: "T1", replacements: [{ find: "may", replace: "must" }] },
    { target: "T1", desiredText: "Rewritten." },
    { target: "T1", delete: true },
    { target: "T1", comment: { content: "Why?", anchorText: "Supplier" }, existingRevisions: "auto" }
  ], resolve);

  assert.deepEqual(operations.map(operation => operation.type), ["redline", "redline", "delete", "comment"]);
  assert.deepEqual(operations[0].replacements, [{ find: "may", replace: "must" }]);
  assert.equal(operations[1].modified, "Rewritten.");
  assert.equal(operations[3].textToComment, "Supplier");
  assert.equal(operations[3].commentContent, "Why?");
  assert.equal(operations[3].existingRevisions, "auto");
  assert.equal(operations[0].operationId, "edit-1");
  assert.equal(operations[0].target.exactText, "The Supplier may terminate.");
  assert.notEqual(operations[0].target, operations[1].target, "descriptors are copied, not shared");
  assert.deepEqual(sources.map(source => source.intent), ["replacements", "desiredText", "delete", "comment"]);
});

test("expectedTextAfter predicts the edited paragraph text", () => {
  const source = { exactText: "The Supplier may terminate this Agreement.", intent: "replacements" };
  assert.equal(
    expectedTextAfter({ ...source, replacements: [{ find: "may", replace: "must" }, { find: "Agreement", replace: "contract" }] }),
    "The Supplier must terminate this contract."
  );
  assert.equal(expectedTextAfter({ ...source, replacements: [{ find: "e", replace: "E", occurrence: 2 }] }), "The SuppliEr may terminate this Agreement.");
  assert.equal(expectedTextAfter({ ...source, replacements: [{ find: "nope", replace: "x" }] }), null);
  assert.equal(expectedTextAfter({ ...source, intent: "desiredText", desiredText: "New." }), "New.");
  assert.equal(expectedTextAfter({ ...source, intent: "comment" }), source.exactText);
  assert.equal(expectedTextAfter({ ...source, intent: "delete" }), null);
});

test("expandOperationHandles replaces handles with descriptors and leaves other operations alone", () => {
  const { resolve } = registryWith([paragraph(1, "One"), paragraph(2, "Two")]);
  const plain = { type: "redline", target: "Free text", modified: "x" };
  const [expanded, untouched] = expandOperationHandles([
    { type: "redline", targetHandle: "T1", targetEndHandle: "T2", modified: "x" },
    plain
  ], resolve);

  assert.equal(expanded.target.exactText, "One");
  assert.equal(expanded.targetEnd.exactText, "Two");
  assert.equal("targetHandle" in expanded, false);
  assert.equal(untouched, plain);
});
