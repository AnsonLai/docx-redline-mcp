import test from "node:test";
import assert from "node:assert/strict";
import { createSessionStore } from "../src/lib/session-store.js";

test("creates, reads, updates, and deletes document sessions", () => {
  const store = createSessionStore();
  const session = store.create({
    sourcePath: "C:/docs/sample.docx",
    zip: { fake: true }
  });

  assert.match(session.handle, /^doc_/);
  assert.equal(store.get(session.handle).sourcePath, "C:/docs/sample.docx");

  store.update(session.handle, current => ({ ...current, dirty: true }));
  assert.equal(store.get(session.handle).dirty, true);

  store.delete(session.handle);
  assert.equal(store.get(session.handle), null);
});
