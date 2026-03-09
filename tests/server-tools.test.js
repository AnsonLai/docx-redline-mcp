import test from "node:test";
import assert from "node:assert/strict";
import { createToolHandlers } from "../src/server.js";

test("exposes the required tool handlers", () => {
  const handlers = createToolHandlers();
  assert.equal(typeof handlers.openDocx, "function");
  assert.equal(typeof handlers.listParts, "function");
  assert.equal(typeof handlers.readPart, "function");
  assert.equal(typeof handlers.writePart, "function");
  assert.equal(typeof handlers.applyOperation, "function");
  assert.equal(typeof handlers.applyOperations, "function");
  assert.equal(typeof handlers.saveDocx, "function");
});
