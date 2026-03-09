import test from "node:test";
import assert from "node:assert/strict";
import { startServer } from "../src/index.js";

test("exports startServer", () => {
  assert.equal(typeof startServer, "function");
});
