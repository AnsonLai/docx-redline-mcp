import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { writeDocxFixture } from "./helpers/fixtures.js";

const serverEntry = fileURLToPath(new URL("../src/index.js", import.meta.url));

async function withClient(run) {
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: [serverEntry],
    env: { ...process.env, DOCX_REDLINE_MCP_LOG_LEVEL: "info" },
    stderr: "pipe"
  });
  const client = new Client({ name: "test-client", version: "0.0.0" });
  await client.connect(transport);
  try {
    return await run(client);
  } finally {
    await client.close();
  }
}

test("stdio server lists the inspection, edit, review and apply tools", async () => {
  await withClient(async client => {
    const { tools } = await client.listTools();
    const names = tools.map(tool => tool.name).sort();
    assert.deepEqual(names, [
      "apply_operation",
      "apply_operations",
      "close_docx",
      "edit_document",
      "inspect_document",
      "list_parts",
      "open_docx",
      "preflight_operations",
      "read_part",
      "resolve_review",
      "save_docx",
      "write_part"
    ]);
  });
});

test("stdio server sends a launch card and honest tool annotations", async () => {
  await withClient(async client => {
    const instructions = client.getInstructions();
    assert.match(instructions, /open_docx -> inspect_document/);
    assert.match(instructions, /requiresUserAuthorization/);

    const { tools } = await client.listTools();
    const byName = Object.fromEntries(tools.map(tool => [tool.name, tool]));
    for (const name of ["open_docx", "inspect_document", "preflight_operations", "list_parts", "read_part"]) {
      assert.equal(byName[name].annotations.readOnlyHint, true, `${name} is read-only`);
    }
    for (const name of ["resolve_review", "save_docx", "close_docx", "write_part", "apply_operations"]) {
      assert.equal(byName[name].annotations.destructiveHint, true, `${name} is flagged destructive`);
    }
    assert.equal(byName.edit_document.annotations.destructiveHint, false);
    assert.equal(byName.edit_document.annotations.readOnlyHint, false);

    for (const tool of tools) {
      assert.ok(tool.title, `${tool.name} has a title`);
      assert.ok(tool.description.length > 20, `${tool.name} has a description`);
    }
    assert.match(byName.edit_document.description, /Atomic/);
    assert.match(byName.resolve_review.description, /explicit/);
  });
});

test("stdio server exposes the library schema and guides as resources", async () => {
  await withClient(async client => {
    const { resources } = await client.listResources();
    const uris = resources.map(resource => resource.uri);
    assert.ok(uris.includes("docx-redline://schema/document-operations"));
    assert.ok(uris.includes("docx-redline://docs/agent-fast-start"));

    const schema = await client.readResource({ uri: "docx-redline://schema/document-operations" });
    const parsed = JSON.parse(schema.contents[0].text);
    assert.equal(typeof parsed, "object");

    const guide = await client.readResource({ uri: "docx-redline://docs/agent-fast-start" });
    assert.match(guide.contents[0].text, /Agent Fast Start/);
  });
});

test("stdio server runs the ordinary edit workflow and reports failures as structured errors", async () => {
  const filePath = await writeDocxFixture("stdio.docx", { paragraphs: ["First clause.", "Second clause."] });

  await withClient(async client => {
    const call = async (name, args) => client.callTool({ name, arguments: args });

    const opened = (await call("open_docx", { path: filePath })).structuredContent;
    assert.ok(opened.handle);

    const inspected = (await call("inspect_document", { handle: opened.handle, search: "Second" })).structuredContent;
    const target = inspected.paragraphs[0];
    assert.match(target.target, /^T\d+$/);

    const edited = await call("edit_document", {
      handle: opened.handle,
      author: "Roo",
      edits: [{ target: target.target, replacements: [{ find: "Second", replace: "Revised" }] }]
    });
    assert.notEqual(edited.isError, true);
    assert.equal(edited.structuredContent.status, "ok");
    assert.equal(edited.structuredContent.completion, true);
    assert.equal(edited.structuredContent.refreshedTargets[0].previousHandle, target.target);

    const stale = await call("edit_document", {
      handle: opened.handle,
      edits: [{ target: target.target, delete: true }]
    });
    assert.equal(stale.isError, true);
    assert.equal(stale.structuredContent.error.code, "STALE_TARGET_HANDLE");
    assert.equal(stale.structuredContent.error.recovery.requiresReinspection, true);

    const failed = await call("apply_operations", {
      handle: opened.handle,
      operations: [{ type: "redline", target: "No such paragraph", modified: "x" }]
    });
    assert.equal(failed.isError, true);
    assert.equal(failed.structuredContent.status, "error");
    assert.ok(failed.structuredContent.error.code);
    assert.deepEqual(failed.structuredContent.retryPlan.failedIndexes, [1]);

    const mismatch = await call("apply_operation", {
      handle: opened.handle,
      expectedRevision: inspected.packageRevision,
      operation: { type: "redline", target: "First clause.", modified: "First changed." }
    });
    assert.equal(mismatch.isError, true);
    assert.equal(mismatch.structuredContent.error.code, "REVISION_MISMATCH");

    const refused = await call("resolve_review", { handle: opened.handle, action: "accept" });
    assert.equal(refused.isError, true);
    assert.equal(refused.structuredContent.error.code, "REVIEW_SCOPE_REQUIRED");

    const saved = (await call("save_docx", { handle: opened.handle })).structuredContent;
    assert.match(saved.savedPath, /stdio\.reviewed\.docx$/);
    assert.equal(saved.completion, true);
  });
});
