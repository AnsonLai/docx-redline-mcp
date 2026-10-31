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

test("stdio server lists the inspection, preflight and apply tools", async () => {
  await withClient(async client => {
    const { tools } = await client.listTools();
    const names = tools.map(tool => tool.name).sort();
    assert.deepEqual(names, [
      "apply_operation",
      "apply_operations",
      "inspect_document",
      "list_parts",
      "open_docx",
      "preflight_operations",
      "read_part",
      "save_docx",
      "write_part"
    ]);
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

test("stdio server keeps stdout clean while editing and reports failures as structured errors", async () => {
  const filePath = await writeDocxFixture("stdio.docx", { paragraphs: ["First clause.", "Second clause."] });

  await withClient(async client => {
    const opened = (await client.callTool({ name: "open_docx", arguments: { path: filePath } })).structuredContent;
    assert.ok(opened.handle);

    const inspected = (await client.callTool({
      name: "inspect_document",
      arguments: { handle: opened.handle, search: "Second" }
    })).structuredContent;
    const target = inspected.paragraphs[0];

    const applied = await client.callTool({
      name: "apply_operations",
      arguments: {
        handle: opened.handle,
        author: "Roo",
        expectedRevision: inspected.packageRevision,
        operations: [{
          type: "redline",
          target: { exactText: target.exactText, fingerprint: target.fingerprint },
          replacements: [{ find: "Second", replace: "Revised" }]
        }]
      }
    });
    assert.notEqual(applied.isError, true);
    assert.equal(applied.structuredContent.status, "ok");

    const failed = await client.callTool({
      name: "apply_operations",
      arguments: {
        handle: opened.handle,
        operations: [{ type: "redline", target: "No such paragraph", modified: "x" }]
      }
    });
    assert.equal(failed.isError, true);
    assert.equal(failed.structuredContent.status, "error");
    assert.ok(failed.structuredContent.error.code);
    assert.deepEqual(failed.structuredContent.retryPlan.failedIndexes, [1]);

    const stale = await client.callTool({
      name: "apply_operation",
      arguments: {
        handle: opened.handle,
        expectedRevision: inspected.packageRevision,
        operation: { type: "redline", target: "First clause.", modified: "First changed." }
      }
    });
    assert.equal(stale.isError, true);
    assert.equal(stale.structuredContent.error.code, "REVISION_MISMATCH");
  });
});
