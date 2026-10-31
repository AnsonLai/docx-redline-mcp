import { readFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

/**
 * Reference documents shipped inside @ansonlai/docx-redline-js, exposed to clients as MCP
 * resources so an agent can read the canonical operation schema and workflow guidance for the
 * exact library release this server runs.
 */
export const LIBRARY_DOCUMENTS = Object.freeze([
  {
    name: "document-operations-schema",
    uri: "docx-redline://schema/document-operations",
    file: "docs/schemas/document-operations.schema.json",
    title: "Document operations JSON schema",
    description: "Canonical JSON schema for every operation accepted by apply_operation(s) and preflight_operations.",
    mimeType: "application/json"
  },
  {
    name: "agent-fast-start",
    uri: "docx-redline://docs/agent-fast-start",
    file: "docs/AGENT_FAST_START.md",
    title: "Agent fast start",
    description: "Inspect once, apply once: the recommended editing workflow, completion and recovery rules.",
    mimeType: "text/markdown"
  },
  {
    name: "agent-knowledge-base",
    uri: "docx-redline://docs/agent-knowledge-base",
    file: "docs/AGENT_KNOWLEDGE_BASE.md",
    title: "Agent knowledge base",
    description: "Detailed operation shapes, list and table handling, revision policies and error recovery.",
    mimeType: "text/markdown"
  }
]);

function libraryRoot() {
  return dirname(fileURLToPath(import.meta.resolve("@ansonlai/docx-redline-js")));
}

export async function readLibraryDocument(document) {
  return readFile(join(libraryRoot(), document.file), "utf8");
}
