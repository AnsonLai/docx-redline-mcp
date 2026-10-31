import { createRequire } from "node:module";
import * as z from "zod/v4";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { createSessionStore } from "./lib/session-store.js";
import { configureDocxRedlineLoggerForMcp } from "./lib/mcp-logger.js";
import { LIBRARY_DOCUMENTS, readLibraryDocument } from "./lib/library-docs.js";
import { OperationFailedError, MAINTENANCE_OPERATION_TYPES } from "./lib/operation-service.js";
import { createApplyOperationHandler, createApplyOperationsHandler } from "./tools/apply-operation.js";
import { createInspectDocumentHandler } from "./tools/inspect-document.js";
import { createListPartsHandler } from "./tools/list-parts.js";
import { createOpenDocxHandler } from "./tools/open-docx.js";
import { createPreflightOperationsHandler } from "./tools/preflight-operations.js";
import { createReadPartHandler } from "./tools/read-part.js";
import { createSaveDocxHandler } from "./tools/save-docx.js";
import { createWritePartHandler } from "./tools/write-part.js";

const require = createRequire(import.meta.url);
export const SERVER_VERSION = require("../package.json").version;

function toToolResult(payload) {
  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(payload, null, 2)
      }
    ],
    structuredContent: payload
  };
}

/**
 * Wraps a handler so failures that carry a structured payload (error code, recovery guidance,
 * per-operation results, retry plan) reach the client as an error result with that payload,
 * instead of a bare message.
 */
function toolCallback(handler) {
  return async args => {
    try {
      return toToolResult(await handler(args));
    } catch (error) {
      if (error instanceof OperationFailedError) {
        const payload = { message: error.message, ...error.payload };
        return { ...toToolResult(payload), isError: true };
      }
      throw error;
    }
  };
}

const revisionTokenSchema = z.object({
  algorithm: z.literal("sha256"),
  version: z.literal(1),
  scope: z.enum(["package", "document-parts"]),
  value: z.string().min(1),
  coveredParts: z.array(z.string()).optional()
}).describe("A packageRevision token returned by open_docx, inspect_document or a previous apply.");

const applyOptionsSchema = z.object({
  atomic: z.boolean().optional().describe(
    "Default true: any failing operation rolls the whole batch back. Set false to keep the operations that succeeded (status \"partial\")."
  ),
  strictTargets: z.boolean().optional().describe(
    "Default true: targets must match exactly (exactText, paragraphId, fingerprint). Set false only to allow looser text matching."
  ),
  existingRevisions: z.enum([
    "merge-same-author",
    "slice-cross-author",
    "auto",
    "accept-all-first",
    "reject-input"
  ]).optional().describe(
    "How to edit paragraphs that already contain tracked changes. Default merge-same-author refuses to touch another author's pending changes; slice-cross-author edits inside them; auto chooses per target."
  ),
  generateRedlines: z.boolean().optional().describe(
    "Default true (tracked changes). false applies clean direct edits with no revision markup."
  ),
  structuredContent: z.boolean().optional().describe(
    "Default true: Markdown tables, headings and lists in replacement text become native Word structures. false keeps replacement text literal."
  ),
  verifyRevisionFidelity: z.enum(["enforce", "warn"]).optional().describe(
    "Default enforce: refuse any edit whose Reject All would not restore the original. warn applies it and reports the violation."
  ),
  strictListConversion: z.boolean().optional().describe(
    "Make unsupported list conversions errors instead of warnings (defaults to atomic)."
  ),
  pairReplacements: z.boolean().optional(),
  removeFormatting: z.boolean().optional(),
  continueOnError: z.boolean().optional()
}).describe("Execution policy for the batch.");

const OPERATION_GUIDE = [
  "Each operation is an object with a `type`. Common types: redline (modified text or localized `replacements`), insert, delete, restore, comment,",
  "comment_reply, comment_resolve, format, character-format, highlight, paragraph-format, list-format, list-change and table-reconciliation.",
  "Target paragraphs with the `target` descriptor from inspect_document ({ exactText, paragraphId, fingerprint }) rather than free text.",
  "Set `part` to edit a header or footer listed by inspect_document.headersFooters.",
  `Package-level maintenance operations: ${MAINTENANCE_OPERATION_TYPES.join(", ")}.`,
  "Accept/reject/delete operations need an explicit request from the user; scope them with `author` (or `ids` for comments).",
  "The full schema is the docx-redline://schema/document-operations resource."
].join(" ");

function createGetSession(store) {
  return function getSession(handle) {
    const session = store.get(handle);
    if (!session) {
      throw new Error(`Unknown handle: ${String(handle || "")}`);
    }
    return session;
  };
}

export function createToolHandlers(deps = {}) {
  configureDocxRedlineLoggerForMcp();

  const store = deps.store || createSessionStore();
  const getSession = deps.getSession || createGetSession(store);

  return {
    openDocx: createOpenDocxHandler({ store }),
    listParts: createListPartsHandler({ getSession }),
    readPart: createReadPartHandler({ getSession }),
    writePart: createWritePartHandler({ getSession, store }),
    inspectDocument: createInspectDocumentHandler({ getSession }),
    preflightOperations: createPreflightOperationsHandler({ getSession }),
    applyOperation: createApplyOperationHandler({ getSession, store }),
    applyOperations: createApplyOperationsHandler({ getSession, store }),
    saveDocx: createSaveDocxHandler({ getSession, store }),
    store
  };
}

export function createMcpServer() {
  const handlers = createToolHandlers();
  const server = new McpServer({
    name: "docx-redline-mcp",
    version: SERVER_VERSION
  });

  server.registerTool("open_docx", {
    description: "Open a DOCX file and create a working document handle. Returns a summary (paragraph and comment counts, existing reviewers, headers/footers) and the packageRevision token.",
    inputSchema: {
      path: z.string().min(1).describe("Absolute or relative path to a .docx file")
    }
  }, toolCallback(handlers.openDocx));

  server.registerTool("list_parts", {
    description: "List package parts for an open DOCX handle.",
    inputSchema: {
      handle: z.string().min(1).describe("Handle returned by open_docx")
    }
  }, toolCallback(handlers.listParts));

  server.registerTool("read_part", {
    description: "Read an OOXML part as XML text or get a read-only binary summary.",
    inputSchema: {
      handle: z.string().min(1),
      partPath: z.string().min(1),
      maxChars: z.number().int().positive().optional(),
      includeBase64: z.boolean().optional()
    }
  }, toolCallback(handlers.readPart));

  server.registerTool("write_part", {
    description: "Replace a writable XML part in the working DOCX package.",
    inputSchema: {
      handle: z.string().min(1),
      partPath: z.string().min(1),
      xml: z.string().min(1)
    }
  }, toolCallback(handlers.writePart));

  server.registerTool("inspect_document", {
    description: [
      "Inspect the open document: exact paragraph text with the target descriptors (paragraphId, fingerprint) to use in operations,",
      "headings, table/list context, tracked-change authors, comments with their anchors, and headers/footers.",
      "Filter with search (case-insensitive), around, indexes, range, revisedOnly, inTable; page with limit/after (follow selection.nextAfter).",
      "Use revisionView \"rejected\" to find deleted text to restore. The returned packageRevision can be passed to apply_operations as expectedRevision."
    ].join(" "),
    inputSchema: {
      handle: z.string().min(1),
      search: z.string().optional().describe("Case-insensitive substring to find"),
      around: z.number().int().min(0).max(20).optional().describe("Paragraphs of context around each match"),
      indexes: z.array(z.number().int().positive()).optional(),
      range: z.object({
        start: z.number().int().positive(),
        end: z.number().int().positive()
      }).optional(),
      revisedOnly: z.boolean().optional().describe("Only paragraphs that contain tracked changes"),
      inTable: z.boolean().optional(),
      skipEmpty: z.boolean().optional(),
      revisionView: z.enum(["accepted", "rejected", "current"]).optional(),
      limit: z.number().int().positive().optional().describe("Direct matches per page (default 100)"),
      after: z.number().int().min(0).optional().describe("Paging cursor from selection.nextAfter"),
      excerptLength: z.number().int().positive().optional(),
      compact: z.boolean().optional().describe("Default true: omit fields that repeat other fields")
    }
  }, toolCallback(handlers.inspectDocument));

  server.registerTool("preflight_operations", {
    description: `Dry-run a batch of operations against the open document without changing it: resolves each target, reports conflicts, the author(s) that would be used and the comment/numbering artifacts required. ${OPERATION_GUIDE}`,
    inputSchema: {
      handle: z.string().min(1),
      operations: z.array(z.record(z.string(), z.unknown())).min(1),
      author: z.string().min(1).optional(),
      options: applyOptionsSchema.optional()
    }
  }, toolCallback(handlers.preflightOperations));

  server.registerTool("apply_operation", {
    description: `Apply one structured operation to the open document. ${OPERATION_GUIDE}`,
    inputSchema: {
      handle: z.string().min(1),
      operation: z.record(z.string(), z.unknown()),
      author: z.string().min(1).optional(),
      options: applyOptionsSchema.optional(),
      expectedRevision: revisionTokenSchema.optional()
    }
  }, toolCallback(handlers.applyOperation));

  server.registerTool("apply_operations", {
    description: `Apply several operations in a single batch. Targets bind at the start of the batch, so every target must come from one inspection. Atomic by default: on any failure nothing is applied and the result carries error.recovery and retryPlan. ${OPERATION_GUIDE}`,
    inputSchema: {
      handle: z.string().min(1),
      operations: z.array(z.record(z.string(), z.unknown())).min(1),
      author: z.string().min(1).optional(),
      options: applyOptionsSchema.optional(),
      expectedRevision: revisionTokenSchema.optional()
    }
  }, toolCallback(handlers.applyOperations));

  server.registerTool("save_docx", {
    description: "Save the working DOCX package in place or to an explicit output path.",
    inputSchema: {
      handle: z.string().min(1),
      outputPath: z.string().min(1).optional()
    }
  }, toolCallback(handlers.saveDocx));

  for (const document of LIBRARY_DOCUMENTS) {
    server.registerResource(document.name, document.uri, {
      title: document.title,
      description: document.description,
      mimeType: document.mimeType
    }, async uri => ({
      contents: [{
        uri: uri.href,
        mimeType: document.mimeType,
        text: await readLibraryDocument(document)
      }]
    }));
  }

  return { server, handlers };
}

export async function startStdioServer() {
  const { server } = createMcpServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  return server;
}
