import { createRequire } from "node:module";
import * as z from "zod/v4";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { warnOnLibraryMismatch } from "./lib/compatibility.js";
import { OperationFailedError } from "./lib/errors.js";
import { LIBRARY_DOCUMENTS, readLibraryDocument } from "./lib/library-docs.js";
import { configureDocxRedlineLoggerForMcp } from "./lib/mcp-logger.js";
import { createSessionStore } from "./lib/session-store.js";
import {
  authorName,
  documentHandle,
  edit,
  executionOptions,
  inspectFilters,
  revisionToken
} from "./schemas.js";
import { SERVER_INSTRUCTIONS, TOOL_DEFINITIONS } from "./tool-definitions.js";
import { createApplyOperationHandler, createApplyOperationsHandler } from "./tools/apply-operation.js";
import { createCloseDocxHandler } from "./tools/close-docx.js";
import { createEditDocumentHandler } from "./tools/edit-document.js";
import { createInspectDocumentHandler } from "./tools/inspect-document.js";
import { createListPartsHandler } from "./tools/list-parts.js";
import { createOpenDocxHandler } from "./tools/open-docx.js";
import { createPreflightOperationsHandler } from "./tools/preflight-operations.js";
import { createReadPartHandler } from "./tools/read-part.js";
import { createResolveReviewHandler } from "./tools/resolve-review.js";
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
 * per-edit results, retry plan) reach the client as an error result with that payload, instead of
 * a bare message.
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
  warnOnLibraryMismatch();

  const store = deps.store || createSessionStore();
  const getSession = deps.getSession || createGetSession(store);

  return {
    openDocx: createOpenDocxHandler({ store }),
    closeDocx: createCloseDocxHandler({ getSession, store }),
    listParts: createListPartsHandler({ getSession }),
    readPart: createReadPartHandler({ getSession }),
    writePart: createWritePartHandler({ getSession, store }),
    inspectDocument: createInspectDocumentHandler({ getSession }),
    preflightOperations: createPreflightOperationsHandler({ getSession }),
    editDocument: createEditDocumentHandler({ getSession, store }),
    applyOperation: createApplyOperationHandler({ getSession, store }),
    applyOperations: createApplyOperationsHandler({ getSession, store }),
    resolveReview: createResolveReviewHandler({ getSession, store }),
    saveDocx: createSaveDocxHandler({ getSession, store }),
    store
  };
}

function register(server, name, inputSchema, callback) {
  const { title, description, annotations } = TOOL_DEFINITIONS[name];
  server.registerTool(name, { title, description, inputSchema, annotations }, callback);
}

const operationList = z.array(z.record(z.string(), z.unknown())).min(1).describe(
  "Canonical operations (see the docx-redline://schema/document-operations resource). Each may use targetHandle / targetEndHandle instead of target / targetEnd."
);

export function createMcpServer() {
  const handlers = createToolHandlers();
  const server = new McpServer(
    { name: "docx-redline-mcp", version: SERVER_VERSION },
    { instructions: SERVER_INSTRUCTIONS }
  );

  register(server, "open_docx", {
    path: z.string().min(1).describe("Absolute or relative path to a .docx file")
  }, toolCallback(handlers.openDocx));

  register(server, "close_docx", {
    handle: documentHandle,
    discardChanges: z.boolean().optional().describe("Drop edits that were never saved")
  }, toolCallback(handlers.closeDocx));

  register(server, "list_parts", {
    handle: documentHandle
  }, toolCallback(handlers.listParts));

  register(server, "read_part", {
    handle: documentHandle,
    partPath: z.string().min(1),
    maxChars: z.number().int().positive().optional(),
    includeBase64: z.boolean().optional()
  }, toolCallback(handlers.readPart));

  register(server, "write_part", {
    handle: documentHandle,
    partPath: z.string().min(1),
    xml: z.string().min(1)
  }, toolCallback(handlers.writePart));

  register(server, "inspect_document", {
    handle: documentHandle,
    ...inspectFilters
  }, toolCallback(handlers.inspectDocument));

  register(server, "preflight_operations", {
    handle: documentHandle,
    operations: operationList,
    author: authorName,
    options: executionOptions.optional()
  }, toolCallback(handlers.preflightOperations));

  register(server, "edit_document", {
    handle: documentHandle,
    edits: z.array(edit).min(1).describe("All edits for this change, in one call"),
    author: authorName,
    options: executionOptions.optional()
  }, toolCallback(handlers.editDocument));

  register(server, "apply_operations", {
    handle: documentHandle,
    operations: operationList,
    author: authorName,
    options: executionOptions.optional(),
    expectedRevision: revisionToken.optional()
  }, toolCallback(handlers.applyOperations));

  register(server, "apply_operation", {
    handle: documentHandle,
    operation: z.record(z.string(), z.unknown()),
    author: authorName,
    options: executionOptions.optional(),
    expectedRevision: revisionToken.optional()
  }, toolCallback(handlers.applyOperation));

  register(server, "resolve_review", {
    handle: documentHandle,
    action: z.enum(["accept", "reject", "delete_comments", "resolve_comment", "reopen_comment"]).describe(
      "accept/reject tracked changes; delete_comments; resolve_comment/reopen_comment a thread"
    ),
    author: z.string().min(1).optional().describe("Limit to this reviewer's tracked changes or comments"),
    allAuthors: z.literal(true).optional().describe("Everyone's. Only when the user said so explicitly."),
    commentIds: z.array(z.union([z.string(), z.number()])).min(1).optional().describe(
      "Comment ids from inspect_document (deleting a thread root removes its replies)"
    ),
    options: executionOptions.optional()
  }, toolCallback(handlers.resolveReview));

  register(server, "save_docx", {
    handle: documentHandle,
    outputPath: z.string().min(1).optional().describe("Write here (must end in .docx)"),
    inPlace: z.literal(true).optional().describe("Overwrite the source file. Only when the user asked for it.")
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
