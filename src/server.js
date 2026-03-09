import * as z from "zod/v4";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { createSessionStore } from "./lib/session-store.js";
import { configureDocxRedlineLoggerForMcp } from "./lib/mcp-logger.js";
import { createApplyOperationHandler, createApplyOperationsHandler } from "./tools/apply-operation.js";
import { createListPartsHandler } from "./tools/list-parts.js";
import { createOpenDocxHandler } from "./tools/open-docx.js";
import { createReadPartHandler } from "./tools/read-part.js";
import { createSaveDocxHandler } from "./tools/save-docx.js";
import { createWritePartHandler } from "./tools/write-part.js";

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
    version: "0.1.0"
  });

  server.registerTool("open_docx", {
    description: "Open a DOCX file and create a working document handle.",
    inputSchema: {
      path: z.string().min(1).describe("Absolute or relative path to a .docx file")
    }
  }, async args => toToolResult(await handlers.openDocx(args)));

  server.registerTool("list_parts", {
    description: "List package parts for an open DOCX handle.",
    inputSchema: {
      handle: z.string().min(1).describe("Handle returned by open_docx")
    }
  }, async args => toToolResult(await handlers.listParts(args)));

  server.registerTool("read_part", {
    description: "Read an OOXML part as XML text or get a read-only binary summary.",
    inputSchema: {
      handle: z.string().min(1),
      partPath: z.string().min(1),
      maxChars: z.number().int().positive().optional(),
      includeBase64: z.boolean().optional()
    }
  }, async args => toToolResult(await handlers.readPart(args)));

  server.registerTool("write_part", {
    description: "Replace a writable XML part in the working DOCX package.",
    inputSchema: {
      handle: z.string().min(1),
      partPath: z.string().min(1),
      xml: z.string().min(1)
    }
  }, async args => toToolResult(await handlers.writePart(args)));

  server.registerTool("apply_operation", {
    description: "Apply a structured redline or maintenance operation to word/document.xml.",
    inputSchema: {
      handle: z.string().min(1),
      operation: z.record(z.string(), z.unknown()),
      author: z.string().min(1).optional()
    }
  }, async args => toToolResult(await handlers.applyOperation(args)));

  server.registerTool("apply_operations", {
    description: "Apply multiple structured redline or maintenance operations to word/document.xml in a single batch.",
    inputSchema: {
      handle: z.string().min(1),
      operations: z.array(z.record(z.string(), z.unknown())).min(1),
      author: z.string().min(1).optional()
    }
  }, async args => toToolResult(await handlers.applyOperations(args)));

  server.registerTool("save_docx", {
    description: "Save the working DOCX package in place or to an explicit output path.",
    inputSchema: {
      handle: z.string().min(1),
      outputPath: z.string().min(1).optional()
    }
  }, async args => toToolResult(await handlers.saveDocx(args)));

  return { server, handlers };
}

export async function startStdioServer() {
  const { server } = createMcpServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  return server;
}
