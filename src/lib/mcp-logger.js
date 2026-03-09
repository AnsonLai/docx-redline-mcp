import { configureLogger } from "@ansonlai/docx-redline-js";
import util from "node:util";

function formatArgs(args) {
  return args
    .map(arg => (typeof arg === "string" ? arg : util.inspect(arg, { depth: 6, breakLength: 120 })))
    .join(" ");
}

function createWriter(stderr, prefix = "") {
  return (...args) => {
    if (!stderr || typeof stderr.write !== "function") return;
    stderr.write(`${prefix}${formatArgs(args)}\n`);
  };
}

export function createMcpSafeLogger(stderr = process.stderr) {
  return {
    log: createWriter(stderr),
    warn: createWriter(stderr, "[warn] "),
    error: createWriter(stderr, "[error] ")
  };
}

export function configureDocxRedlineLoggerForMcp(options = {}) {
  const stderr = options.stderr || process.stderr;
  const env = options.env || process.env;
  const level = env.DOCX_REDLINE_MCP_LOG_LEVEL || "warn";

  configureLogger(createMcpSafeLogger(stderr), { level });

  return { level };
}
