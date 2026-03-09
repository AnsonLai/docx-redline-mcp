import test, { afterEach } from "node:test";
import assert from "node:assert/strict";
import { error, getLogLevel, log, warn } from "@ansonlai/docx-redline-js/adapters/logger.js";
import { configureDocxRedlineLoggerForMcp } from "../src/lib/mcp-logger.js";

function createSink() {
  const writes = [];
  return {
    writes,
    stream: {
      write(chunk) {
        writes.push(String(chunk));
      }
    }
  };
}

afterEach(() => {
  configureDocxRedlineLoggerForMcp({
    stderr: createSink().stream,
    env: {}
  });
});

test("configureDocxRedlineLoggerForMcp defaults to warn to avoid stdout chatter", () => {
  const sink = createSink();

  const result = configureDocxRedlineLoggerForMcp({
    stderr: sink.stream,
    env: {}
  });

  log("info message");
  warn("warn message");
  error("error message");

  assert.equal(result.level, "warn");
  assert.equal(getLogLevel(), "warn");
  assert.equal(sink.writes.some(line => line.includes("info message")), false);
  assert.equal(sink.writes.some(line => line.includes("warn message")), true);
  assert.equal(sink.writes.some(line => line.includes("error message")), true);
});

test("configureDocxRedlineLoggerForMcp can enable info logs on stderr", () => {
  const sink = createSink();

  const result = configureDocxRedlineLoggerForMcp({
    stderr: sink.stream,
    env: { DOCX_REDLINE_MCP_LOG_LEVEL: "info" }
  });

  log("info message");

  assert.equal(result.level, "info");
  assert.equal(getLogLevel(), "info");
  assert.equal(sink.writes.some(line => line.includes("info message")), true);
});
