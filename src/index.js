#!/usr/bin/env node
import { pathToFileURL } from "node:url";
import { startStdioServer } from "./server.js";

export async function startServer() {
  return startStdioServer();
}

const isMain = process.argv[1] && (pathToFileURL(process.argv[1]).href === import.meta.url);
if (isMain) {
  await startServer();
}
