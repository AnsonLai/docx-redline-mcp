import { stat } from "node:fs/promises";
import { createDocxSessionFromPath } from "../lib/docx-package.js";

export function createOpenDocxHandler({ store }) {
  return async function openDocx({ path }) {
    const filePath = String(path || "").trim();
    if (!filePath) {
      throw new Error("Path is required");
    }

    const docxSession = await createDocxSessionFromPath(filePath);
    const fileStats = await stat(filePath);
    const session = store.create({
      sourcePath: filePath,
      zip: docxSession.zip,
      metadata: {
        size: fileStats.size,
        modifiedAt: fileStats.mtime.toISOString()
      }
    });

    return {
      handle: session.handle,
      sourcePath: session.sourcePath,
      dirty: session.dirty,
      metadata: session.metadata,
      parts: docxSession.parts
    };
  };
}
