import { stat } from "node:fs/promises";
import { createDocxSessionFromPath } from "../lib/docx-package.js";
import { openSessionDocument } from "../lib/operation-service.js";

const MAX_REVISED_PARAGRAPHS = 1_000_000;

/**
 * A cheap orientation for the agent: how much is in the document, who has already reviewed it,
 * and which headers/footers can be targeted with an operation's `part`.
 */
async function summarizeDocument(session) {
  const document = await openSessionDocument(session);
  // `counts` only describe the returned paragraphs, so totals come from the selection of a
  // one-paragraph page.
  const everything = document.inspect({ limit: 1, skipEmpty: false });
  const packageRevision = document.getRevisionToken();

  if (everything.status === "error") {
    return { packageRevision, summary: { error: everything.error } };
  }
  // Revision authors are collected from the returned paragraphs only, so ask for every revised one.
  const revised = document.inspect({ revisedOnly: true, skipEmpty: false, limit: MAX_REVISED_PARAGRAPHS });

  return {
    packageRevision,
    summary: {
      counts: {
        paragraphs: everything.selection?.totalMatches ?? 0,
        revisedParagraphs: revised.selection?.totalMatches ?? 0,
        comments: (everything.comments || []).length
      },
      revisionAuthors: revised.revisionAuthors || [],
      commentAuthors: everything.commentAuthors || [],
      headersFooters: (everything.headersFooters || []).map(part => ({
        path: part.path,
        kind: part.kind,
        type: part.type,
        active: part.active,
        hasFields: part.hasFields
      }))
    }
  };
}

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

    let packageRevision = null;
    let summary = null;
    try {
      ({ packageRevision, summary } = await summarizeDocument(session));
    } catch (error) {
      // The raw part tools still work on packages the library cannot inspect.
      summary = { error: { message: error.message } };
    }

    return {
      handle: session.handle,
      sourcePath: session.sourcePath,
      dirty: session.dirty,
      metadata: session.metadata,
      packageRevision,
      summary,
      parts: docxSession.parts
    };
  };
}
