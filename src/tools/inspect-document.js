import { openSessionDocument } from "../lib/operation-service.js";

const DEFAULT_LIMIT = 100;

/** Inspection options accepted from the tool input, forwarded to `DocxDocument.inspect`. */
const INSPECT_OPTION_KEYS = Object.freeze([
  "revisionView",
  "excerptLength",
  "revisedOnly",
  "inTable",
  "skipEmpty",
  "search",
  "indexes",
  "range",
  "around",
  "limit",
  "after"
]);

/**
 * Drops fields that repeat other fields so an inspection stays small enough for an agent to read:
 * `text`/`excerpt` mirror `exactText`, and `segments` only matter when a paragraph has revisions.
 */
function compactParagraph(paragraph) {
  const {
    text,
    excerpt,
    segments,
    structuralReferences,
    revisionAuthors,
    commentIds,
    ...rest
  } = paragraph;
  const compact = { ...rest };
  if (text !== undefined && text !== paragraph.exactText) compact.text = text;
  if (excerpt !== undefined && excerpt !== paragraph.exactText) compact.excerpt = excerpt;
  if (paragraph.hasRevisions) {
    compact.segments = segments;
    compact.revisionAuthors = revisionAuthors;
  }
  if (Array.isArray(commentIds) && commentIds.length > 0) compact.commentIds = commentIds;
  if (Array.isArray(structuralReferences) && structuralReferences.length > 0) {
    compact.structuralReferences = structuralReferences;
  }
  return compact;
}

function compactHeaderFooter(part) {
  const { sections: _sections, ...rest } = part;
  return rest;
}

export function buildInspectOptions(input = {}) {
  const options = {};
  for (const key of INSPECT_OPTION_KEYS) {
    if (input[key] !== undefined) options[key] = input[key];
  }
  if (options.limit === undefined) options.limit = DEFAULT_LIMIT;
  return options;
}

export function createInspectDocumentHandler({ getSession }) {
  return async function inspectDocument(input) {
    const { handle, compact = true } = input;
    const session = getSession(handle);
    const document = await openSessionDocument(session);
    const inspection = document.inspect(buildInspectOptions(input));

    const packageRevision = document.getRevisionToken();
    if (inspection.status === "error") {
      return { handle: session.handle, packageRevision, ...inspection };
    }

    const { revisionToken: _documentRevision, ...rest } = inspection;
    return {
      handle: session.handle,
      packageRevision,
      ...rest,
      paragraphs: compact === false
        ? inspection.paragraphs
        : (inspection.paragraphs || []).map(compactParagraph),
      headersFooters: compact === false
        ? inspection.headersFooters
        : (inspection.headersFooters || []).map(compactHeaderFooter)
    };
  };
}
