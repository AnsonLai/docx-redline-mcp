import { openSessionDocument } from "../lib/operation-service.js";
import { registerParagraph } from "../lib/target-registry.js";

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
 * What an agent needs to read and cite a paragraph. The machine coordinates (`paragraphId`,
 * `fingerprint`, `P12`-style refs) stay behind the handle; `compact: false` returns them.
 */
function compactParagraph(paragraph, target) {
  const compact = {
    target,
    index: paragraph.index,
    exactText: paragraph.exactText,
    humanReference: paragraph.humanReference,
    ...(paragraph.provision ? { provision: paragraph.provision } : {}),
    ...(paragraph.nearestHeading ? { nearestHeading: paragraph.nearestHeading } : {}),
    ...(paragraph.headingLevel ? { headingLevel: paragraph.headingLevel } : {}),
    inTable: paragraph.inTable,
    ...(paragraph.table ? { table: paragraph.table } : {}),
    ...(paragraph.list ? { list: paragraph.list } : {}),
    hasRevisions: paragraph.hasRevisions
  };
  if (paragraph.hasRevisions) {
    compact.revisionAuthors = paragraph.revisionAuthors;
    compact.segments = paragraph.segments;
  }
  if (Array.isArray(paragraph.commentIds) && paragraph.commentIds.length > 0) {
    compact.commentIds = paragraph.commentIds;
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
    const paragraphs = (inspection.paragraphs || []).map(paragraph => {
      const target = registerParagraph(session.targets, paragraph, packageRevision.value);
      return compact === false ? { target, ...paragraph } : compactParagraph(paragraph, target);
    });

    return {
      handle: session.handle,
      packageRevision,
      ...rest,
      paragraphs,
      headersFooters: compact === false
        ? inspection.headersFooters
        : (inspection.headersFooters || []).map(compactHeaderFooter)
    };
  };
}
