import { expectedTextAfter, translateEdits } from "../lib/edit-translator.js";
import { OperationFailedError } from "../lib/errors.js";
import { commitOperations } from "../lib/session-apply.js";
import { registerParagraph, resolveHandle } from "../lib/target-registry.js";

const ALL_PARAGRAPHS = 1_000_000;

/** One compact line per edit, in the order the edits were sent. */
function shapeEdits(results, sources) {
  const byIndex = new Map((results || []).map(result => [result.index, result]));
  return sources.map((source, position) => {
    const result = byIndex.get(position + 1);
    const edit = {
      index: position + 1,
      target: source.handle,
      ...(source.humanReference ? { humanReference: source.humanReference } : {}),
      intent: source.intent,
      status: result?.status || "not_attempted"
    };
    if (result?.change?.replacements) {
      edit.replacements = result.change.replacements.map(replacement => ({
        find: replacement.find,
        replace: replacement.replace,
        occurrence: replacement.occurrence,
        afterExcerpt: replacement.afterExcerpt
      }));
    }
    if (result?.warnings?.length) edit.warnings = result.warnings;
    if (result?.diagnostics?.length) edit.diagnostics = result.diagnostics;
    if (result?.error) edit.error = result.error;
    return edit;
  });
}

function shapePayload(payload, sources) {
  const { results, ...rest } = payload;
  return { ...rest, edits: shapeEdits(results, sources) };
}

/**
 * Finds each successfully edited paragraph again in the edited document so the agent can keep
 * working without re-inspecting. Paragraphs that cannot be identified unambiguously are omitted.
 */
function findRefreshedParagraphs(document, payload, sources) {
  const inspection = document.inspect({ skipEmpty: false, limit: ALL_PARAGRAPHS });
  if (inspection.status === "error") return [];
  const paragraphs = inspection.paragraphs || [];
  const statuses = new Map((payload.results || []).map(result => [result.index, result.status]));

  const found = [];
  sources.forEach((source, position) => {
    if (statuses.get(position + 1) !== "applied") return;
    if (source.intent === "delete" || source.intent === "restoreText") return;

    let match = source.paragraphId
      ? paragraphs.find(paragraph => paragraph.paragraphId === source.paragraphId)
      : null;
    if (!match) {
      const expected = expectedTextAfter(source);
      const candidates = expected == null ? [] : paragraphs.filter(paragraph => paragraph.exactText === expected);
      match = candidates.length === 1 ? candidates[0] : null;
    }
    if (match) found.push({ previousHandle: source.handle, paragraph: match });
  });
  return found;
}

export function createEditDocumentHandler({ getSession, store }) {
  return async function editDocument({ handle, edits, author, options }) {
    const session = getSession(handle);
    let sources = [];
    let refreshed = [];

    try {
      const { session: updated, payload, parts } = await commitOperations({
        session,
        store,
        author,
        options,
        prepare: document => {
          const revision = document.getRevisionToken().value;
          const translation = translateEdits(edits, name => resolveHandle(session.targets, name, revision));
          sources = translation.sources;
          return translation.operations;
        },
        afterApply: (document, result) => {
          refreshed = findRefreshedParagraphs(document, result, sources);
        }
      });

      // Handles for the new revision are issued only once the edit is committed.
      const refreshedTargets = refreshed.map(({ previousHandle, paragraph }) => ({
        previousHandle,
        handle: registerParagraph(session.targets, paragraph, payload.packageRevision.value),
        exactText: paragraph.exactText,
        humanReference: paragraph.humanReference
      }));

      return {
        handle: updated.handle,
        dirty: updated.dirty,
        ...shapePayload(payload, sources),
        ...(refreshedTargets.length ? { refreshedTargets } : {}),
        parts
      };
    } catch (error) {
      if (error instanceof OperationFailedError && error.payload?.results) {
        throw new OperationFailedError(error.message, shapePayload(error.payload, sources));
      }
      throw error;
    }
  };
}
