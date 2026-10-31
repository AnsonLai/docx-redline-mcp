import { requestError } from "./errors.js";

/**
 * Translates the narrow edits accepted by `edit_document` into the library's canonical document
 * operations. Nothing here mutates a document or reimplements targeting: each edit becomes one
 * operation, and the same apply path runs it.
 */

const INTENT_KEYS = ["replacements", "desiredText", "delete", "comment", "restoreText"];

function invalid(index, count, message, field) {
  return requestError({
    code: "INVALID_AGENT_REQUEST",
    message,
    field: field || `edits[${index}]`,
    failedIndexes: [index + 1],
    operationCount: count
  });
}

function handleError(error, index, count) {
  return requestError({
    code: error.code,
    message: `Edit ${index + 1}: ${error.message}`,
    field: `edits[${index}].target`,
    action: "reinspect",
    requiresReinspection: true,
    failedIndexes: [index + 1],
    operationCount: count
  });
}

function validateReplacements(replacements, index, count) {
  const field = `edits[${index}].replacements`;
  if (!Array.isArray(replacements) || replacements.length === 0) {
    throw invalid(index, count, `Edit ${index + 1}: replacements must be a non-empty array.`, field);
  }
  replacements.forEach((replacement, position) => {
    const where = `${field}[${position}]`;
    if (!replacement || typeof replacement.find !== "string" || replacement.find.length === 0) {
      throw invalid(index, count, `Edit ${index + 1}: ${where}.find must be a non-empty string.`, where);
    }
    if (typeof replacement.replace !== "string") {
      throw invalid(index, count, `Edit ${index + 1}: ${where}.replace must be a string.`, where);
    }
    if (replacement.occurrence !== undefined
      && !(Number.isInteger(replacement.occurrence) && replacement.occurrence >= 1)) {
      throw invalid(index, count, `Edit ${index + 1}: ${where}.occurrence must be a positive integer.`, where);
    }
  });
}

/**
 * The paragraph text an edit should leave behind, or null when it cannot be known from the edit
 * alone. Used to find the edited paragraph again after the document changes.
 */
export function expectedTextAfter(source) {
  if (source.intent === "desiredText") return source.desiredText;
  if (source.intent === "comment") return source.exactText;
  if (source.intent !== "replacements") return null;

  // Replacements are resolved against the original text simultaneously.
  const ranges = [];
  for (const replacement of source.replacements) {
    const wanted = replacement.occurrence || 1;
    let from = 0;
    let start = -1;
    for (let found = 0; found < wanted; found += 1) {
      start = source.exactText.indexOf(replacement.find, from);
      if (start < 0) return null;
      from = start + replacement.find.length;
    }
    ranges.push({ start, end: start + replacement.find.length, replace: replacement.replace });
  }
  ranges.sort((a, b) => b.start - a.start);
  let text = source.exactText;
  for (const range of ranges) {
    text = text.slice(0, range.start) + range.replace + text.slice(range.end);
  }
  return text;
}

/**
 * @param {Array<object>} edits
 * @param {(handle: string) => { entry?: object, error?: object }} resolveTarget
 * @returns {{ operations: object[], sources: object[] }}
 */
export function translateEdits(edits, resolveTarget) {
  if (!Array.isArray(edits) || edits.length === 0) {
    throw requestError({
      code: "INVALID_AGENT_REQUEST",
      message: "edits must be a non-empty array.",
      field: "edits"
    });
  }

  const count = edits.length;
  const operations = [];
  const sources = [];

  edits.forEach((edit, index) => {
    if (!edit || typeof edit !== "object" || Array.isArray(edit)) {
      throw invalid(index, count, `Edit ${index + 1} must be an object.`);
    }
    if (typeof edit.target !== "string" || !edit.target.trim()) {
      throw invalid(index, count, `Edit ${index + 1}: target must be a handle returned by inspect_document.`, `edits[${index}].target`);
    }

    const resolution = resolveTarget(edit.target.trim());
    if (resolution.error) throw handleError(resolution.error, index, count);
    const { entry } = resolution;

    const intents = INTENT_KEYS.filter(key => {
      const value = edit[key];
      return key === "delete" ? value === true : value !== undefined && value !== null && value !== false;
    });
    if (intents.length !== 1) {
      throw invalid(
        index,
        count,
        `Edit ${index + 1} must set exactly one of ${INTENT_KEYS.join(", ")} (got ${intents.length === 0 ? "none" : intents.join(", ")}).`
      );
    }
    const intent = intents[0];
    const view = entry.descriptor.revisionView;

    if (intent === "restoreText" && view !== "rejected") {
      throw requestError({
        code: "RESTORE_REQUIRES_REJECTED_VIEW",
        message: `Edit ${index + 1}: restoreText needs a handle from inspect_document with revisionView "rejected", where a reviewer's deleted paragraph is visible.`,
        field: `edits[${index}].target`,
        action: "reinspect",
        requiresReinspection: true,
        failedIndexes: [index + 1],
        operationCount: count
      });
    }
    if (intent !== "restoreText" && view === "rejected") {
      throw requestError({
        code: "EDIT_REQUIRES_ACCEPTED_VIEW",
        message: `Edit ${index + 1}: handle ${entry.handle} came from the rejected view, which only supports restoreText. Inspect the accepted view to edit current text.`,
        field: `edits[${index}].target`,
        action: "reinspect",
        requiresReinspection: true,
        failedIndexes: [index + 1],
        operationCount: count
      });
    }

    const common = {
      operationId: typeof edit.operationId === "string" && edit.operationId ? edit.operationId : `edit-${index + 1}`,
      target: { ...entry.descriptor },
      ...(typeof edit.existingRevisions === "string" ? { existingRevisions: edit.existingRevisions } : {})
    };

    let operation;
    const source = {
      handle: entry.handle,
      humanReference: entry.humanReference,
      exactText: entry.descriptor.exactText,
      paragraphId: entry.descriptor.paragraphId || null,
      index: entry.index,
      intent
    };

    if (intent === "replacements") {
      validateReplacements(edit.replacements, index, count);
      operation = { ...common, type: "redline", replacements: edit.replacements.map(item => ({ ...item })) };
      source.replacements = edit.replacements;
    } else if (intent === "desiredText") {
      if (typeof edit.desiredText !== "string") {
        throw invalid(index, count, `Edit ${index + 1}: desiredText must be a string.`, `edits[${index}].desiredText`);
      }
      operation = { ...common, type: "redline", modified: edit.desiredText };
      source.desiredText = edit.desiredText;
    } else if (intent === "delete") {
      operation = { ...common, type: "delete" };
    } else if (intent === "comment") {
      const comment = edit.comment;
      if (!comment || typeof comment.content !== "string" || !comment.content.trim()) {
        throw invalid(index, count, `Edit ${index + 1}: comment.content must be a non-empty string.`, `edits[${index}].comment`);
      }
      operation = {
        ...common,
        type: "comment",
        commentContent: comment.content,
        ...(typeof comment.anchorText === "string" && comment.anchorText ? { textToComment: comment.anchorText } : {})
      };
    } else {
      if (typeof edit.restoreText !== "string") {
        throw invalid(index, count, `Edit ${index + 1}: restoreText must be a string.`, `edits[${index}].restoreText`);
      }
      operation = { ...common, type: "restore", modified: edit.restoreText };
    }

    operations.push(operation);
    sources.push(source);
  });

  return { operations, sources };
}

/**
 * Lets canonical operations refer to inspected paragraphs by handle: `targetHandle` and
 * `targetEndHandle` are replaced by the descriptors they stand for. Operations that already carry
 * a `target` are left alone.
 */
export function expandOperationHandles(operations, resolveTarget) {
  const count = operations.length;
  return operations.map((operation, index) => {
    if (!operation || typeof operation !== "object") return operation;
    const { targetHandle, targetEndHandle, ...rest } = operation;
    if (targetHandle === undefined && targetEndHandle === undefined) return operation;

    const expanded = { ...rest };
    for (const [handle, field] of [[targetHandle, "target"], [targetEndHandle, "targetEnd"]]) {
      if (handle === undefined) continue;
      const resolution = resolveTarget(String(handle));
      if (resolution.error) {
        throw requestError({
          code: resolution.error.code,
          message: `Operation ${index + 1}: ${resolution.error.message}`,
          field: `operations[${index}].${field}Handle`,
          action: "reinspect",
          requiresReinspection: true,
          failedIndexes: [index + 1],
          operationCount: count
        });
      }
      expanded[field] = { ...resolution.entry.descriptor };
    }
    return expanded;
  });
}
