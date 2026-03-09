import {
  mergeOperationArtifacts,
  applyStructuredOperation,
  applyStructuredOperations
} from "../lib/operation-service.js";
import { refreshSessionParts } from "../lib/docx-package.js";
import { validateSessionPackage } from "../lib/package-validation.js";

async function getDocumentXml(zip) {
  const file = zip.file("word/document.xml");
  if (!file) {
    throw new Error("Part not found: word/document.xml");
  }
  return file.async("string");
}

export function createApplyOperationHandler({ getSession, store }) {
  return async function applyOperationHandler({ handle, operation, author }) {
    return applyOperationsToSession({
      getSession,
      store,
      handle,
      operations: [operation],
      author,
      mode: "single"
    });
  };
}

export function createApplyOperationsHandler({ getSession, store }) {
  return async function applyOperationsHandler({ handle, operations, author }) {
    return applyOperationsToSession({
      getSession,
      store,
      handle,
      operations,
      author,
      mode: "batch"
    });
  };
}

async function applyOperationsToSession({ getSession, store, handle, operations, author, mode }) {
  const normalizedOperations = Array.isArray(operations)
    ? operations.filter(Boolean)
    : [];
  if (normalizedOperations.length === 0) {
    throw new Error("At least one operation is required");
  }

  const session = getSession(handle);
  const previousDocumentXml = await getDocumentXml(session.zip);
  const previousCommentsXml = session.zip.file("word/comments.xml")
    ? await session.zip.file("word/comments.xml").async("string")
    : null;
  const documentXml = await getDocumentXml(session.zip);

  try {
    const result = mode === "batch"
      ? await applyStructuredOperations(
        session.zip,
        documentXml,
        normalizedOperations,
        author || "AI Agent"
      )
      : await applyStructuredOperation(
        documentXml,
        normalizedOperations[0],
        author || "AI Agent"
      );

    session.zip.file("word/document.xml", result.documentXml);
    await mergeOperationArtifacts(session.zip, result);
    await validateSessionPackage(session.zip);

    const updatedSession = store.update(session.handle, current => ({
      ...current,
      dirty: true
    }));

    return {
      handle: updatedSession.handle,
      dirty: updatedSession.dirty,
      hasChanges: Boolean(result?.hasChanges),
      parts: refreshSessionParts(updatedSession)
    };
  } catch (error) {
    session.zip.file("word/document.xml", previousDocumentXml);
    if (typeof previousCommentsXml === "string") {
      session.zip.file("word/comments.xml", previousCommentsXml);
    } else {
      session.zip.remove("word/comments.xml");
    }
    throw error;
  }
}
