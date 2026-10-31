import {
  DEFAULT_AUTHOR,
  applyOperationsToSession
} from "../lib/operation-service.js";
import { refreshSessionParts } from "../lib/docx-package.js";

export function createApplyOperationHandler({ getSession, store }) {
  return async function applyOperationHandler({ handle, operation, author, options, expectedRevision }) {
    return applyOperationsToSessionHandle({
      getSession,
      store,
      handle,
      operations: [operation],
      author,
      options,
      expectedRevision
    });
  };
}

export function createApplyOperationsHandler({ getSession, store }) {
  return async function applyOperationsHandler({ handle, operations, author, options, expectedRevision }) {
    return applyOperationsToSessionHandle({
      getSession,
      store,
      handle,
      operations,
      author,
      options,
      expectedRevision
    });
  };
}

async function applyOperationsToSessionHandle({
  getSession,
  store,
  handle,
  operations,
  author,
  options,
  expectedRevision
}) {
  const normalizedOperations = Array.isArray(operations)
    ? operations.filter(Boolean)
    : [];
  if (normalizedOperations.length === 0) {
    throw new Error("At least one operation is required");
  }

  const session = getSession(handle);
  const { payload, candidate } = await applyOperationsToSession({
    session,
    store,
    operations: normalizedOperations,
    author: author || DEFAULT_AUTHOR,
    options,
    expectedRevision
  });

  let updatedSession = session;
  if (candidate) {
    updatedSession = store.update(session.handle, current => ({
      ...current,
      zip: candidate,
      dirty: true
    }));
  }

  return {
    handle: updatedSession.handle,
    dirty: updatedSession.dirty,
    ...payload,
    parts: refreshSessionParts(updatedSession)
  };
}
