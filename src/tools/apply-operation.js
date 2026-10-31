import { expandOperationHandles } from "../lib/edit-translator.js";
import { resolveHandle } from "../lib/target-registry.js";
import { commitOperations } from "../lib/session-apply.js";

export function createApplyOperationHandler({ getSession, store }) {
  return async function applyOperationHandler({ handle, operation, author, options, expectedRevision }) {
    return applyOperations({
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
    return applyOperations({
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

async function applyOperations({ getSession, store, handle, operations, author, options, expectedRevision }) {
  const requested = Array.isArray(operations) ? operations.filter(Boolean) : [];
  if (requested.length === 0) {
    throw new Error("At least one operation is required");
  }

  const session = getSession(handle);
  const { session: updated, payload, parts } = await commitOperations({
    session,
    store,
    operations: requested,
    author,
    options,
    expectedRevision,
    // `targetHandle` / `targetEndHandle` are checked against the revision being edited.
    prepare: document => {
      const revision = document.getRevisionToken().value;
      return expandOperationHandles(requested, name => resolveHandle(session.targets, name, revision));
    }
  });

  return {
    handle: updated.handle,
    dirty: updated.dirty,
    ...payload,
    parts
  };
}
