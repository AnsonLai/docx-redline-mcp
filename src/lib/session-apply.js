import { DEFAULT_AUTHOR, applyOperationsToSession } from "./operation-service.js";
import { refreshSessionParts } from "./docx-package.js";

/**
 * Runs operations against a session and commits the result only if the whole package validates.
 * Every mutating tool goes through here, so a failed request can never leave a half-applied
 * document behind and a successful one always reports the new package revision.
 *
 * Returns `{ session, payload }` where `payload` is the apply result. Throws `OperationFailedError`
 * (carrying the structured failure payload) when nothing was applied.
 */
export async function commitOperations({
  session,
  store,
  operations = [],
  author,
  options,
  expectedRevision,
  prepare,
  afterApply
}) {
  const { payload, candidate } = await applyOperationsToSession({
    session,
    operations,
    author: author || DEFAULT_AUTHOR,
    options,
    expectedRevision,
    prepare,
    afterApply
  });

  let updated = session;
  if (candidate) {
    updated = store.update(session.handle, current => ({
      ...current,
      zip: candidate,
      dirty: true
    }));
  }

  return { session: updated, payload, parts: refreshSessionParts(updated) };
}
