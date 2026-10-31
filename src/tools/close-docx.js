import { requestError } from "../lib/errors.js";

export function createCloseDocxHandler({ getSession, store }) {
  return async function closeDocx({ handle, discardChanges }) {
    const session = getSession(handle);

    if (session.dirty && session.savedPath === null && discardChanges !== true) {
      throw requestError({
        code: "UNSAVED_CHANGES",
        message: "The document has edits that were never saved. Save it with save_docx, or pass discardChanges: true to drop them.",
        field: "discardChanges",
        action: "save_docx"
      });
    }

    store.delete(session.handle);
    return {
      handle: session.handle,
      closed: true,
      discardedChanges: session.dirty,
      savedPath: session.savedPath
    };
  };
}
