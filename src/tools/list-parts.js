import { refreshSessionParts } from "../lib/docx-package.js";

export function createListPartsHandler({ getSession }) {
  return async function listParts({ handle }) {
    const session = getSession(handle);
    const parts = refreshSessionParts(session);
    return {
      handle: session.handle,
      parts
    };
  };
}
