import { refreshSessionParts, writeXmlPart } from "../lib/docx-package.js";

export function createWritePartHandler({ getSession, store }) {
  return async function writePartHandler({ handle, partPath, xml }) {
    const session = getSession(handle);
    writeXmlPart(session.zip, partPath, xml);

    const updatedSession = store.update(session.handle, current => ({
      ...current,
      dirty: true
    }));

    return {
      handle: updatedSession.handle,
      partPath: String(partPath),
      dirty: updatedSession.dirty,
      parts: refreshSessionParts(updatedSession)
    };
  };
}
