import { saveDocxSession } from "../lib/docx-package.js";

export function createSaveDocxHandler({ getSession, store }) {
  return async function saveDocx({ handle, outputPath }) {
    const session = getSession(handle);
    const savedPath = String(outputPath || session.sourcePath);
    await saveDocxSession(session, savedPath);

    const updatedSession = store.update(session.handle, current => ({
      ...current,
      dirty: false
    }));

    return {
      handle: updatedSession.handle,
      sourcePath: updatedSession.sourcePath,
      savedPath,
      dirty: updatedSession.dirty
    };
  };
}
