import { deriveOutputPath, isSamePath, saveDocxSession } from "../lib/docx-package.js";
import { requestError } from "../lib/errors.js";

export function createSaveDocxHandler({ getSession, store }) {
  return async function saveDocx({ handle, outputPath, inPlace }) {
    const session = getSession(handle);

    if (outputPath && inPlace === true) {
      throw requestError({
        code: "INVALID_AGENT_REQUEST",
        message: "Pass either outputPath or inPlace: true, not both.",
        field: "outputPath"
      });
    }
    if (outputPath && !/\.docx$/i.test(String(outputPath))) {
      throw requestError({
        code: "INVALID_AGENT_REQUEST",
        message: "outputPath must end in .docx.",
        field: "outputPath"
      });
    }

    // The source is only replaced when the caller says so.
    const savedPath = String(
      outputPath || (inPlace === true ? session.sourcePath : await deriveOutputPath(session.sourcePath))
    );
    await saveDocxSession(session, savedPath);

    const replacedSource = isSamePath(savedPath, session.sourcePath);
    const updatedSession = store.update(session.handle, current => ({
      ...current,
      // `dirty` means "differs from the source file", which a copy does not change.
      dirty: replacedSource ? false : current.dirty,
      savedPath
    }));

    return {
      handle: updatedSession.handle,
      sourcePath: updatedSession.sourcePath,
      savedPath,
      replacedSource,
      dirty: updatedSession.dirty,
      completion: true
    };
  };
}
