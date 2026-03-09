import { readPart } from "../lib/docx-package.js";

export function createReadPartHandler({ getSession }) {
  return async function readPartHandler({ handle, partPath, maxChars, includeBase64 }) {
    const session = getSession(handle);
    return readPart(session.zip, partPath, {
      maxChars,
      includeBase64
    });
  };
}
