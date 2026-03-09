import { readFile, writeFile, rename, unlink } from "node:fs/promises";
import { dirname, join } from "node:path";
import JSZip from "jszip";
import { listPartEntries } from "./part-utils.js";
import { assertXmlString, isXmlPartPath, isWritableXmlPartPath } from "./xml-utils.js";

export async function createDocxSessionFromPath(filePath) {
  if (!/\.docx$/i.test(String(filePath || ""))) {
    throw new Error("Only .docx files are supported");
  }

  const buffer = await readFile(filePath);
  const zip = await JSZip.loadAsync(buffer);

  return {
    sourcePath: filePath,
    zip,
    dirty: false,
    parts: listPartEntries(zip)
  };
}

export function refreshSessionParts(session) {
  session.parts = listPartEntries(session.zip);
  return session.parts;
}

export async function readPart(zip, partPath, options = {}) {
  const normalizedPath = String(partPath || "");
  const file = zip?.file(normalizedPath);

  if (!file) {
    throw new Error(`Part not found: ${normalizedPath}`);
  }

  if (isXmlPartPath(normalizedPath)) {
    const text = await file.async("string");
    return {
      kind: "xml",
      path: normalizedPath,
      text: sliceXmlText(text, options),
      fullLength: text.length
    };
  }

  const buffer = await file.async("nodebuffer");
  const result = {
    kind: "binary",
    path: normalizedPath,
    byteLength: buffer.length
  };

  if (options.includeBase64 === true) {
    result.base64 = buffer.toString("base64");
  }

  return result;
}

function sliceXmlText(text, options) {
  const maxChars = Number.isInteger(options.maxChars) && options.maxChars > 0
    ? options.maxChars
    : 0;

  if (maxChars > 0 && text.length > maxChars) {
    return `${text.slice(0, maxChars)}...`;
  }

  return text;
}

export function writeXmlPart(zip, partPath, xmlText) {
  const normalizedPath = String(partPath || "");
  if (!isWritableXmlPartPath(normalizedPath)) {
    throw new Error(`Part is not writable: ${normalizedPath}`);
  }

  assertXmlString(xmlText);
  zip.file(normalizedPath, String(xmlText));
}

export async function saveDocxSession(session, outputPath = session.sourcePath) {
  const targetPath = String(outputPath || session.sourcePath || "");
  if (!targetPath) {
    throw new Error("Output path is required");
  }

  const directoryPath = dirname(targetPath);
  const tempPath = join(directoryPath, `.docx-redline-mcp-${Date.now()}.tmp`);
  const buffer = await session.zip.generateAsync({ type: "nodebuffer" });

  try {
    await writeFile(tempPath, buffer);
    await rename(tempPath, targetPath);
  } catch (error) {
    await unlink(tempPath).catch(() => {});
    throw error;
  }
}
