import { DOMParser } from "@xmldom/xmldom";

const XML_FILE_PATTERN = /\.xml$/i;
const WRITABLE_XML_PARTS = new Set([
  "word/document.xml",
  "word/comments.xml",
  "word/numbering.xml",
  "[Content_Types].xml"
]);

export function isXmlPartPath(partPath) {
  return XML_FILE_PATTERN.test(String(partPath || ""));
}

export function isWritableXmlPartPath(partPath) {
  const value = String(partPath || "");
  if (WRITABLE_XML_PARTS.has(value)) {
    return true;
  }

  return /^word\/(header|footer)\d+\.xml$/i.test(value);
}

export function assertXmlString(xmlText) {
  const xml = String(xmlText || "").trim();
  if (!xml) {
    throw new Error("XML content is required");
  }

  const issues = [];
  const documentNode = new DOMParser({
    errorHandler: {
      warning(message) {
        issues.push(String(message));
      },
      error(message) {
        issues.push(String(message));
      },
      fatalError(message) {
        issues.push(String(message));
      }
    }
  }).parseFromString(xml, "application/xml");

  if (!documentNode?.documentElement || issues.length > 0) {
    throw new Error("Malformed XML");
  }
}
