import { DOMParser, XMLSerializer } from "@xmldom/xmldom";

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function parseXml(xml) {
  return new DOMParser().parseFromString(String(xml || ""), "text/xml");
}

function serializeXml(doc) {
  return new XMLSerializer().serializeToString(doc);
}

function normalizeText(text) {
  return String(text || "").replace(/\s+/g, " ").trim();
}

function paragraphText(paragraph) {
  return Array.from(paragraph.getElementsByTagNameNS(NS_W, "t"))
    .map(node => node.textContent || "")
    .join("");
}

function findParagraphByTarget(documentDoc, targetText) {
  const normalizedTarget = normalizeText(targetText);
  if (!normalizedTarget) return null;

  const paragraphs = Array.from(documentDoc.getElementsByTagNameNS(NS_W, "p"));
  return paragraphs.find(paragraph => normalizeText(paragraphText(paragraph)).includes(normalizedTarget)) || null;
}

function directFieldCharType(node) {
  if (!node || node.nodeType !== 1 || node.namespaceURI !== NS_W || node.localName !== "r") {
    return "";
  }

  const fldChar = Array.from(node.childNodes || []).find(
    child => child.nodeType === 1 && child.namespaceURI === NS_W && child.localName === "fldChar"
  );

  return fldChar?.getAttribute("w:fldCharType") || fldChar?.getAttribute("fldCharType") || "";
}

function hasInstrText(node) {
  if (!node || node.nodeType !== 1) return false;
  return Array.from(node.childNodes || []).some(
    child => child.nodeType === 1 && child.namespaceURI === NS_W && child.localName === "instrText"
  );
}

function removeProofErrNodes(paragraph) {
  for (const node of Array.from(paragraph.getElementsByTagNameNS(NS_W, "proofErr"))) {
    node.parentNode?.removeChild(node);
  }
}

function unlinkFieldsInParagraph(paragraph) {
  const children = Array.from(paragraph.childNodes || []);
  let inField = false;
  let seenSeparate = false;

  for (const child of children) {
    if (child.nodeType !== 1) continue;

    const fieldType = directFieldCharType(child);
    if (fieldType === "begin") {
      inField = true;
      seenSeparate = false;
      child.parentNode?.removeChild(child);
      continue;
    }

    if (!inField) {
      continue;
    }

    if (fieldType === "separate") {
      seenSeparate = true;
      child.parentNode?.removeChild(child);
      continue;
    }

    if (fieldType === "end") {
      inField = false;
      seenSeparate = false;
      child.parentNode?.removeChild(child);
      continue;
    }

    if (!seenSeparate || hasInstrText(child)) {
      child.parentNode?.removeChild(child);
    }
  }
}

/**
 * Text used to locate the target paragraph. Targets are either a plain string or a descriptor
 * returned by inspection (`exactText`/`text`); descriptors addressed only by id, index or capture
 * carry no text and are left to the library.
 */
export function getOperationTargetText(operation) {
  const target = operation?.target;
  if (typeof target === "string") return target;
  if (target && typeof target === "object") {
    const text = target.exactText ?? target.text;
    return typeof text === "string" ? text : "";
  }
  return "";
}

const PREPROCESSED_OPERATION_TYPES = new Set(["redline", "replace"]);

export function preprocessDocumentXmlForOperation(documentXml, operation) {
  if (!PREPROCESSED_OPERATION_TYPES.has(String(operation?.type || "").toLowerCase())) {
    return String(documentXml || "");
  }

  const targetText = getOperationTargetText(operation);
  if (!targetText.trim()) {
    return String(documentXml || "");
  }

  const documentDoc = parseXml(documentXml);
  const paragraph = findParagraphByTarget(documentDoc, targetText);
  if (!paragraph) {
    return String(documentXml || "");
  }

  removeProofErrNodes(paragraph);
  unlinkFieldsInParagraph(paragraph);
  return serializeXml(documentDoc);
}
