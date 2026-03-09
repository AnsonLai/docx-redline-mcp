import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import {
  normalizeBodySectionOrderStandalone,
  sanitizeNestedParagraphsInTables,
  validateDocxPackage
} from "@ansonlai/docx-redline-js";

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const EMPTY_COMMENTS_XML = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  `<w:comments xmlns:w="${NS_W}"></w:comments>`
].join("");

function parseXml(xml, label) {
  const doc = new DOMParser().parseFromString(xml, "text/xml");
  const parseErrors = doc.getElementsByTagName("parsererror");
  if (parseErrors.length > 0) {
    throw new Error(`Invalid XML in ${label}`);
  }
  return doc;
}

function nearestWordAncestor(node, localName) {
  let current = node;
  while (current) {
    if (current.nodeType === 1 && current.namespaceURI === NS_W && current.localName === localName) {
      return current;
    }
    current = current.parentNode;
  }
  return null;
}

function validateCommentRevisionBoundaries(documentDoc) {
  const starts = Array.from(documentDoc.getElementsByTagNameNS(NS_W, "commentRangeStart"));
  const ends = Array.from(documentDoc.getElementsByTagNameNS(NS_W, "commentRangeEnd"));
  const refs = Array.from(documentDoc.getElementsByTagNameNS(NS_W, "commentReference"));

  for (const start of starts) {
    const id = start.getAttribute("w:id") || start.getAttribute("id");
    const end = ends.find(node => (node.getAttribute("w:id") || node.getAttribute("id")) === id);
    const ref = refs.find(node => (node.getAttribute("w:id") || node.getAttribute("id")) === id);

    if (!end || !ref) {
      throw new Error(`Comment marker ${id || "unknown"} is missing a matching end/reference marker.`);
    }

    const startIns = nearestWordAncestor(start, "ins");
    const endIns = nearestWordAncestor(end, "ins");
    const refIns = nearestWordAncestor(ref, "ins");
    const startDel = nearestWordAncestor(start, "del");
    const endDel = nearestWordAncestor(end, "del");
    const refDel = nearestWordAncestor(ref, "del");

    if (!(startIns === endIns && endIns === refIns)) {
      throw new Error(`Comment marker ${id || "unknown"} crosses a tracked insertion boundary.`);
    }

    if (!(startDel === endDel && endDel === refDel)) {
      throw new Error(`Comment marker ${id || "unknown"} crosses a tracked deletion boundary.`);
    }
  }
}

export function normalizeDocumentXml(documentXml) {
  const documentDoc = parseXml(documentXml, "word/document.xml");
  normalizeBodySectionOrderStandalone(documentDoc);
  sanitizeNestedParagraphsInTables(documentDoc);
  return new XMLSerializer().serializeToString(documentDoc);
}

export async function validateSessionPackage(zip) {
  const documentFile = zip.file("word/document.xml");
  if (!documentFile) {
    throw new Error("Part not found: word/document.xml");
  }

  const documentXml = await documentFile.async("string");
  const documentDoc = parseXml(documentXml, "word/document.xml");
  validateCommentRevisionBoundaries(documentDoc);
  await validateDocxPackage(zip);
}

export function getEmptyCommentsXml() {
  return EMPTY_COMMENTS_XML;
}
